use ironcalc::base::Model;
use ironcalc::export::save_to_xlsx;
use serde::Serialize;
use tauri::State;

use std::collections::{HashMap, HashSet};

use crate::hidden::{extract_default_row_height, extract_hidden_col_ranges};
use crate::index::record_open_internal;
use crate::state::{AppState, LoadedFile};
use crate::xls_load::load_xls;
use crate::xlsx_load::{load_xlsx_with_fallback, replicate_my_array_formulas};
use crate::xlsx_save::{extract_sheet_paths, save_preserving, SheetLayoutSnapshot};

#[derive(Serialize)]
pub(crate) struct WorkbookInfo {
    sheet_names: Vec<String>,
    active_sheet: u32,
}

#[derive(Serialize)]
pub(crate) struct SaveResult {
    path: String,
    /// "preserved" — patched the original xlsx in place (charts/pivots/etc kept).
    /// "ironcalc"  — wrote a fresh xlsx via IronCalc (unsupported features lost).
    /// "xls"       — wrote a fresh BIFF8 .xls via fastsheet's xls writer.
    ///               VBA / macros are preserved when the original file was
    ///               loaded with them (see `vba_preserved`); other unsupported
    ///               features (charts, pivots, drawings, conditional
    ///               formatting) are not.
    mode: &'static str,
    cells_patched: usize,
    /// When the save would lose features that exist in the file being
    /// overwritten, we make a `.bak` copy first (or `.bak.N` when
    /// `.bak` already exists). None when the save was lossless or when
    /// no existing file was overwritten (i.e. saving to a brand-new
    /// path). The frontend surfaces this so the user knows where to
    /// recover from on regret.
    #[serde(skip_serializing_if = "Option::is_none")]
    backup_path: Option<String>,
    /// True for .xls saves whose source had VBA / macro storages, so
    /// the UI can report "VBA preserved" instead of the generic
    /// "macros not preserved" copy.
    #[serde(skip_serializing_if = "is_false")]
    vba_preserved: bool,
}

fn is_false(b: &bool) -> bool {
    !*b
}

/// Backup the existing file at `path` to a fresh `.bak` / `.bak.N`, then
/// save the workbook to `path`. Returns the backup path so the UI can
/// report it. The inner save is told a backup already exists so lossy
/// save paths don't create a second backup.
#[derive(Serialize)]
pub(crate) struct BackupResult {
    save: SaveResult,
    backup_path: String,
}

#[derive(Serialize)]
pub(crate) struct WorkbookRange {
    sheet_name: String,
    rows: Vec<Vec<String>>,
    source_rows: u32,
    source_cols: u32,
    cells_read: usize,
}

fn load_model_for_import(path: &str) -> Result<Model<'static>, String> {
    let p = std::path::Path::new(path);
    if !p.exists() {
        return Err(format!("file does not exist: {path}"));
    }
    if !p.is_file() {
        return Err(format!("not a file: {path}"));
    }
    let is_xls = p
        .extension()
        .and_then(|s| s.to_str())
        .map(|e| e.eq_ignore_ascii_case("xls"))
        .unwrap_or(false);
    let mut model = if is_xls {
        let (m, _, _) = load_xls(path)?;
        m
    } else {
        let bytes = std::fs::read(path).ok();
        let mut m = load_xlsx_with_fallback(path)?;
        if let Some(b) = &bytes {
            let _ = replicate_my_array_formulas(&mut m, b);
        }
        m
    };
    model.evaluate();
    Ok(model)
}

#[tauri::command]
pub(crate) fn open_workbook(
    path: String,
    state: State<'_, AppState>,
) -> Result<WorkbookInfo, String> {
    // Phase timing for the full open-file flow. Active when
    // FASTSHEET_PROFILE_LOAD is set. Includes everything from file
    // detection through state install — i.e. the wall-clock the user
    // sees from clicking Open to the GUI being ready to render.
    let total = std::time::Instant::now();
    crate::util::profile_log(&format!("[open_workbook] === path={path}"));
    let mut t = std::time::Instant::now();
    let lap = |t: &mut std::time::Instant, label: &str| {
        crate::util::profile_log(&format!(
            "[open_workbook] {:>20} {:>7.1}ms",
            label,
            t.elapsed().as_secs_f64() * 1000.0
        ));
        *t = std::time::Instant::now();
    };

    let is_xls = std::path::Path::new(&path)
        .extension()
        .and_then(|s| s.to_str())
        .map(|e| e.eq_ignore_ascii_case("xls"))
        .unwrap_or(false);
    let mut xls_hidden_cols: HashMap<u32, HashSet<i32>> = HashMap::new();
    let mut xls_preserved: crate::xls_preserve::PreservedXlsData = Default::default();
    // For xlsx we read the file bytes once and reuse them for both
    // MY* array-formula replication and the in-memory LoadedFile
    // snapshot used by save_preserving. Two reads of the same large
    // file across a WSL UNC share is the slowest path on cold open.
    let xlsx_bytes: Option<Vec<u8>> = if is_xls { None } else { std::fs::read(&path).ok() };
    let mut model = if is_xls {
        let (m, hc, preserved) = load_xls(&path)?;
        xls_hidden_cols = hc;
        xls_preserved = preserved;
        m
    } else {
        let mut m = load_xlsx_with_fallback(&path)?;
        // Replicate `<f t="array" ref="X:Y">MY*(...)</f>` across each spill cell
        // BEFORE evaluate() so dependents see fresh per-cell formulas instead of
        // POI's stale cached `<v>` daughters. Best-effort — non-fatal if the
        // file can't be re-read (we already loaded it once). Doesn't apply to
        // .xls (no `t="array"` markers; calamine returns each spill cell's
        // formula text directly).
        if let Some(b) = &xlsx_bytes {
            let _ = replicate_my_array_formulas(&mut m, b);
        }
        m
    };
    lap(&mut t, "load+replicate");
    model.evaluate();
    lap(&mut t, "evaluate");
    let names: Vec<String> = model
        .workbook
        .worksheets
        .iter()
        .map(|w| w.name.clone())
        .collect();
    let info = WorkbookInfo {
        sheet_names: names,
        active_sheet: 0,
    };
    *state.model.lock().unwrap() = Some(model);
    // Snapshot the original bytes + sheet path mapping for in-place save
    // preservation. Best-effort — preservation just won't be available if
    // this fails (we'll fall back to ironcalc save_to_xlsx).
    let mut hidden_cols_init = std::collections::HashMap::new();
    // .xls files don't have an xlsx zip to snapshot for save_preserving;
    // leave state.loaded as None so save_workbook falls through to the
    // IronCalc save_to_xlsx path (which writes a fresh .xlsx). The
    // hidden-col map we got from xls_load still flows into state.hidden_cols.
    if is_xls {
        *state.loaded.lock().unwrap() = None;
        hidden_cols_init = xls_hidden_cols;
        // Stash the captured VBA / macro storages from the source so
        // the writer can replay them on save. Empty when the source
        // has no macros — common case.
        *state.xls_preserved.lock().unwrap() = if xls_preserved.is_empty() {
            None
        } else {
            Some(xls_preserved)
        };
    } else if let Some(bytes) = xlsx_bytes {
        let sheet_paths = extract_sheet_paths(&bytes).unwrap_or_default();
        // Seed the in-memory hidden-column state from the original xlsx so
        // get_layout doesn't have to re-scrape the zip on every refresh,
        // and so set_column_hidden has somewhere to mutate.
        let mut default_rh: HashMap<u32, f64> = HashMap::new();
        for (idx, sheet_path) in sheet_paths.iter().enumerate() {
            let ranges = extract_hidden_col_ranges(&bytes, sheet_path);
            let cols: HashSet<i32> = ranges.iter().flat_map(|(lo, hi)| *lo..=*hi).collect();
            if !cols.is_empty() {
                hidden_cols_init.insert(idx as u32, cols);
            }
            // Capture the file's per-sheet defaultRowHeight (in
            // points). Used by get_layout to size rows that have no
            // explicit Row entry — see state::default_row_heights.
            if let Some(pt) = extract_default_row_height(&bytes, sheet_path) {
                default_rh.insert(idx as u32, pt);
            }
        }
        *state.default_row_heights.lock().unwrap() = default_rh;
        *state.loaded.lock().unwrap() = Some(LoadedFile {
            path: path.clone(),
            bytes,
            sheet_paths,
        });
    } else {
        *state.loaded.lock().unwrap() = None;
        state.default_row_heights.lock().unwrap().clear();
    }
    lap(&mut t, "snapshot+hidden");
    *state.hidden_cols.lock().unwrap() = hidden_cols_init;
    state.dirty.lock().unwrap().clear();
    state.style_dirty.lock().unwrap().clear();
    *state.structural_dirty.lock().unwrap() = false;
    *state.workbook_dirty.lock().unwrap() = false;
    // Loading a fresh workbook invalidates any active compare —
    // diffing the new model against the previous right side would
    // confuse more than help.
    *state.compare.lock().unwrap() = None;
    state.protected_ranges.lock().unwrap().clear();
    state.input_ranges.lock().unwrap().clear();
    let _ = record_open_internal(&state, &path);
    lap(&mut t, "state_install");
    crate::util::profile_log(&format!(
        "[open_workbook] {:>20} {:>7.1}ms",
        "TOTAL",
        total.elapsed().as_secs_f64() * 1000.0
    ));
    Ok(info)
}

#[tauri::command]
pub(crate) fn new_workbook(state: State<'_, AppState>) -> Result<WorkbookInfo, String> {
    let model = Model::new_empty("untitled", "en", "UTC", "en").map_err(|e| e)?;
    let names: Vec<String> = model
        .workbook
        .worksheets
        .iter()
        .map(|w| w.name.clone())
        .collect();
    let info = WorkbookInfo {
        sheet_names: names,
        active_sheet: 0,
    };
    *state.model.lock().unwrap() = Some(model);
    *state.loaded.lock().unwrap() = None;
    *state.xls_preserved.lock().unwrap() = None;
    state.dirty.lock().unwrap().clear();
    state.hidden_cols.lock().unwrap().clear();
    state.default_row_heights.lock().unwrap().clear();
    state.style_dirty.lock().unwrap().clear();
    *state.structural_dirty.lock().unwrap() = false;
    *state.workbook_dirty.lock().unwrap() = false;
    *state.compare.lock().unwrap() = None;
    state.protected_ranges.lock().unwrap().clear();
    state.input_ranges.lock().unwrap().clear();
    Ok(info)
}

#[tauri::command]
pub(crate) fn save_workbook(
    path: String,
    state: State<'_, AppState>,
) -> Result<SaveResult, String> {
    save_workbook_inner(path, state, false)
}

#[tauri::command]
pub(crate) fn extract_cells_to_workbook(
    path: String,
    rows: Vec<Vec<String>>,
) -> Result<SaveResult, String> {
    if rows.is_empty() {
        return Err("no cells selected".into());
    }
    let mut model = Model::new_empty("extract", "en", "UTC", "en").map_err(|e| e)?;
    let mut cells_written = 0usize;
    for (row_idx, row) in rows.iter().enumerate() {
        for (col_idx, value) in row.iter().enumerate() {
            if value.is_empty() {
                continue;
            }
            model
                .set_user_input(0, row_idx as i32 + 1, col_idx as i32 + 1, value.clone())
                .map_err(|e| e.to_string())?;
            cells_written += 1;
        }
    }
    model.evaluate();
    let backup_path = crate::atomic::backup_if_exists(std::path::Path::new(&path))?
        .map(|p| p.to_string_lossy().into_owned());
    crate::atomic::write(std::path::Path::new(&path), |tmp| {
        save_to_xlsx(&model, &tmp.to_string_lossy()).map_err(|e| e.to_string())
    })?;
    Ok(SaveResult {
        path,
        mode: "ironcalc",
        cells_patched: cells_written,
        backup_path,
        vba_preserved: false,
    })
}

#[tauri::command]
pub(crate) fn read_workbook_first_sheet(path: String) -> Result<WorkbookRange, String> {
    let model = load_model_for_import(&path)?;
    let ws = model
        .workbook
        .worksheets
        .first()
        .ok_or("workbook has no sheets")?;
    let mut min_row = i32::MAX;
    let mut min_col = i32::MAX;
    let mut max_row = 0i32;
    let mut max_col = 0i32;
    let mut cells = Vec::new();
    for (row, cols) in &ws.sheet_data {
        for col in cols.keys() {
            let input = model
                .get_localized_cell_content(0, *row, *col)
                .unwrap_or_default();
            if input.is_empty() {
                continue;
            }
            min_row = min_row.min(*row);
            min_col = min_col.min(*col);
            max_row = max_row.max(*row);
            max_col = max_col.max(*col);
            cells.push((*row, *col, input));
        }
    }
    if cells.is_empty() {
        return Err(format!("no populated cells in first sheet of {path}"));
    }
    let source_rows = (max_row - min_row + 1) as u32;
    let source_cols = (max_col - min_col + 1) as u32;
    let mut rows = vec![vec![String::new(); source_cols as usize]; source_rows as usize];
    let cells_read = cells.len();
    for (row, col, input) in cells {
        rows[(row - min_row) as usize][(col - min_col) as usize] = input;
    }
    Ok(WorkbookRange {
        sheet_name: ws.name.clone(),
        rows,
        source_rows,
        source_cols,
        cells_read,
    })
}

#[tauri::command]
pub(crate) fn workbook_has_unsaved_changes(state: State<'_, AppState>) -> bool {
    *state.workbook_dirty.lock().unwrap()
}

fn save_workbook_inner(
    path: String,
    state: State<'_, AppState>,
    backup_already_created: bool,
) -> Result<SaveResult, String> {
    // .xls target → BIFF writer. Skips the xlsx preservation / IronCalc
    // paths entirely. Greenfield-only for now — we don't try to patch
    // the original .xls bytes the way save_preserving does for xlsx.
    let is_xls_target = std::path::Path::new(&path)
        .extension()
        .and_then(|s| s.to_str())
        .map(|e| e.eq_ignore_ascii_case("xls"))
        .unwrap_or(false);
    if is_xls_target {
        let (bytes, preserved, vba_preserved) = {
            let model_guard = state.model.lock().unwrap();
            let model = model_guard.as_ref().ok_or("no workbook open")?;
            // Inject the VBA / macro storages we captured on load (if
            // any) so Excel-side macros survive the save. Save-as is OK
            // too — VBA streams are self-contained, no offset
            // cross-references with the workbook stream.
            let preserved = state.xls_preserved.lock().unwrap().clone();
            let vba_preserved = preserved
                .as_ref()
                .map(|p| !p.is_empty())
                .unwrap_or(false);
            // IronCalc's `Col` has no hidden field, so hidden-col state
            // lives in AppState's side-channel — thread it through so
            // COLINFO emits the hidden bit on save.
            let hidden = state.hidden_cols.lock().unwrap().clone();
            let bytes = crate::xls_save::build_xls_bytes(model, Some(&hidden));
            (bytes, preserved, vba_preserved)
        };
        let backup_path = if backup_already_created {
            None
        } else {
            crate::atomic::backup_if_exists(std::path::Path::new(&path))?
                .map(|p| p.to_string_lossy().into_owned())
        };
        crate::xls_save::write_xls_bytes_with_preserved(
            &path,
            &bytes,
            preserved.as_ref().filter(|p| !p.is_empty()),
        )?;
        state.dirty.lock().unwrap().clear();
        state.style_dirty.lock().unwrap().clear();
        *state.structural_dirty.lock().unwrap() = false;
        *state.workbook_dirty.lock().unwrap() = false;
        let _ = record_open_internal(&state, &path);
        return Ok(SaveResult {
            path,
            mode: "xls",
            cells_patched: 0,
            backup_path,
            vba_preserved,
        });
    }
    // Preservation eligibility (xlsx only):
    //   1. We have an in-memory snapshot of the original xlsx
    //   2. The save target is the SAME file we loaded from (otherwise it's
    //      a Save As / new file — start fresh from IronCalc)
    //   3. We were able to map sheet_idx → zip entry path on load
    //   4. No style changes (preservation can't patch styles.xml safely)
    //   5. No structural changes (insert/delete row/col/sheet shifts
    //      coordinates — the cell-XML patcher would land edits at the
    //      wrong row/col, or worse, leave stale rows behind)
    let loaded_guard = state.loaded.lock().unwrap();
    let dirty_guard = state.dirty.lock().unwrap();
    let style_dirty_present = !state.style_dirty.lock().unwrap().is_empty();
    let structural_dirty_present = *state.structural_dirty.lock().unwrap();
    let mut try_preserve = false;
    if let Some(loaded) = loaded_guard.as_ref() {
        if !loaded.sheet_paths.is_empty()
            && std::path::Path::new(&loaded.path) == std::path::Path::new(&path)
            && !style_dirty_present
            && !structural_dirty_present
        {
            try_preserve = true;
        }
    }
    if try_preserve {
        if let Some(loaded) = loaded_guard.as_ref() {
            let loaded_snapshot = loaded.clone();
            let dirty_snapshot = dirty_guard.clone();
            let n = dirty_snapshot.len();
            // Snapshot the in-memory layout state (cols, rows, frozen
            // panes, hidden col side-channel) for every sheet so
            // save_preserving can project it back into the XML alongside
            // the dirty cells. Cheap clones — Col / Row are small.
            let layouts = collect_layout_snapshots(&state);
            drop(loaded_guard);
            drop(dirty_guard);
            save_preserving(&loaded_snapshot, &dirty_snapshot, &layouts, &path)?;
            state.dirty.lock().unwrap().clear();
            // Defensive: preservation path is gated on structural_dirty
            // being false, so this would already be false; clear anyway
            // to keep all save paths uniform.
            *state.structural_dirty.lock().unwrap() = false;
            *state.workbook_dirty.lock().unwrap() = false;
            // Refresh our in-memory snapshot so subsequent saves keep
            // patching the latest version of the file.
            if let Ok(bytes) = std::fs::read(&path) {
                let sheet_paths = extract_sheet_paths(&bytes).unwrap_or_default();
                *state.loaded.lock().unwrap() = Some(LoadedFile {
                    path: path.clone(),
                    bytes,
                    sheet_paths,
                });
            }
            let _ = record_open_internal(&state, &path);
            return Ok(SaveResult {
                path,
                mode: "preserved",
                cells_patched: n,
                backup_path: None,
                vba_preserved: false,
            });
        }
    }
    drop(loaded_guard);
    drop(dirty_guard);
    // Fallback path — IronCalc save (loses unsupported features:
    // charts, pivots, comments, conditional formatting, drawings,
    // shared formulas in some cases). If a file already exists at
    // the target, snapshot it as `.bak` first so anything we'd lose
    // is recoverable.
    {
        let model_guard = state.model.lock().unwrap();
        if model_guard.is_none() {
            return Err("no workbook open".into());
        }
    }
    let backup_path = if backup_already_created {
        None
    } else {
        crate::atomic::backup_if_exists(std::path::Path::new(&path))?
            .map(|p| p.to_string_lossy().into_owned())
    };
    let model_guard = state.model.lock().unwrap();
    let model = model_guard.as_ref().ok_or("no workbook open")?;
    crate::atomic::write(std::path::Path::new(&path), |tmp| {
        save_to_xlsx(model, &tmp.to_string_lossy()).map_err(|e| e.to_string())
    })?;
    drop(model_guard);
    state.dirty.lock().unwrap().clear();
    state.style_dirty.lock().unwrap().clear();
    *state.structural_dirty.lock().unwrap() = false;
    *state.workbook_dirty.lock().unwrap() = false;
    let _ = record_open_internal(&state, &path);
    Ok(SaveResult {
        path,
        mode: "ironcalc",
        cells_patched: 0,
        backup_path,
        vba_preserved: false,
    })
}

/// Append a new sheet (auto-named "SheetN") and return its name + index.
/// Marks the workbook structural-dirty: sheet add/delete shifts every
/// subsequent sheet's index, which the xlsx preservation path can't
/// patch safely (the loaded sheet_paths mapping would desync).
#[tauri::command]
pub(crate) fn add_sheet(state: State<'_, AppState>) -> Result<(String, u32), String> {
    let mut guard = state.model.lock().unwrap();
    let model = guard.as_mut().ok_or("no workbook open")?;
    let (name, idx) = model.new_sheet();
    drop(guard);
    *state.structural_dirty.lock().unwrap() = true;
    *state.workbook_dirty.lock().unwrap() = true;
    Ok((name, idx))
}

/// Add a sheet with a specific name. Errors if the name already exists.
#[tauri::command]
pub(crate) fn add_sheet_named(name: String, state: State<'_, AppState>) -> Result<(), String> {
    let mut guard = state.model.lock().unwrap();
    let model = guard.as_mut().ok_or("no workbook open")?;
    model.add_sheet(&name)?;
    drop(guard);
    *state.structural_dirty.lock().unwrap() = true;
    *state.workbook_dirty.lock().unwrap() = true;
    Ok(())
}

#[tauri::command]
pub(crate) fn rename_sheet(
    sheet: u32,
    new_name: String,
    state: State<'_, AppState>,
) -> Result<(), String> {
    let mut guard = state.model.lock().unwrap();
    let model = guard.as_mut().ok_or("no workbook open")?;
    model.rename_sheet_by_index(sheet, &new_name)?;
    drop(guard);
    // Sheet rename changes the worksheet's name but not its index in
    // `sheet_paths` — preservation could in principle survive it. But
    // the original xlsx's workbook.xml + cell formulas reference the
    // old name; patching that consistently across every sheet's
    // shared-formulas / defined-names is more work than just routing
    // through save_to_xlsx for now.
    *state.structural_dirty.lock().unwrap() = true;
    *state.workbook_dirty.lock().unwrap() = true;
    Ok(())
}

#[tauri::command]
pub(crate) fn delete_sheet(sheet: u32, state: State<'_, AppState>) -> Result<(), String> {
    let mut guard = state.model.lock().unwrap();
    let model = guard.as_mut().ok_or("no workbook open")?;
    model.delete_sheet(sheet)?;
    drop(guard);
    *state.structural_dirty.lock().unwrap() = true;
    *state.workbook_dirty.lock().unwrap() = true;
    Ok(())
}

/// Refresh the WorkbookInfo (sheet name list) — frontend calls this after
/// any sheet add/delete/rename to resync the tab bar without re-opening
/// the file.
#[tauri::command]
pub(crate) fn workbook_info(state: State<'_, AppState>) -> Result<WorkbookInfo, String> {
    let guard = state.model.lock().unwrap();
    let model = guard.as_ref().ok_or("no workbook open")?;
    let names: Vec<String> = model
        .workbook
        .worksheets
        .iter()
        .map(|w| w.name.clone())
        .collect();
    Ok(WorkbookInfo {
        sheet_names: names,
        active_sheet: 0,
    })
}

/// Snapshot the in-memory layout for every sheet so save_preserving can
/// project it back into the saved xlsx. We snapshot ALL sheets (not just
/// the ones with dirty cells) because layout changes might be the only
/// thing the user did. Empty when no model is loaded — caller treats that
/// as "no layout patches needed".
fn collect_layout_snapshots(state: &State<'_, AppState>) -> HashMap<u32, SheetLayoutSnapshot> {
    let model_guard = state.model.lock().unwrap();
    let Some(model) = model_guard.as_ref() else {
        return HashMap::new();
    };
    let hidden_guard = state.hidden_cols.lock().unwrap();
    let mut out = HashMap::new();
    for (idx, ws) in model.workbook.worksheets.iter().enumerate() {
        let hidden_cols = hidden_guard.get(&(idx as u32)).cloned().unwrap_or_default();
        out.insert(
            idx as u32,
            SheetLayoutSnapshot {
                cols: ws.cols.clone(),
                hidden_cols,
                rows: ws.rows.clone(),
                frozen_rows: ws.frozen_rows,
                frozen_cols: ws.frozen_columns,
            },
        );
    }
    out
}

#[tauri::command]
pub(crate) fn file_exists(path: String) -> bool {
    std::path::Path::new(&path).exists()
}

#[tauri::command]
pub(crate) fn erase_file(path: String) -> Result<(), String> {
    let p = std::path::Path::new(&path);
    if !p.exists() {
        return Err(format!("file does not exist: {path}"));
    }
    if !p.is_file() {
        return Err(format!("not a file: {path}"));
    }
    std::fs::remove_file(p).map_err(|e| e.to_string())
}

#[tauri::command]
pub(crate) fn read_text_file(path: String) -> Result<String, String> {
    let p = std::path::Path::new(&path);
    if !p.exists() {
        return Err(format!("file does not exist: {path}"));
    }
    if !p.is_file() {
        return Err(format!("not a file: {path}"));
    }
    std::fs::read_to_string(p).map_err(|e| e.to_string())
}

/// Load a second workbook from `path` and diff it against the active
/// model. Returns the diff list + missing-sheet summary; the right
/// model stays in `state.compare` for trace integration. The active
/// workbook is unchanged (this is a read-only side-load).
///
/// Loading uses the same pipeline as `open_workbook` so palette
/// preprocessing, array-marker stripping, MY* replication, and the
/// xls VBA-strip recovery all apply. We do NOT replicate the right-
/// side workbook's xls hidden cols / preserved VBA — compare doesn't
/// need them and they'd just consume memory.
#[tauri::command]
pub(crate) fn compare_open(
    path: String,
    state: State<'_, AppState>,
) -> Result<crate::compare::CompareResult, String> {
    let is_xls = std::path::Path::new(&path)
        .extension()
        .and_then(|s| s.to_str())
        .map(|e| e.eq_ignore_ascii_case("xls"))
        .unwrap_or(false);
    let mut right = if is_xls {
        let (m, _hc, _preserved) = load_xls(&path)?;
        m
    } else {
        let mut m = load_xlsx_with_fallback(&path)?;
        if let Ok(bytes) = std::fs::read(&path) {
            let _ = replicate_my_array_formulas(&mut m, &bytes);
        }
        m
    };
    right.evaluate();

    let model_guard = state.model.lock().unwrap();
    let left = model_guard.as_ref().ok_or("no workbook open")?;
    let (session, result) = crate::compare::diff_workbooks(left, right, path);
    drop(model_guard);
    *state.compare.lock().unwrap() = Some(session);
    Ok(result)
}

#[tauri::command]
pub(crate) fn compare_close(state: State<'_, AppState>) -> Result<(), String> {
    *state.compare.lock().unwrap() = None;
    Ok(())
}

/// Returns the right-side formatted value at a left-coords address,
/// or null when no compare session is active or the cell isn't on
/// the right side. Used by trace nodes that want to render
/// `left | right` without the frontend re-walking the dep tree.
#[tauri::command]
pub(crate) fn compare_value_at(
    sheet: u32,
    row: i32,
    col: i32,
    state: State<'_, AppState>,
) -> Result<Option<String>, String> {
    // Canonical AppState lock order is model → compare (also used by
    // cells::trace_formula). Acquiring in opposite directions across
    // commands risks a deadlock if any command path ever runs
    // concurrently with another that hits the same two locks.
    let model_guard = state.model.lock().unwrap();
    let Some(left) = model_guard.as_ref() else {
        return Ok(None);
    };
    let compare_guard = state.compare.lock().unwrap();
    let Some(session) = compare_guard.as_ref() else {
        return Ok(None);
    };
    Ok(session.right_value_at(left, sheet, row, col))
}

#[tauri::command]
pub(crate) fn backup_and_save(
    path: String,
    state: State<'_, AppState>,
) -> Result<BackupResult, String> {
    let p = std::path::Path::new(&path);
    let bak = crate::atomic::backup_if_exists(p)?
        .ok_or_else(|| format!("no existing file to back up: {}", p.display()))?;
    let save = save_workbook_inner(path, state, true)?;
    Ok(BackupResult {
        save,
        backup_path: bak.to_string_lossy().into_owned(),
    })
}
