use ironcalc::base::Model;
use ironcalc::export::save_to_xlsx;
use serde::Serialize;
use tauri::State;

use std::collections::{HashMap, HashSet};

use crate::hidden::extract_hidden_col_ranges;
use crate::index::record_open_internal;
use crate::state::{AppState, LoadedFile};
use crate::xls_load::load_xls;
use crate::xls_save::save_xls;
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
    /// "xls"       — wrote a fresh BIFF8 .xls via fastsheet's xls writer
    ///               (formulas currently round-trip as cached values; see
    ///               xls_save.rs phase notes).
    mode: &'static str,
    cells_patched: usize,
}

/// Backup the existing file at `path` to `path.bak` (overwriting any
/// previous .bak), then save the workbook to `path`. Returns the .bak
/// path so the UI can report it.
#[derive(Serialize)]
pub(crate) struct BackupResult {
    save: SaveResult,
    backup_path: String,
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
        if let Ok(bytes) = std::fs::read(&path) {
            let _ = replicate_my_array_formulas(&mut m, &bytes);
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
    } else if let Ok(bytes) = std::fs::read(&path) {
        let sheet_paths = extract_sheet_paths(&bytes).unwrap_or_default();
        // Seed the in-memory hidden-column state from the original xlsx so
        // get_layout doesn't have to re-scrape the zip on every refresh,
        // and so set_column_hidden has somewhere to mutate.
        for (idx, sheet_path) in sheet_paths.iter().enumerate() {
            let ranges = extract_hidden_col_ranges(&bytes, sheet_path);
            let cols: HashSet<i32> = ranges.iter().flat_map(|(lo, hi)| *lo..=*hi).collect();
            if !cols.is_empty() {
                hidden_cols_init.insert(idx as u32, cols);
            }
        }
        *state.loaded.lock().unwrap() = Some(LoadedFile {
            path: path.clone(),
            bytes,
            sheet_paths,
        });
    } else {
        *state.loaded.lock().unwrap() = None;
    }
    lap(&mut t, "snapshot+hidden");
    *state.hidden_cols.lock().unwrap() = hidden_cols_init;
    state.dirty.lock().unwrap().clear();
    state.style_dirty.lock().unwrap().clear();
    // Loading a fresh workbook invalidates any active compare —
    // diffing the new model against the previous right side would
    // confuse more than help.
    *state.compare.lock().unwrap() = None;
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
    state.style_dirty.lock().unwrap().clear();
    *state.compare.lock().unwrap() = None;
    Ok(info)
}

#[tauri::command]
pub(crate) fn save_workbook(
    path: String,
    state: State<'_, AppState>,
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
        let model_guard = state.model.lock().unwrap();
        let model = model_guard.as_ref().ok_or("no workbook open")?;
        // Inject the VBA / macro storages we captured on load (if
        // any) so Excel-side macros survive the save. Save-as is OK
        // too — VBA streams are self-contained, no offset
        // cross-references with the workbook stream.
        let preserved_guard = state.xls_preserved.lock().unwrap();
        crate::xls_save::save_xls_with_preserved(model, &path, preserved_guard.as_ref())?;
        drop(preserved_guard);
        drop(model_guard);
        state.dirty.lock().unwrap().clear();
        state.style_dirty.lock().unwrap().clear();
        return Ok(SaveResult {
            path,
            mode: "xls",
            cells_patched: 0,
        });
    }
    // Preservation eligibility (xlsx only):
    //   1. We have an in-memory snapshot of the original xlsx
    //   2. The save target is the SAME file we loaded from (otherwise it's
    //      a Save As / new file — start fresh from IronCalc)
    //   3. We were able to map sheet_idx → zip entry path on load
    let loaded_guard = state.loaded.lock().unwrap();
    let dirty_guard = state.dirty.lock().unwrap();
    let style_dirty_present = !state.style_dirty.lock().unwrap().is_empty();
    let mut try_preserve = false;
    if let Some(loaded) = loaded_guard.as_ref() {
        if !loaded.sheet_paths.is_empty()
            && std::path::Path::new(&loaded.path) == std::path::Path::new(&path)
            // Style changes can't be patched into the original styles.xml
            // safely yet, so route through save_to_xlsx when present.
            && !style_dirty_present
        {
            try_preserve = true;
        }
    }
    if try_preserve {
        if let Some(loaded) = loaded_guard.as_ref() {
            let n = dirty_guard.len();
            // Snapshot the in-memory layout state (cols, rows, frozen
            // panes, hidden col side-channel) for every sheet so
            // save_preserving can project it back into the XML alongside
            // the dirty cells. Cheap clones — Col / Row are small.
            let layouts = collect_layout_snapshots(&state);
            save_preserving(loaded, &dirty_guard, &layouts, &path)?;
            drop(loaded_guard);
            drop(dirty_guard);
            state.dirty.lock().unwrap().clear();
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
            return Ok(SaveResult {
                path,
                mode: "preserved",
                cells_patched: n,
            });
        }
    }
    drop(loaded_guard);
    drop(dirty_guard);
    // Fallback path — IronCalc save (loses unsupported features).
    let model_guard = state.model.lock().unwrap();
    let model = model_guard.as_ref().ok_or("no workbook open")?;
    let _ = std::fs::remove_file(&path);
    save_to_xlsx(model, &path).map_err(|e| e.to_string())?;
    drop(model_guard);
    state.dirty.lock().unwrap().clear();
    state.style_dirty.lock().unwrap().clear();
    Ok(SaveResult {
        path,
        mode: "ironcalc",
        cells_patched: 0,
    })
}

/// Append a new sheet (auto-named "SheetN") and return its name + index.
#[tauri::command]
pub(crate) fn add_sheet(state: State<'_, AppState>) -> Result<(String, u32), String> {
    let mut guard = state.model.lock().unwrap();
    let model = guard.as_mut().ok_or("no workbook open")?;
    let (name, idx) = model.new_sheet();
    Ok((name, idx))
}

/// Add a sheet with a specific name. Errors if the name already exists.
#[tauri::command]
pub(crate) fn add_sheet_named(name: String, state: State<'_, AppState>) -> Result<(), String> {
    let mut guard = state.model.lock().unwrap();
    let model = guard.as_mut().ok_or("no workbook open")?;
    model.add_sheet(&name)
}

#[tauri::command]
pub(crate) fn rename_sheet(
    sheet: u32,
    new_name: String,
    state: State<'_, AppState>,
) -> Result<(), String> {
    let mut guard = state.model.lock().unwrap();
    let model = guard.as_mut().ok_or("no workbook open")?;
    model.rename_sheet_by_index(sheet, &new_name)
}

#[tauri::command]
pub(crate) fn delete_sheet(sheet: u32, state: State<'_, AppState>) -> Result<(), String> {
    let mut guard = state.model.lock().unwrap();
    let model = guard.as_mut().ok_or("no workbook open")?;
    model.delete_sheet(sheet)
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
    let compare_guard = state.compare.lock().unwrap();
    let Some(session) = compare_guard.as_ref() else {
        return Ok(None);
    };
    let model_guard = state.model.lock().unwrap();
    let Some(left) = model_guard.as_ref() else {
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
    let bak = p.with_extension({
        let cur = p.extension().and_then(|s| s.to_str()).unwrap_or("");
        if cur.is_empty() {
            "bak".to_string()
        } else {
            format!("{cur}.bak")
        }
    });
    if p.exists() {
        let _ = std::fs::remove_file(&bak);
        std::fs::copy(p, &bak).map_err(|e| format!("backup copy failed: {e}"))?;
    }
    let save = save_workbook(path, state)?;
    Ok(BackupResult {
        save,
        backup_path: bak.to_string_lossy().into_owned(),
    })
}
