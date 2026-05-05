//! .xls (HSSF / BIFF) support via calamine.
//!
//! IronCalc 0.7.1 only reads .xlsx. To open the older binary format we
//! parse with calamine into a per-sheet matrix of (value, formula) pairs,
//! then construct a fresh IronCalc Model and replay each cell as a
//! `set_user_input` call. Formulas come through as `=...` strings;
//! literals as their text representation.
//!
//! Format coverage is limited by what calamine exposes:
//!   * Cell values + formulas: yes.
//!   * Merged cells: yes (via `worksheet_merge_cells`).
//!   * Styles (font, fill, borders, alignment, number format): NO.
//!     calamine parses these for internal value formatting but does not
//!     expose them on its public API, so we can't round-trip them
//!     through IronCalc. Affected sheets render with IronCalc defaults.
//!   * Column widths, row heights, hidden rows/cols: NO (same reason).
//!   * Custom indexed palette: NO (not exposed).
//!   * Array-formula spill ranges: NO — calamine treats PtgExp records
//!     as opaque (see xls.rs line 1216). Spilled cells return their
//!     cached value only; the formula isn't replicated. Task #44.
//!
//! Closing the gap properly needs either a patch to calamine or a
//! direct BIFF record scanner; both are significant efforts tracked as
//! task #82 (format parity).
//!
//! Save-preservation doesn't apply: there's no original .xlsx zip to
//! patch. `state.loaded` is left None for .xls workbooks, so save_workbook
//! falls back to the IronCalc save path (always produces .xlsx output).

use std::collections::{HashMap, HashSet};
use std::io::Cursor;
use std::time::Instant;

use calamine::{Data, Reader, Xls};
use ironcalc::base::types::{
    Alignment, BorderItem, BorderStyle, HorizontalAlignment, VerticalAlignment,
};
use ironcalc::base::Model;

use crate::util::col_letter_i;
use crate::xlsx_load::strip_trailing_anchor;
use crate::xls_biff::{
    decode_array_formula as fastsheet_lib_decode_array, decode_full_formula,
    scan_xls_shape,
};
use crate::xls_preserve::{extract as extract_preserved, PreservedXlsData};

/// Second return value is the hidden-column map that should be stored in
/// AppState — it lives outside the Model because IronCalc's Col struct
/// doesn't have a hidden field (see CLAUDE.md "Non-obvious behaviours").
/// Open an xls file with calamine, transparently stripping any
/// embedded VBA project on first-attempt failure. Calamine's xls
/// reader (xls.rs:189) calls `VbaProject::from_cfb` whenever the CFB
/// container has a `_VBA_PROJECT_CUR` directory, and a corrupt or
/// unsupported VBA stream panics inside the RLE decompressor
/// (cfb.rs:362) rather than returning an error. Real-world files
/// with non-standard VBA streams trigger this. We:
///   1. catch_unwind around the normal open;
///   2. if it panicked AND the file has a VBA directory, copy it to
///      a temp file with that directory removed, and retry.
/// We never use the VBA payload, so the fallback path produces a
/// workbook indistinguishable from the original for our purposes.
fn open_xls_with_vba_fallback(
    bytes: &[u8],
) -> Result<Xls<Cursor<Vec<u8>>>, String> {
    // Silence the default panic-trace print — we're converting the
    // panic into an Err and don't want stderr noise. Restore the
    // previous hook before returning.
    let prev_hook = std::panic::take_hook();
    std::panic::set_hook(Box::new(|_| {}));
    let opened = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        Xls::new(Cursor::new(bytes.to_vec()))
    }));
    std::panic::set_hook(prev_hook);
    match opened {
        Ok(Ok(wb)) => return Ok(wb),
        Ok(Err(e)) => return Err(format!("calamine open: {e}")),
        Err(_) => {
            // Fall through to the VBA-strip retry below.
        }
    }
    let stripped = strip_vba_in_memory(bytes).map_err(|e| {
        format!(
            "calamine panicked decoding the workbook AND the VBA-strip \
             fallback failed: {e}"
        )
    })?;
    let prev_hook = std::panic::take_hook();
    std::panic::set_hook(Box::new(|_| {}));
    let retry = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        Xls::new(Cursor::new(stripped))
    }));
    std::panic::set_hook(prev_hook);
    match retry {
        Ok(Ok(wb)) => Ok(wb),
        Ok(Err(e)) => Err(format!("calamine open after VBA strip: {e}")),
        Err(_) => Err(
            "calamine panicked decoding the workbook even after \
             stripping the VBA stream — file may be corrupt"
                .to_string(),
        ),
    }
}

fn strip_vba_in_memory(bytes: &[u8]) -> Result<Vec<u8>, String> {
    let mut buf = Cursor::new(bytes.to_vec());
    {
        let mut comp = cfb::CompoundFile::open(&mut buf)
            .map_err(|e| format!("cfb open: {e}"))?;
        // Remove any VBA-related storage we know about. `_VBA_PROJECT_CUR`
        // is the canonical location; some files have additional
        // co-resident `VBA` / `Macros` storages.
        for name in ["/_VBA_PROJECT_CUR", "/VBA", "/Macros", "/_VBA_PROJECT"] {
            if comp.exists(name) && comp.is_storage(name) {
                let _ = comp.remove_storage_all(name);
            }
        }
        comp.flush().map_err(|e| format!("cfb flush: {e}"))?;
    }
    Ok(buf.into_inner())
}

/// Phase timing for load_xls. Active when FASTSHEET_PROFILE_LOAD is set in
/// the environment. Each phase appends its wall-clock elapsed to the
/// profile log (see `util::profile_log`).
struct PhaseTimer {
    start: Instant,
}

impl PhaseTimer {
    fn new() -> Self {
        Self { start: Instant::now() }
    }
    fn lap(&mut self, label: &str) {
        crate::util::profile_log(&format!(
            "[load_xls] {:>20} {:>7.1}ms",
            label,
            self.start.elapsed().as_secs_f64() * 1000.0
        ));
        self.start = Instant::now();
    }
}

pub fn load_xls(
    path: &str,
) -> Result<(Model<'static>, HashMap<u32, HashSet<i32>>, PreservedXlsData), String> {
    let mut timer = PhaseTimer::new();
    crate::util::profile_log(&format!("[load_xls] === path={path}"));
    // Read the file ONCE, then drive both calamine and the BIFF
    // scanner from the in-memory bytes. The previous version opened
    // the file twice (calamine via path, scanner via cfb path) and
    // each path-based open did many small reads to walk CFB headers
    // / FAT / sectors. On a Windows .xls served from the WSL
    // `\\wsl.localhost` share that's a 9P round-trip per read, and
    // biff_scan ballooned from ~100ms to ~3700ms. Reading the whole
    // file linearly into a Vec is one syscall pair in practice and
    // lets cfb / calamine do all their seeks in RAM.
    let bytes = std::fs::read(path)
        .map_err(|e| format!("read {path}: {e}"))?;
    timer.lap("read_bytes");

    // Capture the original VBA / macro storages from the source CFB
    // so the writer can replay them into a saved file. Best-effort —
    // returns an empty PreservedXlsData when the file has no VBA
    // (which is most files).
    let preserved = extract_preserved(&bytes);
    timer.lap("preserve_extract");

    // Calamine 0.26's xls reader unconditionally parses any embedded
    // `_VBA_PROJECT_CUR` directory during `Xls::new`, and its CFB
    // RLE decompressor (cfb.rs:362) `assert_eq!`s the chunk signature
    // — corrupt or non-standard VBA streams crash the whole process.
    // We don't use the VBA payload anyway, so the open helper catches
    // the unwind and falls back to a VBA-stripped in-memory copy.
    let mut wb: Xls<Cursor<Vec<u8>>> = match open_xls_with_vba_fallback(&bytes) {
        Ok(wb) => wb,
        Err(e) => return Err(format!("xls open: {e}")),
    };
    let sheet_names: Vec<String> = wb.sheet_names().to_vec();
    if sheet_names.is_empty() {
        return Err("xls file has no sheets".to_string());
    }
    timer.lap("calamine_open");

    // BIFF scan runs over the same in-memory bytes — it walks record
    // types calamine doesn't surface (COLINFO, ROW, PANE, WINDOW2,
    // FONT, XF, PALETTE, NAME, EXTERNNAME, ARRAY, SHRFMLA). Best-
    // effort; an empty XlsShape just means everything falls back to
    // IronCalc defaults.
    let shape = scan_xls_shape(&bytes);
    timer.lap("biff_scan");

    // Sheet names that need quoting when referenced from formulas.
    // Excel's grammar: any non-alphanumeric-underscore char, or a name
    // starting with a digit, must be wrapped in single quotes. Formulas
    // that calamine returns are the raw BIFF decoded text — unquoted —
    // which IronCalc then fails to parse, turning cells into #ERROR!.
    let names_needing_quotes: Vec<String> = sheet_names
        .iter()
        .filter(|n| needs_sheet_quoting(n))
        .cloned()
        .collect();

    // Model::new_empty's name/locale/tz/language are &'a str — to keep
    // the returned Model<'static> we hardcode literals. The internal
    // workbook name is cosmetic; the file path is tracked separately.
    let mut model = Model::new_empty("workbook", "en", "UTC", "en")
        .map_err(|e| format!("Model::new_empty: {e}"))?;

    // new_empty seeds a single sheet ("Sheet1"). Rename the first to match
    // the source's first sheet, then add the rest.
    model
        .rename_sheet_by_index(0, &sheet_names[0])
        .map_err(|e| format!("rename_sheet_by_index: {e}"))?;
    for name in &sheet_names[1..] {
        model
            .add_sheet(name)
            .map_err(|e| format!("add_sheet({name}): {e}"))?;
    }
    timer.lap("model_init");

    // Load defined names before cells so formulas that reference them
    // resolve rather than evaluating to #REF!. Two BIFF quirks to
    // handle:
    //   * calamine surfaces internal Excel placeholders like
    //     `_xlfn.AVERAGEIF` with a "=Unsupported ptg: 1c" stub —
    //     those aren't real formulas; skip them.
    //   * IronCalc's parser for defined-name formulas wants the raw
    //     reference WITHOUT the leading `=`. calamine is inconsistent
    //     about including it — strip defensively.
    // new_defined_name is strict: ParsedReference only accepts simple
    // range references, so OFFSET / expression-based dynamic names
    // will fail registration. That's fine — the formulas that use them
    // would error anyway. We just log-and-skip.
    for (raw_name, raw_formula) in wb.defined_names() {
        if raw_name.is_empty() || raw_formula.is_empty() {
            continue;
        }
        if raw_name.starts_with("_xlfn") {
            continue;
        }
        if raw_formula.starts_with("Unsupported ptg")
            || raw_formula.starts_with("=Unsupported ptg")
        {
            continue;
        }
        let stripped = raw_formula.trim_start_matches('=');
        // Prefer our own decoded refs entirely. calamine's
        // `parse_defined_names` has multiple overlapping bugs — 2D
        // / 3D col masking, and for 3D it picks the wrong sheet via
        // xti. If we have a single-ref decoding from our BIFF NAME
        // scan we just use it verbatim; that's what 99% of names
        // are. If we have multiple refs (range expressions joined by
        // commas etc.), fall back to calamine + patch_refs.
        let own = shape.defined_name_refs.get(&raw_name.to_lowercase());
        let patched = match own {
            Some(refs) if refs.len() == 1 => refs[0].clone(),
            Some(refs) if refs.len() > 1 => {
                let tuples: Vec<(bool, String)> =
                    refs.iter().map(|s| (true, s.clone())).collect();
                patch_refs(stripped, &tuples)
            }
            _ => stripped.to_string(),
        };
        let unwrapped = unwrap_user_xlfn(&patched);
        let normalized = normalize_range_order(&unwrapped);
        let quoted = quote_sheet_refs(&normalized, &names_needing_quotes);
        let _ = model.new_defined_name(raw_name, None, &quoted);
    }
    timer.lap("defined_names");

    for (sheet_idx, name) in sheet_names.iter().enumerate() {
        let formulas = wb.worksheet_formula(name).ok();
        let merges = wb.worksheet_merge_cells(name);
        let range = wb
            .worksheet_range(name)
            .map_err(|e| format!("worksheet_range({name}): {e}"))?;
        let (start_row, start_col) = range.start().unwrap_or((0, 0));
        let mut max_row_1 = 0i32;
        let mut max_col_1 = 0i32;

        for (rel_row, row) in range.rows().enumerate() {
            for (rel_col, cell) in row.iter().enumerate() {
                let abs_row = start_row as usize + rel_row;
                let abs_col = start_col as usize + rel_col;

                // Prefer the formula text when present — IronCalc will
                // re-parse and re-evaluate it, so cached values from the
                // .xls don't matter.
                let r_1based_early = abs_row as i32 + 1;
                let c_1based_early = abs_col as i32 + 1;
                let sheet_key_early = sheet_idx as u32;
                // For cells where calamine returns empty formula text
                // but we have raw rgce (because the cell starts with
                // PtgExp pointing to a shared formula), decode the
                // shared formula ourselves.
                let own_decoded: Option<String> = shape
                    .ptgexp_cells
                    .get(&(sheet_key_early, r_1based_early, c_1based_early))
                    .and_then(|rgce| {
                        let raw_defined: Vec<String> = wb
                            .defined_names()
                            .iter()
                            .map(|(n, _)| n.clone())
                            .collect();
                        let result = crate::xls_biff::decode_full_formula(
                            rgce,
                            &raw_defined,
                            &shape.xti_table,
                            &shape.biff_sheet_names,
                            &shape.extern_names,
                            &shape.shared_formulas,
                            sheet_key_early,
                            r_1based_early,
                            c_1based_early,
                        );
                        if std::env::var("FASTSHEET_XLS_DEBUG_DECODE").is_ok() {
                            eprintln!(
                                "DECODE s{} r{} c{}: rgce={:02x?} → {:?}",
                                sheet_key_early, r_1based_early, c_1based_early,
                                rgce, result
                            );
                        }
                        result
                    });
                let calamine_text = formulas
                    .as_ref()
                    .and_then(|f| f.get_value((abs_row as u32, abs_col as u32)))
                    .filter(|s| !s.is_empty())
                    .cloned();
                // Prefer our own decoded text for shared-formula cells
                // (calamine returns empty); fall back to patched
                // calamine text otherwise.
                let value = match (calamine_text.clone(), own_decoded.clone()) {
                    (_, Some(ours)) => Some(if ours.starts_with('=') { ours } else { format!("={ours}") }),
                    (Some(cal), None) => Some(cal),
                    (None, None) => None,
                };
                let value = value
                    .map(|s| {
                        let with_eq = if s.starts_with('=') { s } else { format!("={s}") };
                        // Fix calamine's PtgRef3d/PtgArea3d column-
                        // index corruption by splicing in the correctly-
                        // decoded refs from our BIFF scan. Skip for
                        // own-decoded shared formulas — those are
                        // already correct.
                        let r_1based = abs_row as i32 + 1;
                        let c_1based = abs_col as i32 + 1;
                        let key = (sheet_idx as u32, r_1based, c_1based);
                        let used_own = own_decoded.is_some();
                        let patched = if used_own {
                            with_eq
                        } else {
                            // Fix calamine's 0x0C/0x0D comparison-op
                            // swap: every `>` in the original comes
                            // out as `>=` and vice-versa. We scanned
                            // the rgce for the real op sequence —
                            // apply by ordinal position.
                            // Fix calamine's PtgStr emitter FIRST: it
                            // wraps the raw bytes in `"..."` without
                            // doubling internal `"` characters, so
                            // formulas containing a literal quote
                            // come out syntactically broken (e.g.
                            // `"" Hexes"` parses as empty-string +
                            // the rest as garbage). patch_refs and
                            // patch_cmp_ops walk the text expecting
                            // valid string syntax — if we run them
                            // first they'd mis-locate string bounds
                            // and swap refs between the wrong
                            // positions, producing garbage like
                            // `LEFT(C163,2)` where the original was
                            // `LEFT(C167,2)`.
                            let str_fixed = shape
                                .formula_strings
                                .get(&key)
                                .map(|ss| patch_ptg_strings(&with_eq, ss))
                                .unwrap_or(with_eq);
                            let cmp_fixed = shape
                                .formula_cmp_ops
                                .get(&key)
                                .map(|ops| patch_cmp_ops(&str_fixed, ops))
                                .unwrap_or(str_fixed);
                            let ref_fixed = shape
                                .formula_refs
                                .get(&key)
                                .map(|refs| patch_refs(&cmp_fixed, refs))
                                .unwrap_or(cmp_fixed);
                            // Resolve PtgNameX placeholders via the
                            // EXTERNNAME table. Calamine emits
                            // `[PtgNameX]` for each PtgNameX token
                            // (xls.rs ~line 1444); we substitute the
                            // real function name (e.g. MROUND) in
                            // ordinal order using indices captured
                            // from the rgce.
                            shape
                                .formula_name_xs
                                .get(&key)
                                .map(|idxs| patch_ptg_name_x(&ref_fixed, idxs, &shape.extern_names))
                                .unwrap_or(ref_fixed) };
                        let unwrapped = unwrap_user_xlfn(&patched);
                        quote_sheet_refs(&unwrapped, &names_needing_quotes)
                    })
                    .or_else(|| literal_to_input(cell));

                let Some(v) = value else { continue };
                let r = abs_row as i32 + 1;
                let c = abs_col as i32 + 1;
                if r > max_row_1 { max_row_1 = r; }
                if c > max_col_1 { max_col_1 = c; }
                let _ = model.set_user_input(sheet_idx as u32, r, c, v);
                // For DateTime cells we wrote the raw Excel serial — apply
                // a date number-format so it renders as a date instead of
                // a number. (calamine doesn't surface the cell's original
                // format string, so we use a sensible default.)
                if matches!(cell, Data::DateTime(_)) {
                    if let Ok(mut style) = model.get_style_for_cell(sheet_idx as u32, r, c) {
                        if !is_likely_date_format(&style.num_fmt) {
                            style.num_fmt = "yyyy-mm-dd".to_string();
                            let _ = model.set_cell_style(sheet_idx as u32, r, c, &style);
                        }
                    }
                }
            }
        }

        // Apply merged ranges. calamine gives us (start, end) as 0-indexed
        // (row, col) pairs; IronCalc expects A1-style strings in
        // worksheet.merge_cells (e.g. "A1:B2").
        if let Some(dims) = merges {
            let ws = match model.workbook.worksheets.get_mut(sheet_idx) {
                Some(w) => w,
                None => continue,
            };
            for d in dims {
                let r1 = d.start.0 as i32 + 1;
                let c1 = d.start.1 as i32 + 1;
                let r2 = d.end.0 as i32 + 1;
                let c2 = d.end.1 as i32 + 1;
                ws.merge_cells.push(format!(
                    "{}{}:{}{}",
                    col_letter_i(c1),
                    r1,
                    col_letter_i(c2),
                    r2,
                ));
            }
        }

        // Apply column widths from the BIFF scan. IronCalc's
        // set_column_width takes a width in chars × COLUMN_WIDTH_FACTOR
        // (=12), so convert px → chars × 12 using the same 7-px-per-char
        // assumption the frontend uses (see utils.ts colWidthPx).
        let sheet_key = sheet_idx as u32;
        if let Some(widths) = shape.col_widths.get(&sheet_key) {
            for (&col, &px) in widths {
                let internal = (px * 12.0 / 7.0).max(0.0);
                let _ = model.set_column_width(sheet_key, col, internal);
            }
        }

        // Apply row heights: set_row_height takes pt × ROW_HEIGHT_FACTOR
        // (=2), and the scanner gave us points directly.
        if let Some(heights) = shape.row_heights.get(&sheet_key) {
            for (&row, &pt) in heights {
                let internal = (pt * 2.0).max(0.0);
                let _ = model.set_row_height(sheet_key, row, internal);
            }
        }

        // Hidden rows — stored on Worksheet.rows with the row's hidden
        // flag set. IronCalc exposes set_row_hidden for this.
        if let Some(rows) = shape.hidden_rows.get(&sheet_key) {
            for &row in rows {
                let _ = model.set_row_hidden(sheet_key, row, true);
            }
        }

        // Frozen panes — mirror the xlsx path (cells::set_frozen_panes
        // writes straight to worksheet.frozen_rows/columns).
        if let Some(&(fr_rows, fr_cols)) = shape.frozen_panes.get(&sheet_key) {
            if let Some(ws) = model.workbook.worksheets.get_mut(sheet_idx) {
                ws.frozen_rows = fr_rows.max(0);
                ws.frozen_columns = fr_cols.max(0);
            }
        }

        // Write the actual used dimension so get_sheet_dim returns
        // something meaningful. Model::new_empty leaves dimension = "A1"
        // which would force the frontend viewport to the MIN clamp and
        // truncate large xls sheets on first paint (growth-on-nav would
        // eventually recover, but first-paint reach is awful).
        if max_row_1 > 0 && max_col_1 > 0 {
            if let Some(ws) = model.workbook.worksheets.get_mut(sheet_idx) {
                ws.dimension = format!("A1:{}{}", col_letter_i(max_col_1), max_row_1);
            }
        }
    }
    timer.lap("cell_loop");

    // Per-cell styling from the BIFF XF + FONT + PALETTE scan. For each
    // cell with an XF, pull the referenced font and apply bold/italic/
    // underline/strike/size/color to its IronCalc Style. Safe to run by
    // default — these fields round-trip cleanly; it's only num_fmt that
    // gave us trouble (see the env-gated block below).
    //
    // Default-workbook font: FONT record 0 is always the workbook
    // default. Overwrite IronCalc's own default (Calibri 13) with the
    // xls file's default so cells that DON'T have an explicit XF still
    // render at the correct size/face. Without this, every unstyled
    // cell would display as 13pt Calibri even though the xls used 10pt
    // Arial throughout — which is the "fonts are the wrong size" bug.
    if let Some(f0) = shape.fonts.first() {
        if let Some(default_font) = model.workbook.styles.fonts.get_mut(0) {
            if f0.size_pt > 0 {
                default_font.sz = f0.size_pt;
            }
            if !f0.name.is_empty() {
                default_font.name = f0.name.clone();
            }
        }
    }
    let default_font_size = shape
        .fonts
        .first()
        .map(|f| f.size_pt)
        .filter(|s| *s > 0)
        .unwrap_or(10);
    let default_font_name = shape
        .fonts
        .first()
        .map(|f| f.name.clone())
        .filter(|n| !n.is_empty())
        .unwrap_or_else(|| "Arial".to_string());
    // Style + number-format application, in a single pass with a per-ixfe
    // index cache.
    //
    // The naive version (the previous two loops) called
    // `Model::set_cell_style(&Style)` once per cell, and that internally
    // does an O(unique_styles) linear scan with a Style clone+compare per
    // step in `get_style_index_or_create`. On a 30k-styled-cell workbook
    // with ~200 unique XFs that's ~6M Style clones; Hetairos paid 2.5s
    // total for this dance.
    //
    // Now: build the merged Style fresh ONCE per unique ixfe, register
    // it once via `get_style_index_or_create`, cache the returned i32,
    // and write the index directly to each subsequent cell via the cheap
    // worksheet-level setter. The cache key is just `ixfe` because all
    // cells start at default style 0 at this point in load — the only
    // exception is date-typed cells that already had a numfmt applied
    // earlier (see Data::DateTime branch in the cell loop), which take
    // the slow per-cell path so their pre-existing fields aren't lost.
    //
    // Numfmt application is folded into the same pass — the previous
    // separate loop was paying the same get/set cost a second time.
    //
    // IronCalc's formatter rejects some BIFF format strings (conditional
    // `[Red]`, locale prefixes, custom text literals) and turns the cell
    // into `#ERROR!`. Heuristic-gate to "safe" strings: digit-grouping,
    // decimals, percent, currency, and date tokens. Force-apply
    // everything via FASTSHEET_XLS_APPLY_NUMFMT=all for diagnostics.
    let force_all_numfmt = std::env::var("FASTSHEET_XLS_APPLY_NUMFMT")
        .map(|v| v == "all")
        .unwrap_or(false);
    let mut style_cache: HashMap<u16, i32> = HashMap::new();
    for (&(sheet, row, col), &ixfe) in &shape.cell_xfs {
        let xf = match shape.xfs.get(ixfe as usize) {
            Some(x) => x,
            None => continue,
        };

        // Resolve the numfmt to apply for this XF, if any.
        let numfmt_override: Option<&str> = shape
            .formats
            .get(&xf.fmt_idx)
            .map(|s| s.as_str())
            .filter(|s| !s.is_empty() && *s != "General")
            .filter(|s| force_all_numfmt || is_safe_numfmt(s));

        // Fast path: cell hasn't been styled yet, and we already built
        // the merged Style for this ixfe. Just write the cached index.
        let pre_existing_idx = model.get_cell_style_index(sheet, row, col).unwrap_or(0);
        if pre_existing_idx == 0 {
            if let Some(&cached) = style_cache.get(&ixfe) {
                if let Ok(ws) = model.workbook.worksheet_mut(sheet) {
                    let _ = ws.set_cell_style(row, col, cached);
                }
                continue;
            }
        }

        // Slow path: build the merged Style for this XF, register it,
        // and cache the resulting index for subsequent cells with the
        // same ixfe. Resolve the effective font_idx via XF inheritance.
        // If this cell XF has fAtrFnt=0 it doesn't override the parent
        // style XF's font — many real-world templates use that pattern
        // for "bordered but otherwise default" cells. Tiered rule:
        //   1. cell.font_idx == 0 → use workbook default font directly,
        //      ignore the parent style chain. Dominant case for cost-
        //      summary numeric cells whose style XFs walk to a font
        //      that was used for borders/fills only.
        //   2. cell.font_idx > 0 → bold/italic/underline/strike/color
        //      from cell font, size+face from parent style font (so
        //      cells with stale Arial 14 pointers don't visually
        //      balloon next to 11pt neighbours).
        //   3. Style XFs themselves keep their resolved font as-is.
        let resolved_font_idx = resolve_xf_font(&shape.xfs, ixfe);
        let resolved_font = match shape.fonts.get(resolved_font_idx as usize) {
            Some(f) => f.clone(),
            None => continue,
        };
        let workbook_default = shape.fonts.get(0).cloned().unwrap_or_else(|| resolved_font.clone());
        let cell_font = shape.fonts.get(xf.font_idx as usize).cloned();
        let font = if !xf.is_style && xf.font_idx == 0 {
            workbook_default.clone()
        } else if let (false, Some(c)) = (xf.is_style, cell_font) {
            crate::xls_biff::FontEntry {
                bold: c.bold,
                italic: c.italic,
                underline: c.underline,
                strike: c.strike,
                color_idx: c.color_idx,
                size_pt: resolved_font.size_pt,
                name: resolved_font.name.clone(),
            }
        } else {
            resolved_font.clone()
        };
        let font = &font;
        // Fills, borders, and alignment read from the XF directly —
        // many BIFF writers emit these attributes on cell XFs with
        // fAtrPat/fAtrBdr/fAtrAlc cleared but still expect them rendered
        // (the flag-clearing is a "serialized from a theme" artifact,
        // not a "use the parent" directive).
        let effective_fill = ResolvedFill {
            fill_pattern: xf.fill_pattern,
            fill_fg: xf.fill_fg,
            fill_bg: xf.fill_bg,
        };
        let effective_border = ResolvedBorder {
            border_left: xf.border_left,
            border_right: xf.border_right,
            border_top: xf.border_top,
            border_bottom: xf.border_bottom,
        };
        let effective_align = ResolvedAlign {
            h_align: xf.h_align,
            v_align: xf.v_align,
            wrap: xf.wrap,
        };
        let has_fill = effective_fill.fill_pattern == 1
            && effective_fill.fill_fg != 64
            && effective_fill.fill_fg != 65;
        let has_border = effective_border.border_left != 0
            || effective_border.border_right != 0
            || effective_border.border_top != 0
            || effective_border.border_bottom != 0;
        let has_align = effective_align.h_align != 0
            || effective_align.v_align != 2
            || effective_align.wrap;
        let has_nondefault_font = font.bold
            || font.italic
            || font.underline
            || font.strike
            || (font.size_pt > 0 && font.size_pt != default_font_size)
            || (font.color_idx != 0x7FFF && font.color_idx != 8 && font.color_idx != 64)
            || (!font.name.is_empty() && font.name != default_font_name);
        if !has_fill && !has_border && !has_align && !has_nondefault_font && numfmt_override.is_none() {
            continue;
        }
        let Ok(mut style) = model.get_style_for_cell(sheet, row, col) else { continue };
        // Always set size + name from the XF's font — IronCalc's default
        // is Calibri 13, but xls defaults are typically Arial 10 or
        // Calibri 11. Bold FONT records sometimes have dy_height=0 (size
        // unset — inherit the workbook default).
        style.font.sz = if font.size_pt > 0 {
            font.size_pt
        } else {
            default_font_size
        };
        if !font.name.is_empty() {
            style.font.name = font.name.clone();
        } else {
            style.font.name = default_font_name.clone();
        }
        if font.bold { style.font.b = true; }
        if font.italic { style.font.i = true; }
        if font.underline { style.font.u = true; }
        if font.strike { style.font.strike = true; }
        if font.color_idx != 0x7FFF && font.color_idx != 8 && font.color_idx != 64 {
            if let Some(hex) = shape.palette.get(&font.color_idx) {
                style.font.color = Some(hex.clone());
            }
        }
        // Fill color — only apply for solid pattern (fls == 1). Indices
        // 64 ("system window") and 65 ("system bg") mean "use theme /
        // none" — skip those.
        if has_fill {
            if let Some(hex) = shape.palette.get(&effective_fill.fill_fg) {
                style.fill.pattern_type = "solid".to_string();
                style.fill.fg_color = Some(hex.clone());
                style.fill.bg_color = None;
            }
        }
        // Alignment — map BIFF's 3-bit h/v codes to IronCalc's enums.
        let h = match effective_align.h_align {
            1 => Some(HorizontalAlignment::Left),
            2 => Some(HorizontalAlignment::Center),
            3 => Some(HorizontalAlignment::Right),
            4 => Some(HorizontalAlignment::Fill),
            5 => Some(HorizontalAlignment::Justify),
            6 => Some(HorizontalAlignment::CenterContinuous),
            7 => Some(HorizontalAlignment::Distributed),
            _ => None,
        };
        let v = match effective_align.v_align {
            0 => Some(VerticalAlignment::Top),
            1 => Some(VerticalAlignment::Center),
            2 => Some(VerticalAlignment::Bottom),
            3 => Some(VerticalAlignment::Justify),
            4 => Some(VerticalAlignment::Distributed),
            _ => None,
        };
        if h.is_some() || (v.is_some() && effective_align.v_align != 2) || effective_align.wrap {
            let mut a = style.alignment.clone().unwrap_or_default();
            if let Some(hv) = h { a.horizontal = hv; }
            if let Some(vv) = v { a.vertical = vv; }
            if effective_align.wrap { a.wrap_text = true; }
            if a != Alignment::default() {
                style.alignment = Some(a);
            }
        }
        // Borders — only flag presence; xlsx path renders any
        // Some(BorderItem) as a thin black border.
        let border_item = || BorderItem {
            style: BorderStyle::Thin,
            color: Some("#000000".to_string()),
        };
        if effective_border.border_left != 0 && style.border.left.is_none() {
            style.border.left = Some(border_item());
        }
        if effective_border.border_right != 0 && style.border.right.is_none() {
            style.border.right = Some(border_item());
        }
        if effective_border.border_top != 0 && style.border.top.is_none() {
            style.border.top = Some(border_item());
        }
        if effective_border.border_bottom != 0 && style.border.bottom.is_none() {
            style.border.bottom = Some(border_item());
        }
        // Numfmt — folded into the same Style write so the cached index
        // captures both styling and number format.
        if let Some(fmt) = numfmt_override {
            if style.num_fmt != fmt {
                style.num_fmt = fmt.to_string();
            }
        }

        // Register the merged Style and write its index into the cell.
        let style_idx = model.workbook.styles.get_style_index_or_create(&style);
        if let Ok(ws) = model.workbook.worksheet_mut(sheet) {
            let _ = ws.set_cell_style(row, col, style_idx);
        }
        // Cache by ixfe — only for cells that started at default style 0.
        // Date-formatted cells have a non-zero pre_existing_idx and take
        // the slow path so their pre-applied numfmt isn't dropped.
        if pre_existing_idx == 0 {
            style_cache.insert(ixfe, style_idx);
        }
    }
    timer.lap("style_apply");

    // Replicate MY* array formulas across their spill ranges. Mirrors
    // the xlsx path (replicate_my_array_formulas) but sourced from the
    // BIFF ARRAY records the scanner collected. For each array range,
    // look up the anchor cell's formula text — if it's a MY* call, use
    // the same strip_trailing_anchor trick to transform per-cell with
    // dr/dc offsets so IronCalc doesn't short-circuit spill cells to
    // #CIRC! (each cell would otherwise reference itself as anchor).
    // Array-formula replication. The template uses MY*() UDF calls
    // (MYUNIQUE, MYSORT, MYFILTER, MYTRANSPOSE, MYSORTBLANK) as the
    // anchor formula in every ARRAY record; each spill cell then
    // needs the same call with the anchor replaced by per-cell dr/dc
    // offsets (see fastsheet_udfs.rs for the eval semantics).
    //
    // Calamine ignores PtgExp-referenced ARRAY formulas, so the
    // anchor cell we load only has the cached VALUE, not the MY*
    // formula. We decode the ARRAY record's rgce ourselves (via
    // scan_xls_shape's array_formulas map) and lift it into text
    // form. Defined-name indices come from calamine's defined_names
    // iteration (same BIFF order as our scanner's resolution).
    let raw_defined_names: Vec<String> = wb
        .defined_names()
        .iter()
        .map(|(n, _)| n.clone())
        .collect();
    for (sheet, r1, r2, c1, c2) in &shape.array_ranges {
        let key = (*sheet, *r1, *c1);
        let rgce = match shape.array_formulas.get(&key) {
            Some(b) => b,
            None => continue,
        };
        let decoded = match fastsheet_lib_decode_array(
            rgce,
            &raw_defined_names,
            &shape.xti_table,
            &shape.biff_sheet_names,
            &shape.extern_names,
        ) {
            Some(t) => t,
            None => continue,
        };
        // Only replicate formulas that start with a `MY*` UDF name.
        let upper = decoded.to_uppercase();
        if !upper.starts_with("MY") { continue; }
        // Build the per-cell formula by stripping the trailing anchor
        // and appending dr, dc offsets — same trick as the xlsx path.
        // `strip_trailing_anchor` works on a formula string that
        // already starts with `=`, and returns the prefix (including
        // the leading `=`). Our decoded text has no `=`, so we add
        // one before stripping and don't add another after.
        let head_only = match strip_trailing_anchor(&format!("={decoded}")) {
            Some(s) => s,
            None => continue,
        };
        for r in *r1..=*r2 {
            for c in *c1..=*c2 {
                let dr = r - r1;
                let dc = c - c1;
                let formula = format!("{head_only}, {dr}, {dc})");
                let _ = model.set_user_input(*sheet, r, c, formula);
            }
        }
    }
    timer.lap("my_replicate");

    // Hidden columns live in the AppState side-channel (IronCalc's Col
    // struct has no hidden field) — hand them back so the caller can
    // seed state.hidden_cols the same way xlsx does.
    Ok((model, shape.hidden_cols, preserved))
}

/// Normalize `A$1:B$1` and `$K$1:$B$65536`-style ranges so the
/// top-left corner comes first. calamine's BIFF parser emits defined
/// names in the order stored in the PtgArea structure, which can have
/// the column / row order flipped for ranges that were originally
/// anchored from the bottom-right. IronCalc's VLOOKUP (and Excel
/// itself) expect colFirst ≤ colLast and rowFirst ≤ rowLast, so we
/// swap if needed.
///
/// Conservative: only touches patterns that look exactly like a
/// sheet-qualified or unqualified A1:B2 range. Doesn't try to parse
/// whole formulas.
/// Follow the fAtrFnt / ixfParent chain to find the XF whose font
/// actually applies to cells bound to `ixfe`. Cell XFs commonly set
/// fAtrFnt=0, which means "use my parent style's font"; without
/// this resolution we'd render bold cells at whatever loud font the
/// unused XF slot references (the BIFF "Cost Overview / Discount!B1"
/// inflation case).
fn resolve_xf_font(
    xfs: &[crate::xls_biff::XfEntry],
    start_ixfe: u16,
) -> u16 {
    let mut cur = start_ixfe;
    for _ in 0..8 {
        let xf = match xfs.get(cur as usize) {
            Some(x) => x,
            None => return 0,
        };
        if xf.atr_fnt || xf.is_style {
            return xf.font_idx;
        }
        if xf.ixf_parent == cur || xf.ixf_parent == 0xFFF {
            return xf.font_idx;
        }
        cur = xf.ixf_parent;
    }
    xfs.get(cur as usize).map(|x| x.font_idx).unwrap_or(0)
}

#[derive(Default, Clone, Copy)]
struct ResolvedFill { fill_pattern: u16, fill_fg: u16, fill_bg: u16 }

// Retained for reference — strict fAtr*-based inheritance of fill/
// border/alignment matched Excel's documented spec but dropped
// rendered fills on real-world BIFF8 output (writers often emit the
// attributes on cell XFs with fAtrPat cleared but still expect
// rendering). Current loader reads these attrs directly off the XF.
#[allow(dead_code)]
fn resolve_xf_fill(
    xfs: &[crate::xls_biff::XfEntry],
    start_ixfe: u16,
) -> ResolvedFill {
    let mut cur = start_ixfe;
    for _ in 0..8 {
        let xf = match xfs.get(cur as usize) {
            Some(x) => x,
            None => return ResolvedFill::default(),
        };
        if xf.atr_pat || xf.is_style {
            return ResolvedFill {
                fill_pattern: xf.fill_pattern,
                fill_fg: xf.fill_fg,
                fill_bg: xf.fill_bg,
            };
        }
        if xf.ixf_parent == cur { break; }
        cur = xf.ixf_parent;
    }
    xfs.get(cur as usize)
        .map(|xf| ResolvedFill {
            fill_pattern: xf.fill_pattern,
            fill_fg: xf.fill_fg,
            fill_bg: xf.fill_bg,
        })
        .unwrap_or_default()
}

#[derive(Default, Clone, Copy)]
struct ResolvedBorder { border_left: u8, border_right: u8, border_top: u8, border_bottom: u8 }

#[allow(dead_code)]
fn resolve_xf_border(
    xfs: &[crate::xls_biff::XfEntry],
    start_ixfe: u16,
) -> ResolvedBorder {
    let mut cur = start_ixfe;
    for _ in 0..8 {
        let xf = match xfs.get(cur as usize) {
            Some(x) => x,
            None => return ResolvedBorder::default(),
        };
        if xf.atr_bdr || xf.is_style {
            return ResolvedBorder {
                border_left: xf.border_left,
                border_right: xf.border_right,
                border_top: xf.border_top,
                border_bottom: xf.border_bottom,
            };
        }
        if xf.ixf_parent == cur { break; }
        cur = xf.ixf_parent;
    }
    xfs.get(cur as usize)
        .map(|xf| ResolvedBorder {
            border_left: xf.border_left,
            border_right: xf.border_right,
            border_top: xf.border_top,
            border_bottom: xf.border_bottom,
        })
        .unwrap_or_default()
}

#[derive(Default, Clone, Copy)]
struct ResolvedAlign { h_align: u8, v_align: u8, wrap: bool }

#[allow(dead_code)]
fn resolve_xf_align(
    xfs: &[crate::xls_biff::XfEntry],
    start_ixfe: u16,
) -> ResolvedAlign {
    let mut cur = start_ixfe;
    for _ in 0..8 {
        let xf = match xfs.get(cur as usize) {
            Some(x) => x,
            None => return ResolvedAlign::default(),
        };
        if xf.atr_alc || xf.is_style {
            return ResolvedAlign {
                h_align: xf.h_align,
                v_align: xf.v_align,
                wrap: xf.wrap,
            };
        }
        if xf.ixf_parent == cur { break; }
        cur = xf.ixf_parent;
    }
    xfs.get(cur as usize)
        .map(|xf| ResolvedAlign {
            h_align: xf.h_align,
            v_align: xf.v_align,
            wrap: xf.wrap,
        })
        .unwrap_or_default()
}

/// Splice in correctly-decoded references over calamine's output.
///
/// `correct_refs` is an ordered list of `(is_3d, ref_text)` tuples
/// from our BIFF FORMULA-record scan. For each ref we find in
/// calamine's output (in order), we replace the ref portion with our
/// decoded text. `is_3d` distinguishes `Sheet!A1` refs from bare
/// `A1` refs so we can match calamine's patterns confidently.
///
/// Needed because calamine's formula decoder has col-decoding bugs
/// on every PtgRef/PtgArea variant it handles — see the note in
/// `xls_biff.rs::XlsShape::formula_refs` for the full picture.
///
/// The ordinal-position matching works because ptg evaluation order
/// matches infix-output order for calamine's single-cell decoder.
/// Replace calamine's comparison operators with the real ones from
/// our BIFF rgce scan. calamine maps PtgGE (0x0C) to `>` and PtgGT
/// (0x0D) to `>=` — inverted. `correct_ops` is the N operators in
/// rgce order; we scan the formula text for `<`, `<=`, `=`, `>=`,
/// `>`, `<>` in order and replace each by position.
fn patch_cmp_ops(formula: &str, correct_ops: &[&str]) -> String {
    if correct_ops.is_empty() { return formula.to_string(); }
    let bytes = formula.as_bytes();
    let mut out = String::with_capacity(formula.len());
    let mut i = 0usize;
    let mut op_idx = 0usize;
    while i < bytes.len() {
        // Skip string literals.
        if bytes[i] == b'"' {
            out.push('"');
            i += 1;
            while i < bytes.len() {
                if bytes[i] == b'"' {
                    if i + 1 < bytes.len() && bytes[i + 1] == b'"' {
                        out.push_str("\"\"");
                        i += 2;
                        continue;
                    }
                    out.push('"');
                    i += 1;
                    break;
                }
                out.push(bytes[i] as char);
                i += 1;
            }
            continue;
        }
        // Detect an operator. Prefer longer matches.
        let b0 = bytes[i];
        let b1 = bytes.get(i + 1).copied();
        let (orig_len, orig_op): (usize, &str) = if b0 == b'<' && b1 == Some(b'=') {
            (2, "<=")
        } else if b0 == b'<' && b1 == Some(b'>') {
            (2, "<>")
        } else if b0 == b'>' && b1 == Some(b'=') {
            (2, ">=")
        } else if b0 == b'<' {
            (1, "<")
        } else if b0 == b'>' {
            (1, ">")
        } else if b0 == b'=' && i > 0 {
            // Don't treat a formula-leading `=` as a comparison.
            (1, "=")
        } else {
            out.push(b0 as char);
            i += 1;
            continue;
        };
        // For `=` at i==0 (formula start) we already fell through
        // above. Otherwise `=` is only a comparison if NOT the
        // leading character after `(`, `,`, or operator — actually
        // it's simpler to always treat it as comparison and trust
        // the ordinal match to realign.
        if op_idx < correct_ops.len() {
            // Only substitute if calamine's decoding is wrong. For
            // non-`>=`/`>` ops our scan and calamine agree, so this
            // will be a no-op replacement.
            if correct_ops[op_idx] != orig_op {
                out.push_str(correct_ops[op_idx]);
            } else {
                out.push_str(orig_op);
            }
            op_idx += 1;
        } else {
            out.push_str(orig_op);
        }
        i += orig_len;
    }
    out
}

/// Rewrite calamine's string literals in a formula with the correct
/// on-wire escape form. Calamine's `0x17` (PtgStr) handler wraps the
/// raw bytes in `"..."` but never doubles internal `"` chars, so an
/// Excel string whose content is `" Hexes` (leading quote char) comes
/// out as `"" Hexes"` — which IronCalc correctly parses as "empty
/// string" + the rest as a syntax error.
///
/// Because calamine's output is ambiguous (an embedded `"` inside a
/// string is indistinguishable from "close current string, open new
/// string" with just surface-text inspection), we use the N-th
/// PtgStr's known CONTENT LENGTH to split calamine's text. Find the
/// next `"`, then consume exactly `N` chars of raw body (where N is
/// the char-length of the correct content), skip the closing `"`,
/// and emit the properly-doubled replacement. The assumption is that
/// calamine emits each PtgStr as `"` + raw_content + `"`, in order,
/// never escaping — which matches its source at xls.rs:1265-1271.
fn patch_ptg_strings(formula: &str, correct_strings: &[String]) -> String {
    if correct_strings.is_empty() {
        return formula.to_string();
    }
    let chars: Vec<char> = formula.chars().collect();
    let mut out = String::with_capacity(formula.len());
    let mut i = 0usize;
    let mut idx = 0usize;
    while i < chars.len() {
        if chars[i] != '"' || idx >= correct_strings.len() {
            out.push(chars[i]);
            i += 1;
            continue;
        }
        // Opening quote. Consume exactly len chars of the correct
        // content (calamine emitted them verbatim from the rgce).
        let content = &correct_strings[idx];
        let content_chars = content.chars().count();
        // Guard: make sure calamine's output actually has this many
        // chars available + a closing `"`. If not, fall back to
        // copying the lone `"` and continuing — should never happen
        // on a well-formed calamine output.
        if i + 1 + content_chars >= chars.len()
            || chars[i + 1 + content_chars] != '"'
        {
            out.push(chars[i]);
            i += 1;
            continue;
        }
        out.push('"');
        for ch in content.chars() {
            if ch == '"' {
                out.push_str("\"\"");
            } else {
                out.push(ch);
            }
        }
        out.push('"');
        i += 2 + content_chars;
        idx += 1;
    }
    out
}

fn patch_refs(formula: &str, correct_refs: &[(bool, String)]) -> String {
    if correct_refs.is_empty() {
        return formula.to_string();
    }
    let bytes = formula.as_bytes();
    let mut out = String::with_capacity(formula.len());
    let mut i = 0usize;
    let mut ref_idx = 0usize;
    while i < bytes.len() {
        // Skip over string literals. Excel formulas embed strings as
        // "…" with "" as the escape for a literal quote. If we walk
        // into string contents we'd happily replace "F10" → "J12"
        // because the text LOOKS like a cell ref, corrupting the
        // formula's logic. Copy strings out verbatim.
        if bytes[i] == b'"' {
            out.push('"');
            i += 1;
            while i < bytes.len() {
                if bytes[i] == b'"' {
                    // Escape check: "" is a literal quote.
                    if i + 1 < bytes.len() && bytes[i + 1] == b'"' {
                        out.push_str("\"\"");
                        i += 2;
                        continue;
                    }
                    out.push('"');
                    i += 1;
                    break;
                }
                out.push(bytes[i] as char);
                i += 1;
            }
            continue;
        }
        // Also skip quoted sheet names '…' when they're NOT a ref
        // prefix. A sheet name followed by `!` is the ref-start path
        // below and is handled there; a standalone `'…'` (e.g. inside
        // a CHAR(39) or a weird literal) gets copied.
        let at_boundary = i == 0
            || !(bytes[i - 1].is_ascii_alphanumeric()
                || bytes[i - 1] == b'_'
                || bytes[i - 1] == b'.');
        if !at_boundary {
            out.push(bytes[i] as char);
            i += 1;
            continue;
        }
        // Calamine renders PtgRef3d/PtgArea3d with unresolvable xti
        // as `#REF!ref_text` — the leading `#` on the sheet name is
        // significant to us because it'd be left behind if we only
        // replaced the `REF!ref_text` portion and would produce
        // `=#SurfaceTable!G1`-style output. Special-case it.
        if bytes[i] == b'#' && formula[i..].starts_with("#REF!") {
            let ref_start = i + 5;
            let ref_end = scan_ref_text(formula, ref_start);
            if ref_end > ref_start {
                if ref_idx < correct_refs.len() && correct_refs[ref_idx].0 {
                    out.push_str(&correct_refs[ref_idx].1);
                    ref_idx += 1;
                } else {
                    out.push_str(&formula[i..ref_end]);
                    if ref_idx < correct_refs.len() { ref_idx += 1; }
                }
                i = ref_end;
                continue;
            }
        }
        // Try a sheet-qualified ref first (3D). Our correct ref
        // already includes the sheet name (possibly fixed from
        // calamine's wrong xti resolution), so we replace the
        // ENTIRE `SheetName!ref` span.
        if let Some((sheet_end, _quoted)) = try_scan_sheet_name(formula, i) {
            if bytes.get(sheet_end) == Some(&b'!') {
                let ref_start = sheet_end + 1;
                let ref_end = scan_ref_text(formula, ref_start);
                if ref_end > ref_start {
                    if ref_idx < correct_refs.len() && correct_refs[ref_idx].0 {
                        // 3D replacement.
                        out.push_str(&correct_refs[ref_idx].1);
                        ref_idx += 1;
                    } else {
                        // Fallback: copy calamine's full text.
                        out.push_str(&formula[i..ref_end]);
                        if ref_idx < correct_refs.len() {
                            ref_idx += 1;
                        }
                    }
                    i = ref_end;
                    continue;
                }
            }
        }
        // Try a bare A1 ref (2D).
        let ref_end = scan_ref_text(formula, i);
        if ref_end > i {
            if ref_idx < correct_refs.len() && !correct_refs[ref_idx].0 {
                out.push_str(&correct_refs[ref_idx].1);
                ref_idx += 1;
            } else {
                out.push_str(&formula[i..ref_end]);
            }
            i = ref_end;
            continue;
        }
        out.push(bytes[i] as char);
        i += 1;
    }
    out
}

fn try_scan_sheet_name(formula: &str, start: usize) -> Option<(usize, bool)> {
    let (end, quoted) = scan_sheet_name(formula, start);
    if end == start {
        None
    } else {
        Some((end, quoted))
    }
}

/// Return the end offset of a sheet-name token starting at `start`.
/// If the name is quoted (`'Foo Bar'`), include the closing quote.
/// If it's an unquoted identifier, scan while alphanumeric/underscore.
fn scan_sheet_name(formula: &str, start: usize) -> (usize, bool) {
    let bytes = formula.as_bytes();
    if bytes.get(start) == Some(&b'\'') {
        let mut i = start + 1;
        while i < bytes.len() {
            if bytes[i] == b'\'' {
                // Could be doubled `''` inside the name.
                if bytes.get(i + 1) == Some(&b'\'') { i += 2; continue; }
                return (i + 1, true);
            }
            i += 1;
        }
        return (start, true); // unterminated
    }
    let mut i = start;
    while i < bytes.len()
        && (bytes[i].is_ascii_alphanumeric() || bytes[i] == b'_' || bytes[i] == b' ')
    {
        i += 1;
    }
    (i, false)
}

/// From `start`, scan as much A1/range reference text as possible:
/// `[$]COL[$]ROW` optionally followed by `:[$]COL[$]ROW`.
fn scan_ref_text(formula: &str, start: usize) -> usize {
    let bytes = formula.as_bytes();
    let after_a1 = match scan_a1(formula, start) {
        Some(i) => i,
        None => return start,
    };
    if bytes.get(after_a1) != Some(&b':') { return after_a1; }
    match scan_a1(formula, after_a1 + 1) {
        Some(i) => i,
        None => after_a1,
    }
}

fn scan_a1(formula: &str, start: usize) -> Option<usize> {
    let bytes = formula.as_bytes();
    let mut i = start;
    if bytes.get(i) == Some(&b'$') { i += 1; }
    let col_start = i;
    while matches!(bytes.get(i), Some(b) if b.is_ascii_uppercase()) { i += 1; }
    if i - col_start == 0 || i - col_start > 3 { return None; }
    if bytes.get(i) == Some(&b'$') { i += 1; }
    let row_start = i;
    while matches!(bytes.get(i), Some(b) if b.is_ascii_digit()) { i += 1; }
    if i == row_start { return None; }
    Some(i)
}

fn normalize_range_order(formula: &str) -> String {
    let bytes = formula.as_bytes();
    let mut out = String::with_capacity(formula.len());
    let mut i = 0usize;
    while i < bytes.len() {
        // Look for a range pattern: [$]COL[$]ROW:[$]COL[$]ROW starting
        // at a non-ident boundary. Ranges are usually preceded by `!`,
        // `(`, `,`, `=`, or start of string.
        let at_boundary = i == 0 || !is_ident_boundary_char(bytes[i - 1]);
        if !at_boundary {
            out.push(bytes[i] as char);
            i += 1;
            continue;
        }
        if let Some((a1_end, c1, r1, c1_abs, r1_abs)) = parse_a1(formula, i) {
            // Need a `:` and another A1 to be a range.
            if formula.as_bytes().get(a1_end) == Some(&b':') {
                if let Some((a2_end, c2, r2, c2_abs, r2_abs)) =
                    parse_a1(formula, a1_end + 1)
                {
                    let c1_i = col_to_index(&c1);
                    let c2_i = col_to_index(&c2);
                    let (col_lo, col_lo_abs, col_hi, col_hi_abs) = if c1_i > c2_i {
                        (&c2, c2_abs, &c1, c1_abs)
                    } else {
                        (&c1, c1_abs, &c2, c2_abs)
                    };
                    let (row_lo, row_lo_abs, row_hi, row_hi_abs) = if r1 > r2 {
                        (r2, r2_abs, r1, r1_abs)
                    } else {
                        (r1, r1_abs, r2, r2_abs)
                    };
                    if col_lo != &c1 || row_lo != r1 {
                        let a = if col_lo_abs { "$" } else { "" };
                        let b = if row_lo_abs { "$" } else { "" };
                        let c = if col_hi_abs { "$" } else { "" };
                        let d = if row_hi_abs { "$" } else { "" };
                        out.push_str(&format!("{a}{col_lo}{b}{row_lo}:{c}{col_hi}{d}{row_hi}"));
                        i = a2_end;
                        continue;
                    }
                    out.push_str(&formula[i..a2_end]);
                    i = a2_end;
                    continue;
                }
            }
        }
        out.push(bytes[i] as char);
        i += 1;
    }
    out
}

fn is_ident_boundary_char(b: u8) -> bool {
    b.is_ascii_alphanumeric() || b == b'_'
}

/// Parse `[$]COL[$]ROW` at `start`. Returns (end, col_str, row, col_abs, row_abs).
fn parse_a1(formula: &str, start: usize) -> Option<(usize, String, u32, bool, bool)> {
    let bytes = formula.as_bytes();
    let mut i = start;
    let c_abs = bytes.get(i) == Some(&b'$');
    if c_abs { i += 1; }
    let col_start = i;
    while let Some(&b) = bytes.get(i) {
        if b.is_ascii_uppercase() { i += 1; } else { break; }
    }
    if i == col_start || i - col_start > 3 { return None; }
    let col = formula[col_start..i].to_string();
    let r_abs = bytes.get(i) == Some(&b'$');
    if r_abs { i += 1; }
    let row_start = i;
    while let Some(&b) = bytes.get(i) {
        if b.is_ascii_digit() { i += 1; } else { break; }
    }
    if i == row_start { return None; }
    let row: u32 = formula[row_start..i].parse().ok()?;
    Some((i, col, row, c_abs, r_abs))
}

fn col_to_index(s: &str) -> u32 {
    let mut n = 0u32;
    for b in s.bytes() {
        n = n * 26 + (b - b'A' + 1) as u32;
    }
    n
}

/// Unwrap calamine's `User(_xlfn.FOO, args...)` encoding for post-2007
/// functions. BIFF stores these as FTAB entry 255 ("User") with the
/// function name as the first argument; calamine emits them literally
/// as `User(_xlfn.IFERROR, ...)`, which IronCalc's parser doesn't
/// understand. Rewrite the outer call so `FOO(args...)` is what
/// IronCalc sees.
///
/// Nested `User(...)` calls are handled one layer at a time via the
/// outer-pass loop — we walk the string, and for every `User(` we find,
/// rewrite to `FOO(` if the first argument is `_xlfn.FOO` or `_xlfn.FOO(`.
/// Replace each `[PtgNameX]` placeholder calamine emits (xls.rs
/// ~1444) with the matching extern-name resolved via the BIFF
/// EXTERNNAME table. We rewrite to the `_xlfn.<name>` shape so
/// `unwrap_user_xlfn` (the next step) handles the surrounding
/// `User(...)` call uniformly with the regular `_xlfn.IFERROR` /
/// `_xlfn.STDEV.S` cases. Indices are 1-based; missing or
/// out-of-range entries (shouldn't happen on well-formed files)
/// fall back to a name-less marker so the formula errors loudly
/// instead of silently picking the wrong function.
fn patch_ptg_name_x(formula: &str, indices: &[u16], extern_names: &[String]) -> String {
    if indices.is_empty() || !formula.contains("[PtgNameX]") {
        return formula.to_string();
    }
    let mut out = String::with_capacity(formula.len());
    let mut rest = formula;
    let mut i = 0usize;
    while let Some(pos) = rest.find("[PtgNameX]") {
        out.push_str(&rest[..pos]);
        let name = indices
            .get(i)
            .and_then(|idx| {
                if *idx == 0 {
                    None
                } else {
                    extern_names.get((*idx as usize).saturating_sub(1))
                }
            })
            .filter(|s| !s.is_empty())
            .cloned()
            .unwrap_or_else(|| "_unresolved_externname_".to_string());
        out.push_str("_xlfn.");
        out.push_str(&name);
        rest = &rest[pos + "[PtgNameX]".len()..];
        i += 1;
    }
    out.push_str(rest);
    out
}

fn unwrap_user_xlfn(formula: &str) -> String {
    // Fast path: no "User(" → nothing to do.
    if !formula.contains("User(") && !formula.contains("user(") {
        return formula.to_string();
    }
    let bytes = formula.as_bytes();
    let mut out = String::with_capacity(formula.len());
    let mut i = 0usize;
    while i < bytes.len() {
        // Match `User(` at a word boundary.
        let matches = (i == 0 || !is_ident_char(bytes[i - 1]))
            && i + 5 <= bytes.len()
            && (formula[i..].starts_with("User(") || formula[i..].starts_with("user("));
        if !matches {
            out.push(bytes[i] as char);
            i += 1;
            continue;
        }
        // `User(` — look for `_xlfn.FOO` as the first argument. The
        // argument is terminated by either `,` (no-arg call) or `(`
        // (arg call), whichever comes first at depth 0 from the `(` we
        // just consumed.
        let after_paren = i + 5;
        let (fn_name, tail_start) = match parse_xlfn_head(formula, after_paren) {
            Some(x) => x,
            None => {
                // Not an _xlfn call — copy `User(` verbatim and continue.
                out.push_str(&formula[i..after_paren]);
                i = after_paren;
                continue;
            }
        };
        // Emit `FOO(` and continue from tail_start (which points past
        // the `,` separating name from first real arg, OR at the `)` if
        // the _xlfn call was parenthesized and had its own args that
        // we've already committed to).
        out.push_str(&fn_name);
        out.push('(');
        i = tail_start;
    }
    out
}

fn is_ident_char(b: u8) -> bool {
    b.is_ascii_alphanumeric() || b == b'_' || b == b'.'
}

/// Given formula text and an index pointing at the first char after
/// `User(`, try to identify an `_xlfn.FOO` first argument. Returns
/// (bare_function_name, resume_index) on success. resume_index points
/// at the first character of the remaining args, with the leading
/// comma already consumed — or at the position just past the close of
/// the User() call if _xlfn.FOO has no args.
fn parse_xlfn_head(formula: &str, start: usize) -> Option<(String, usize)> {
    let bytes = formula.as_bytes();
    // Expect `_xlfn.`
    let prefix = "_xlfn.";
    if !formula[start..].starts_with(prefix) {
        return None;
    }
    let name_start = start + prefix.len();
    // Read identifier chars for the function name.
    let mut j = name_start;
    while j < bytes.len() && is_ident_name_char(bytes[j]) {
        j += 1;
    }
    if j == name_start {
        return None;
    }
    let fn_name = formula[name_start..j].to_string();
    // Three shapes to handle:
    //   User(_xlfn.FOO,<args>)           ← most common
    //   User(_xlfn.FOO(<args>))          ← also seen
    //   User(_xlfn.FOO)                  ← no-arg call
    let ch = bytes.get(j).copied();
    match ch {
        Some(b',') => Some((fn_name, j + 1)),
        Some(b')') => {
            // User(_xlfn.FOO) — zero-arg call. Emit FOO( then let caller
            // hit the `)` which the outer loop emits verbatim. We consume
            // through the `)` so that the outer `User(` `)` pair gets
            // rewritten to `FOO()`.
            Some((fn_name, j))
        }
        // `User(_xlfn.FOO(...))` is theoretically possible but I
        // haven't seen it in practice. Skip rewriting to avoid
        // producing an unbalanced paren (the trailing `)` of the
        // outer User() would be left dangling). Fall through to the
        // caller, which copies `User(` verbatim.
        _ => None,
    }
}

fn is_ident_name_char(b: u8) -> bool {
    b.is_ascii_alphanumeric() || b == b'_'
}

/// True if a sheet name needs to be wrapped in single quotes when
/// referenced from a formula (per Excel's grammar). Covers spaces,
/// hyphens, and any non-alphanumeric-non-underscore character, plus
/// names that start with a digit.
fn needs_sheet_quoting(name: &str) -> bool {
    if name.is_empty() {
        return false;
    }
    let first = name.chars().next().unwrap();
    if first.is_ascii_digit() {
        return true;
    }
    name.chars()
        .any(|c| !(c.is_ascii_alphanumeric() || c == '_'))
}

/// Rewrite `SheetName!A1` → `'SheetName'!A1` for every known name that
/// needs quoting. Skips references already wrapped in quotes. Conservative
/// word-boundary check: only matches when the preceding char is NOT an
/// alphanumeric / underscore / apostrophe (so we don't quote inside an
/// already-quoted name or inside another identifier).
fn quote_sheet_refs(formula: &str, names: &[String]) -> String {
    if names.is_empty() {
        return formula.to_string();
    }
    // Longest names first so `Cable Workings` matches before `Cable`.
    let mut sorted: Vec<&str> = names.iter().map(|s| s.as_str()).collect();
    sorted.sort_by_key(|s| std::cmp::Reverse(s.len()));

    let mut result = formula.to_string();
    for name in sorted {
        let needle = format!("{name}!");
        let replacement = format!("'{name}'!");
        let mut out = String::new();
        let mut last = 0usize;
        let bytes = result.as_bytes();
        let mut i = 0usize;
        while i + needle.len() <= bytes.len() {
            if result[i..].starts_with(&needle) {
                // Guard: don't double-quote. If preceding char is `'` the
                // name is already quoted (ends with `!` after closing
                // apostrophe → handled by a different branch, not here).
                let prev = if i == 0 { 0u8 } else { bytes[i - 1] };
                let already_quoted = prev == b'\'';
                // Also guard against accidentally quoting mid-identifier
                // (e.g. "MyCable Workings!" would match "Cable Workings!"
                // in the middle — rare but possible).
                let at_ident_start = prev == 0
                    || !(prev.is_ascii_alphanumeric() || prev == b'_' || prev == b'\'');
                if !already_quoted && at_ident_start {
                    out.push_str(&result[last..i]);
                    out.push_str(&replacement);
                    last = i + needle.len();
                    i = last;
                    continue;
                }
            }
            i += 1;
        }
        out.push_str(&result[last..]);
        result = out;
    }
    result
}

/// Conservative filter for number-format strings we're willing to apply
/// by default. IronCalc's formatter blows up on a subset of valid BIFF
/// formats (conditional sections, bracketed color/locale, text-literal
/// patterns). These heuristic-exclude the patterns that tripped probe.
/// Anything not matching here can still be force-applied by setting
/// `FASTSHEET_XLS_APPLY_NUMFMT=all`.
fn is_safe_numfmt(s: &str) -> bool {
    // Reject bracketed sections ([Red], [>0], [$-409], [h], etc.) —
    // these are the common source of parse errors.
    if s.contains('[') {
        return false;
    }
    // Reject text-literal markers (`@` meaning "original text") — these
    // tend to collide with IronCalc's own text handling.
    if s.contains('@') {
        return false;
    }
    // Reject padding markers (`_`) — these produce space padding in
    // Excel but IronCalc doesn't always honor them.
    if s.contains('_') {
        return false;
    }
    // Reject text in quotes — rare but trips formatter on some files.
    if s.contains('"') {
        return false;
    }
    true
}

/// Quick heuristic for "is this format string a date/time format?" — the
/// definitive list per ECMA-376 is large; we just check for the common
/// y/m/d/h/s tokens. Used so we don't stomp on a cell that already has a
/// reasonable date format (e.g. survived from an earlier load).
fn is_likely_date_format(s: &str) -> bool {
    let lower = s.to_lowercase();
    if lower.is_empty() || lower == "general" {
        return false;
    }
    lower.contains('y') || lower.contains('d') || lower.contains('h') || lower.contains('s')
        || (lower.contains('m') && !lower.contains("0.")) // m is ambiguous (month vs minute) — exclude pure-decimal fmts
}

fn literal_to_input(cell: &Data) -> Option<String> {
    match cell {
        Data::Empty => None,
        Data::String(s) => {
            if s.is_empty() {
                None
            } else {
                Some(s.clone())
            }
        }
        Data::Float(f) => Some(format_number(*f)),
        Data::Int(i) => Some(i.to_string()),
        Data::Bool(b) => Some(if *b { "TRUE".to_string() } else { "FALSE".to_string() }),
        Data::DateTime(dt) => {
            // Excel serial date — pass through the raw f64 so IronCalc
            // applies its own date format. Loses the .xls's display
            // formatting, but the value stays correct.
            Some(dt.as_f64().to_string())
        }
        Data::DateTimeIso(s) | Data::DurationIso(s) => Some(s.clone()),
        Data::Error(e) => {
            // Use the built-in Excel error functions so IronCalc
            // stores the cell as an actual error value (not a text
            // literal of the error code). Without this, ISNA() /
            // ISERROR() on the cell return FALSE and chains that
            // expect error-valued sentinels break.
            Some(match e {
                calamine::CellErrorType::NA => "=NA()".to_string(),
                calamine::CellErrorType::Div0 => "=1/0".to_string(),
                calamine::CellErrorType::Ref => "=#REF!".to_string(),
                calamine::CellErrorType::Name => "=#NAME?".to_string(),
                calamine::CellErrorType::Num => "=#NUM!".to_string(),
                calamine::CellErrorType::Value => "=#VALUE!".to_string(),
                calamine::CellErrorType::Null => "=#NULL!".to_string(),
                calamine::CellErrorType::GettingData => "=NA()".to_string(),
            })
        }
    }
}

/// Compact float formatting — strip trailing zeros for integer-valued floats
/// so set_user_input sees "42" instead of "42.0" (IronCalc treats those
/// differently when round-tripping cell types).
fn format_number(f: f64) -> String {
    if f.is_finite() && f.fract() == 0.0 && f.abs() < 1e15 {
        format!("{}", f as i64)
    } else {
        format!("{}", f)
    }
}
