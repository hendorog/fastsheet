use ironcalc::base::types::{Alignment, BorderItem, BorderStyle, HorizontalAlignment, VerticalAlignment};
use serde::{Deserialize, Serialize};
use tauri::State;

use crate::state::{AppState, ProtectedRange};
use crate::util::col_letter;

const LAST_ROW: i32 = 1_048_576;
const LAST_COLUMN: i32 = 16_384;

fn cell_is_protected(state: &AppState, sheet: u32, row: u32, col: u32) -> bool {
    state
        .protected_ranges
        .lock()
        .unwrap()
        .get(&sheet)
        .map(|ranges| ranges.iter().any(|range| range.contains(row, col)))
        .unwrap_or(false)
}

fn cell_allowed_by_input_ranges(state: &AppState, sheet: u32, row: u32, col: u32) -> bool {
    state
        .input_ranges
        .lock()
        .unwrap()
        .get(&sheet)
        .map(|ranges| ranges.iter().any(|range| range.contains(row, col)))
        .unwrap_or(true)
}

/// Flat snapshot of a cell's visual style — only what the frontend renders.
/// `None`-style cells use the workbook default (no inline CSS needed).
#[derive(Serialize, Default)]
pub(crate) struct CellStyleView {
    #[serde(skip_serializing_if = "is_false_b")]
    bold: bool,
    #[serde(skip_serializing_if = "is_false_b")]
    italic: bool,
    #[serde(skip_serializing_if = "is_false_b")]
    underline: bool,
    #[serde(skip_serializing_if = "is_false_b")]
    strike: bool,
    #[serde(skip_serializing_if = "Option::is_none")]
    size_pt: Option<i32>,
    #[serde(skip_serializing_if = "Option::is_none")]
    family: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    color: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    bg: Option<String>,
    /// "left" | "center" | "right" (general handled by frontend per cell type)
    #[serde(skip_serializing_if = "Option::is_none")]
    align_h: Option<&'static str>,
    /// "top" | "middle" | "bottom"
    #[serde(skip_serializing_if = "Option::is_none")]
    align_v: Option<&'static str>,
    #[serde(skip_serializing_if = "is_false_b")]
    wrap: bool,
    /// Border presence flags — one bit per side. Frontend renders thin
    /// black borders on whichever sides are set.
    #[serde(skip_serializing_if = "is_false_b")]
    border_top: bool,
    #[serde(skip_serializing_if = "is_false_b")]
    border_bottom: bool,
    #[serde(skip_serializing_if = "is_false_b")]
    border_left: bool,
    #[serde(skip_serializing_if = "is_false_b")]
    border_right: bool,
}

fn is_false_b(b: &bool) -> bool {
    !*b
}

#[derive(Serialize)]
pub(crate) struct CellView {
    row: u32,
    col: u32,
    /// Display string as IronCalc would format it (formulas → evaluated value).
    text: String,
    /// Original input — formula like "=SUM(A1:A5)" or the raw entered value.
    input: String,
    is_formula: bool,
    #[serde(skip_serializing_if = "Option::is_none")]
    style: Option<CellStyleView>,
}

fn build_style_view(
    s: &ironcalc::base::types::Style,
    default_sz: i32,
    default_name: &str,
) -> Option<CellStyleView> {
    use ironcalc::base::types::{HorizontalAlignment, VerticalAlignment};
    let f = &s.font;
    let fill = &s.fill;
    let mut sv = CellStyleView::default();
    let mut any = false;
    if f.b { sv.bold = true; any = true; }
    if f.i { sv.italic = true; any = true; }
    if f.u { sv.underline = true; any = true; }
    if f.strike { sv.strike = true; any = true; }
    // Only emit if it deviates from the workbook's default font size.
    // For xls loaded via load_xls we overwrote fonts[0].sz with the
    // file's FONT-record-0 size (e.g. 10 or 11). The old hardcoded
    // 13 check left cells with the file's default size EMITTING it
    // as a style override, so bold cells at 11pt would show fine
    // while unstyled 11pt cells would also render at 11pt (app-root
    // CSS default is 11pt) — no issue there. The bigger problem was
    // bold FONT records with dy_height=0 silently inheriting
    // IronCalc's own 13pt, which made bold cells render 2-3pt bigger
    // than the surrounding non-bold text. That's fixed in load_xls
    // by always applying default_font_size when size_pt is 0.
    if f.sz != default_sz { sv.size_pt = Some(f.sz); any = true; }
    if f.name != default_name {
        sv.family = Some(f.name.clone());
        any = true;
    }
    if let Some(c) = &f.color {
        if c != "#000000" {
            sv.color = Some(c.clone());
            any = true;
        }
    }
    // Background — only solid pattern with a colour we can render
    // straightforwardly. Patterns like "darkGrid" we ignore for now.
    if fill.pattern_type == "solid" {
        if let Some(c) = fill.fg_color.as_ref().or(fill.bg_color.as_ref()) {
            sv.bg = Some(c.clone());
            any = true;
        }
    }
    if let Some(a) = &s.alignment {
        let h = match a.horizontal {
            HorizontalAlignment::Left => Some("left"),
            HorizontalAlignment::Center => Some("center"),
            HorizontalAlignment::CenterContinuous => Some("center"),
            HorizontalAlignment::Right => Some("right"),
            HorizontalAlignment::Justify => Some("justify"),
            HorizontalAlignment::Fill => Some("left"),
            HorizontalAlignment::Distributed => Some("justify"),
            HorizontalAlignment::General => None,
        };
        if h.is_some() { sv.align_h = h; any = true; }
        let v = match a.vertical {
            VerticalAlignment::Top => Some("top"),
            VerticalAlignment::Center => Some("middle"),
            VerticalAlignment::Bottom => Some("bottom"),
            VerticalAlignment::Justify => Some("middle"),
            VerticalAlignment::Distributed => Some("middle"),
        };
        if v.is_some() && !matches!(a.vertical, VerticalAlignment::Bottom) {
            sv.align_v = v;
            any = true;
        }
        if a.wrap_text { sv.wrap = true; any = true; }
    }
    if s.border.top.is_some() { sv.border_top = true; any = true; }
    if s.border.bottom.is_some() { sv.border_bottom = true; any = true; }
    if s.border.left.is_some() { sv.border_left = true; any = true; }
    if s.border.right.is_some() { sv.border_right = true; any = true; }
    if any { Some(sv) } else { None }
}

/// Per-column / per-row sizing for the requested viewport. Excel's column
/// width unit is "characters of the default font"; the canonical conversion
/// is `px = floor(width * 7 + 5)` for sans-serif. Row heights are in points
/// and convert to px via `pt * 96/72`.
#[derive(Serialize)]
pub(crate) struct LayoutData {
    /// (col_index_1based, width_in_chars)
    col_widths: Vec<(u32, f64)>,
    /// (row_index_1based, height_in_points)
    row_heights: Vec<(u32, f64)>,
    /// Number of frozen header rows (top), from worksheet.frozen_rows.
    frozen_rows: i32,
    /// Number of frozen header cols (left), from worksheet.frozen_columns.
    frozen_cols: i32,
    /// Merged-cell ranges as A1-style strings (e.g. "A1:B2"). Frontend
    /// renders the anchor with colspan/rowspan and skips the others.
    merged_ranges: Vec<String>,
    /// Whether worksheet grid lines should be drawn.
    show_grid_lines: bool,
}

/// Read a rectangular block of cells [start_row..=end_row, start_col..=end_col] (1-indexed).
#[tauri::command]
pub(crate) fn get_cells(
    sheet: u32,
    start_row: u32,
    end_row: u32,
    start_col: u32,
    end_col: u32,
    state: State<'_, AppState>,
) -> Result<Vec<CellView>, String> {
    let t0 = std::time::Instant::now();
    let guard = state.model.lock().unwrap();
    let model = guard.as_ref().ok_or("no workbook open")?;
    let mut out =
        Vec::with_capacity(((end_row - start_row + 1) * (end_col - start_col + 1)) as usize);
    // Workbook-default font size + face are invariant for the
    // duration of this call. Hoist out of the per-cell loops so a
    // 200×100 viewport doesn't repeat the same .first() lookup
    // 20 000 times.
    let default_font = model.workbook.styles.fonts.first();
    let default_sz = default_font.map(|f| f.sz).unwrap_or(13);
    let default_name = default_font.map(|f| f.name.as_str()).unwrap_or("Calibri");
    for row in start_row..=end_row {
        for col in start_col..=end_col {
            let text = model
                .get_formatted_cell_value(sheet, row as i32, col as i32)
                .unwrap_or_default();
            let input = model
                .get_localized_cell_content(sheet, row as i32, col as i32)
                .unwrap_or_default();
            let is_formula = input.starts_with('=');
            // Per-cell style — None when the cell uses the workbook default,
            // saving a chunk of payload size on big viewports.
            let style = model
                .get_style_for_cell(sheet, row as i32, col as i32)
                .ok()
                .as_ref()
                .and_then(|s| build_style_view(s, default_sz, default_name));
            // Elide truly empty cells from the wire: no text, no input, no
            // style. Frontend `cells.get(key)` returns undefined for these
            // and renders them as transparent passthrough — exactly what
            // we want for spill rendering anyway. Cuts the get_cells
            // payload by ~10× on typical sheets.
            if text.is_empty() && input.is_empty() && style.is_none() {
                continue;
            }
            out.push(CellView {
                row,
                col,
                text,
                input,
                is_formula,
                style,
            });
        }
    }
    let cells_count = out.len();
    let area = (end_row - start_row + 1) * (end_col - start_col + 1);
    crate::util::profile_log(&format!(
        "[get_cells] sheet={} {}x{} area={} kept={} {:>7.1}ms",
        sheet,
        end_row - start_row + 1,
        end_col - start_col + 1,
        area,
        cells_count,
        t0.elapsed().as_secs_f64() * 1000.0
    ));
    Ok(out)
}

#[tauri::command]
pub(crate) fn get_layout(
    sheet: u32,
    start_row: u32,
    end_row: u32,
    start_col: u32,
    end_col: u32,
    state: State<'_, AppState>,
) -> Result<LayoutData, String> {
    let t0 = std::time::Instant::now();
    let guard = state.model.lock().unwrap();
    let model = guard.as_ref().ok_or("no workbook open")?;
    let ws = model
        .workbook
        .worksheets
        .get(sheet as usize)
        .ok_or("bad sheet index")?;
    // Hidden columns live in a side-channel set populated on workbook open
    // (see workbook::open_workbook) and updated by set_column_hidden.
    // IronCalc's `Col` struct has no `hidden` field, so this state is the
    // authoritative source.
    let hidden_cols_guard = state.hidden_cols.lock().unwrap();
    let empty_set = std::collections::HashSet::new();
    let hidden_cols = hidden_cols_guard.get(&sheet).unwrap_or(&empty_set);
    let is_col_hidden = |c: i32| hidden_cols.contains(&c);
    let mut col_widths = Vec::new();
    for col in start_col..=end_col {
        let w = if is_col_hidden(col as i32) {
            0.0
        } else {
            model.get_column_width(sheet, col as i32).unwrap_or(0.0)
        };
        col_widths.push((col, w));
    }
    let mut row_heights = Vec::new();
    for row in start_row..=end_row {
        let hidden = ws
            .rows
            .iter()
            .find(|r| r.r == row as i32)
            .map(|r| r.hidden)
            .unwrap_or(false);
        let h = if hidden {
            0.0
        } else {
            model.get_row_height(sheet, row as i32).unwrap_or(0.0)
        };
        row_heights.push((row, h));
    }
    let layout = LayoutData {
        col_widths,
        row_heights,
        frozen_rows: ws.frozen_rows,
        frozen_cols: ws.frozen_columns,
        merged_ranges: ws.merge_cells.clone(),
        show_grid_lines: ws.show_grid_lines,
    };
    crate::util::profile_log(&format!(
        "[get_layout] sheet={} rows={} cols={} {:>7.1}ms",
        sheet,
        end_row - start_row + 1,
        end_col - start_col + 1,
        t0.elapsed().as_secs_f64() * 1000.0
    ));
    Ok(layout)
}

#[tauri::command]
pub(crate) fn set_cell(
    sheet: u32,
    row: u32,
    col: u32,
    value: String,
    state: State<'_, AppState>,
) -> Result<String, String> {
    if !cell_allowed_by_input_ranges(&state, sheet, row, col) {
        return Err(format!("cell {}{} is outside the input range", col_letter(col), row));
    }
    if cell_is_protected(&state, sheet, row, col) {
        return Err(format!("cell {}{} is protected", col_letter(col), row));
    }
    let mut guard = state.model.lock().unwrap();
    let model = guard.as_mut().ok_or("no workbook open")?;
    let previous = model
        .get_localized_cell_content(sheet, row as i32, col as i32)
        .unwrap_or_default();
    let changed = previous != value;
    if value.is_empty() {
        model
            .cell_clear_contents(sheet, row as i32, col as i32)
            .map_err(|e| e)?;
    } else {
        model
            .set_user_input(sheet, row as i32, col as i32, value.clone())
            .map_err(|e| e)?;
    }
    // Recalc when auto-recalc is on — without this, formula edits
    // leave the cell as `Cell::CellFormula` (un-evaluated), which
    // `cell.value()` reports as the literal string "#ERROR!" until
    // the next manual F9. With auto-recalc on the user gets the real
    // value back immediately and downstream operations like number-
    // format changes display correctly. The lock is held across the
    // evaluate; evaluator is single-threaded by design.
    let auto = *state.auto_recalc.lock().unwrap();
    if auto {
        model.evaluate();
    }
    // Track the edit for in-place preservation on save. We store the user's
    // raw input so the saver can re-classify it (number / formula / string).
    state
        .dirty
        .lock()
        .unwrap()
        .insert((sheet, row as i32, col as i32), value);
    if changed {
        *state.workbook_dirty.lock().unwrap() = true;
    }
    Ok(model
        .get_formatted_cell_value(sheet, row as i32, col as i32)
        .unwrap_or_default())
}

#[tauri::command]
pub(crate) fn protect_range(
    sheet: u32,
    r1: u32,
    c1: u32,
    r2: u32,
    c2: u32,
    state: State<'_, AppState>,
) -> usize {
    let range = ProtectedRange::normalized(r1, c1, r2, c2);
    let mut protected = state.protected_ranges.lock().unwrap();
    let ranges = protected.entry(sheet).or_default();
    ranges.push(range);
    ranges.len()
}

#[tauri::command]
pub(crate) fn unprotect_range(
    sheet: u32,
    r1: u32,
    c1: u32,
    r2: u32,
    c2: u32,
    state: State<'_, AppState>,
) -> usize {
    let target = ProtectedRange::normalized(r1, c1, r2, c2);
    let mut protected = state.protected_ranges.lock().unwrap();
    let Some(ranges) = protected.get_mut(&sheet) else {
        return 0;
    };
    let before = ranges.len();
    ranges.retain(|range| !range.overlaps(&target));
    before - ranges.len()
}

#[tauri::command]
pub(crate) fn restrict_input_range(
    sheet: u32,
    r1: u32,
    c1: u32,
    r2: u32,
    c2: u32,
    state: State<'_, AppState>,
) {
    let range = ProtectedRange::normalized(r1, c1, r2, c2);
    state.input_ranges.lock().unwrap().insert(sheet, vec![range]);
}

#[tauri::command]
pub(crate) fn clear_input_restriction(sheet: u32, state: State<'_, AppState>) -> bool {
    state.input_ranges.lock().unwrap().remove(&sheet).is_some()
}

#[tauri::command]
pub(crate) fn set_show_grid_lines(
    sheet: u32,
    show: bool,
    state: State<'_, AppState>,
) -> Result<(), String> {
    let mut guard = state.model.lock().unwrap();
    let model = guard.as_mut().ok_or("no workbook open")?;
    model.set_show_grid_lines(sheet, show)?;
    drop(guard);
    *state.workbook_dirty.lock().unwrap() = true;
    Ok(())
}

/// Get / set auto-recalc state. The frontend uses these for the
/// `/W G R` menu items and the status-bar indicator. Setting auto-
/// recalc ON does NOT itself evaluate — pair with `recalc` if the
/// caller wants stale formulas refreshed at the same time.
#[tauri::command]
pub(crate) fn get_auto_recalc(state: State<'_, AppState>) -> bool {
    *state.auto_recalc.lock().unwrap()
}

#[tauri::command]
pub(crate) fn set_auto_recalc(enabled: bool, state: State<'_, AppState>) -> Result<(), String> {
    *state.auto_recalc.lock().unwrap() = enabled;
    Ok(())
}

#[tauri::command]
pub(crate) fn recalc(state: State<'_, AppState>) -> Result<u128, String> {
    use std::time::Instant;
    let mut guard = state.model.lock().unwrap();
    let model = guard.as_mut().ok_or("no workbook open")?;
    let t0 = Instant::now();
    model.evaluate();
    let ms = t0.elapsed().as_millis();
    crate::util::profile_log(&format!("[recalc] {:>7}ms", ms));
    Ok(ms)
}

#[tauri::command]
pub(crate) fn cell_addr(row: u32, col: u32) -> String {
    format!("{}{}", col_letter(col), row)
}

/// Trace the dependency chain for the formula at (sheet, row, col).
/// Returns a tree the frontend renders as a popup. Every cell node
/// includes its address, formula text (if any), evaluated value, and
/// recursively the cells/ranges/named-ranges its formula depends on.
#[tauri::command]
pub(crate) fn trace_formula(
    sheet: u32,
    row: i32,
    col: i32,
    state: State<'_, AppState>,
) -> Result<crate::trace::TraceNode, String> {
    let guard = state.model.lock().unwrap();
    let model = guard.as_ref().ok_or("no workbook open")?;
    let compare_guard = state.compare.lock().unwrap();
    Ok(crate::trace::trace(
        model,
        sheet,
        row,
        col,
        compare_guard.as_ref(),
    ))
}

/// List every defined name in the workbook with its resolved formula
/// string and (best-effort) jump target. Used by the /Formula Names
/// menu option for browse + jump-to.
#[tauri::command]
pub(crate) fn list_named_ranges(
    state: State<'_, AppState>,
) -> Result<Vec<crate::trace::NamedRangeInfo>, String> {
    let guard = state.model.lock().unwrap();
    let model = guard.as_ref().ok_or("no workbook open")?;
    Ok(crate::trace::list_named_ranges(model))
}

/// Create a workbook-scoped named range pointing at the given range.
#[tauri::command]
pub(crate) fn define_name(
    name: String,
    sheet: u32,
    r1: i32,
    c1: i32,
    r2: i32,
    c2: i32,
    state: State<'_, AppState>,
) -> Result<(), String> {
    let mut guard = state.model.lock().unwrap();
    let model = guard.as_mut().ok_or("no workbook open")?;
    let sheet_name = model
        .workbook
        .worksheets
        .get(sheet as usize)
        .ok_or("bad sheet index")?
        .name
        .clone();
    let needs_quotes = sheet_name.contains(' ') || sheet_name.contains('-');
    let qualified = if needs_quotes {
        format!("'{}'", sheet_name.replace('\'', "''"))
    } else {
        sheet_name
    };
    let formula = format!(
        "={}!${}${}:${}${}",
        qualified,
        col_letter(c1 as u32),
        r1,
        col_letter(c2 as u32),
        r2
    );
    model.new_defined_name(&name, None, &formula)?;
    drop(guard);
    *state.workbook_dirty.lock().unwrap() = true;
    Ok(())
}

#[tauri::command]
pub(crate) fn delete_name(name: String, state: State<'_, AppState>) -> Result<(), String> {
    let mut guard = state.model.lock().unwrap();
    let model = guard.as_mut().ok_or("no workbook open")?;
    model.delete_defined_name(&name, None)?;
    drop(guard);
    *state.workbook_dirty.lock().unwrap() = true;
    Ok(())
}

#[tauri::command]
pub(crate) fn list_names(state: State<'_, AppState>) -> Result<Vec<(String, String)>, String> {
    let guard = state.model.lock().unwrap();
    let model = guard.as_ref().ok_or("no workbook open")?;
    Ok(model
        .workbook
        .defined_names
        .iter()
        .map(|n| (n.name.clone(), n.formula.clone()))
        .collect())
}

/// One generic style mutation applied to every cell in a rectangle.
/// `kind` selects the field; the optional `value` carries the
/// argument for setters that take one (e.g. fill colour). Toggle ops
/// (bold/italic/underline) flip every cell to the OPPOSITE of the
/// first cell's current state — Excel/Google-Sheets convention so the
/// whole selection ends up consistent.
#[derive(Deserialize)]
#[serde(tag = "kind", rename_all = "snake_case")]
pub(crate) enum StyleOp {
    ToggleBold,
    ToggleItalic,
    ToggleUnderline,
    ToggleStrike,
    SetBold { enabled: bool },
    SetItalic { enabled: bool },
    SetUnderline { enabled: bool },
    SetStrike { enabled: bool },
    /// Reset font attributes — clears bold, italic, underline, strike,
    /// and font color back to the cell's defaults. Preserves number
    /// format, borders, fill, and alignment so the user can strip
    /// just the typography without rebuilding the rest of the cell's
    /// formatting.
    ResetAttributes,
    AlignLeft,
    AlignCenter,
    AlignRight,
    AlignJustify,
    AlignGeneral,
    AlignVerticalTop,
    AlignVerticalMiddle,
    AlignVerticalBottom,
    SetWrap { enabled: bool },
    SetFillColor { color: String },
    SetTextColor { color: String },
    ClearFillColor,
    ClearTextColor,
    ClearFormat,
    /// Apply a thin black border to one or more sides of every cell. The
    /// `where` field selects which sides get the border (combinations
    /// expressed as comma-separated tokens for simplicity).
    SetBorder {
        sides: String, // "all" | "outline" | "top" | "bottom" | "left" | "right" | "none"
    },
}

#[derive(Serialize)]
pub(crate) struct StyleEditResult {
    pub count: usize,
    /// Style index per cell BEFORE the op (row-major: r1..r2 outer, c1..c2 inner).
    pub prev_indices: Vec<i32>,
    /// Style index per cell AFTER the op (same order).
    pub next_indices: Vec<i32>,
}

#[tauri::command]
pub(crate) fn set_range_style(
    sheet: u32,
    r1: i32,
    c1: i32,
    r2: i32,
    c2: i32,
    op: StyleOp,
    state: State<'_, AppState>,
) -> Result<StyleEditResult, String> {
    let mut guard = state.model.lock().unwrap();
    let model = guard.as_mut().ok_or("no workbook open")?;

    // For toggle ops, flip to !first-cell's-state so the whole selection
    // ends up consistent rather than alternating.
    let toggle_target: Option<bool> = match &op {
        StyleOp::ToggleBold => Some(!model.get_style_for_cell(sheet, r1, c1)?.font.b),
        StyleOp::ToggleItalic => Some(!model.get_style_for_cell(sheet, r1, c1)?.font.i),
        StyleOp::ToggleUnderline => Some(!model.get_style_for_cell(sheet, r1, c1)?.font.u),
        StyleOp::ToggleStrike => Some(!model.get_style_for_cell(sheet, r1, c1)?.font.strike),
        _ => None,
    };

    let mut n = 0;
    let mut prev_indices = Vec::new();
    let mut next_indices = Vec::new();
    for r in r1..=r2 {
        for c in c1..=c2 {
            let prev_idx = model.get_cell_style_index(sheet, r, c)?;
            if matches!(&op, StyleOp::ClearFormat) {
                model.workbook.worksheet_mut(sheet)?.set_cell_style(r, c, 0)?;
                let next_idx = model.get_cell_style_index(sheet, r, c)?;
                prev_indices.push(prev_idx);
                next_indices.push(next_idx);
                n += 1;
                continue;
            }
            let mut s = model.get_style_for_cell(sheet, r, c)?;
            match &op {
                StyleOp::ToggleBold => s.font.b = toggle_target.unwrap(),
                StyleOp::ToggleItalic => s.font.i = toggle_target.unwrap(),
                StyleOp::ToggleUnderline => s.font.u = toggle_target.unwrap(),
                StyleOp::ToggleStrike => s.font.strike = toggle_target.unwrap(),
                StyleOp::SetBold { enabled } => s.font.b = *enabled,
                StyleOp::SetItalic { enabled } => s.font.i = *enabled,
                StyleOp::SetUnderline { enabled } => s.font.u = *enabled,
                StyleOp::SetStrike { enabled } => s.font.strike = *enabled,
                StyleOp::ResetAttributes => {
                    s.font.b = false;
                    s.font.i = false;
                    s.font.u = false;
                    s.font.strike = false;
                    s.font.color = None;
                }
                StyleOp::AlignLeft => {
                    let mut a = s.alignment.clone().unwrap_or_default();
                    a.horizontal = HorizontalAlignment::Left;
                    s.alignment = Some(a);
                }
                StyleOp::AlignCenter => {
                    let mut a = s.alignment.clone().unwrap_or_default();
                    a.horizontal = HorizontalAlignment::Center;
                    s.alignment = Some(a);
                }
                StyleOp::AlignRight => {
                    let mut a = s.alignment.clone().unwrap_or_default();
                    a.horizontal = HorizontalAlignment::Right;
                    s.alignment = Some(a);
                }
                StyleOp::AlignJustify => {
                    let mut a = s.alignment.clone().unwrap_or_default();
                    a.horizontal = HorizontalAlignment::Justify;
                    s.alignment = Some(a);
                }
                StyleOp::AlignGeneral => {
                    let mut a = s.alignment.clone().unwrap_or_default();
                    a.horizontal = HorizontalAlignment::General;
                    if a == Alignment::default() {
                        s.alignment = None;
                    } else {
                        s.alignment = Some(a);
                    }
                }
                StyleOp::AlignVerticalTop => {
                    let mut a = s.alignment.clone().unwrap_or_default();
                    a.vertical = VerticalAlignment::Top;
                    s.alignment = Some(a);
                }
                StyleOp::AlignVerticalMiddle => {
                    let mut a = s.alignment.clone().unwrap_or_default();
                    a.vertical = VerticalAlignment::Center;
                    s.alignment = Some(a);
                }
                StyleOp::AlignVerticalBottom => {
                    let mut a = s.alignment.clone().unwrap_or_default();
                    a.vertical = VerticalAlignment::Bottom;
                    if a == Alignment::default() {
                        s.alignment = None;
                    } else {
                        s.alignment = Some(a);
                    }
                }
                StyleOp::SetWrap { enabled } => {
                    let mut a = s.alignment.clone().unwrap_or_default();
                    a.wrap_text = *enabled;
                    if a == Alignment::default() {
                        s.alignment = None;
                    } else {
                        s.alignment = Some(a);
                    }
                }
                StyleOp::SetFillColor { color } => {
                    s.fill.pattern_type = "solid".to_string();
                    s.fill.fg_color = Some(color.clone());
                    s.fill.bg_color = None;
                }
                StyleOp::ClearFillColor => {
                    s.fill.pattern_type = "none".to_string();
                    s.fill.fg_color = None;
                    s.fill.bg_color = None;
                }
                StyleOp::SetTextColor { color } => {
                    s.font.color = Some(color.clone());
                }
                StyleOp::ClearTextColor => {
                    s.font.color = None;
                }
                StyleOp::ClearFormat => {}
                StyleOp::SetBorder { sides } => {
                    let item = BorderItem {
                        style: BorderStyle::Thin,
                        color: Some("#000000".to_string()),
                    };
                    let on_outline = sides == "outline";
                    let on_all = sides == "all";
                    let none = sides == "none";
                    let on_top = on_all || sides == "top" || (on_outline && r == r1);
                    let on_bottom = on_all || sides == "bottom" || (on_outline && r == r2);
                    let on_left = on_all || sides == "left" || (on_outline && c == c1);
                    let on_right = on_all || sides == "right" || (on_outline && c == c2);
                    if none {
                        s.border.top = None;
                        s.border.bottom = None;
                        s.border.left = None;
                        s.border.right = None;
                    } else {
                        if on_top { s.border.top = Some(item.clone()); }
                        if on_bottom { s.border.bottom = Some(item.clone()); }
                        if on_left { s.border.left = Some(item.clone()); }
                        if on_right { s.border.right = Some(item); }
                    }
                }
            }
            model.set_cell_style(sheet, r, c, &s)?;
            let next_idx = model.get_cell_style_index(sheet, r, c)?;
            prev_indices.push(prev_idx);
            next_indices.push(next_idx);
            n += 1;
        }
    }
    drop(guard);
    state.style_dirty.lock().unwrap().insert(sheet);
    *state.workbook_dirty.lock().unwrap() = true;
    Ok(StyleEditResult { count: n, prev_indices, next_indices })
}

/// Restore per-cell style indices captured by a previous set_range_style
/// call. Used by undo/redo to roll the styles back / forward.
#[tauri::command]
pub(crate) fn apply_style_indices(
    sheet: u32,
    r1: i32,
    c1: i32,
    r2: i32,
    c2: i32,
    indices: Vec<i32>,
    state: State<'_, AppState>,
) -> Result<(), String> {
    let mut guard = state.model.lock().unwrap();
    let model = guard.as_mut().ok_or("no workbook open")?;
    let ws = model.workbook.worksheet_mut(sheet)?;
    let mut i = 0;
    for r in r1..=r2 {
        for c in c1..=c2 {
            let idx = indices.get(i).copied().unwrap_or(0);
            ws.set_cell_style(r, c, idx)?;
            i += 1;
        }
    }
    drop(guard);
    state.style_dirty.lock().unwrap().insert(sheet);
    *state.workbook_dirty.lock().unwrap() = true;
    Ok(())
}

/// Apply a number-format string to every cell in a rectangle. Other style
/// attributes (font, fill, alignment, borders) are preserved per-cell —
/// we round-trip through get_style_for_cell so the format change doesn't
/// stomp on existing styling.
#[tauri::command]
pub(crate) fn set_range_number_format(
    sheet: u32,
    r1: i32,
    c1: i32,
    r2: i32,
    c2: i32,
    format: String,
    state: State<'_, AppState>,
) -> Result<usize, String> {
    let mut guard = state.model.lock().unwrap();
    let model = guard.as_mut().ok_or("no workbook open")?;
    let mut n = 0;
    for r in r1..=r2 {
        for c in c1..=c2 {
            let mut style = model.get_style_for_cell(sheet, r, c)?;
            style.num_fmt = format.clone();
            model.set_cell_style(sheet, r, c, &style)?;
            n += 1;
        }
    }
    drop(guard);
    state.style_dirty.lock().unwrap().insert(sheet);
    *state.workbook_dirty.lock().unwrap() = true;
    Ok(n)
}

/// Distinct hex colors actually used anywhere in the workbook's
/// style table. Used by the ColorPicker as the "recents" row so the
/// user can match an existing palette without retyping a hex string.
/// Only `#RRGGBB` values pass through — theme references and indexed
/// palette entries (legacy xlsx) are skipped, since the picker's
/// custom editor works in HSL space and needs a literal RGB.
#[tauri::command]
pub(crate) fn list_workbook_colors(state: State<'_, AppState>) -> Result<Vec<String>, String> {
    use std::collections::BTreeSet;
    let guard = state.model.lock().unwrap();
    let model = guard.as_ref().ok_or("no workbook open")?;
    let mut seen = BTreeSet::new();
    let mut push = |c: &Option<String>| {
        if let Some(s) = c {
            if s.starts_with('#') && s.len() == 7 {
                seen.insert(s.to_uppercase());
            }
        }
    };
    for f in &model.workbook.styles.fonts {
        push(&f.color);
    }
    for f in &model.workbook.styles.fills {
        push(&f.fg_color);
        push(&f.bg_color);
    }
    for b in &model.workbook.styles.borders {
        if let Some(s) = &b.left { push(&s.color); }
        if let Some(s) = &b.right { push(&s.color); }
        if let Some(s) = &b.top { push(&s.color); }
        if let Some(s) = &b.bottom { push(&s.color); }
        if let Some(s) = &b.diagonal { push(&s.color); }
    }
    Ok(seen.into_iter().collect())
}

/// Comprehensive style snapshot for the Format Cells modal. Pulls
/// num_fmt + alignment + font + fill + borders from a single cell so
/// the dialog can populate all tabs in one read.
#[derive(Serialize)]
pub(crate) struct CellFormatInfo {
    num_fmt: String,
    bold: bool,
    italic: bool,
    underline: bool,
    strike: bool,
    font_size: i32,
    font_name: String,
    font_color: Option<String>,
    fill_color: Option<String>,
    /// "general" | "left" | "center" | "right" | "justify"
    align_h: &'static str,
    /// "top" | "middle" | "bottom"
    align_v: &'static str,
    wrap: bool,
    border_top: bool,
    border_bottom: bool,
    border_left: bool,
    border_right: bool,
}

#[tauri::command]
pub(crate) fn get_cell_format(
    sheet: u32,
    row: i32,
    col: i32,
    state: State<'_, AppState>,
) -> Result<CellFormatInfo, String> {
    use ironcalc::base::types::{HorizontalAlignment, VerticalAlignment};
    let guard = state.model.lock().unwrap();
    let model = guard.as_ref().ok_or("no workbook open")?;
    let style = model.get_style_for_cell(sheet, row, col)?;
    let f = &style.font;
    let align_h = style
        .alignment
        .as_ref()
        .map(|a| match a.horizontal {
            HorizontalAlignment::Left => "left",
            HorizontalAlignment::Center | HorizontalAlignment::CenterContinuous => "center",
            HorizontalAlignment::Right => "right",
            HorizontalAlignment::Justify | HorizontalAlignment::Distributed => "justify",
            HorizontalAlignment::Fill | HorizontalAlignment::General => "general",
        })
        .unwrap_or("general");
    let align_v = style
        .alignment
        .as_ref()
        .map(|a| match a.vertical {
            VerticalAlignment::Top => "top",
            VerticalAlignment::Center => "middle",
            VerticalAlignment::Justify | VerticalAlignment::Distributed => "middle",
            VerticalAlignment::Bottom => "bottom",
        })
        .unwrap_or("bottom");
    let wrap = style.alignment.as_ref().map(|a| a.wrap_text).unwrap_or(false);
    let fill_color = if style.fill.pattern_type == "solid" {
        style.fill.fg_color.clone().or(style.fill.bg_color.clone())
    } else {
        None
    };
    Ok(CellFormatInfo {
        num_fmt: style.num_fmt.clone(),
        bold: f.b,
        italic: f.i,
        underline: f.u,
        strike: f.strike,
        font_size: f.sz,
        font_name: f.name.clone(),
        font_color: f.color.clone(),
        fill_color,
        align_h,
        align_v,
        wrap,
        border_top: style.border.top.is_some(),
        border_bottom: style.border.bottom.is_some(),
        border_left: style.border.left.is_some(),
        border_right: style.border.right.is_some(),
    })
}

/// Insert `count` blank rows at `row`, shifting subsequent rows down.
/// Sets `structural_dirty` so the next save bypasses the xlsx
/// preservation path — patching by absolute (row, col) coords would
/// silently desync from data after a row shift.
#[tauri::command]
pub(crate) fn insert_rows(
    sheet: u32,
    row: i32,
    count: i32,
    state: State<'_, AppState>,
) -> Result<(), String> {
    let mut guard = state.model.lock().unwrap();
    let model = guard.as_mut().ok_or("no workbook open")?;
    model.insert_rows(sheet, row, count)?;
    drop(guard);
    *state.structural_dirty.lock().unwrap() = true;
    *state.workbook_dirty.lock().unwrap() = true;
    Ok(())
}

/// Delete `count` rows starting at `row`, shifting subsequent rows up.
#[tauri::command]
pub(crate) fn delete_rows(
    sheet: u32,
    row: i32,
    count: i32,
    state: State<'_, AppState>,
) -> Result<(), String> {
    let mut guard = state.model.lock().unwrap();
    let model = guard.as_mut().ok_or("no workbook open")?;
    model.delete_rows(sheet, row, count)?;
    drop(guard);
    *state.structural_dirty.lock().unwrap() = true;
    *state.workbook_dirty.lock().unwrap() = true;
    Ok(())
}

/// Insert `count` blank cols at `col`, shifting subsequent cols right.
#[tauri::command]
pub(crate) fn insert_columns(
    sheet: u32,
    col: i32,
    count: i32,
    state: State<'_, AppState>,
) -> Result<(), String> {
    let mut guard = state.model.lock().unwrap();
    let model = guard.as_mut().ok_or("no workbook open")?;
    model.insert_columns(sheet, col, count)?;
    drop(guard);
    *state.structural_dirty.lock().unwrap() = true;
    *state.workbook_dirty.lock().unwrap() = true;
    Ok(())
}

/// Delete `count` cols starting at `col`, shifting subsequent cols left.
#[tauri::command]
pub(crate) fn delete_columns(
    sheet: u32,
    col: i32,
    count: i32,
    state: State<'_, AppState>,
) -> Result<(), String> {
    let mut guard = state.model.lock().unwrap();
    let model = guard.as_mut().ok_or("no workbook open")?;
    model.delete_columns(sheet, col, count)?;
    drop(guard);
    *state.structural_dirty.lock().unwrap() = true;
    *state.workbook_dirty.lock().unwrap() = true;
    Ok(())
}

fn worksheet_used_bounds(model: &ironcalc::base::Model<'_>, sheet: u32) -> Result<(i32, i32), String> {
    let ws = model.workbook.worksheet(sheet)?;
    let dim = ws.dimension();
    Ok((dim.max_row.max(1), dim.max_column.max(1)))
}

fn normalize_cell_range(
    r1: i32,
    c1: i32,
    r2: i32,
    c2: i32,
) -> Result<(i32, i32, i32, i32), String> {
    let top = r1.min(r2);
    let bottom = r1.max(r2);
    let left = c1.min(c2);
    let right = c1.max(c2);
    if top < 1 || left < 1 || bottom > LAST_ROW || right > LAST_COLUMN {
        return Err("Invalid cell range".to_string());
    }
    Ok((top, left, bottom, right))
}

fn format_cell_range(r1: i32, c1: i32, r2: i32, c2: i32) -> String {
    let start = format!("{}{}", col_letter(c1 as u32), r1);
    let end = format!("{}{}", col_letter(c2 as u32), r2);
    if start == end {
        start
    } else {
        format!("{start}:{end}")
    }
}

fn parse_merge_range(range: &str) -> Option<(i32, i32, i32, i32)> {
    let mut parts = range.split(':');
    let first = parts.next()?.trim();
    let second = parts.next().unwrap_or(first).trim();
    if parts.next().is_some() {
        return None;
    }
    let (r1, c1) = parse_a1_addr(first)?;
    let (r2, c2) = parse_a1_addr(second)?;
    normalize_cell_range(r1 as i32, c1 as i32, r2 as i32, c2 as i32).ok()
}

fn ranges_overlap(a: (i32, i32, i32, i32), b: (i32, i32, i32, i32)) -> bool {
    let (ar1, ac1, ar2, ac2) = a;
    let (br1, bc1, br2, bc2) = b;
    ar1 <= br2 && ar2 >= br1 && ac1 <= bc2 && ac2 >= bc1
}

/// Merge the selected rectangle. Existing overlapping merges are rejected
/// rather than rewritten implicitly, which avoids silently destroying a
/// separate layout decision.
#[tauri::command]
pub(crate) fn merge_cells(
    sheet: u32,
    r1: i32,
    c1: i32,
    r2: i32,
    c2: i32,
    state: State<'_, AppState>,
) -> Result<(), String> {
    let range = normalize_cell_range(r1, c1, r2, c2)?;
    let (top, left, bottom, right) = range;
    if top == bottom && left == right {
        return Err("Select at least two cells to merge".to_string());
    }
    let range_label = format_cell_range(top, left, bottom, right);
    let mut guard = state.model.lock().unwrap();
    let model = guard.as_mut().ok_or("no workbook open")?;
    let ws = model
        .workbook
        .worksheets
        .get_mut(sheet as usize)
        .ok_or("bad sheet index")?;
    for existing in &ws.merge_cells {
        if let Some(existing_range) = parse_merge_range(existing) {
            if existing_range == range {
                return Ok(());
            }
            if ranges_overlap(existing_range, range) {
                return Err(format!("Merge overlaps existing merged range {existing}"));
            }
        }
    }
    ws.merge_cells.push(range_label);
    drop(guard);
    *state.structural_dirty.lock().unwrap() = true;
    *state.workbook_dirty.lock().unwrap() = true;
    Ok(())
}

/// Unmerge any merged ranges that overlap the selected rectangle.
#[tauri::command]
pub(crate) fn unmerge_cells(
    sheet: u32,
    r1: i32,
    c1: i32,
    r2: i32,
    c2: i32,
    state: State<'_, AppState>,
) -> Result<usize, String> {
    let range = normalize_cell_range(r1, c1, r2, c2)?;
    let mut guard = state.model.lock().unwrap();
    let model = guard.as_mut().ok_or("no workbook open")?;
    let ws = model
        .workbook
        .worksheets
        .get_mut(sheet as usize)
        .ok_or("bad sheet index")?;
    let before = ws.merge_cells.len();
    ws.merge_cells.retain(|existing| {
        parse_merge_range(existing).map_or(true, |r| !ranges_overlap(r, range))
    });
    let removed = before - ws.merge_cells.len();
    drop(guard);
    if removed > 0 {
        *state.structural_dirty.lock().unwrap() = true;
        *state.workbook_dirty.lock().unwrap() = true;
    }
    Ok(removed)
}

/// Move one cell's content + style from (sr, sc) to (tr, tc), going
/// through `set_user_input` so any formula gets re-parsed at the new
/// anchor. Without this round-trip a relative ref like `=A5+1` at B5
/// would still mean "one column left" after moving to C5 — i.e. now
/// pointing at B5 (empty) instead of A5. Mirrors the IronCalc-internal
/// `move_cell` (which is private) so we don't have to widen the
/// vendor patch beyond `displace_cells` / `shift_cell_formula`.
fn move_cell_via_input(
    model: &mut ironcalc::base::Model<'_>,
    sheet: u32,
    sr: i32,
    sc: i32,
    tr: i32,
    tc: i32,
) -> Result<(), String> {
    // get_localized_cell_content returns "=FORMULA" for formula cells,
    // the value text otherwise, "" for missing cells. That's exactly
    // the input string set_user_input wants — re-parsing it at the
    // target anchor rebases relative refs correctly.
    let input = model.get_localized_cell_content(sheet, sr, sc)?;
    if input.is_empty() {
        // Source is empty. Clear the target so a previous iteration's
        // write doesn't leak through.
        model.cell_clear_contents(sheet, tr, tc)?;
        return Ok(());
    }
    let style_idx = model.workbook.worksheet(sheet)?.get_style(sr, sc);
    model.set_user_input(sheet, tr, tc, input)?;
    model
        .workbook
        .worksheet_mut(sheet)?
        .set_cell_style(tr, tc, style_idx)?;
    model.cell_clear_all(sheet, sr, sc)?;
    Ok(())
}

/// Insert blank cells and shift the affected row segment right.
/// After moving the cells, `displace_cells` walks every formula in
/// the workbook and rewrites refs into the shifted region — that's
/// the bit the previous `set_raw_cell`-only implementation skipped,
/// silently leaving downstream `=B5` formulas pointing at whatever
/// landed at B5 instead of where B5's content actually went.
#[tauri::command]
pub(crate) fn insert_cells_shift_right(
    sheet: u32,
    r1: i32,
    c1: i32,
    r2: i32,
    c2: i32,
    state: State<'_, AppState>,
) -> Result<(), String> {
    use ironcalc::base::expressions::parser::stringify::DisplaceData;
    if r1 < 1 || c1 < 1 || r2 < r1 || c2 < c1 || r2 > LAST_ROW || c2 > LAST_COLUMN {
        return Err("Invalid cell range".to_string());
    }
    let width = c2 - c1 + 1;
    let mut guard = state.model.lock().unwrap();
    let model = guard.as_mut().ok_or("no workbook open")?;
    let (_, max_col) = worksheet_used_bounds(model, sheet)?;
    if max_col + width > LAST_COLUMN {
        return Err("Cannot shift cells beyond the last column".to_string());
    }
    // Right-to-left so each move reads source data before any
    // earlier-iteration write could overwrite it.
    for r in r1..=r2 {
        for c in (c1..=max_col).rev() {
            move_cell_via_input(model, sheet, r, c, r, c + width)?;
        }
    }
    for r in r1..=r2 {
        model
            .displace_cells(&DisplaceData::CellHorizontal {
                sheet,
                row: r,
                column: c1,
                delta: width,
            })
            .map_err(|e| format!("displace cells: {e}"))?;
    }
    if *state.auto_recalc.lock().unwrap() {
        model.evaluate();
    }
    drop(guard);
    *state.structural_dirty.lock().unwrap() = true;
    *state.workbook_dirty.lock().unwrap() = true;
    Ok(())
}

/// Insert blank cells and shift the affected column segment down.
#[tauri::command]
pub(crate) fn insert_cells_shift_down(
    sheet: u32,
    r1: i32,
    c1: i32,
    r2: i32,
    c2: i32,
    state: State<'_, AppState>,
) -> Result<(), String> {
    use ironcalc::base::expressions::parser::stringify::DisplaceData;
    if r1 < 1 || c1 < 1 || r2 < r1 || c2 < c1 || r2 > LAST_ROW || c2 > LAST_COLUMN {
        return Err("Invalid cell range".to_string());
    }
    let height = r2 - r1 + 1;
    let mut guard = state.model.lock().unwrap();
    let model = guard.as_mut().ok_or("no workbook open")?;
    let (max_row, _) = worksheet_used_bounds(model, sheet)?;
    if max_row + height > LAST_ROW {
        return Err("Cannot shift cells beyond the last row".to_string());
    }
    for c in c1..=c2 {
        for r in (r1..=max_row).rev() {
            move_cell_via_input(model, sheet, r, c, r + height, c)?;
        }
    }
    for c in c1..=c2 {
        model
            .displace_cells(&DisplaceData::CellVertical {
                sheet,
                row: r1,
                column: c,
                delta: height,
            })
            .map_err(|e| format!("displace cells: {e}"))?;
    }
    if *state.auto_recalc.lock().unwrap() {
        model.evaluate();
    }
    drop(guard);
    *state.structural_dirty.lock().unwrap() = true;
    *state.workbook_dirty.lock().unwrap() = true;
    Ok(())
}

/// Delete cells and shift the affected row segment left. Refs that
/// pointed at deleted cells become `#REF!`; refs past the deletion
/// shift left by `width`. Both come from `displace_cells` with a
/// negative delta — that's the part the prior naive impl skipped.
#[tauri::command]
pub(crate) fn delete_cells_shift_left(
    sheet: u32,
    r1: i32,
    c1: i32,
    r2: i32,
    c2: i32,
    state: State<'_, AppState>,
) -> Result<(), String> {
    use ironcalc::base::expressions::parser::stringify::DisplaceData;
    if r1 < 1 || c1 < 1 || r2 < r1 || c2 < c1 || r2 > LAST_ROW || c2 > LAST_COLUMN {
        return Err("Invalid cell range".to_string());
    }
    let width = c2 - c1 + 1;
    let mut guard = state.model.lock().unwrap();
    let model = guard.as_mut().ok_or("no workbook open")?;
    let (_, max_col) = worksheet_used_bounds(model, sheet)?;
    // Clear the deleted cells first; left-to-right move from c1+width.
    for r in r1..=r2 {
        for c in c1..=c2 {
            model.cell_clear_all(sheet, r, c)?;
        }
        for c in c1..=max_col {
            let src = c + width;
            if src > LAST_COLUMN {
                model.cell_clear_contents(sheet, r, c)?;
            } else {
                move_cell_via_input(model, sheet, r, src, r, c)?;
            }
        }
    }
    for r in r1..=r2 {
        model
            .displace_cells(&DisplaceData::CellHorizontal {
                sheet,
                row: r,
                column: c1,
                delta: -width,
            })
            .map_err(|e| format!("displace cells: {e}"))?;
    }
    if *state.auto_recalc.lock().unwrap() {
        model.evaluate();
    }
    drop(guard);
    *state.structural_dirty.lock().unwrap() = true;
    *state.workbook_dirty.lock().unwrap() = true;
    Ok(())
}

/// Delete cells and shift the affected column segment up.
#[tauri::command]
pub(crate) fn delete_cells_shift_up(
    sheet: u32,
    r1: i32,
    c1: i32,
    r2: i32,
    c2: i32,
    state: State<'_, AppState>,
) -> Result<(), String> {
    use ironcalc::base::expressions::parser::stringify::DisplaceData;
    if r1 < 1 || c1 < 1 || r2 < r1 || c2 < c1 || r2 > LAST_ROW || c2 > LAST_COLUMN {
        return Err("Invalid cell range".to_string());
    }
    let height = r2 - r1 + 1;
    let mut guard = state.model.lock().unwrap();
    let model = guard.as_mut().ok_or("no workbook open")?;
    let (max_row, _) = worksheet_used_bounds(model, sheet)?;
    for c in c1..=c2 {
        for r in r1..=r2 {
            model.cell_clear_all(sheet, r, c)?;
        }
        for r in r1..=max_row {
            let src = r + height;
            if src > LAST_ROW {
                model.cell_clear_contents(sheet, r, c)?;
            } else {
                move_cell_via_input(model, sheet, src, c, r, c)?;
            }
        }
    }
    for c in c1..=c2 {
        model
            .displace_cells(&DisplaceData::CellVertical {
                sheet,
                row: r1,
                column: c,
                delta: -height,
            })
            .map_err(|e| format!("displace cells: {e}"))?;
    }
    if *state.auto_recalc.lock().unwrap() {
        model.evaluate();
    }
    drop(guard);
    *state.structural_dirty.lock().unwrap() = true;
    *state.workbook_dirty.lock().unwrap() = true;
    Ok(())
}

/// Set the displayed width of a column. `px` is the display pixel value
/// the user dragged to; we reverse the colWidthPx scaling factor (7/12)
/// to get IronCalc's internal "char × 12" unit.
#[tauri::command]
pub(crate) fn set_column_width(
    sheet: u32,
    col: i32,
    px: f64,
    state: State<'_, AppState>,
) -> Result<(), String> {
    let mut guard = state.model.lock().unwrap();
    let model = guard.as_mut().ok_or("no workbook open")?;
    let internal = (px * 12.0 / 7.0).max(0.0);
    model.set_column_width(sheet, col, internal)?;
    drop(guard);
    *state.workbook_dirty.lock().unwrap() = true;
    Ok(())
}

/// Set the displayed height of a row. `px` is the display pixel value;
/// reverse the rowHeightPx scaling (96/72 / 2) to IronCalc's "pt × 2".
#[tauri::command]
pub(crate) fn set_row_height(
    sheet: u32,
    row: i32,
    px: f64,
    state: State<'_, AppState>,
) -> Result<(), String> {
    let mut guard = state.model.lock().unwrap();
    let model = guard.as_mut().ok_or("no workbook open")?;
    let internal = (px * (72.0 / 96.0) * 2.0).max(0.0);
    model.set_row_height(sheet, row, internal)?;
    drop(guard);
    *state.workbook_dirty.lock().unwrap() = true;
    Ok(())
}

/// Toggle hidden state on a row. Persists into the IronCalc model only
/// (not back into the saved xlsx — that's a follow-up); next refresh
/// will see the new state via get_layout.
#[tauri::command]
pub(crate) fn set_row_hidden(
    sheet: u32,
    row: i32,
    hidden: bool,
    state: State<'_, AppState>,
) -> Result<(), String> {
    let mut guard = state.model.lock().unwrap();
    let model = guard.as_mut().ok_or("no workbook open")?;
    model.set_row_hidden(sheet, row, hidden)?;
    drop(guard);
    *state.workbook_dirty.lock().unwrap() = true;
    Ok(())
}

/// Toggle hidden state on a column. The state is shadowed in
/// AppState::hidden_cols since IronCalc's Col struct lacks a hidden field.
#[tauri::command]
pub(crate) fn set_column_hidden(
    sheet: u32,
    col: i32,
    hidden: bool,
    state: State<'_, AppState>,
) -> Result<(), String> {
    let mut hc = state.hidden_cols.lock().unwrap();
    let entry = hc.entry(sheet).or_default();
    if hidden {
        entry.insert(col);
    } else {
        entry.remove(&col);
    }
    drop(hc);
    *state.workbook_dirty.lock().unwrap() = true;
    Ok(())
}

/// Set the frozen pane counts for a sheet. Either argument can be 0
/// (no freeze in that direction). Mirrors Lotus /Worksheet/Titles
/// (Both / Horizontal / Vertical / Clear).
#[tauri::command]
pub(crate) fn set_frozen_panes(
    sheet: u32,
    rows: i32,
    cols: i32,
    state: State<'_, AppState>,
) -> Result<(), String> {
    let mut guard = state.model.lock().unwrap();
    let model = guard.as_mut().ok_or("no workbook open")?;
    let ws = model
        .workbook
        .worksheets
        .get_mut(sheet as usize)
        .ok_or("bad sheet index")?;
    ws.frozen_rows = rows.max(0);
    ws.frozen_columns = cols.max(0);
    drop(guard);
    *state.workbook_dirty.lock().unwrap() = true;
    Ok(())
}

/// Unhide every hidden row in the active sheet — Lotus
/// /Worksheet/Row/Display.
#[tauri::command]
pub(crate) fn show_all_rows(sheet: u32, state: State<'_, AppState>) -> Result<usize, String> {
    let mut guard = state.model.lock().unwrap();
    let model = guard.as_mut().ok_or("no workbook open")?;
    let ws = model
        .workbook
        .worksheets
        .get_mut(sheet as usize)
        .ok_or("bad sheet index")?;
    let mut n = 0;
    for r in ws.rows.iter_mut() {
        if r.hidden {
            r.hidden = false;
            n += 1;
        }
    }
    drop(guard);
    if n > 0 {
        *state.workbook_dirty.lock().unwrap() = true;
    }
    Ok(n)
}

/// Unhide every hidden column in the active sheet — Lotus
/// /Worksheet/Column/Display.
#[tauri::command]
pub(crate) fn show_all_cols(sheet: u32, state: State<'_, AppState>) -> Result<usize, String> {
    let mut hc = state.hidden_cols.lock().unwrap();
    let n = hc.get(&sheet).map(|s| s.len()).unwrap_or(0);
    if let Some(set) = hc.get_mut(&sheet) {
        set.clear();
    }
    drop(hc);
    if n > 0 {
        *state.workbook_dirty.lock().unwrap() = true;
    }
    Ok(n)
}

/// True (un-clamped) used range of the sheet — the bottom-right of
/// worksheet.dimension, no min/max bounds applied. Used by Ctrl+End.
#[tauri::command]
pub(crate) fn get_used_range(
    sheet: u32,
    state: State<'_, AppState>,
) -> Result<(u32, u32), String> {
    let guard = state.model.lock().unwrap();
    let model = guard.as_ref().ok_or("no workbook open")?;
    let ws = model
        .workbook
        .worksheets
        .get(sheet as usize)
        .ok_or("bad sheet index")?;
    let (rows, cols) = parse_dimension(&ws.dimension).unwrap_or((1, 1));
    Ok((rows, cols))
}

/// Used range of a sheet in (rows, cols), parsed from the worksheet's
/// `dimension` attribute (e.g. "A1:CK127"). MIN gives empty sheets a
/// usable workspace; MAX caps pathological dimensions to Excel's own
/// limits — both axes are virtualised in the frontend (only the visible
/// band of rows × the visible band of cols hits the DOM) so huge
/// dimensions don't translate into upfront render work.
#[tauri::command]
pub(crate) fn get_sheet_dim(
    sheet: u32,
    state: State<'_, AppState>,
) -> Result<(u32, u32), String> {
    const MIN_ROWS: u32 = 100;
    const MIN_COLS: u32 = 60;
    const MAX_ROWS: u32 = 1_048_576;
    const MAX_COLS: u32 = 16_384;
    let guard = state.model.lock().unwrap();
    let model = guard.as_ref().ok_or("no workbook open")?;
    let ws = model
        .workbook
        .worksheets
        .get(sheet as usize)
        .ok_or("bad sheet index")?;
    let (rows, cols) = parse_dimension(&ws.dimension).unwrap_or((MIN_ROWS, MIN_COLS));
    Ok((rows.clamp(MIN_ROWS, MAX_ROWS), cols.clamp(MIN_COLS, MAX_COLS)))
}

/// Parse "A1:CK127" / "A1" / "$AB$45:$CK$127" → (max_row, max_col).
/// Tolerant: returns None on malformed input so the caller can fall back.
fn parse_dimension(s: &str) -> Option<(u32, u32)> {
    let mut max_r = 0u32;
    let mut max_c = 0u32;
    for part in s.split(':') {
        let (r, c) = parse_a1_addr(part.trim())?;
        max_r = max_r.max(r);
        max_c = max_c.max(c);
    }
    if max_r > 0 && max_c > 0 {
        Some((max_r, max_c))
    } else {
        None
    }
}

fn parse_a1_addr(s: &str) -> Option<(u32, u32)> {
    let s = s.replace('$', "");
    let mut col = 0u32;
    let mut row_start = None;
    for (i, c) in s.char_indices() {
        if c.is_ascii_alphabetic() {
            col = col * 26 + (c.to_ascii_uppercase() as u32 - 'A' as u32 + 1);
        } else {
            row_start = Some(i);
            break;
        }
    }
    let row_start = row_start?;
    let row: u32 = s[row_start..].parse().ok()?;
    if col == 0 || row == 0 {
        return None;
    }
    Some((row, col))
}

/// Excel-style Ctrl+Arrow jump.
///
/// `dr`/`dc` are -1, 0, or +1 indicating the direction (exactly one is
/// non-zero in practice). Semantics match Excel:
///   * If the current cell is empty: skip empties until the first non-empty
///     in that direction; if none found, stop at the last visited cell.
///   * If the current cell is populated AND the next cell is populated:
///     stop at the LAST populated cell in the contiguous run.
///   * If the current cell is populated AND the next cell is empty: skip
///     empties until the next non-empty; if none, stop at the boundary.
///
/// Bounded by `max_step` cells walked to avoid pathological scans on
/// near-empty sheets — Excel itself caps at 1048576 rows / 16384 cols.
#[tauri::command]
pub(crate) fn jump_edge(
    sheet: u32,
    row: u32,
    col: u32,
    dr: i32,
    dc: i32,
    state: State<'_, AppState>,
) -> Result<(u32, u32), String> {
    let guard = state.model.lock().unwrap();
    let model = guard.as_ref().ok_or("no workbook open")?;
    if (dr == 0 && dc == 0) || dr.abs() > 1 || dc.abs() > 1 {
        return Ok((row, col));
    }
    let max_row: i32 = 1_048_576;
    let max_col: i32 = 16_384;
    let max_step: i32 = 16_384;

    let is_empty = |r: i32, c: i32| -> bool {
        if r < 1 || c < 1 || r > max_row || c > max_col {
            return true;
        }
        model
            .get_formatted_cell_value(sheet, r, c)
            .map(|s| s.is_empty())
            .unwrap_or(true)
    };

    let in_bounds = |r: i32, c: i32| -> bool { r >= 1 && c >= 1 && r <= max_row && c <= max_col };

    let r0 = row as i32;
    let c0 = col as i32;
    let started_empty = is_empty(r0, c0);
    let next_r = r0 + dr;
    let next_c = c0 + dc;
    if !in_bounds(next_r, next_c) {
        return Ok((row, col));
    }

    let (mut r, mut c) = (r0, c0);

    if started_empty {
        let (mut nr, mut nc) = (r0 + dr, c0 + dc);
        let mut steps = 0;
        while in_bounds(nr, nc) && is_empty(nr, nc) && steps < max_step {
            r = nr;
            c = nc;
            nr += dr;
            nc += dc;
            steps += 1;
        }
        if in_bounds(nr, nc) && !is_empty(nr, nc) {
            return Ok((nr as u32, nc as u32));
        }
        return Ok((r as u32, c as u32));
    }

    let next_is_empty = is_empty(next_r, next_c);

    if next_is_empty {
        let (mut nr, mut nc) = (next_r + dr, next_c + dc);
        let mut steps = 0;
        while in_bounds(nr, nc) && is_empty(nr, nc) && steps < max_step {
            nr += dr;
            nc += dc;
            steps += 1;
        }
        if in_bounds(nr, nc) && !is_empty(nr, nc) {
            return Ok((nr as u32, nc as u32));
        }
        return Ok(((nr - dr) as u32, (nc - dc) as u32));
    }

    let mut steps = 0;
    while in_bounds(r + dr, c + dc) && !is_empty(r + dr, c + dc) && steps < max_step {
        r += dr;
        c += dc;
        steps += 1;
    }
    Ok((r as u32, c as u32))
}
