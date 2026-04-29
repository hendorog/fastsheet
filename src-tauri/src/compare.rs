//! Workbook comparison.
//!
//! Diffs the active workbook against a "right-side" workbook loaded
//! from disk. Sheet matching is by name; for matched sheets we walk
//! the union of populated cells and report any cell whose displayed
//! value or formula text differs. Sheets present on only one side
//! become a header entry in the diff list — the user sees the
//! asymmetry rather than silently dropping it.
//!
//! "Equal as formatted" semantics: we compare via
//! `get_formatted_cell_value` on both sides. This is what the user
//! sees in the grid — small float jitter that gets rounded away by
//! number-format display doesn't show up as a diff. Formula-text
//! diffs are reported even when the formatted values match (a cell
//! that says `=A1+B1` vs `=A1+B2` and happens to evaluate the same
//! is still informative).

use std::collections::{BTreeSet, HashMap};

use ironcalc::base::types::Cell;
use ironcalc::base::Model;
use serde::Serialize;

use crate::util::col_letter;

/// One reported difference. The frontend renders these in a docked
/// list; clicking jumps the cursor.
#[derive(Serialize, Debug, Clone)]
pub struct CompareDiff {
    pub sheet: String,
    /// Sheet index in the LEFT (active) model, when the sheet exists
    /// there. None for sheets that exist only on the right.
    pub sheet_idx: Option<u32>,
    pub row: i32,
    pub col: i32,
    /// "A1"-style address for display in the panel.
    pub address: String,
    /// Formatted display value on each side. Empty string when the
    /// cell is missing on that side.
    pub left_value: String,
    pub right_value: String,
    /// Formula text (with leading `=`) when the cell carried a
    /// formula on that side. None for value cells or empty cells.
    pub left_formula: Option<String>,
    pub right_formula: Option<String>,
    /// "value" — formatted value differs.
    /// "formula" — formatted value matches but formula text differs.
    /// "missing-left" / "missing-right" — cell present on only one side.
    pub kind: &'static str,
}

/// Sheet-level summary entries placed at the top of the diff list
/// when sheets exist on only one side.
#[derive(Serialize, Debug, Clone)]
pub struct CompareSheetMissing {
    pub sheet: String,
    pub side: &'static str,
}

/// Result returned to the frontend on compare_open.
#[derive(Serialize, Debug, Clone)]
pub struct CompareResult {
    pub right_path: String,
    pub diffs: Vec<CompareDiff>,
    pub missing_sheets: Vec<CompareSheetMissing>,
    /// Counts so the UI can show "237 cell diffs across 4 sheets"
    /// without re-counting. Capped count = `diffs.len()`; total =
    /// the actual number found (may exceed the cap).
    pub total_diffs: usize,
    pub diffs_capped: bool,
}

/// Hard cap so a wildly different file doesn't OOM the UI. Above this
/// the result truncates and `diffs_capped: true` is returned.
const MAX_DIFFS: usize = 5000;

/// In-memory session held in AppState for the duration of compare
/// mode. The right model is only used to look up
/// `get_formatted_cell_value` / formula text during trace; we don't
/// re-diff on cell edits (the active model is what's mutating, the
/// right side is read-only context).
pub struct CompareSession {
    pub right_path: String,
    pub right_model: Model<'static>,
    /// Pre-computed sheet-name → right_sheet_idx map (since the right
    /// model can have a different sheet ordering than the left). Used
    /// by `right_value_at` to project a left address onto the right.
    pub right_sheet_by_name: HashMap<String, u32>,
}

impl CompareSession {
    /// Look up the right-side formatted value at an address given by
    /// left-side coordinates. Returns None when the sheet name has no
    /// match on the right.
    pub fn right_value_at(
        &self,
        left_model: &Model,
        sheet_idx: u32,
        row: i32,
        col: i32,
    ) -> Option<String> {
        let name = left_model
            .workbook
            .worksheets
            .get(sheet_idx as usize)?
            .name
            .clone();
        let right_idx = *self.right_sheet_by_name.get(&name)?;
        self.right_model
            .get_formatted_cell_value(right_idx, row, col)
            .ok()
    }

    /// Right-side formula text (with leading `=`) at left-coords, or None.
    pub fn right_formula_at(
        &self,
        left_model: &Model,
        sheet_idx: u32,
        row: i32,
        col: i32,
    ) -> Option<String> {
        let name = left_model
            .workbook
            .worksheets
            .get(sheet_idx as usize)?
            .name
            .clone();
        let right_idx = *self.right_sheet_by_name.get(&name)?;
        self.right_model
            .get_cell_formula(right_idx, row, col)
            .ok()
            .flatten()
    }
}

/// Build a CompareSession from a freshly-loaded right model and run
/// the diff against the active (left) model. The right model is moved
/// into the session.
pub fn diff_workbooks(
    left: &Model,
    right: Model<'static>,
    right_path: String,
) -> (CompareSession, CompareResult) {
    let left_sheet_by_name: HashMap<String, u32> = left
        .workbook
        .worksheets
        .iter()
        .enumerate()
        .map(|(i, w)| (w.name.clone(), i as u32))
        .collect();
    let right_sheet_by_name: HashMap<String, u32> = right
        .workbook
        .worksheets
        .iter()
        .enumerate()
        .map(|(i, w)| (w.name.clone(), i as u32))
        .collect();

    let mut missing_sheets = Vec::new();
    for name in left_sheet_by_name.keys() {
        if !right_sheet_by_name.contains_key(name) {
            missing_sheets.push(CompareSheetMissing {
                sheet: name.clone(),
                side: "right",
            });
        }
    }
    for name in right_sheet_by_name.keys() {
        if !left_sheet_by_name.contains_key(name) {
            missing_sheets.push(CompareSheetMissing {
                sheet: name.clone(),
                side: "left",
            });
        }
    }
    missing_sheets.sort_by(|a, b| a.sheet.to_lowercase().cmp(&b.sheet.to_lowercase()));

    let mut diffs = Vec::new();
    let mut total_diffs = 0usize;
    // Walk shared sheets in left-side order so the panel ordering
    // tracks what the user sees in the tab bar.
    for (left_idx, ws) in left.workbook.worksheets.iter().enumerate() {
        let Some(&right_idx) = right_sheet_by_name.get(&ws.name) else {
            continue;
        };
        diff_sheet(
            left,
            &right,
            left_idx as u32,
            right_idx,
            &ws.name,
            &mut diffs,
            &mut total_diffs,
        );
    }

    let capped = total_diffs > diffs.len();
    let result = CompareResult {
        right_path: right_path.clone(),
        diffs,
        missing_sheets,
        total_diffs,
        diffs_capped: capped,
    };
    let session = CompareSession {
        right_path,
        right_model: right,
        right_sheet_by_name,
    };
    (session, result)
}

fn diff_sheet(
    left: &Model,
    right: &Model,
    left_idx: u32,
    right_idx: u32,
    sheet_name: &str,
    out: &mut Vec<CompareDiff>,
    total: &mut usize,
) {
    // Union of populated (row, col) on both sides — BTreeSet so we
    // emit diffs in row-major order without an extra sort pass.
    let mut keys: BTreeSet<(i32, i32)> = BTreeSet::new();
    if let Some(ws) = left.workbook.worksheets.get(left_idx as usize) {
        for (row, cols) in &ws.sheet_data {
            for col in cols.keys() {
                keys.insert((*row, *col));
            }
        }
    }
    if let Some(ws) = right.workbook.worksheets.get(right_idx as usize) {
        for (row, cols) in &ws.sheet_data {
            for col in cols.keys() {
                keys.insert((*row, *col));
            }
        }
    }

    for (row, col) in keys {
        let left_present = cell_present(left, left_idx, row, col);
        let right_present = cell_present(right, right_idx, row, col);
        let left_value = if left_present {
            left.get_formatted_cell_value(left_idx, row, col).unwrap_or_default()
        } else {
            String::new()
        };
        let right_value = if right_present {
            right
                .get_formatted_cell_value(right_idx, row, col)
                .unwrap_or_default()
        } else {
            String::new()
        };
        let left_formula = left.get_cell_formula(left_idx, row, col).ok().flatten();
        let right_formula = right.get_cell_formula(right_idx, row, col).ok().flatten();

        // "Empty on one side" is treated as a value-of-empty match
        // when the other side is also empty/blank. Cells that are
        // populated as empty strings on both sides are equal.
        let values_equal = left_value == right_value;
        let formulas_equal = left_formula == right_formula;

        let kind: &'static str = if values_equal && formulas_equal {
            continue;
        } else if !left_present && right_present {
            "missing-left"
        } else if left_present && !right_present {
            "missing-right"
        } else if !values_equal {
            "value"
        } else {
            // Values match but formulas differ — informative even
            // though the displayed answer is the same.
            "formula"
        };

        *total += 1;
        if out.len() >= MAX_DIFFS {
            // Keep counting so the UI can show "X+ diffs (capped)"
            // but stop pushing.
            continue;
        }

        out.push(CompareDiff {
            sheet: sheet_name.to_string(),
            sheet_idx: Some(left_idx),
            row,
            col,
            address: format!("{}{}", col_letter(col as u32), row),
            left_value,
            right_value,
            left_formula,
            right_formula,
            kind,
        });
    }
}

/// True iff the cell at (sheet, row, col) is present in the model's
/// sheet_data (any variant — including EmptyCell with style).
fn cell_present(model: &Model, sheet: u32, row: i32, col: i32) -> bool {
    let Some(ws) = model.workbook.worksheets.get(sheet as usize) else {
        return false;
    };
    let Some(cols) = ws.sheet_data.get(&row) else {
        return false;
    };
    match cols.get(&col) {
        // Treat EmptyCell as "not really present for compare" — these
        // exist only because of style indices on otherwise blank
        // cells, and reporting style-only diffs isn't in scope.
        Some(Cell::EmptyCell { .. }) | None => false,
        Some(_) => true,
    }
}
