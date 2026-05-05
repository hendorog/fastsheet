//! Formula dependency tracer.
//!
//! Walks the parsed Node tree of a cell's formula, collects every
//! reference / range / defined-name it touches, and recurses into
//! cells that themselves have formulas. The result is a serializable
//! tree the frontend renders as a dependency-chain popup so the user
//! can troubleshoot why a formula evaluates to what it does without
//! firing up Excel and chasing named ranges by hand.
//!
//! Walk semantics:
//!   * Cell references resolve to (sheet, row, col) at the absolute
//!     position. IronCalc stores relative refs as offsets from the
//!     formula's anchor cell, so we convert here.
//!   * Range references are reported as one node showing the range
//!     bounds and a value summary; we don't expand every cell of a
//!     2000-cell lookup table.
//!   * Defined names are reported with their declared formula text
//!     (the resolved range), not recursed further unless the user
//!     drills in.
//!   * Cycle detection: a `(sheet, row, col)` set tracks ancestors;
//!     when we'd recurse into a cell already on the path we mark
//!     `cycle: true` and stop. This catches both genuine circular
//!     refs and cases where two formulas alternate.
//!   * Depth cap: hard-coded at 8 by default. Beyond that the tree
//!     gets too dense to read in the popup; user can re-trace from a
//!     deeper cell to see further.

use ironcalc::base::expressions::parser::Node;
use ironcalc::base::types::Cell;
use ironcalc::base::Model;
use serde::Serialize;
use std::collections::HashSet;

use crate::util::col_letter;

/// One node in the dependency tree. The frontend renders this with
/// indentation; `deps` are the children.
#[derive(Serialize, Debug, Clone)]
pub struct TraceNode {
    /// Display address — "Discount!G4" for cells, "Discount!B24:W35"
    /// for ranges, the defined-name string for defined names.
    pub address: String,
    /// "cell" | "range" | "name" | "literal" | "error"
    pub kind: &'static str,
    /// Sheet index for "cell" kind so the frontend can jump there.
    /// None for ranges + names (could resolve to either; frontend
    /// uses `address` as the goto target instead).
    pub sheet: Option<u32>,
    pub row: Option<i32>,
    pub col: Option<i32>,
    /// Formula text for cell kind (when the cell carries a formula).
    /// Pre-formatted: leading "=" included so it's display-ready.
    pub formula: Option<String>,
    /// The cell's evaluated/displayed value, or a value-summary for
    /// ranges/names.
    pub value: String,
    /// Optional explanatory note — e.g. "→ Discount!$B$24:$W$35" for
    /// a defined name; "12 cells, e.g. \"foo\", \"bar\", \"baz\"" for
    /// a range; "circular reference" for cycle hits.
    pub note: Option<String>,
    /// True iff the cell value is an error (#N/A, #VALUE!, etc.).
    pub is_error: bool,
    /// True if recursion stopped here because of cycle detection.
    pub cycle: bool,
    /// True if recursion stopped here because the depth cap was hit.
    pub truncated: bool,
    /// Right-side formatted value, when a compare session is active.
    /// None outside compare mode; Some("") for cells that exist in
    /// the left model but are empty/missing on the right.
    pub compare_value: Option<String>,
    /// True iff `compare_value` differs from `value`. Lets the
    /// frontend tint trace nodes whose right-side disagrees without
    /// re-doing the comparison client-side.
    pub compare_differs: bool,
    pub deps: Vec<TraceNode>,
}

/// Top-level entry: trace the formula at (sheet, row, col). Returns a
/// rooted tree. Cells without formulas come back as a single node
/// with `kind: "literal"` and no deps.
///
/// `compare`, when present, populates each cell-kind node's
/// `compare_value` with the right-side formatted value so the
/// frontend can render `left | right` next to deps. Range / name /
/// literal nodes don't carry a compare value (the right-side address
/// resolution would be ambiguous).
pub fn trace(
    model: &Model,
    sheet: u32,
    row: i32,
    col: i32,
    compare: Option<&crate::compare::CompareSession>,
) -> TraceNode {
    let mut visited = HashSet::new();
    visited.insert((sheet, row, col));
    trace_cell(model, sheet, row, col, &mut visited, 0, MAX_DEPTH, compare)
}

const MAX_DEPTH: u32 = 8;
const MAX_SIBLINGS: usize = 32;

/// Build a TraceNode for one specific cell.
fn trace_cell(
    model: &Model,
    sheet: u32,
    row: i32,
    col: i32,
    visited: &mut HashSet<(u32, i32, i32)>,
    depth: u32,
    max_depth: u32,
    compare: Option<&crate::compare::CompareSession>,
) -> TraceNode {
    let sheet_name = model
        .workbook
        .worksheets
        .get(sheet as usize)
        .map(|w| w.name.clone())
        .unwrap_or_else(|| format!("Sheet{sheet}"));
    let address = format_addr(&sheet_name, row, col);
    let value = model
        .get_formatted_cell_value(sheet, row, col)
        .unwrap_or_default();
    let is_error = is_error_value(&value);

    // Look up the cell's formula (if any) via parsed_formulas.
    let formula_text = model
        .get_cell_formula(sheet, row, col)
        .ok()
        .flatten();
    let parsed = formula_index(model, sheet, row, col)
        .and_then(|f| model.parsed_formulas.get(sheet as usize).and_then(|v| v.get(f)));

    let compare_value = compare.and_then(|c| c.right_value_at(model, sheet, row, col));
    let compare_differs = compare_value
        .as_ref()
        .is_some_and(|cv| cv != &value);
    let mut node = TraceNode {
        address,
        kind: if formula_text.is_some() { "cell" } else { "literal" },
        sheet: Some(sheet),
        row: Some(row),
        col: Some(col),
        formula: formula_text,
        value,
        note: None,
        is_error,
        cycle: false,
        truncated: false,
        compare_value,
        compare_differs,
        deps: Vec::new(),
    };

    if depth >= max_depth {
        node.truncated = true;
        return node;
    }

    // Walk the parsed node tree of THIS cell's formula, collecting
    // every reference / range / defined-name it touches. Each one
    // becomes a child TraceNode.
    if let Some(n) = parsed {
        let mut refs = Vec::new();
        collect_refs(n, sheet, row, col, &mut refs);
        // Dedupe — a formula that mentions the same cell twice
        // shouldn't generate two child nodes.
        refs = dedupe_refs(refs);
        for r in refs.into_iter().take(MAX_SIBLINGS) {
            node.deps.push(build_dep_node(model, &r, visited, depth, max_depth, compare));
        }
    }
    node
}

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
enum CollectedRef {
    Cell { sheet: u32, row: i32, col: i32 },
    Range { sheet: u32, r1: i32, c1: i32, r2: i32, c2: i32 },
    Name { name: String, scope: Option<u32>, formula: String },
}

/// Walk a parsed Node tree and append every reference-like leaf to
/// `out`. `anchor_*` is the formula's host cell — used to convert
/// IronCalc's relative offsets to absolute coords.
fn collect_refs(
    node: &Node,
    anchor_sheet: u32,
    anchor_row: i32,
    anchor_col: i32,
    out: &mut Vec<CollectedRef>,
) {
    match node {
        Node::ReferenceKind {
            sheet_index,
            absolute_row,
            absolute_column,
            row,
            column,
            ..
        } => {
            let r = if *absolute_row { *row } else { anchor_row + *row };
            let c = if *absolute_column { *column } else { anchor_col + *column };
            out.push(CollectedRef::Cell {
                sheet: *sheet_index,
                row: r,
                col: c,
            });
        }
        Node::RangeKind {
            sheet_index,
            absolute_row1, absolute_column1, row1, column1,
            absolute_row2, absolute_column2, row2, column2,
            ..
        } => {
            let r1 = if *absolute_row1 { *row1 } else { anchor_row + *row1 };
            let c1 = if *absolute_column1 { *column1 } else { anchor_col + *column1 };
            let r2 = if *absolute_row2 { *row2 } else { anchor_row + *row2 };
            let c2 = if *absolute_column2 { *column2 } else { anchor_col + *column2 };
            out.push(CollectedRef::Range {
                sheet: *sheet_index,
                r1: r1.min(r2), c1: c1.min(c2),
                r2: r1.max(r2), c2: c1.max(c2),
            });
        }
        Node::DefinedNameKind((name, scope, formula)) => {
            out.push(CollectedRef::Name {
                name: name.clone(),
                scope: *scope,
                formula: formula.clone(),
            });
        }
        Node::FunctionKind { args, .. } | Node::InvalidFunctionKind { args, .. } => {
            for a in args {
                collect_refs(a, anchor_sheet, anchor_row, anchor_col, out);
            }
        }
        Node::OpRangeKind { left, right }
        | Node::OpConcatenateKind { left, right }
        | Node::OpSumKind { left, right, .. }
        | Node::OpProductKind { left, right, .. }
        | Node::OpPowerKind { left, right }
        | Node::CompareKind { left, right, .. } => {
            collect_refs(left, anchor_sheet, anchor_row, anchor_col, out);
            collect_refs(right, anchor_sheet, anchor_row, anchor_col, out);
        }
        Node::UnaryKind { right, .. } | Node::ImplicitIntersection { child: right, .. } => {
            collect_refs(right, anchor_sheet, anchor_row, anchor_col, out);
        }
        // Literals + error variants don't reference anything.
        _ => {}
    }
}

fn dedupe_refs(refs: Vec<CollectedRef>) -> Vec<CollectedRef> {
    let mut seen = HashSet::new();
    let mut out = Vec::with_capacity(refs.len());
    for r in refs {
        if seen.insert(r.clone()) {
            out.push(r);
        }
    }
    out
}

/// Convert a CollectedRef into a TraceNode, recursing for Cell kinds
/// that themselves carry a formula.
fn build_dep_node(
    model: &Model,
    r: &CollectedRef,
    visited: &mut HashSet<(u32, i32, i32)>,
    depth: u32,
    max_depth: u32,
    compare: Option<&crate::compare::CompareSession>,
) -> TraceNode {
    match r {
        CollectedRef::Cell { sheet, row, col } => {
            let key = (*sheet, *row, *col);
            if visited.contains(&key) {
                let sheet_name = sheet_name_of(model, *sheet);
                let address = format_addr(&sheet_name, *row, *col);
                let value = model
                    .get_formatted_cell_value(*sheet, *row, *col)
                    .unwrap_or_default();
                let compare_value = compare.and_then(|c| c.right_value_at(model, *sheet, *row, *col));
                let compare_differs = compare_value.as_ref().is_some_and(|cv| cv != &value);
                return TraceNode {
                    address,
                    kind: "cell",
                    sheet: Some(*sheet),
                    row: Some(*row),
                    col: Some(*col),
                    formula: None,
                    is_error: is_error_value(&value),
                    value,
                    note: Some("(circular — already on path)".into()),
                    cycle: true,
                    truncated: false,
                    compare_value,
                    compare_differs,
                    deps: Vec::new(),
                };
            }
            visited.insert(key);
            let n = trace_cell(model, *sheet, *row, *col, visited, depth + 1, max_depth, compare);
            visited.remove(&key);
            n
        }
        CollectedRef::Range { sheet, r1, c1, r2, c2 } => {
            let sheet_name = sheet_name_of(model, *sheet);
            let address = format!(
                "{}!{}{}:{}{}",
                sheet_name,
                col_letter(*c1 as u32), r1,
                col_letter(*c2 as u32), r2,
            );
            let preview = range_preview(model, *sheet, *r1, *c1, *r2, *c2);
            TraceNode {
                address,
                kind: "range",
                sheet: Some(*sheet),
                row: Some(*r1),
                col: Some(*c1),
                formula: None,
                value: preview.summary,
                note: preview.preview,
                is_error: false,
                cycle: false,
                truncated: false,
                compare_value: None,
                compare_differs: false,
                deps: Vec::new(),
            }
        }
        CollectedRef::Name { name, scope, formula } => {
            let scope_label = match scope {
                Some(idx) => format!(" (sheet-local: {})", sheet_name_of(model, *idx)),
                None => String::new(),
            };
            // Resolve the name's target. A defined name's formula is
            // typically a sheet-qualified A1 ref or range like
            // "Discount!$B$24:$W$35" or "Sheet1!$A$1". Parse it so the
            // value column shows the actual cell value (single-cell
            // names) or a range summary (multi-cell names) instead of
            // the address text. Falls back to the address when the
            // formula isn't a simple A1 ref (e.g. dynamic OFFSET names).
            let target = parse_name_target(model, formula);
            let (value, mut note_extra, sheet, row, col, is_error) = match target {
                Some((sheet_idx, r1, c1, r2, c2)) if r1 == r2 && c1 == c2 => {
                    let v = model
                        .get_formatted_cell_value(sheet_idx, r1, c1)
                        .unwrap_or_default();
                    let err = is_error_value(&v);
                    (v, Some(formula.clone()), Some(sheet_idx), Some(r1), Some(c1), err)
                }
                Some((sheet_idx, r1, c1, r2, c2)) => {
                    let preview = range_preview(model, sheet_idx, r1, c1, r2, c2);
                    let extra = match preview.preview {
                        Some(p) => format!("{} — {}", formula, p),
                        None => formula.clone(),
                    };
                    (preview.summary, Some(extra), None, None, None, false)
                }
                None => (formula.clone(), None, None, None, None, false),
            };
            // Compose the note: defined-name label + scope, then the
            // target address (or preview) on a follow-up line.
            let mut note = format!("defined name{scope_label}");
            if let Some(extra) = note_extra.take() {
                note.push_str(" → ");
                note.push_str(&extra);
            }
            TraceNode {
                address: name.clone(),
                kind: "name",
                sheet,
                row,
                col,
                formula: None,
                value,
                note: Some(note),
                is_error,
                cycle: false,
                truncated: false,
                compare_value: None,
                compare_differs: false,
                deps: Vec::new(),
            }
        }
    }
}

/// Parse a defined-name's formula text into a (sheet_idx, r1, c1, r2,
/// c2) range, 1-based. Handles single-cell ("Sheet1!$A$1"), range
/// ("Sheet1!$A$1:$B$5"), and quoted-sheet ("'My Sheet'!$A$1") forms.
/// Returns None for anything more complex (dynamic OFFSET formulas,
/// multi-area unions, cross-sheet ranges, etc.) — caller falls back to
/// showing the formula text.
fn parse_name_target(model: &Model, formula: &str) -> Option<(u32, i32, i32, i32, i32)> {
    let s = formula.trim_start_matches('=').trim();
    let (sheet_name, rest) = if let Some(rest) = s.strip_prefix('\'') {
        let (name, after) = rest.split_once("'!")?;
        (name.to_string(), after)
    } else {
        let (name, after) = s.split_once('!')?;
        (name.to_string(), after)
    };
    let sheet_idx = model
        .workbook
        .worksheets
        .iter()
        .position(|w| w.name == sheet_name)
        .map(|i| i as u32)?;
    let (start, end) = match rest.split_once(':') {
        Some((a, b)) => (a, b),
        None => (rest, rest),
    };
    let (r1, c1) = parse_a1(start)?;
    let (r2, c2) = parse_a1(end)?;
    Some((sheet_idx, r1.min(r2), c1.min(c2), r1.max(r2), c1.max(c2)))
}

struct RangePreview {
    summary: String,
    preview: Option<String>,
}

fn range_preview(
    model: &Model,
    sheet: u32,
    r1: i32,
    c1: i32,
    r2: i32,
    c2: i32,
) -> RangePreview {
    let total = ((r2 - r1 + 1) * (c2 - c1 + 1)) as usize;
    let mut samples: Vec<String> = Vec::new();
    'outer: for r in r1..=r2 {
        for c in c1..=c2 {
            let v = model.get_formatted_cell_value(sheet, r, c).unwrap_or_default();
            if !v.is_empty() {
                samples.push(v);
                if samples.len() >= 3 {
                    break 'outer;
                }
            }
        }
    }
    let summary = format!("{} cell{}", total, if total == 1 { "" } else { "s" });
    let preview = if samples.is_empty() {
        None
    } else {
        Some(format!("preview: {}", samples.join(", ")))
    };
    RangePreview { summary, preview }
}

// ---------------------------------------------------------------------------
// Named ranges listing — for the /Formula Names menu option.
// ---------------------------------------------------------------------------

#[derive(Serialize, Debug, Clone)]
pub struct NamedRangeInfo {
    pub name: String,
    /// "Discount!$B$24:$W$35" — the resolved formula string IronCalc
    /// has stored. Could also be a multi-range expression for some
    /// names, but we treat it as opaque text.
    pub formula: String,
    /// "(global)" or "(local: <sheet name>)" for sheet-scoped names.
    pub scope: String,
    /// Best-effort first-cell sheet/row/col so the Goto action knows
    /// where to land. None when we can't parse the formula text into
    /// a single A1-style ref.
    pub jump_sheet: Option<u32>,
    pub jump_row: Option<i32>,
    pub jump_col: Option<i32>,
}

pub fn list_named_ranges(model: &Model) -> Vec<NamedRangeInfo> {
    let mut out = Vec::new();
    for n in &model.workbook.defined_names {
        let scope = match n.sheet_id {
            Some(idx) => format!("(local: {})", sheet_name_of(model, idx)),
            None => "(global)".into(),
        };
        let (jump_sheet, jump_row, jump_col) = parse_jump_target(model, &n.formula);
        out.push(NamedRangeInfo {
            name: n.name.clone(),
            formula: n.formula.clone(),
            scope,
            jump_sheet,
            jump_row,
            jump_col,
        });
    }
    out.sort_by(|a, b| a.name.to_lowercase().cmp(&b.name.to_lowercase()));
    out
}

/// Top-left of the parsed range — used by the Names list's Goto
/// action. Returns (None, None, None) when the formula isn't a simple
/// sheet-qualified A1 ref.
fn parse_jump_target(model: &Model, formula: &str) -> (Option<u32>, Option<i32>, Option<i32>) {
    match parse_name_target(model, formula) {
        Some((sheet_idx, r1, c1, _, _)) => (Some(sheet_idx), Some(r1), Some(c1)),
        None => (None, None, None),
    }
}

/// Parse an A1-style cell address (with optional `$` absolute markers)
/// into (row, col), 1-based. Returns None on malformed input.
fn parse_a1(s: &str) -> Option<(i32, i32)> {
    let s = s.trim_start_matches('$');
    let mut col = 0i32;
    let mut bytes = s.bytes().peekable();
    while let Some(&b) = bytes.peek() {
        if b == b'$' {
            bytes.next();
            continue;
        }
        if !b.is_ascii_alphabetic() {
            break;
        }
        col = col * 26 + (b.to_ascii_uppercase() - b'A' + 1) as i32;
        bytes.next();
    }
    if col == 0 {
        return None;
    }
    let row_bytes: Vec<u8> = bytes.collect();
    let row_str = std::str::from_utf8(&row_bytes).ok()?;
    let row_str = row_str.trim_start_matches('$');
    let row: i32 = row_str.parse().ok()?;
    Some((row, col))
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

fn sheet_name_of(model: &Model, sheet: u32) -> String {
    model
        .workbook
        .worksheets
        .get(sheet as usize)
        .map(|w| w.name.clone())
        .unwrap_or_else(|| format!("Sheet{sheet}"))
}

fn format_addr(sheet_name: &str, row: i32, col: i32) -> String {
    let needs_quotes = sheet_name
        .chars()
        .any(|c| !c.is_alphanumeric() && c != '_');
    if needs_quotes {
        format!("'{}'!{}{}", sheet_name, col_letter(col as u32), row)
    } else {
        format!("{}!{}{}", sheet_name, col_letter(col as u32), row)
    }
}

fn is_error_value(v: &str) -> bool {
    v.starts_with('#') && v.ends_with('!') || v == "#N/A"
}

fn formula_index(model: &Model, sheet: u32, row: i32, col: i32) -> Option<usize> {
    let ws = model.workbook.worksheets.get(sheet as usize)?;
    let cell = ws.sheet_data.get(&row)?.get(&col)?;
    match cell {
        Cell::CellFormula { f, .. }
        | Cell::CellFormulaNumber { f, .. }
        | Cell::CellFormulaBoolean { f, .. }
        | Cell::CellFormulaString { f, .. }
        | Cell::CellFormulaError { f, .. } => Some(*f as usize),
        _ => None,
    }
}
