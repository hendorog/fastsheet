//! Data analysis commands — `/D Summary`, `/D Filter`, `/D Distribution`,
//! `/D Regression`, `/D Parse`. Each command takes a selection
//! rectangle, computes a result grid, and returns it as
//! `Vec<Vec<String>>`. The frontend writes the result via the
//! existing per-cell `set_cell` path so undo/dirty-tracking flows
//! through the normal channels.
//!
//! Why backend? `code_quality.md` rule 4: "Reusable logic should live
//! in modules that can be called from Tauri command handlers,
//! headless probes, integration tests." The data ops were originally
//! in `+page.svelte` doing per-cell `fetchBand` + `set_cell` —
//! testable only by driving the GUI, slow on ranges with many
//! formulas, and duplicated effort across commands.
//!
//! The numeric-value extraction here goes via `Cell::value()` on the
//! evaluated cell, not the input text. So a column of `=A1*2`
//! formulas summarises by their *evaluated* numbers, not by trying
//! to f64-parse the formula text (which the prior frontend impl did
//! and silently dropped).

use ironcalc::base::types::Cell;
use ironcalc::base::Model;
use serde::Serialize;
use tauri::State;

use crate::state::AppState;

/// Friendly error wrapper. Keeps the Tauri command signatures tight.
type Out = Result<Vec<Vec<String>>, String>;

fn validate_range(r1: i32, c1: i32, r2: i32, c2: i32) -> Result<(i32, i32, i32, i32), String> {
    if r1 < 1 || c1 < 1 || r2 < r1 || c2 < c1 {
        return Err("Invalid selection".to_string());
    }
    Ok((r1, c1, r2, c2))
}

/// Pull a cell's evaluated value as a numeric `Option<f64>`. Numbers
/// (literal or formula-cached) yield Some; booleans yield Some(0.0/1.0)
/// since spreadsheets historically treat them as numeric in math
/// contexts; strings, errors, empties yield None.
fn cell_as_f64(cell: &Cell) -> Option<f64> {
    match cell {
        Cell::NumberCell { v, .. } => Some(*v),
        Cell::CellFormulaNumber { v, .. } => Some(*v),
        Cell::BooleanCell { v, .. } => Some(if *v { 1.0 } else { 0.0 }),
        Cell::CellFormulaBoolean { v, .. } => Some(if *v { 1.0 } else { 0.0 }),
        _ => None,
    }
}

/// Read a cell as a display-suitable string. For numbers we use the
/// f64's lossless Display (so f64 round-trips), for strings the raw
/// content, for empties an empty string. We avoid `get_formatted_cell_value`
/// here because callers want round-trippable text, not locale-formatted.
fn cell_as_string(cell: &Cell, shared: &[String]) -> String {
    match cell {
        Cell::EmptyCell { .. } => String::new(),
        Cell::BooleanCell { v, .. } | Cell::CellFormulaBoolean { v, .. } => {
            if *v { "TRUE".into() } else { "FALSE".into() }
        }
        Cell::NumberCell { v, .. } | Cell::CellFormulaNumber { v, .. } => format!("{v}"),
        Cell::ErrorCell { .. } | Cell::CellFormulaError { .. } => String::new(),
        Cell::SharedString { si, .. } => {
            shared.get(*si as usize).cloned().unwrap_or_default()
        }
        Cell::CellFormula { .. } => String::new(),
        Cell::CellFormulaString { v, .. } => v.clone(),
    }
}

/// Read the range as a 2D grid of display strings. Empty cells come
/// back as empty strings so callers can index by (r-r1, c-c1) without
/// gaps.
fn read_range_strings(
    model: &Model,
    sheet: u32,
    r1: i32,
    c1: i32,
    r2: i32,
    c2: i32,
) -> Result<Vec<Vec<String>>, String> {
    let ws = model.workbook.worksheet(sheet)?;
    let shared = &model.workbook.shared_strings;
    let mut grid = Vec::with_capacity((r2 - r1 + 1) as usize);
    for r in r1..=r2 {
        let mut row = Vec::with_capacity((c2 - c1 + 1) as usize);
        for c in c1..=c2 {
            let s = ws
                .cell(r, c)
                .map(|cell| cell_as_string(cell, shared))
                .unwrap_or_default();
            row.push(s);
        }
        grid.push(row);
    }
    Ok(grid)
}

/// Read the range as a 2D grid of f64 options. None for cells that
/// can't be coerced to a number.
fn read_range_numbers(
    model: &Model,
    sheet: u32,
    r1: i32,
    c1: i32,
    r2: i32,
    c2: i32,
) -> Result<Vec<Vec<Option<f64>>>, String> {
    let ws = model.workbook.worksheet(sheet)?;
    let mut grid = Vec::with_capacity((r2 - r1 + 1) as usize);
    for r in r1..=r2 {
        let mut row = Vec::with_capacity((c2 - c1 + 1) as usize);
        for c in c1..=c2 {
            row.push(ws.cell(r, c).and_then(cell_as_f64));
        }
        grid.push(row);
    }
    Ok(grid)
}

fn col_letter(col: i32) -> String {
    let mut n = col;
    let mut s = String::new();
    while n > 0 {
        n -= 1;
        s.insert(0, (b'A' + (n % 26) as u8) as char);
        n /= 26;
    }
    s
}

fn fmt_number(n: f64) -> String {
    // Compact display — strip trailing zeros after a decimal point so
    // counts and integers don't get cluttered with ".0".
    if n.fract() == 0.0 && n.abs() < 1e15 {
        format!("{}", n as i64)
    } else {
        format!("{n}")
    }
}

#[tauri::command]
pub(crate) fn data_summary(
    sheet: u32,
    r1: i32,
    c1: i32,
    r2: i32,
    c2: i32,
    state: State<'_, AppState>,
) -> Out {
    let (r1, c1, r2, c2) = validate_range(r1, c1, r2, c2)?;
    if r2 - r1 < 1 {
        return Err("Select a header row and at least one data row".into());
    }
    let guard = state.model.lock().unwrap();
    let model = guard.as_ref().ok_or("no workbook open")?;
    let strings = read_range_strings(model, sheet, r1, c1, r2, c2)?;
    let numbers = read_range_numbers(model, sheet, r1, c1, r2, c2)?;
    let mut out: Vec<Vec<String>> = vec![vec![
        "Column".into(),
        "Count".into(),
        "Numeric".into(),
        "Sum".into(),
        "Average".into(),
        "Min".into(),
        "Max".into(),
    ]];
    for offset in 0..=(c2 - c1) as usize {
        let header = strings[0]
            .get(offset)
            .filter(|s| !s.trim().is_empty())
            .cloned()
            .unwrap_or_else(|| col_letter(c1 + offset as i32));
        let mut count = 0usize;
        let mut nums: Vec<f64> = Vec::new();
        for row in 1..strings.len() {
            let s = strings[row].get(offset).cloned().unwrap_or_default();
            if !s.trim().is_empty() {
                count += 1;
            }
            if let Some(Some(n)) = numbers.get(row).and_then(|r| r.get(offset)) {
                nums.push(*n);
            }
        }
        let sum: f64 = nums.iter().copied().sum();
        let n_count = nums.len();
        let mut row = vec![header, count.to_string(), n_count.to_string()];
        if n_count > 0 {
            let min = nums.iter().copied().fold(f64::INFINITY, f64::min);
            let max = nums.iter().copied().fold(f64::NEG_INFINITY, f64::max);
            row.push(fmt_number(sum));
            row.push(fmt_number(sum / n_count as f64));
            row.push(fmt_number(min));
            row.push(fmt_number(max));
        } else {
            row.extend(["".into(), "".into(), "".into(), "".into()]);
        }
        out.push(row);
    }
    Ok(out)
}

#[tauri::command]
pub(crate) fn data_filter(
    sheet: u32,
    r1: i32,
    c1: i32,
    r2: i32,
    c2: i32,
    filter_col: i32,
    needle: String,
    state: State<'_, AppState>,
) -> Out {
    let (r1, c1, r2, c2) = validate_range(r1, c1, r2, c2)?;
    if filter_col < c1 || filter_col > c2 {
        return Err(format!("Column {filter_col} not in selection"));
    }
    if r2 - r1 < 1 {
        return Err("Select a header row and at least one data row".into());
    }
    let guard = state.model.lock().unwrap();
    let model = guard.as_ref().ok_or("no workbook open")?;
    let strings = read_range_strings(model, sheet, r1, c1, r2, c2)?;
    let needle_lower = needle.to_lowercase();
    let idx = (filter_col - c1) as usize;
    let mut out = Vec::with_capacity(strings.len());
    out.push(strings[0].clone());
    for row in &strings[1..] {
        let cell = row.get(idx).cloned().unwrap_or_default();
        if cell.to_lowercase().contains(&needle_lower) {
            out.push(row.clone());
        }
    }
    Ok(out)
}

#[tauri::command]
pub(crate) fn data_distribution(
    sheet: u32,
    r1: i32,
    c1: i32,
    r2: i32,
    c2: i32,
    state: State<'_, AppState>,
) -> Out {
    let (r1, c1, r2, c2) = validate_range(r1, c1, r2, c2)?;
    let guard = state.model.lock().unwrap();
    let model = guard.as_ref().ok_or("no workbook open")?;
    let numbers = read_range_numbers(model, sheet, r1, c1, r2, c2)?;
    // Bucket on raw f64 bit pattern to keep equal numbers grouped.
    // Most use cases land on integers so this is fine; for genuine
    // float distributions the user would want histograms, not
    // exact-match counts (a future enhancement).
    let mut counts: std::collections::BTreeMap<u64, (f64, usize)> = Default::default();
    for row in &numbers {
        for cell in row {
            if let Some(n) = cell {
                let key = n.to_bits();
                let entry = counts.entry(key).or_insert((*n, 0));
                entry.1 += 1;
            }
        }
    }
    if counts.is_empty() {
        return Err("No numeric values in selection".into());
    }
    let mut entries: Vec<(f64, usize)> = counts.values().copied().collect();
    entries.sort_by(|a, b| a.0.partial_cmp(&b.0).unwrap_or(std::cmp::Ordering::Equal));
    let mut out: Vec<Vec<String>> = vec![vec!["Value".into(), "Frequency".into()]];
    for (v, c) in entries {
        out.push(vec![fmt_number(v), c.to_string()]);
    }
    Ok(out)
}

#[tauri::command]
pub(crate) fn data_regression(
    sheet: u32,
    r1: i32,
    c1: i32,
    r2: i32,
    c2: i32,
    state: State<'_, AppState>,
) -> Out {
    let (r1, c1, r2, c2) = validate_range(r1, c1, r2, c2)?;
    if c2 - c1 + 1 < 2 {
        return Err("Select at least two columns: X then Y".into());
    }
    let guard = state.model.lock().unwrap();
    let model = guard.as_ref().ok_or("no workbook open")?;
    let numbers = read_range_numbers(model, sheet, r1, c1, r2, c2)?;
    let mut pairs: Vec<(f64, f64)> = Vec::new();
    for row in &numbers {
        if let (Some(Some(x)), Some(Some(y))) = (row.first(), row.get(1)) {
            pairs.push((*x, *y));
        }
    }
    if pairs.len() < 2 {
        return Err("Regression needs at least two numeric X/Y pairs".into());
    }
    let n = pairs.len() as f64;
    let sx: f64 = pairs.iter().map(|p| p.0).sum();
    let sy: f64 = pairs.iter().map(|p| p.1).sum();
    let sxx: f64 = pairs.iter().map(|p| p.0 * p.0).sum();
    let sxy: f64 = pairs.iter().map(|p| p.0 * p.1).sum();
    let syy: f64 = pairs.iter().map(|p| p.1 * p.1).sum();
    let denom = n * sxx - sx * sx;
    if denom == 0.0 {
        return Err("Regression failed: X values have no variance".into());
    }
    let slope = (n * sxy - sx * sy) / denom;
    let intercept = (sy - slope * sx) / n;
    let r_denom = ((n * sxx - sx * sx) * (n * syy - sy * sy)).sqrt();
    let r = if r_denom == 0.0 {
        0.0
    } else {
        (n * sxy - sx * sy) / r_denom
    };
    Ok(vec![
        vec!["Regression".into(), "".into()],
        vec!["Count".into(), pairs.len().to_string()],
        vec!["Slope".into(), fmt_number(slope)],
        vec!["Intercept".into(), fmt_number(intercept)],
        vec!["R".into(), fmt_number(r)],
        vec!["R^2".into(), fmt_number(r * r)],
    ])
}

#[derive(serde::Deserialize)]
pub(crate) struct ParseDelimiter {
    /// "comma" | "tab" | "semicolon" | "space" | a single character
    pub kind: String,
    pub literal: Option<String>,
}

#[tauri::command]
pub(crate) fn data_parse(
    sheet: u32,
    r1: i32,
    c1: i32,
    r2: i32,
    c2: i32,
    delimiter: ParseDelimiter,
    state: State<'_, AppState>,
) -> Out {
    let (r1, c1, r2, c2) = validate_range(r1, c1, r2, c2)?;
    if c1 != c2 {
        return Err("Select one column to parse".into());
    }
    let sep = match delimiter.kind.as_str() {
        "comma" => ",",
        "tab" => "\t",
        "semicolon" => ";",
        "space" => " ",
        _ => match delimiter.literal.as_deref() {
            Some(s) if s.chars().count() == 1 => s,
            _ => return Err("Invalid delimiter".into()),
        },
    };
    let guard = state.model.lock().unwrap();
    let model = guard.as_ref().ok_or("no workbook open")?;
    let strings = read_range_strings(model, sheet, r1, c1, r2, c2)?;
    let mut out: Vec<Vec<String>> = Vec::with_capacity(strings.len());
    for row in strings {
        let source = row.into_iter().next().unwrap_or_default();
        out.push(split_delimited(&source, sep));
    }
    Ok(out)
}

/// CSV-aware split. Recognises double-quoted fields (with `""` as an
/// embedded literal quote) for the comma/tab/semicolon/single-char
/// delimiters; whitespace is treated as `\s+` so runs collapse into
/// one separator.
fn split_delimited(text: &str, sep: &str) -> Vec<String> {
    if sep == " " {
        return text.split_whitespace().map(|s| s.to_string()).collect();
    }
    let sep_char = sep.chars().next().unwrap_or(',');
    let mut out = Vec::new();
    let mut cur = String::new();
    let mut quoted = false;
    let chars: Vec<char> = text.chars().collect();
    let mut i = 0;
    while i < chars.len() {
        let ch = chars[i];
        if ch == '"' {
            if quoted && chars.get(i + 1) == Some(&'"') {
                cur.push('"');
                i += 2;
                continue;
            }
            quoted = !quoted;
            i += 1;
        } else if !quoted && ch == sep_char {
            out.push(std::mem::take(&mut cur));
            i += 1;
        } else {
            cur.push(ch);
            i += 1;
        }
    }
    out.push(cur);
    out
}

/// Re-serialize `Output` so the frontend doesn't need a second type.
#[derive(Serialize)]
#[allow(dead_code)]
pub(crate) struct DataAnalysisResult {
    pub rows: Vec<Vec<String>>,
}

#[cfg(test)]
mod tests {
    use super::split_delimited;

    #[test]
    fn split_handles_simple_csv() {
        let parts = split_delimited("a,b,c", ",");
        assert_eq!(parts, vec!["a", "b", "c"]);
    }

    #[test]
    fn split_preserves_quoted_separator() {
        let parts = split_delimited(r#""a,b",c"#, ",");
        assert_eq!(parts, vec!["a,b", "c"]);
    }

    #[test]
    fn split_unescapes_doubled_quote() {
        let parts = split_delimited(r#""he said ""hi""",x"#, ",");
        assert_eq!(parts, vec![r#"he said "hi""#, "x"]);
    }

    #[test]
    fn split_collapses_whitespace_for_space_delim() {
        let parts = split_delimited("foo   bar\tbaz", " ");
        assert_eq!(parts, vec!["foo", "bar", "baz"]);
    }

    #[test]
    fn split_with_empty_fields_for_csv() {
        let parts = split_delimited("a,,b", ",");
        assert_eq!(parts, vec!["a", "", "b"]);
    }
}
