//! oracle_probe — Excel-fidelity oracle for .xls workbooks.
//!
//! Loads each workbook through the app's xls loader, recalculates every
//! formula with IronCalc, and diffs the results against the cached values
//! Excel stored in the file (read directly with calamine). Cached values are
//! what Excel itself computed, so any mismatch is either a fidelity bug or a
//! formula the loader mis-translated.
//!
//! Usage: cargo run --release --features dev-bins --example oracle_probe -- <file.xls> [...]
//!
//! Cells whose formula is volatile (NOW/TODAY/RAND/RANDBETWEEN) are skipped;
//! cells that merely depend on volatile cells will still be reported, so
//! files built around TODAY() are expected to show cascades.

use std::io::Cursor;

use calamine::{Data, Reader, Xls};
use fastsheet_lib::load_xls;
use ironcalc::base::cell::CellValue;

fn col_letter(mut col: u32) -> String {
    let mut out = String::new();
    while col > 0 {
        let r = ((col - 1) % 26) as u8;
        out.insert(0, (b'A' + r) as char);
        col = (col - 1) / 26;
    }
    out
}

fn is_volatile(formula: &str) -> bool {
    let up = formula.to_uppercase();
    ["NOW(", "TODAY(", "RAND(", "RANDBETWEEN("]
        .iter()
        .any(|f| up.contains(f))
}

/// Excel-style equality: agree to ~15 significant digits.
fn numbers_match(expected: f64, got: f64) -> bool {
    if expected == got {
        return true;
    }
    (expected - got).abs() <= 1e-9 * expected.abs().max(1.0)
}

struct FileReport {
    checked: usize,
    volatile_skipped: usize,
    no_oracle: usize,
    diffs: Vec<String>,
}

fn probe(path: &str, max_diffs_shown: usize) -> Result<FileReport, String> {
    let bytes = std::fs::read(path).map_err(|e| format!("read: {e}"))?;
    let mut wb: Xls<_> =
        calamine::open_workbook_from_rs(Cursor::new(&bytes)).map_err(|e| format!("calamine: {e}"))?;
    let sheet_names = wb.sheet_names().to_owned();

    let mut model = load_xls(path)?.0;
    model.evaluate();

    // Map calamine sheet names to model sheet indexes by name
    let model_sheets: Vec<String> = model
        .workbook
        .worksheets
        .iter()
        .map(|ws| ws.name.clone())
        .collect();

    let mut report = FileReport {
        checked: 0,
        volatile_skipped: 0,
        no_oracle: 0,
        diffs: Vec::new(),
    };

    for name in &sheet_names {
        let Some(sheet_idx) = model_sheets.iter().position(|n| n == name) else {
            report
                .diffs
                .push(format!("{name}: sheet missing from loaded model"));
            continue;
        };
        let sheet_idx = sheet_idx as u32;
        let Ok(formulas) = wb.worksheet_formula(name) else {
            continue;
        };
        let Ok(values) = wb.worksheet_range(name) else {
            continue;
        };
        let (frow, fcol) = formulas.start().unwrap_or((0, 0));
        for (rel_row, row) in formulas.rows().enumerate() {
            for (rel_col, formula) in row.iter().enumerate() {
                if formula.is_empty() {
                    continue;
                }
                if is_volatile(formula) {
                    report.volatile_skipped += 1;
                    continue;
                }
                let abs_row = frow + rel_row as u32;
                let abs_col = fcol + rel_col as u32;
                let expected = values.get_value((abs_row, abs_col));
                let got = model
                    .get_cell_value_by_index(sheet_idx, abs_row as i32 + 1, abs_col as i32 + 1)
                    .unwrap_or(CellValue::None);

                let matches = match (expected, &got) {
                    // No cached value stored: nothing to compare against
                    (None | Some(Data::Empty), _) => {
                        report.no_oracle += 1;
                        continue;
                    }
                    (Some(Data::Float(x)), CellValue::Number(y)) => numbers_match(*x, *y),
                    (Some(Data::Int(i)), CellValue::Number(y)) => numbers_match(*i as f64, *y),
                    (Some(Data::DateTime(dt)), CellValue::Number(y)) => {
                        numbers_match(dt.as_f64(), *y)
                    }
                    (Some(Data::String(s)), CellValue::String(t)) => s == t,
                    (Some(Data::Bool(b)), CellValue::Boolean(c)) => b == c,
                    (Some(Data::Error(e)), CellValue::String(t)) => &e.to_string() == t,
                    // Excel stores an empty string for some blank-result formulas
                    (Some(Data::String(s)), CellValue::None) => s.is_empty(),
                    _ => false,
                };
                report.checked += 1;
                if !matches {
                    let cell_ref =
                        format!("{name}!{}{}", col_letter(abs_col + 1), abs_row + 1);
                    let loaded_formula = model
                        .get_cell_formula(sheet_idx, abs_row as i32 + 1, abs_col as i32 + 1)
                        .ok()
                        .flatten()
                        .unwrap_or_else(|| "<no formula in model>".to_string());
                    report.diffs.push(format!(
                        "{cell_ref}: ={formula}\n      loaded: {loaded_formula}\n      excel: {expected:?}\n      ours:  {got:?}"
                    ));
                }
            }
        }
    }

    println!(
        "  checked {} formula cells; {} diffs; {} volatile skipped; {} without stored value",
        report.checked,
        report.diffs.len(),
        report.volatile_skipped,
        report.no_oracle
    );
    for d in report.diffs.iter().take(max_diffs_shown) {
        println!("    {d}");
    }
    if report.diffs.len() > max_diffs_shown {
        println!("    ... {} more", report.diffs.len() - max_diffs_shown);
    }
    Ok(report)
}

fn main() {
    let args: Vec<String> = std::env::args().skip(1).collect();
    if args.is_empty() {
        eprintln!("usage: oracle_probe <file.xls> [...]");
        std::process::exit(2);
    }
    let max_shown: usize = std::env::var("ORACLE_MAX_DIFFS")
        .ok()
        .and_then(|v| v.parse().ok())
        .unwrap_or(10);
    let mut total_checked = 0usize;
    let mut total_diffs = 0usize;
    let mut failed_files = 0usize;
    for path in &args {
        println!("=== {path} ===");
        match probe(path, max_shown) {
            Ok(r) => {
                total_checked += r.checked;
                total_diffs += r.diffs.len();
            }
            Err(e) => {
                println!("  LOAD ERROR: {e}");
                failed_files += 1;
            }
        }
    }
    println!(
        "\nTOTAL: {} files, {} formula cells checked, {} diffs, {} files failed to load",
        args.len(),
        total_checked,
        total_diffs,
        failed_files
    );
}
