//! probe — load each fixture .xlsx and print cell values + formulas.
//! This validates IronCalc fidelity for the fastsheet spike, independent of the UI.
//!
//! Usage: cargo run --bin probe -- <fixture.xlsx> [<fixture.xlsx> ...]
//!
//! Verifies for each file:
//!   - load succeeds
//!   - sheet count
//!   - cells in the bounded region: input (formula or value text) + evaluated text
//!   - any cell errors visible as #REF! / #NAME? / #VALUE! etc

use std::time::Instant;

use fastsheet_lib::{extract_hidden_col_ranges, load_xls, load_xlsx_with_fallback, replicate_my_array_formulas};

fn col_letter(mut col: u32) -> String {
    let mut out = String::new();
    while col > 0 {
        let r = ((col - 1) % 26) as u8;
        out.insert(0, (b'A' + r) as char);
        col = (col - 1) / 26;
    }
    out
}
fn col_letter_i_pb(col: i32) -> String {
    if col < 1 { return String::new(); }
    col_letter(col as u32)
}

fn probe(path: &str) -> Result<(), String> {
    println!("=== {path} ===");
    let t0 = Instant::now();
    let is_xls = std::path::Path::new(path)
        .extension()
        .and_then(|s| s.to_str())
        .map(|e| e.eq_ignore_ascii_case("xls"))
        .unwrap_or(false);
    let mut model = if is_xls { load_xls(path)?.0 } else { load_xlsx_with_fallback(path)? };
    // Match the GUI: replicate MY* array formulas before evaluate (xlsx only).
    let replicated = if is_xls {
        0
    } else {
        std::fs::read(path)
            .ok()
            .and_then(|b| replicate_my_array_formulas(&mut model, &b).ok())
            .unwrap_or(0)
    };
    let load_ms = t0.elapsed().as_millis();
    let t1 = Instant::now();
    model.evaluate();
    let eval_ms = t1.elapsed().as_millis();
    if replicated > 0 {
        println!("  replicated {replicated} MY* spill cells");
    }
    println!(
        "  load: {load_ms}ms  evaluate: {eval_ms}ms  sheets: {}",
        model.workbook.worksheets.len()
    );
    for (i, ws) in model.workbook.worksheets.iter().enumerate() {
        println!("  [{i}] sheet {:?}  dimension {:?}", ws.name, ws.dimension);
    }
    // Probe a 12x14 window — covers every fixture's used range.
    let mut error_count = 0usize;
    let mut nonempty = 0usize;
    for sheet_idx in 0..model.workbook.worksheets.len() as u32 {
        println!(
            "  >>> sheet [{sheet_idx}] {:?}",
            model.workbook.worksheets[sheet_idx as usize].name
        );
        // Match the existing behaviour: 14×12 window on every sheet, plus
        // an extra 22..50 × 1..12 strip so the MY* spill targets in
        // SurfaceSelector / Temp / SurfaceTable are covered too.
        let row_ranges: [(i32, i32); 2] = [(1, 14), (22, 50)];
        for (rs, re) in row_ranges {
        for row in rs..=re {
            for col in 1..=12_i32 {
                let val = model
                    .get_formatted_cell_value(sheet_idx, row, col)
                    .unwrap_or_default();
                if val.is_empty() {
                    continue;
                }
                let input = model
                    .get_localized_cell_content(sheet_idx, row, col)
                    .unwrap_or_default();
                nonempty += 1;
                let is_err = val.starts_with('#') && val.ends_with('!');
                if is_err {
                    error_count += 1;
                }
                let marker = if is_err { " ⚠" } else { "" };
                println!(
                    "  {}{:>2}: {:<28}  →  {}{}",
                    col_letter(col as u32),
                    row,
                    input,
                    val,
                    marker
                );
            }
        }
        }
    }
    println!(
        "  -- {nonempty} non-empty cells, {error_count} errors --"
    );
    // Error-type breakdown — useful for spotting regressions between
    // loader changes (e.g. whether #REF! count jumps).
    {
        use std::collections::BTreeMap;
        let mut by_type: BTreeMap<String, usize> = BTreeMap::new();
        for sheet_idx in 0..model.workbook.worksheets.len() as u32 {
            for row in 1..=60 {
                for col in 1..=14_i32 {
                    let val = model.get_formatted_cell_value(sheet_idx, row, col).unwrap_or_default();
                    if val.starts_with('#') && val.ends_with('!') {
                        *by_type.entry(val).or_insert(0) += 1;
                    }
                }
            }
        }
        println!("  error-type breakdown: {by_type:?}");
    }
    // Count fills + font sizes per sheet — sanity check for the xls
    // load's style pipeline. On a typical xls these should be non-zero.
    {
        use std::collections::BTreeMap;
        let mut total_fills = 0usize;
        let mut histogram: BTreeMap<i32, usize> = BTreeMap::new();
        for (idx, _ws) in model.workbook.worksheets.iter().enumerate() {
            for row in 1..=60 {
                for col in 1..=15 {
                    if let Ok(s) = model.get_style_for_cell(idx as u32, row, col) {
                        if s.fill.pattern_type == "solid" && s.fill.fg_color.is_some() {
                            total_fills += 1;
                        }
                        if !model
                            .get_formatted_cell_value(idx as u32, row, col)
                            .unwrap_or_default()
                            .is_empty()
                        {
                            *histogram.entry(s.font.sz).or_insert(0) += 1;
                        }
                    }
                }
            }
        }
        println!("  style-summary: {total_fills} filled cells, font-size histogram {histogram:?}");
    }
    // Dump fill colors for a few cells in each sheet so we can see what
    // IronCalc resolved theme/indexed colours to.
    for (idx, ws) in model.workbook.worksheets.iter().enumerate().take(6) {
        if ws.name == "Genoa" || ws.name == "Main" || ws.name == "Downwind" {
            println!("  ---- {} cell styles ----", ws.name);
            for (r, c) in [(13, 3), (13, 4), (14, 3), (14, 4), (3, 3), (5, 3)] {
                if let Ok(s) = model.get_style_for_cell(idx as u32, r, c) {
                    let f = &s.fill;
                    println!(
                        "    {}{}  pattern={} fg={:?} bg={:?}  text_color={:?}",
                        col_letter_i_pb(c),
                        r,
                        f.pattern_type,
                        f.fg_color,
                        f.bg_color,
                        s.font.color
                    );
                }
            }
        }
    }
    // Dump column widths + hidden-col ranges for first few sheets so we
    // can see which sheets have custom widths / hidden cols.
    let bytes = std::fs::read(path).ok();
    for (idx, ws) in model.workbook.worksheets.iter().enumerate().take(5) {
        print!("  s{idx} {:<20} cols:", ws.name);
        for c in 1..=10 {
            let w = model.get_column_width(idx as u32, c).unwrap_or(0.0);
            print!(" {w:.0}");
        }
        if let Some(b) = bytes.as_deref() {
            let sheet_path = format!("xl/worksheets/sheet{}.xml", idx + 1);
            let hidden = extract_hidden_col_ranges(b, &sheet_path);
            print!("  hidden={hidden:?}");
        }
        println!();
    }
    println!();
    Ok(())
}

fn main() {
    let args: Vec<String> = std::env::args().skip(1).collect();
    if args.is_empty() {
        eprintln!("usage: probe <fixture.xlsx> [...]");
        std::process::exit(2);
    }
    let mut failed = 0;
    for p in &args {
        if let Err(e) = probe(p) {
            eprintln!("FAIL {p}: {e}");
            failed += 1;
        }
    }
    if failed > 0 {
        std::process::exit(1);
    }
}
