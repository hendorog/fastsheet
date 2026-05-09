//! save_probe — load an xls, save via xls_save, attempt to re-open,
//! diff cell values vs the original. Diagnostic for the xls write
//! pipeline.
//!
//! Usage: cargo run --bin save_probe -- <input.xls>
//!
//! Output: load summary, structural info on the saved file (cfb +
//! BIFF header sanity), and a per-cell value comparison.
//!   total / matching / mismatching / missing_in_rt
//! Up to 5 sample diffs are printed with the source cell's resolved
//! style index + format code, and the round-tripped equivalent, to
//! help triage what's getting lost.

use std::io::Read;

fn main() {
    let path = std::env::args().nth(1).expect("usage: save_probe <input.xls>");
    println!("=== load {path}");
    let (mut model, _hidden, _preserved, raw_rgce) = fastsheet_lib::load_xls(&path).expect("load_xls");
    model.evaluate();
    let prgce = fastsheet_lib::xls_load::finalize_preserved_rgce(&model, raw_rgce);
    if !prgce.is_empty() {
        println!("  preserved rgce for {} #ERROR!-displaying cells", prgce.len());
    }

    // Optional cell-level inspection: set FASTSHEET_PROBE_CELL=sheet:row:col
    // to dump the cell's formula text + parsed Node + cell variant.
    if let Ok(target) = std::env::var("FASTSHEET_PROBE_CELL") {
        let parts: Vec<&str> = target.split(':').collect();
        if parts.len() == 3 {
            let s: usize = parts[0].parse().unwrap();
            let r: i32 = parts[1].parse().unwrap();
            let c: i32 = parts[2].parse().unwrap();
            if let Some(ws) = model.workbook.worksheets.get(s) {
                if let Some(cell) = ws.sheet_data.get(&r).and_then(|cols| cols.get(&c)) {
                    use ironcalc::base::types::Cell;
                    let f_idx = match cell {
                        Cell::CellFormula { f, .. }
                        | Cell::CellFormulaNumber { f, .. }
                        | Cell::CellFormulaBoolean { f, .. }
                        | Cell::CellFormulaString { f, .. }
                        | Cell::CellFormulaError { f, .. } => Some(*f),
                        _ => None,
                    };
                    if let Some(f) = f_idx {
                        if let Some(formulas) = ws.shared_formulas.get(f as usize) {
                            println!("  PROBE sheet[{s}] r{r}c{c}: formula=\"{formulas}\"");
                        }
                        if let Some(node) = model.parsed_formulas.get(s).and_then(|v| v.get(f as usize)) {
                            println!("  PROBE node: {node:?}");
                        }
                    }
                    println!("  PROBE cell: {cell:?}");
                }
                let v = model.get_formatted_cell_value(s as u32, r, c).unwrap_or_default();
                println!("  PROBE formatted: {v:?}");
            }
        }
    }

    println!(
        "  loaded: {} sheets, shared_strings={}, cell_xfs={}, num_fmts={}, fonts={}, fills={}, borders={}",
        model.workbook.worksheets.len(),
        model.workbook.shared_strings.len(),
        model.workbook.styles.cell_xfs.len(),
        model.workbook.styles.num_fmts.len(),
        model.workbook.styles.fonts.len(),
        model.workbook.styles.fills.len(),
        model.workbook.styles.borders.len(),
    );

    let out = std::env::temp_dir().join("save_probe_out.xls");
    println!("=== save_xls → {}", out.display());
    let preserved = fastsheet_lib::xls_preserve::extract(&std::fs::read(&path).expect("read"));
    if !preserved.workbook_substreams.is_empty() {
        println!("  preserved {} substreams from source workbook stream", preserved.workbook_substreams.len());
    }
    let bytes = fastsheet_lib::xls_save::build_xls_bytes_with_options(&model, None, Some(&prgce), Some(&preserved));
    fastsheet_lib::xls_save::write_xls_bytes_with_preserved(&out, &bytes, None).expect("save_xls");
    let bytes = std::fs::read(&out).expect("read back");
    println!("  wrote {} bytes", bytes.len());

    if bytes.len() >= 32 {
        let sec_log = u16::from_le_bytes([bytes[30], bytes[31]]);
        println!("  cfb sector size: 2^{} = {} bytes", sec_log, 1 << sec_log);
    }

    println!("=== re-open via cfb");
    match cfb::open(&out) {
        Ok(mut comp) => {
            let entries: Vec<_> = comp.walk().map(|e| e.path().to_owned()).collect();
            println!("  cfb entries: {entries:?}");
            if let Ok(mut s) = comp.open_stream("/Workbook") {
                let mut buf = Vec::new();
                if let Err(e) = s.read_to_end(&mut buf) {
                    println!("  /Workbook read error: {e}");
                } else {
                    println!("  /Workbook read OK: {} bytes", buf.len());
                }
            } else {
                println!("  /Workbook stream missing");
            }
        }
        Err(e) => {
            println!("  cfb::open FAILED: {e}");
        }
    }

    println!("=== re-open via load_xls");
    let (mut reloaded, _, _, _) = match fastsheet_lib::load_xls(&out.to_string_lossy()) {
        Ok(t) => t,
        Err(e) => {
            println!("  load_xls FAILED: {e}");
            return;
        }
    };
    println!("  load_xls OK: {} sheets", reloaded.workbook.worksheets.len());
    reloaded.evaluate();

    // Per-cell diff against the original. get_formatted_cell_value returns
    // the displayed text (post-format), so this catches both value-loss
    // and format-loss bugs.
    let mut total = 0u64;
    let mut matching = 0u64;
    let mut mismatching = 0u64;
    let mut missing_in_rt = 0u64;
    let mut sample_misses: Vec<String> = Vec::new();
    for (sheet_idx, ws) in model.workbook.worksheets.iter().enumerate() {
        for (row, cols) in ws.sheet_data.iter() {
            for (col, _cell) in cols.iter() {
                total += 1;
                let orig = model
                    .get_formatted_cell_value(sheet_idx as u32, *row, *col)
                    .unwrap_or_default();
                let rt = reloaded
                    .get_formatted_cell_value(sheet_idx as u32, *row, *col)
                    .unwrap_or_default();
                if orig == rt {
                    matching += 1;
                } else if rt.is_empty() && !orig.is_empty() {
                    missing_in_rt += 1;
                    if sample_misses.len() < 5 {
                        sample_misses.push(format!(
                            "  sheet[{sheet_idx}] r{row}c{col}: orig={orig:?} rt=(empty)"
                        ));
                    }
                } else {
                    mismatching += 1;
                    if sample_misses.len() < 5 {
                        sample_misses.push(format!(
                            "  sheet[{sheet_idx}] r{row}c{col}: orig={orig:?} rt={rt:?}"
                        ));
                    }
                }
            }
        }
    }
    println!(
        "  cells: total={total} matching={matching} mismatching={mismatching} missing_in_rt={missing_in_rt}"
    );
    if !sample_misses.is_empty() {
        println!("  sample diffs:");
        for s in &sample_misses {
            println!("{s}");
        }
    }

    // Optional: with FASTSHEET_PROBE_FULL_DIFF=1 set, dump all
    // mismatches grouped by the kind of diff (numeric drift / error
    // mismatch / N/A propagation / string format).
    // Compare orig vs rt for a list of cells; useful for tracing
    // a single cascade. Cells specified as "sheet:row:col" comma-sep.
    if let Ok(cells) = std::env::var("FASTSHEET_PROBE_COMPARE") {
        for spec in cells.split(',') {
            let parts: Vec<&str> = spec.split(':').collect();
            if parts.len() != 3 { continue; }
            let s: u32 = parts[0].parse().unwrap();
            let r: i32 = parts[1].parse().unwrap();
            let c: i32 = parts[2].parse().unwrap();
            let orig = model.get_formatted_cell_value(s, r, c).unwrap_or_default();
            let rt = reloaded.get_formatted_cell_value(s, r, c).unwrap_or_default();
            let orig_cell = model.workbook.worksheets.get(s as usize)
                .and_then(|w| w.sheet_data.get(&r).and_then(|cols| cols.get(&c)));
            let rt_cell = reloaded.workbook.worksheets.get(s as usize)
                .and_then(|w| w.sheet_data.get(&r).and_then(|cols| cols.get(&c)));
            println!("COMPARE [{s}]r{r}c{c}: orig={orig:?} rt={rt:?}");
            println!("  orig cell: {orig_cell:?}");
            println!("  rt   cell: {rt_cell:?}");
            // For formula cells, also dump the formula text.
            use ironcalc::base::types::Cell;
            let f_index = |c: &Cell| match c {
                Cell::CellFormula { f, .. } | Cell::CellFormulaNumber { f, .. }
                | Cell::CellFormulaBoolean { f, .. } | Cell::CellFormulaString { f, .. }
                | Cell::CellFormulaError { f, .. } => Some(*f),
                _ => None,
            };
            if let Some(fi) = orig_cell.and_then(f_index) {
                if let Some(f) = model.workbook.worksheets[s as usize].shared_formulas.get(fi as usize) {
                    println!("  orig formula: {f:?}");
                }
            }
            if let Some(fi) = rt_cell.and_then(f_index) {
                if let Some(f) = reloaded.workbook.worksheets[s as usize].shared_formulas.get(fi as usize) {
                    println!("  rt   formula: {f:?}");
                }
            }
        }
    }
    if std::env::var("FASTSHEET_PROBE_FULL_DIFF").is_ok() {
        use std::collections::HashMap;
        let mut buckets: HashMap<&'static str, Vec<String>> = HashMap::new();
        for (sheet_idx, ws) in model.workbook.worksheets.iter().enumerate() {
            for (row, cols) in ws.sheet_data.iter() {
                for (col, _cell) in cols.iter() {
                    let orig = model
                        .get_formatted_cell_value(sheet_idx as u32, *row, *col)
                        .unwrap_or_default();
                    let rt = reloaded
                        .get_formatted_cell_value(sheet_idx as u32, *row, *col)
                        .unwrap_or_default();
                    if orig == rt { continue; }
                    let bucket = classify_diff(&orig, &rt);
                    let entry = format!(
                        "  sheet[{sheet_idx}] r{row}c{col}: orig={orig:?} rt={rt:?}"
                    );
                    buckets.entry(bucket).or_default().push(entry);
                }
            }
        }
        let mut bucket_names: Vec<_> = buckets.keys().copied().collect();
        bucket_names.sort();
        println!();
        println!("=== Full diff breakdown");
        for name in &bucket_names {
            let entries = &buckets[name];
            println!("[{name}] — {} cells", entries.len());
            for e in entries.iter().take(5) {
                println!("{e}");
            }
            if entries.len() > 5 {
                println!("  ... and {} more", entries.len() - 5);
            }
        }
    }
}

fn classify_diff(orig: &str, rt: &str) -> &'static str {
    let is_error = |s: &str| s.starts_with('#') && s.ends_with('!') || s == "#N/A";
    let orig_err = is_error(orig);
    let rt_err = is_error(rt);
    match (orig_err, rt_err) {
        (false, true) if rt == "#N/A" => "rt-na (cascade or lookup miss)",
        (false, true) if rt == "#DIV/0!" => "rt-div0",
        (false, true) if rt == "#VALUE!" => "rt-value",
        (false, true) => "rt-other-error",
        (true, false) => "orig-error-rt-clean",
        (true, true) if orig != rt => "error-code-changed",
        (false, false) => {
            // Both values; check if numeric or string difference.
            let orig_n = orig.parse::<f64>();
            let rt_n = rt.parse::<f64>();
            match (orig_n, rt_n) {
                (Ok(a), Ok(b)) if (a - b).abs() < 0.01 * a.abs().max(1.0) => {
                    "numeric-drift-<1pct"
                }
                (Ok(_), Ok(_)) => "numeric-drift-large",
                _ => "string-diff",
            }
        }
        _ => "other",
    }
}
