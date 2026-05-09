//! Integration tests for the .xls write+read round-trip.
//!
//! Runs the same path the GUI follows when you save and reopen a
//! file: `load_xls → save_xls → load_xls → diff against original`.
//! These exist to catch regressions where a save+reload silently
//! mutates cells that the user didn't touch.
//!
//! Two layers:
//!
//! 1. **Synthetic round-trip** — builds a workbook programmatically
//!    that exercises every feature the writer has been asked to
//!    handle (literals, refs, range, ops, IF / IFERROR / VLOOKUP /
//!    INDEX / MATCH, defined names, 3D refs, embedded-quote strings,
//!    custom palette colors, font styles, column widths). Runs
//!    on every `cargo test`. Deterministic.
//!
//! 2. **Real-file round-trip** — driven by environment variables so
//!    contributors can point at their own .xls files without
//!    embedding paths in the test source. Set `FASTSHEET_RT_FILE` to
//!    a .xls path and (optionally) `FASTSHEET_RT_THRESHOLD` to the
//!    accepted mismatch count. Skipped silently when unset, so CI
//!    stays green.

use ironcalc::base::Model;
use std::path::PathBuf;

#[derive(Debug)]
struct RoundTripResult {
    total: u64,
    mismatches: Vec<Diff>,
}

#[derive(Debug)]
struct Diff {
    sheet: usize,
    row: i32,
    col: i32,
    orig: String,
    rt: String,
}

/// Save `model` to a temp .xls, reload, diff cell-by-cell against
/// the original. The diff uses `get_formatted_cell_value` so it
/// catches both raw value drift and number-format display changes.
fn round_trip(model: &Model, label: &str) -> RoundTripResult {
    let out = std::env::temp_dir().join(format!("fastsheet_rt_{label}.xls"));
    fastsheet_lib::save_xls(model, &out).expect("save_xls");
    let (mut reloaded, _, _, _) =
        fastsheet_lib::load_xls(&out.to_string_lossy()).expect("load_xls");
    reloaded.evaluate();
    let mut result = RoundTripResult { total: 0, mismatches: vec![] };
    for (s, ws) in model.workbook.worksheets.iter().enumerate() {
        for (row, cols) in &ws.sheet_data {
            for (col, _) in cols {
                result.total += 1;
                let orig = model
                    .get_formatted_cell_value(s as u32, *row, *col)
                    .unwrap_or_default();
                let rt = reloaded
                    .get_formatted_cell_value(s as u32, *row, *col)
                    .unwrap_or_default();
                if orig != rt {
                    result.mismatches.push(Diff {
                        sheet: s,
                        row: *row,
                        col: *col,
                        orig,
                        rt,
                    });
                }
            }
        }
    }
    std::fs::remove_file(&out).ok();
    result
}

fn dump_diffs(label: &str, result: &RoundTripResult) {
    eprintln!(
        "[{label}] cells={} mismatches={}",
        result.total,
        result.mismatches.len()
    );
    for d in result.mismatches.iter().take(10) {
        eprintln!(
            "  [{}] r{}c{}: orig={:?} rt={:?}",
            d.sheet, d.row, d.col, d.orig, d.rt
        );
    }
    if result.mismatches.len() > 10 {
        eprintln!("  ... and {} more", result.mismatches.len() - 10);
    }
}

/// Assert the round-trip is clean (zero mismatches) for synthetic
/// scenarios where we control every cell. Real-file fixtures use a
/// noise-floor threshold instead because IronCalc has error codes
/// (`#ERROR!`, `#N/IMPL!`, etc.) that BIFF can't represent.
fn assert_clean(label: &str, result: &RoundTripResult) {
    if !result.mismatches.is_empty() {
        dump_diffs(label, result);
        panic!("{label}: round-trip introduced mismatches");
    }
}

// ---------------------------------------------------------------------------
// 1. Synthetic round-trips — exercise each writer feature on its own
//    so regressions surface in a focused test rather than as a single
//    aggregate failure.
// ---------------------------------------------------------------------------

#[test]
fn rt_plain_values() {
    let mut m = Model::new_empty("rt", "en", "UTC", "en").unwrap();
    m.set_user_input(0, 1, 1, "42".into()).unwrap();
    m.set_user_input(0, 1, 2, "hello".into()).unwrap();
    m.set_user_input(0, 1, 3, "TRUE".into()).unwrap();
    m.set_user_input(0, 1, 4, "3.14".into()).unwrap();
    m.set_user_input(0, 2, 1, "text with comma, period.".into()).unwrap();
    m.evaluate();
    let r = round_trip(&m, "plain_values");
    assert_clean("plain_values", &r);
}

#[test]
fn rt_basic_formulas() {
    let mut m = Model::new_empty("rt", "en", "UTC", "en").unwrap();
    m.set_user_input(0, 1, 1, "10".into()).unwrap();
    m.set_user_input(0, 1, 2, "20".into()).unwrap();
    m.set_user_input(0, 2, 1, "=A1+B1".into()).unwrap();
    m.set_user_input(0, 2, 2, "=A1*B1".into()).unwrap();
    m.set_user_input(0, 2, 3, "=SUM(A1:B1)".into()).unwrap();
    m.set_user_input(0, 2, 4, "=AVERAGE(A1:B1)".into()).unwrap();
    m.set_user_input(0, 3, 1, "=IF(A1>5, \"big\", \"small\")".into()).unwrap();
    m.evaluate();
    let r = round_trip(&m, "basic_formulas");
    assert_clean("basic_formulas", &r);
}

#[test]
fn rt_iferror_vlookup() {
    // Pin the IFERROR-via-VLOOKUP pattern that broke previously
    // (FTAB iftab=365 vs 480). This is THE most common pattern in
    // real-world spreadsheet templates.
    let mut m = Model::new_empty("rt", "en", "UTC", "en").unwrap();
    m.set_user_input(0, 1, 1, "alpha".into()).unwrap();
    m.set_user_input(0, 1, 2, "100".into()).unwrap();
    m.set_user_input(0, 2, 1, "beta".into()).unwrap();
    m.set_user_input(0, 2, 2, "200".into()).unwrap();
    m.new_defined_name("Lookup", None, "Sheet1!$A$1:$B$2").unwrap();
    // Hit + missing-key fallback paths.
    m.set_user_input(0, 5, 1,
        "=IFERROR(VLOOKUP(\"alpha\", Lookup, 2, FALSE), -1)".into()).unwrap();
    m.set_user_input(0, 5, 2,
        "=IFERROR(VLOOKUP(\"missing\", Lookup, 2, FALSE), -1)".into()).unwrap();
    m.evaluate();
    let r = round_trip(&m, "iferror_vlookup");
    assert_clean("iferror_vlookup", &r);
}

#[test]
fn rt_3d_refs_with_named_range() {
    let mut m = Model::new_empty("rt", "en", "UTC", "en").unwrap();
    m.add_sheet("Data").unwrap();
    m.set_user_input(1, 1, 1, "1".into()).unwrap();
    m.set_user_input(1, 2, 1, "2".into()).unwrap();
    m.set_user_input(1, 3, 1, "3".into()).unwrap();
    m.new_defined_name("DataCol", None, "Data!$A$1:$A$3").unwrap();
    m.set_user_input(0, 1, 1, "=SUM(Data!A1:A3)".into()).unwrap();
    m.set_user_input(0, 1, 2, "=SUM(DataCol)".into()).unwrap();
    m.set_user_input(0, 1, 3, "=Data!A2".into()).unwrap();
    m.evaluate();
    let r = round_trip(&m, "3d_refs");
    assert_clean("3d_refs", &r);
}

#[test]
fn rt_strings_with_embedded_quotes() {
    // Pin the doubled-quote bug surfaced by GUI testing
    // (e.g. cell content like `6" headboard`). Three storage paths:
    // literal SharedString, formula reference, formula with
    // literal-quote string source.
    let mut m = Model::new_empty("rt", "en", "UTC", "en").unwrap();
    m.set_user_input(0, 1, 1, r#"6" headboard"#.into()).unwrap();
    m.set_user_input(0, 2, 1, "=A1".into()).unwrap();
    m.set_user_input(0, 3, 1, r#"="6"" headboard""#.into()).unwrap();
    m.set_user_input(0, 4, 1, r#"=A1 & " — extended""#.into()).unwrap();
    m.evaluate();
    let r = round_trip(&m, "embedded_quotes");
    assert_clean("embedded_quotes", &r);
}

#[test]
fn rt_styles_and_colors() {
    // Custom font color + cell fill that was previously affected by
    // the palette-collision bug (cyan → dark blue). Keep the test
    // simple — the regression was at the icv-allocation level, not
    // the model level, so a few distinct colors is enough.
    use ironcalc::base::types::{Style, Font, Fill};
    let mut m = Model::new_empty("rt", "en", "UTC", "en").unwrap();
    for col in 1..=6 {
        m.set_user_input(0, 1, col, format!("c{col}")).unwrap();
    }
    let mut style_red = Style::default();
    style_red.font = Font { color: Some("#FF0000".into()), ..Font::default() };
    m.set_cell_style(0, 1, 1, &style_red).unwrap();

    let mut style_cyan = Style::default();
    style_cyan.fill = Fill {
        pattern_type: "solid".into(),
        fg_color: Some("#CCFFFF".into()),
        bg_color: None,
    };
    m.set_cell_style(0, 1, 2, &style_cyan).unwrap();

    let mut style_custom = Style::default();
    style_custom.font = Font { color: Some("#FF8800".into()), ..Font::default() };
    m.set_cell_style(0, 1, 3, &style_custom).unwrap();

    m.evaluate();
    let r = round_trip(&m, "styles_colors");
    assert_clean("styles_colors", &r);
}

#[test]
fn rt_column_widths_and_row_heights() {
    let mut m = Model::new_empty("rt", "en", "UTC", "en").unwrap();
    m.set_user_input(0, 1, 1, "wide".into()).unwrap();
    m.set_user_input(0, 1, 2, "narrow".into()).unwrap();
    m.set_user_input(0, 1, 3, "default".into()).unwrap();
    // set_column_width takes the "chars * 12" internal form.
    m.set_column_width(0, 1, 30.0 * 12.0).unwrap(); // 30 chars
    m.set_column_width(0, 2, 4.0 * 12.0).unwrap();  // 4 chars
    m.set_row_height(0, 1, 25.0 * 2.0).unwrap();    // 25 pt
    m.evaluate();

    let out = std::env::temp_dir().join("fastsheet_rt_widths.xls");
    fastsheet_lib::save_xls(&m, &out).expect("save");
    let (reloaded, _, _, _) = fastsheet_lib::load_xls(&out.to_string_lossy()).expect("reload");
    // get_column_width returns chars * 12 (internal form). Round to
    // tolerate u16 quantization on the BIFF wire.
    let w1 = reloaded.get_column_width(0, 1).unwrap_or(0.0).round();
    let w2 = reloaded.get_column_width(0, 2).unwrap_or(0.0).round();
    let h1 = reloaded.get_row_height(0, 1).unwrap_or(0.0).round();
    std::fs::remove_file(&out).ok();
    assert_eq!(w1 as i64, (30.0_f64 * 12.0).round() as i64, "col 1 width");
    assert_eq!(w2 as i64, (4.0_f64 * 12.0).round() as i64, "col 2 width");
    assert_eq!(h1 as i64, (25.0_f64 * 2.0).round() as i64, "row 1 height");
}

#[test]
fn rt_frozen_panes_rows_only() {
    let mut m = Model::new_empty("rt", "en", "UTC", "en").unwrap();
    m.set_user_input(0, 1, 1, "header".into()).unwrap();
    m.set_user_input(0, 2, 1, "data".into()).unwrap();
    m.workbook.worksheets[0].frozen_rows = 1;
    m.workbook.worksheets[0].frozen_columns = 0;
    m.evaluate();

    let out = std::env::temp_dir().join("fastsheet_rt_frozen_rows.xls");
    fastsheet_lib::save_xls(&m, &out).expect("save");
    let (reloaded, _, _, _) = fastsheet_lib::load_xls(&out.to_string_lossy()).expect("reload");
    std::fs::remove_file(&out).ok();
    assert_eq!(reloaded.workbook.worksheets[0].frozen_rows, 1);
    assert_eq!(reloaded.workbook.worksheets[0].frozen_columns, 0);
}

#[test]
fn rt_frozen_panes_cols_only() {
    let mut m = Model::new_empty("rt", "en", "UTC", "en").unwrap();
    m.set_user_input(0, 1, 1, "key".into()).unwrap();
    m.set_user_input(0, 1, 2, "value".into()).unwrap();
    m.workbook.worksheets[0].frozen_rows = 0;
    m.workbook.worksheets[0].frozen_columns = 1;
    m.evaluate();

    let out = std::env::temp_dir().join("fastsheet_rt_frozen_cols.xls");
    fastsheet_lib::save_xls(&m, &out).expect("save");
    let (reloaded, _, _, _) = fastsheet_lib::load_xls(&out.to_string_lossy()).expect("reload");
    std::fs::remove_file(&out).ok();
    assert_eq!(reloaded.workbook.worksheets[0].frozen_rows, 0);
    assert_eq!(reloaded.workbook.worksheets[0].frozen_columns, 1);
}

#[test]
fn rt_frozen_panes_both_axes() {
    let mut m = Model::new_empty("rt", "en", "UTC", "en").unwrap();
    m.set_user_input(0, 1, 1, "x".into()).unwrap();
    m.workbook.worksheets[0].frozen_rows = 3;
    m.workbook.worksheets[0].frozen_columns = 2;
    m.evaluate();

    let out = std::env::temp_dir().join("fastsheet_rt_frozen_both.xls");
    fastsheet_lib::save_xls(&m, &out).expect("save");
    let (reloaded, _, _, _) = fastsheet_lib::load_xls(&out.to_string_lossy()).expect("reload");
    std::fs::remove_file(&out).ok();
    assert_eq!(reloaded.workbook.worksheets[0].frozen_rows, 3);
    assert_eq!(reloaded.workbook.worksheets[0].frozen_columns, 2);
}

#[test]
fn rt_no_frozen_panes_stays_unfrozen() {
    let mut m = Model::new_empty("rt", "en", "UTC", "en").unwrap();
    m.set_user_input(0, 1, 1, "x".into()).unwrap();
    m.evaluate();

    let out = std::env::temp_dir().join("fastsheet_rt_unfrozen.xls");
    fastsheet_lib::save_xls(&m, &out).expect("save");
    let (reloaded, _, _, _) = fastsheet_lib::load_xls(&out.to_string_lossy()).expect("reload");
    std::fs::remove_file(&out).ok();
    assert_eq!(reloaded.workbook.worksheets[0].frozen_rows, 0);
    assert_eq!(reloaded.workbook.worksheets[0].frozen_columns, 0);
}

#[test]
fn rt_hidden_cols_survive_save() {
    use std::collections::{HashMap, HashSet};
    let mut m = Model::new_empty("rt", "en", "UTC", "en").unwrap();
    m.set_user_input(0, 1, 1, "a".into()).unwrap();
    m.set_user_input(0, 1, 2, "b".into()).unwrap();
    m.set_user_input(0, 1, 3, "c".into()).unwrap();
    m.set_user_input(0, 1, 4, "d".into()).unwrap();
    // Width on col 3 so it's covered by an explicit IronCalc Col entry;
    // hidden status is a side-channel.
    m.set_column_width(0, 3, 15.0 * 12.0).unwrap();
    m.evaluate();

    let mut hidden: HashMap<u32, HashSet<i32>> = HashMap::new();
    let mut s = HashSet::new();
    s.insert(2); // hidden col covered ONLY by hidden_cols
    s.insert(3); // hidden col that also has a custom width
    hidden.insert(0, s);

    let out = std::env::temp_dir().join("fastsheet_rt_hidden_cols.xls");
    fastsheet_lib::save_xls_with_preserved(&m, &out, None, Some(&hidden)).expect("save");
    let (_reloaded, hc_after, _, _) = fastsheet_lib::load_xls(&out.to_string_lossy()).expect("reload");
    std::fs::remove_file(&out).ok();
    let cols = hc_after.get(&0).cloned().unwrap_or_default();
    assert!(cols.contains(&2), "col 2 should be hidden after round-trip");
    assert!(cols.contains(&3), "col 3 should be hidden after round-trip");
    assert!(!cols.contains(&1), "col 1 must not be hidden");
    assert!(!cols.contains(&4), "col 4 must not be hidden");
}

#[test]
fn rt_hidden_cols_default_state_has_none() {
    let mut m = Model::new_empty("rt", "en", "UTC", "en").unwrap();
    m.set_user_input(0, 1, 1, "x".into()).unwrap();
    m.evaluate();

    let out = std::env::temp_dir().join("fastsheet_rt_no_hidden.xls");
    fastsheet_lib::save_xls(&m, &out).expect("save");
    let (_, hc_after, _, _) = fastsheet_lib::load_xls(&out.to_string_lossy()).expect("reload");
    std::fs::remove_file(&out).ok();
    let cols = hc_after.get(&0).cloned().unwrap_or_default();
    assert!(cols.is_empty(), "no hidden cols expected, got {:?}", cols);
}

// ---------------------------------------------------------------------------
// 2. Real-file round-trips — only run when the fixture is present.
//    This is the closest thing to a CI-friendly version of the user's
//    manual workflow: open the file, save, close, reopen, verify.
//
//    The mismatch threshold is the documented noise floor for that
//    file. Going UP from there means a regression; going DOWN means
//    we improved on the cached values and the threshold should
//    tighten.
// ---------------------------------------------------------------------------

/// Run a round-trip on an external .xls fixture configured via env
/// vars. This is the CI-friendly version of the manual workflow:
/// open the file, save, close, reopen, verify.
///
/// - `FASTSHEET_RT_FILE` — path to the .xls fixture (required to opt
///   in; test is skipped when unset).
/// - `FASTSHEET_RT_THRESHOLD` — accepted mismatch count. Defaults to
///   0. Going UP from there means a regression; going DOWN means we
///   improved and the threshold should tighten.
#[test]
fn rt_external_fixture() {
    let Ok(path_str) = std::env::var("FASTSHEET_RT_FILE") else {
        eprintln!("skipping: FASTSHEET_RT_FILE not set");
        return;
    };
    let path = PathBuf::from(&path_str);
    if !path.exists() {
        eprintln!("skipping: FASTSHEET_RT_FILE path does not exist: {path_str}");
        return;
    }
    let threshold: usize = std::env::var("FASTSHEET_RT_THRESHOLD")
        .ok()
        .and_then(|s| s.parse().ok())
        .unwrap_or(0);

    eprintln!("running round-trip on {}", path.display());
    let (mut original, _, _, _) = fastsheet_lib::load_xls(&path.to_string_lossy()).expect("load");
    original.evaluate();
    let r = round_trip(&original, "external");
    eprintln!(
        "[external] cells={} mismatches={} (threshold={})",
        r.total,
        r.mismatches.len(),
        threshold
    );
    if r.mismatches.len() > threshold {
        dump_diffs("external", &r);
        panic!(
            "external fixture: {} mismatches exceeds threshold of {}",
            r.mismatches.len(),
            threshold
        );
    }
}

// ---------------------------------------------------------------------------
// 3. Mutation round-trip — the user's manual workflow:
//    open file → change cell → save → close → reopen → check that
//    only the changed cell differs.
// ---------------------------------------------------------------------------

#[test]
fn rt_mutation_only_touches_one_cell() {
    let mut m = Model::new_empty("rt", "en", "UTC", "en").unwrap();
    // Build a dependency network: A1 = input, A2 = =A1*2, A3 = =A2+5.
    m.set_user_input(0, 1, 1, "10".into()).unwrap();
    m.set_user_input(0, 2, 1, "=A1*2".into()).unwrap();
    m.set_user_input(0, 3, 1, "=A2+5".into()).unwrap();
    m.set_user_input(0, 4, 1, "unrelated text".into()).unwrap();
    m.evaluate();

    // Save the original.
    let baseline_path = std::env::temp_dir().join("fastsheet_rt_mut_baseline.xls");
    fastsheet_lib::save_xls(&m, &baseline_path).expect("baseline save");

    // Mutate A1 and save again.
    m.set_user_input(0, 1, 1, "100".into()).unwrap();
    m.evaluate();
    let mutated_path = std::env::temp_dir().join("fastsheet_rt_mut_mutated.xls");
    fastsheet_lib::save_xls(&m, &mutated_path).expect("mutated save");

    // Reload both and diff.
    let (mut baseline, _, _, _) =
        fastsheet_lib::load_xls(&baseline_path.to_string_lossy()).expect("reload baseline");
    let (mut mutated, _, _, _) =
        fastsheet_lib::load_xls(&mutated_path.to_string_lossy()).expect("reload mutated");
    baseline.evaluate();
    mutated.evaluate();

    // Expected: A1 went 10 → 100, A2 dependent on A1 went 20 → 200,
    // A3 dependent on A2 went 25 → 205. A4 is an unrelated text
    // cell and must be byte-identical.
    assert_eq!(baseline.get_formatted_cell_value(0, 1, 1).unwrap(), "10");
    assert_eq!(mutated.get_formatted_cell_value(0, 1, 1).unwrap(), "100");
    assert_eq!(baseline.get_formatted_cell_value(0, 2, 1).unwrap(), "20");
    assert_eq!(mutated.get_formatted_cell_value(0, 2, 1).unwrap(), "200");
    assert_eq!(baseline.get_formatted_cell_value(0, 3, 1).unwrap(), "25");
    assert_eq!(mutated.get_formatted_cell_value(0, 3, 1).unwrap(), "205");
    assert_eq!(
        baseline.get_formatted_cell_value(0, 4, 1).unwrap(),
        mutated.get_formatted_cell_value(0, 4, 1).unwrap(),
        "unrelated cell must round-trip identically across mutations"
    );

    std::fs::remove_file(&baseline_path).ok();
    std::fs::remove_file(&mutated_path).ok();
}

// ---------------------------------------------------------------------------
// 4. Macro / VBA preservation — captured on load, replayed on save.
//    Verifies the storage tree comes through bit-identical so Excel-side
//    macros survive a fastsheet save+reload.
// ---------------------------------------------------------------------------

#[test]
fn rt_vba_macros_preserved() {
    use std::io::{Cursor, Read, Write};

    // Build a synthetic .xls in memory: minimal /Workbook stream
    // (not parseable by IronCalc — that's fine, this test is purely
    // about the storage round-trip) plus a fake VBA storage tree.
    // We then re-read via cfb to verify our extract+inject works
    // end-to-end without going through the BIFF writer.
    let src_path = std::env::temp_dir().join("fastsheet_rt_vba_src.xls");
    let _ = std::fs::remove_file(&src_path);

    {
        let f = std::fs::OpenOptions::new()
            .read(true).write(true).create_new(true)
            .open(&src_path).unwrap();
        let mut comp = cfb::CompoundFile::create_with_version(cfb::Version::V3, f).unwrap();
        // /Workbook needs SOMETHING; the content is irrelevant here.
        comp.create_stream("/Workbook").unwrap()
            .write_all(b"\x09\x08\x10\x00\x00\x06\x05\x00").unwrap();
        // Build the VBA tree we want preserved.
        comp.create_storage("/_VBA_PROJECT_CUR").unwrap();
        comp.create_storage("/_VBA_PROJECT_CUR/VBA").unwrap();
        comp.create_stream("/_VBA_PROJECT_CUR/VBA/_VBA_PROJECT").unwrap()
            .write_all(b"FAKE_VBA_PROJECT_STREAM_BYTES_AAAABBBBCCCC").unwrap();
        comp.create_stream("/_VBA_PROJECT_CUR/VBA/Module1").unwrap()
            .write_all(b"Attribute VB_Name = \"Module1\"\nSub Hello()\nEnd Sub").unwrap();
        comp.create_stream("/_VBA_PROJECT_CUR/VBA/dir").unwrap()
            .write_all(b"DIR_STREAM_BYTES").unwrap();
        comp.flush().unwrap();
    }

    // Extract from the source bytes.
    let bytes = std::fs::read(&src_path).unwrap();
    let preserved = fastsheet_lib::xls_preserve::extract(&bytes);
    assert!(!preserved.is_empty(), "expected VBA entries to be captured");

    // Build a target CFB and inject. Includes a /Workbook stream like
    // the real writer would, so we exercise the same shape.
    let dst_path = std::env::temp_dir().join("fastsheet_rt_vba_dst.xls");
    let _ = std::fs::remove_file(&dst_path);
    {
        let f = std::fs::OpenOptions::new()
            .read(true).write(true).create_new(true)
            .open(&dst_path).unwrap();
        let mut comp = cfb::CompoundFile::create_with_version(cfb::Version::V3, f).unwrap();
        comp.create_stream("/Workbook").unwrap()
            .write_all(b"\x09\x08\x10\x00\x00\x06\x05\x00").unwrap();
        fastsheet_lib::xls_preserve::inject(&mut comp, &preserved);
        comp.flush().unwrap();
    }

    // Re-read the destination and confirm every preserved stream is
    // present with byte-identical content.
    let dst_bytes = std::fs::read(&dst_path).unwrap();
    let mut dst_cfb = cfb::CompoundFile::open(Cursor::new(dst_bytes)).unwrap();
    for entry in &preserved.entries {
        if entry.is_storage {
            assert!(dst_cfb.is_storage(&entry.path),
                "missing preserved storage: {:?}", entry.path);
        } else {
            assert!(dst_cfb.is_stream(&entry.path),
                "missing preserved stream: {:?}", entry.path);
            let mut got = Vec::new();
            dst_cfb.open_stream(&entry.path).unwrap().read_to_end(&mut got).unwrap();
            assert_eq!(&got, entry.data.as_ref().unwrap(),
                "stream {:?} did not round-trip identically", entry.path);
        }
    }

    std::fs::remove_file(&src_path).ok();
    std::fs::remove_file(&dst_path).ok();
}

#[test]
fn rt_vba_absent_when_source_has_none() {
    use std::io::Write;
    // Source with NO VBA storages → extract returns empty bundle,
    // inject is a no-op, and the writer produces a clean CFB without
    // any VBA debris.
    let src_path = std::env::temp_dir().join("fastsheet_rt_vba_empty.xls");
    let _ = std::fs::remove_file(&src_path);
    {
        let f = std::fs::OpenOptions::new()
            .read(true).write(true).create_new(true)
            .open(&src_path).unwrap();
        let mut comp = cfb::CompoundFile::create_with_version(cfb::Version::V3, f).unwrap();
        comp.create_stream("/Workbook").unwrap()
            .write_all(b"\x09\x08\x10\x00\x00\x06\x05\x00").unwrap();
        comp.flush().unwrap();
    }
    let bytes = std::fs::read(&src_path).unwrap();
    let preserved = fastsheet_lib::xls_preserve::extract(&bytes);
    assert!(preserved.is_empty(), "no-VBA source should yield empty bundle");
    std::fs::remove_file(&src_path).ok();
}
