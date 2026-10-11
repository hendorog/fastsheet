//! Excel-semantics validation for function edge cases (pending item #2
//! in CLAUDE.md): INT, INDIRECT (dynamic / R1C1 refs), COLUMN /
//! COLUMNS, SHEET / SHEETS (reference arg), OFFSET.
//!
//! All functions are already implemented in vendored IronCalc 0.7.1 —
//! these tests pin the non-trivial cases to Excel's documented
//! behaviour so an upstream upgrade (or our own vendor patch) can't
//! silently drift. Expected values were derived from the Excel
//! function reference, not from IronCalc's current output.

use ironcalc::base::Model;

/// Sheet1 ("Sheet1") holds a small numeric block, sheet 2 ("Data")
/// one marker value. The formula under test always lands in E5 so
/// COLUMN() / ROW() / relative-R1C1 cases have a known anchor.
fn test_model() -> Model<'static> {
    let mut m = Model::new_empty("edge", "en", "UTC", "en").expect("model");
    m.add_sheet("Data").expect("add sheet");
    // Sheet1!A1:B3 = [[10,40],[20,50],[30,60]]
    m.set_user_input(0, 1, 1, "10".to_string()).unwrap();
    m.set_user_input(0, 2, 1, "20".to_string()).unwrap();
    m.set_user_input(0, 3, 1, "30".to_string()).unwrap();
    m.set_user_input(0, 1, 2, "40".to_string()).unwrap();
    m.set_user_input(0, 2, 2, "50".to_string()).unwrap();
    m.set_user_input(0, 3, 2, "60".to_string()).unwrap();
    // Anchor-relative targets around the E5 formula cell:
    // F6 for R[1]C[1], F5 for bare-row RC[1], E1 for bare-column R1C.
    m.set_user_input(0, 6, 6, "77".to_string()).unwrap();
    m.set_user_input(0, 5, 6, "88".to_string()).unwrap();
    m.set_user_input(0, 1, 5, "55".to_string()).unwrap();
    // Data!A1
    m.set_user_input(1, 1, 1, "99".to_string()).unwrap();
    m
}

fn assert_cases(cases: &[(&str, &str)]) {
    let mut m = test_model();
    let mut failures = vec![];
    for (formula, expected) in cases {
        m.set_user_input(0, 5, 5, formula.to_string()).unwrap();
        m.evaluate();
        let got = m.get_formatted_cell_value(0, 5, 5).unwrap_or_default();
        if got != *expected {
            failures.push(format!("{formula}: expected {expected:?}, got {got:?}"));
        }
    }
    assert!(failures.is_empty(), "\n{}", failures.join("\n"));
}

#[test]
fn int_floors_toward_negative_infinity() {
    assert_cases(&[
        ("=INT(8.9)", "8"),
        ("=INT(-8.5)", "-9"), // floor, NOT truncate: Excel gives -9
        ("=INT(-0.1)", "-1"),
        ("=INT(0)", "0"),
        ("=INT(\"8.9\")", "8"), // numeric-text coercion
        ("=INT(TRUE)", "1"),
    ]);
}

#[test]
fn column_and_columns() {
    assert_cases(&[
        ("=COLUMN()", "5"),      // formula lives in E5
        ("=COLUMN(C7)", "3"),    // explicit ref wins over own position
        ("=COLUMN(B2:D4)", "2"), // range → first (leftmost) column
        ("=COLUMN(Data!A1)", "1"),
        ("=COLUMNS(B2:D4)", "3"),
        ("=COLUMNS(A1)", "1"),
        ("=COLUMNS(A:C)", "3"), // whole-column range
    ]);
}

#[test]
fn sheet_and_sheets() {
    assert_cases(&[
        ("=SHEET()", "1"),         // 1-based index of the formula's own sheet
        ("=SHEET(Data!A1)", "2"),  // reference arg → sheet of the reference
        ("=SHEET(\"Data\")", "2"), // sheet-name text arg
        ("=SHEETS()", "2"),        // workbook sheet count
        ("=SHEET(\"NoSuchSheet\")", "#N/A"),
    ]);
}

#[test]
fn indirect_static_dynamic_and_r1c1() {
    assert_cases(&[
        ("=INDIRECT(\"A1\")", "10"),
        ("=INDIRECT(\"B1\",TRUE)", "40"),
        ("=INDIRECT(\"Data!A1\")", "99"),        // cross-sheet text ref
        ("=INDIRECT(\"A\"&(1+2))", "30"),        // dynamically built address
        ("=SUM(INDIRECT(\"A1:B2\"))", "120"),    // range-shaped result
        ("=INDIRECT(\"R2C1\",FALSE)", "20"),     // absolute R1C1
        ("=INDIRECT(\"r2c1\",FALSE)", "20"),     // case-insensitive
        ("=INDIRECT(\"R[1]C[1]\",FALSE)", "77"), // relative R1C1 from E5 → F6
        ("=INDIRECT(\"R[-4]C[-4]\",FALSE)", "10"), // negative offsets → A1
        ("=INDIRECT(\"R1C\",FALSE)", "55"),      // bare C = own column (E) → E1
        ("=INDIRECT(\"RC[1]\",FALSE)", "88"),    // bare R = own row (5) → F5
        ("=SUM(INDIRECT(\"R1C1:R2C2\",FALSE))", "120"), // R1C1 range
        ("=INDIRECT(\"Data!R1C1\",FALSE)", "99"), // sheet-prefixed R1C1
        ("=INDIRECT(\"R0C1\",FALSE)", "#REF!"),  // row 0 is out of bounds
        ("=INDIRECT(\"C1R1\",FALSE)", "#REF!"),  // wrong axis order
        ("=INDIRECT(\"bogus name\")", "#REF!"),  // unparsable → #REF!
    ]);
}

#[test]
fn offset_moves_and_resizes() {
    assert_cases(&[
        ("=OFFSET(A1,1,1)", "50"),
        ("=OFFSET(A1,2,0)", "30"),
        ("=SUM(OFFSET(A1,0,0,2,2))", "120"),   // resized to A1:B2
        ("=SUM(OFFSET(B3,-2,-1,2,1))", "30"),  // negative offsets stay in-grid
        ("=OFFSET(A1,-1,0)", "#REF!"),         // walks off the top
        ("=OFFSET(A1,0,-1)", "#REF!"),         // walks off the left
        ("=OFFSET(A1,0,0,0,1)", "#REF!"),      // zero height is invalid
        ("=OFFSET(A1,0,0,1,0)", "#REF!"),      // zero width is invalid
        ("=OFFSET(A1,1048576,0)", "#REF!"),    // beyond the last row
    ]);
}

#[test]
fn lookup_blank_value_semantics() {
    // Vendor patch (found via PricerPRU 2026-08): a blank lookup cell
    // behaves as the number 0, and blank cells in the lookup vector
    // never match it. Without the patch VLOOKUP/HLOOKUP "matched"
    // blank-to-blank and returned data from a blank-keyed row/column
    // instead of #N/A. Model layout: D1 blank, D2:D4 = [Leech, 7, 0].
    let mut m = test_model();
    m.set_user_input(0, 2, 4, "Leech".to_string()).unwrap();
    m.set_user_input(0, 3, 4, "7".to_string()).unwrap();
    m.set_user_input(0, 4, 4, "0".to_string()).unwrap();
    let cases = [
        // Z9 is blank; D1 (blank) must NOT match → but D4 holds an
        // actual 0, which a blank lookup (coerced to 0) DOES match.
        ("=VLOOKUP(Z9,D1:E4,1,FALSE)", "0"),
        // Blank lookup against a vector with no zeros → #N/A.
        ("=VLOOKUP(Z9,D1:E2,1,FALSE)", "#N/A"),
        ("=HLOOKUP(Z9,A1:B2,2,FALSE)", "#N/A"),
        // Non-blank lookups are unaffected.
        ("=VLOOKUP(\"Leech\",D1:E4,1,FALSE)", "Leech"),
        ("=VLOOKUP(7,D1:E4,1,FALSE)", "7"),
    ];
    let mut failures = vec![];
    for (formula, expected) in cases {
        m.set_user_input(0, 5, 5, formula.to_string()).unwrap();
        m.evaluate();
        let got = m.get_formatted_cell_value(0, 5, 5).unwrap_or_default();
        if got != expected {
            failures.push(format!("{formula}: expected {expected:?}, got {got:?}"));
        }
    }
    assert!(failures.is_empty(), "\n{}", failures.join("\n"));
}
