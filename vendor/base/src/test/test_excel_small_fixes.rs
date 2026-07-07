#![allow(clippy::unwrap_used)]

// Excel-fidelity fixes: RATE guess arg, NaN/Inf -> #NUM!, TIME truncation,
// FIND/SEARCH start_num in characters, LEFT/RIGHT/MID numeric-text coercion,
// TRIM space-only collapse, COUNTIF "<>" numeric-text criteria.
// Every expected value below is the value Excel returns.

use crate::test::util::new_empty_model;

#[test]
fn rate_accepts_guess_argument() {
    let mut model = new_empty_model();
    model._set("A1", "=RATE(24,-250,5000)");
    model._set("A2", "=RATE(24,-250,5000,0,0,0.1)");
    // 7 arguments is still an error
    model._set("A3", "=RATE(24,-250,5000,0,0,0.1,1)");
    model.evaluate();

    assert_eq!(model._get_text("A1"), model._get_text("A2"));
    assert_eq!(model._get_text("A3"), *"#ERROR!");
}

#[test]
fn out_of_domain_results_are_num_errors() {
    let mut model = new_empty_model();
    model._set("A1", "=ASIN(2)");
    model._set("A2", "=ACOS(-1.5)");
    model._set("A3", "=ATANH(2)");
    model._set("A4", "=ATANH(1)");
    model._set("A5", "=ACOSH(0.5)");
    model._set("A6", "=EXP(1000)");
    // still-valid domains are untouched
    model._set("B1", "=ASIN(1)*2");
    model.evaluate();

    for cell in ["A1", "A2", "A3", "A4", "A5", "A6"] {
        assert_eq!(model._get_text(cell), *"#NUM!", "cell {cell}");
    }
    assert_eq!(model._get_text("B1"), *"3.141592654");
}

#[test]
fn time_truncates_components_toward_zero() {
    let mut model = new_empty_model();
    // trunc(-0.5) = 0, so this is TIME(12,0,0)
    model._set("A1", "=TIME(12,-0.5,0)");
    model._set("A2", "=TIME(12,0,0)");
    model.evaluate();

    assert_eq!(model._get_text("A1"), model._get_text("A2"));
}

#[test]
fn find_search_start_num_counts_characters() {
    let mut model = new_empty_model();
    // "été" has 3 characters (5 bytes): start 4 is out of range
    model._set("A1", "=FIND(\"t\",\"été\",4)");
    model._set("A2", "=FIND(\"t\",\"été\",2)");
    model._set("A3", "=SEARCH(\"T\",\"été\",4)");
    model._set("A4", "=SEARCH(\"T\",\"été\",2)");
    model.evaluate();

    assert_eq!(model._get_text("A1"), *"#VALUE!");
    assert_eq!(model._get_text("A2"), *"2");
    assert_eq!(model._get_text("A3"), *"#VALUE!");
    assert_eq!(model._get_text("A4"), *"2");
}

#[test]
fn left_right_mid_coerce_numeric_text_counts() {
    let mut model = new_empty_model();
    model._set("A1", "=LEFT(\"abcd\",\"2\")");
    model._set("A2", "=RIGHT(\"abcd\",TRUE)");
    model._set("A3", "=MID(\"abcd\",\"2\",\"2\")");
    model._set("A4", "=LEFT(\"abcd\",\"x\")");
    model._set("A5", "=MID(\"abcd\",0,1)");
    model.evaluate();

    assert_eq!(model._get_text("A1"), *"ab");
    assert_eq!(model._get_text("A2"), *"d");
    assert_eq!(model._get_text("A3"), *"bc");
    assert_eq!(model._get_text("A4"), *"#VALUE!");
    assert_eq!(model._get_text("A5"), *"#VALUE!");
}

#[test]
fn trim_collapses_only_ascii_spaces() {
    let mut model = new_empty_model();
    model._set("A1", "=TRIM(\"  a   b  \")");
    model._set("A2", "=TRIM(\"a b\")");
    model._set("A3", "=TRIM(\"   \")");
    // NBSP (160) and tab (9) are not touched by Excel's TRIM
    model._set("A4", "=LEN(TRIM(\"a\"&CHAR(160)&CHAR(160)&\"b\"))");
    model._set("A5", "=LEN(TRIM(CHAR(9)&\"a\"))");
    model.evaluate();

    assert_eq!(model._get_text("A1"), *"a b");
    assert_eq!(model._get_text("A2"), *"a b");
    assert_eq!(model._get_text("A3"), *"");
    assert_eq!(model._get_text("A4"), *"4");
    assert_eq!(model._get_text("A5"), *"2");
}

#[test]
fn countif_not_equal_matches_numeric_text() {
    let mut model = new_empty_model();
    model._set("A1", "5");
    model._set("A2", "=\"5\"");
    model._set("A3", "=\"five\"");
    model._set("B1", "=COUNTIF(A1:A3,\"<>5\")");
    model._set("B2", "=COUNTIF(A1:A3,5)");
    model.evaluate();

    // numeric text "5" equals 5 for COUNTIF in Excel, in both branches
    assert_eq!(model._get_text("B1"), *"1");
    assert_eq!(model._get_text("B2"), *"2");
}
