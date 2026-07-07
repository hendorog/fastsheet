#![allow(clippy::unwrap_used)]

// Excel treats two numbers as equal when they agree to 15 significant digits,
// and approximate-match lookups return the LAST of equal keys.
// Every expected value below is the value Excel returns.

use crate::test::util::new_empty_model;

#[test]
fn exact_match_absorbs_float_noise() {
    let mut model = new_empty_model();
    model._set("A1", "=0.1+0.2");
    model._set("A2", "0.4");
    model._set("B1", "=MATCH(0.3,A1:A2,0)");
    model._set("B2", "=COUNTIF(A1:A2,0.3)");
    model._set("B3", "=SUMIF(A1:A2,0.3,A1:A2)");
    model.evaluate();

    assert_eq!(model._get_text("B1"), *"1");
    assert_eq!(model._get_text("B2"), *"1");
    assert_eq!(model._get_text("B3"), *"0.3");
}

#[test]
fn vlookup_exact_match_absorbs_float_noise() {
    let mut model = new_empty_model();
    model._set("A1", "=1.1+2.2");
    model._set("B1", "found");
    model._set("C1", "=VLOOKUP(3.3,A1:B1,2,FALSE)");
    model.evaluate();

    assert_eq!(model._get_text("C1"), *"found");
}

#[test]
fn equality_scales_with_magnitude() {
    let mut model = new_empty_model();
    // 1e6 + 0.1 accumulated differently: absolute diff >> f64::EPSILON
    model._set("A1", "=1000000+0.1");
    model._set("A2", "0");
    model._set("B1", "=COUNTIF(A1:A2,1000000.1)");
    model.evaluate();

    assert_eq!(model._get_text("B1"), *"1");
}

#[test]
fn approximate_match_returns_last_duplicate() {
    let mut model = new_empty_model();
    for (row, v) in [(1, "1"), (2, "2"), (3, "2"), (4, "2"), (5, "3")] {
        model._set(&format!("A{row}"), v);
        model._set(&format!("B{row}"), &format!("r{row}"));
    }
    model._set("C1", "=MATCH(2,A1:A5,1)");
    model._set("C2", "=VLOOKUP(2,A1:B5,2)");
    model._set("C3", "=VLOOKUP(2.5,A1:B5,2)");
    model._set("C4", "=VLOOKUP(0.5,A1:B5,2)");
    model.evaluate();

    assert_eq!(model._get_text("C1"), *"4");
    assert_eq!(model._get_text("C2"), *"r4");
    assert_eq!(model._get_text("C3"), *"r4");
    assert_eq!(model._get_text("C4"), *"#N/A");
}

#[test]
fn comparison_operators_still_agree() {
    let mut model = new_empty_model();
    model._set("A1", "=0.1+0.2=0.3");
    model._set("A2", "=0.1+0.2>0.3");
    model._set("A3", "=1.1+2.2<=3.3");
    model.evaluate();

    assert_eq!(model._get_text("A1"), *"TRUE");
    assert_eq!(model._get_text("A2"), *"FALSE");
    assert_eq!(model._get_text("A3"), *"TRUE");
}
