#![allow(clippy::unwrap_used)]

// Excel coerces numbers to text at 15 significant digits, with plain decimal
// notation for 1e-4 <= |v| < 1e21 and uppercase scientific outside.
// Every expected value below is the value Excel returns.

use crate::test::util::new_empty_model;

#[test]
fn concatenation_uses_15_significant_digits() {
    let mut model = new_empty_model();
    model._set("A1", "=0.1+0.2&\"\"");
    model._set("A2", "=\"x\"&1/3");
    model._set("A3", "=CONCATENATE(\"x\",1/3)");
    model._set("A4", "=EXACT(0.1+0.2&\"\",\"0.3\")");
    model.evaluate();

    assert_eq!(model._get_text("A1"), *"0.3");
    assert_eq!(model._get_text("A2"), *"x0.333333333333333");
    assert_eq!(model._get_text("A3"), *"x0.333333333333333");
    assert_eq!(model._get_text("A4"), *"TRUE");
}

#[test]
fn text_functions_see_the_15_digit_view() {
    let mut model = new_empty_model();
    model._set("A1", "=LEN(0.1+0.2)");
    model._set("A2", "=LEFT(1/3,5)");
    // note: =4.1-4 is 0.0999999999999996 even at 15 digits, in Excel too
    model._set("A3", "=TEXTJOIN(\",\",TRUE,0.1+0.2,4.1-4)");
    model.evaluate();

    assert_eq!(model._get_text("A1"), *"3");
    assert_eq!(model._get_text("A2"), *"0.333");
    assert_eq!(model._get_text("A3"), *"0.3,0.0999999999999996");
}

#[test]
fn magnitude_edges_match_excel_notation() {
    let mut model = new_empty_model();
    model._set("A1", "=10^16&\"\"");
    model._set("A2", "=10^21&\"\"");
    model._set("A3", "=0.0001&\"\"");
    model._set("A4", "=0.00001&\"\"");
    model.evaluate();

    assert_eq!(model._get_text("A1"), *"10000000000000000");
    assert_eq!(model._get_text("A2"), *"1E+21");
    assert_eq!(model._get_text("A3"), *"0.0001");
    assert_eq!(model._get_text("A4"), *"1E-05");
}
