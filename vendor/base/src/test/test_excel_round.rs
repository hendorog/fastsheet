#![allow(clippy::unwrap_used)]

// ROUND/ROUNDUP/ROUNDDOWN operate on the 15-significant-digit decimal
// representation, like Excel; binary scaling gets 2.675 & friends wrong.
// Every expected value below is the value Excel returns.

use crate::test::util::new_empty_model;

#[test]
fn round_decimal_boundaries() {
    let mut model = new_empty_model();
    model._set("A1", "=ROUND(2.675,2)");
    model._set("A2", "=ROUND(-2.675,2)");
    model._set("A3", "=ROUND(1.005,2)");
    model._set("A4", "=ROUND(2.5,0)");
    model._set("A5", "=ROUND(-2.5,0)");
    model._set("A6", "=ROUND(1.45,1)");
    model.evaluate();

    assert_eq!(model._get_text("A1"), *"2.68");
    assert_eq!(model._get_text("A2"), *"-2.68");
    assert_eq!(model._get_text("A3"), *"1.01");
    assert_eq!(model._get_text("A4"), *"3");
    assert_eq!(model._get_text("A5"), *"-3");
    assert_eq!(model._get_text("A6"), *"1.5");
}

#[test]
fn round_negative_digits() {
    let mut model = new_empty_model();
    model._set("A1", "=ROUND(12345.6789,-2)");
    model._set("A2", "=ROUND(12350,-2)");
    model._set("A3", "=ROUNDUP(3.141,-3)");
    model._set("A4", "=ROUNDDOWN(987,-1)");
    model._set("A5", "=ROUND(12,-5)");
    model.evaluate();

    assert_eq!(model._get_text("A1"), *"12300");
    assert_eq!(model._get_text("A2"), *"12400");
    assert_eq!(model._get_text("A3"), *"1000");
    assert_eq!(model._get_text("A4"), *"980");
    assert_eq!(model._get_text("A5"), *"0");
}

#[test]
fn roundup_rounddown_directions() {
    let mut model = new_empty_model();
    model._set("A1", "=ROUNDUP(0.0001,0)");
    model._set("A2", "=ROUNDUP(-0.0001,0)");
    model._set("A3", "=ROUNDDOWN(2.675,2)");
    model._set("A4", "=ROUNDUP(2.671,2)");
    model._set("A5", "=ROUNDDOWN(-3.99999,0)");
    model._set("A6", "=ROUNDUP(-3.00001,0)");
    model.evaluate();

    assert_eq!(model._get_text("A1"), *"1");
    assert_eq!(model._get_text("A2"), *"-1");
    assert_eq!(model._get_text("A3"), *"2.67");
    assert_eq!(model._get_text("A4"), *"2.68");
    assert_eq!(model._get_text("A5"), *"-3");
    assert_eq!(model._get_text("A6"), *"-4");
}

#[test]
fn round_extremes_and_identity() {
    let mut model = new_empty_model();
    model._set("A1", "=ROUND(2.675,20)");
    model._set("A2", "=ROUND(0,5)");
    model._set("A3", "=ROUND(0.5,0)");
    model._set("A4", "=ROUND(0.4,0)");
    model._set("A5", "=ROUND(1.05*(0.0284+0.0046)-0.0284,4)");
    model.evaluate();

    assert_eq!(model._get_text("A1"), *"2.675");
    assert_eq!(model._get_text("A2"), *"0");
    assert_eq!(model._get_text("A3"), *"1");
    assert_eq!(model._get_text("A4"), *"0");
    assert_eq!(model._get_text("A5"), *"0.0063");
}
