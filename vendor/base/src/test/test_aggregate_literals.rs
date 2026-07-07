#![allow(clippy::unwrap_used)]

// Excel's aggregates coerce DIRECT scalar arguments (TRUE -> 1, "2" -> 2)
// but ignore text/logicals reached through references, and they iterate
// inline array constants. Every expected value below is Excel's.

use crate::test::util::new_empty_model;

#[test]
fn sum_coerces_direct_literals_but_not_references() {
    let mut model = new_empty_model();
    model._set("B1", "=\"2\"");
    model._set("B2", "TRUE");
    model._set("A1", "=SUM(\"2\",TRUE,1)");
    model._set("A2", "=SUM(B1)");
    model._set("A3", "=SUM(B1:B2)");
    model._set("A4", "=SUM(1=1,2=2)");
    model._set("A5", "=SUM(\"abc\")");
    model.evaluate();

    assert_eq!(model._get_text("A1"), *"4");
    assert_eq!(model._get_text("A2"), *"0");
    assert_eq!(model._get_text("A3"), *"0");
    assert_eq!(model._get_text("A4"), *"2");
    assert_eq!(model._get_text("A5"), *"#VALUE!");
}

#[test]
fn min_max_product_handle_arrays_and_literals() {
    let mut model = new_empty_model();
    model._set("A1", "=MIN({3,1,2})");
    model._set("A2", "=MAX({1,2,3})");
    model._set("A3", "=MAX(\"5\",2)");
    model._set("A4", "=PRODUCT({2,3},2)");
    model._set("A5", "=MIN(TRUE,2)");
    model.evaluate();

    assert_eq!(model._get_text("A1"), *"1");
    assert_eq!(model._get_text("A2"), *"3");
    assert_eq!(model._get_text("A3"), *"5");
    assert_eq!(model._get_text("A4"), *"12");
    assert_eq!(model._get_text("A5"), *"1");
}

#[test]
fn count_counta_iterate_arrays() {
    let mut model = new_empty_model();
    model._set("B1", "=\"2\"");
    model._set("A1", "=COUNT({1,2,\"a\",TRUE})");
    model._set("A2", "=COUNTA({1,\"a\";2,3})");
    model._set("A3", "=COUNT(\"7\",B1)");
    model.evaluate();

    assert_eq!(model._get_text("A1"), *"2");
    assert_eq!(model._get_text("A2"), *"4");
    assert_eq!(model._get_text("A3"), *"1");
}

#[test]
fn averagea_arrays_and_reference_text() {
    let mut model = new_empty_model();
    model._set("B1", "=\"txt\"");
    model._set("B2", "4");
    model._set("A1", "=AVERAGEA({1,\"a\",5})");
    // AVERAGEA counts referenced text as 0: (0 + 4) / 2
    model._set("A2", "=AVERAGEA(B1:B2)");
    model._set("A3", "=AVERAGEA(B1)");
    model.evaluate();

    assert_eq!(model._get_text("A1"), *"2");
    assert_eq!(model._get_text("A2"), *"2");
    assert_eq!(model._get_text("A3"), *"0");
}

#[test]
fn npv_coerces_direct_scalars() {
    let mut model = new_empty_model();
    model._set("A1", "=NPV(0.1,100,\"200\")");
    model._set("A2", "=NPV(0.1,100,200)");
    model.evaluate();

    assert_eq!(model._get_text("A1"), model._get_text("A2"));
}
