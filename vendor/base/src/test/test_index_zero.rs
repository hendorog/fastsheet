#![allow(clippy::unwrap_used)]

// INDEX with row_num/col_num of 0 returns the whole row/column in Excel.
// Every expected value below is the value Excel returns.

use crate::test::util::new_empty_model;

fn model_with_grid<'a>() -> crate::model::Model<'a> {
    let mut model = new_empty_model();
    // A1:B3 = 1..6 row-major
    model._set("A1", "1");
    model._set("B1", "2");
    model._set("A2", "3");
    model._set("B2", "4");
    model._set("A3", "5");
    model._set("B3", "6");
    model
}

#[test]
fn index_zero_row_or_column() {
    let mut model = model_with_grid();
    model._set("D1", "=SUM(INDEX(A1:B3,0,2))");
    model._set("D2", "=SUM(INDEX(A1:B3,2,0))");
    model._set("D3", "=SUM(INDEX(A1:B3,0,0))");
    model._set("D4", "=AVERAGE(INDEX(A1:B3,0,1))");
    model.evaluate();

    assert_eq!(model._get_text("D1"), *"12");
    assert_eq!(model._get_text("D2"), *"7");
    assert_eq!(model._get_text("D3"), *"21");
    assert_eq!(model._get_text("D4"), *"3");
}

#[test]
fn index_zero_out_of_bounds() {
    let mut model = model_with_grid();
    model._set("D1", "=INDEX(A1:B3,0,3)");
    model._set("D2", "=INDEX(A1:B3,4,0)");
    model._set("D3", "=INDEX(A1:B3,-1,1)");
    model._set("D4", "=INDEX(A1:B3,1,-1)");
    model.evaluate();

    assert_eq!(model._get_text("D1"), *"#REF!");
    assert_eq!(model._get_text("D2"), *"#REF!");
    assert_eq!(model._get_text("D3"), *"#VALUE!");
    assert_eq!(model._get_text("D4"), *"#VALUE!");
}

#[test]
fn index_single_cell_still_works() {
    let mut model = model_with_grid();
    model._set("D1", "=INDEX(A1:B3,2,2)");
    model._set("D2", "=INDEX(A1:A3,3)");
    model.evaluate();

    assert_eq!(model._get_text("D1"), *"4");
    assert_eq!(model._get_text("D2"), *"5");
}
