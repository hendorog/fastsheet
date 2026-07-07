#![allow(clippy::unwrap_used)]

// Excel's 1900 date system (inherited from Lotus 1-2-3) believes 1900 was a
// leap year: serial 1 = 1900-01-01, serial 60 = the fictitious 1900-02-29,
// serial 61 = 1900-03-01. We emulate the phantom day exactly.
// Every expected value below is the value Excel returns.

use crate::cell::CellValue;
use crate::test::util::new_empty_model;

fn serial(model: &crate::model::Model, cell: &str) -> f64 {
    match model.get_cell_value_by_ref(&format!("Sheet1!{cell}")) {
        Ok(CellValue::Number(f)) => f,
        other => panic!("expected a number in {cell}, got {other:?}"),
    }
}

#[test]
fn serials_1_to_61_match_excel() {
    let mut model = new_empty_model();
    model._set("A1", "=DATE(1900,1,1)");
    model._set("A2", "=DATE(1900,2,28)");
    model._set("A3", "=DATE(1900,2,29)");
    model._set("A4", "=DATE(1900,3,1)");
    model._set("A5", "=DATE(1900,1,60)");
    model.evaluate();

    assert_eq!(serial(&model, "A1"), 1.0);
    assert_eq!(serial(&model, "A2"), 59.0);
    assert_eq!(serial(&model, "A3"), 60.0);
    assert_eq!(serial(&model, "A4"), 61.0);
    assert_eq!(serial(&model, "A5"), 60.0);
}

#[test]
fn day_month_year_of_early_serials() {
    let mut model = new_empty_model();
    model._set("A1", "=YEAR(1)");
    model._set("A2", "=MONTH(1)");
    model._set("A3", "=DAY(1)");
    model._set("A4", "=DAY(59)");
    model._set("A5", "=DAY(60)");
    model._set("A6", "=MONTH(60)");
    model._set("A7", "=DAY(61)");
    model._set("A8", "=YEAR(0.5)");
    model.evaluate();

    assert_eq!(model._get_text("A1"), *"1900");
    assert_eq!(model._get_text("A2"), *"1");
    assert_eq!(model._get_text("A3"), *"1");
    assert_eq!(model._get_text("A4"), *"28");
    assert_eq!(model._get_text("A5"), *"29");
    assert_eq!(model._get_text("A6"), *"2");
    assert_eq!(model._get_text("A7"), *"1");
    assert_eq!(model._get_text("A8"), *"1900");
}

#[test]
fn weekday_cycle_is_continuous() {
    let mut model = new_empty_model();
    // Excel says serial 1 is a Sunday (1900-01-01 was really a Monday)
    model._set("A1", "=WEEKDAY(1)");
    model._set("A2", "=WEEKDAY(60)");
    model._set("A3", "=WEEKDAY(61)");
    // A modern date: 2020-01-01 was a Wednesday
    model._set("A4", "=WEEKDAY(DATE(2020,1,1))");
    model.evaluate();

    assert_eq!(model._get_text("A1"), *"1");
    assert_eq!(model._get_text("A2"), *"4");
    assert_eq!(model._get_text("A3"), *"5");
    assert_eq!(model._get_text("A4"), *"4");
}

#[test]
fn date_differences_count_the_phantom_day() {
    let mut model = new_empty_model();
    model._set("A1", "=DATE(1901,1,1)-DATE(1900,1,1)");
    model._set("A2", "=DATE(2021,1,1)-DATE(2020,1,1)");
    model.evaluate();

    // Excel counts 366 days for 1900 because of the phantom Feb 29
    assert_eq!(serial(&model, "A1"), 366.0);
    assert_eq!(serial(&model, "A2"), 366.0);
}

#[test]
fn date_year_rule_0_to_1899_adds_1900() {
    let mut model = new_empty_model();
    model._set("A1", "=DATE(108,1,2)");
    model._set("A2", "=DATE(2008,1,2)");
    model._set("A3", "=YEAR(DATE(99,7,4))");
    model.evaluate();

    assert_eq!(model._get_text("A1"), model._get_text("A2"));
    assert_eq!(model._get_text("A3"), *"1999");
}

#[test]
fn modern_serials_are_unchanged() {
    let mut model = new_empty_model();
    model._set("A1", "=DATE(2020,1,1)");
    model._set("A2", "=DATE(2026,7,7)");
    model._set("A3", "=TEXT(43831,\"yyyy-mm-dd\")");
    model.evaluate();

    assert_eq!(serial(&model, "A1"), 43831.0);
    assert_eq!(serial(&model, "A2"), 46210.0);
    assert_eq!(model._get_text("A3"), *"2020-01-01");
}
