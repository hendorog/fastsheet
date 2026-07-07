#![allow(clippy::unwrap_used)]

// Excel defined names may contain '?' and '\' after the first character
// (e.g. `Is_Ulx_Allowed?`, seen in real-world workbooks).

use crate::test::util::new_empty_model;

#[test]
fn defined_names_may_contain_question_mark() {
    let mut model = new_empty_model();
    model._set("C1", "TRUE");
    model
        .new_defined_name("Is_Ulx_Allowed?", None, "Sheet1!$C$1")
        .unwrap();
    model._set("B1", "=IF(Is_Ulx_Allowed?,\"yes\",\"no\")");
    model.evaluate();

    assert_eq!(model._get_text("B1"), *"yes");
}

#[test]
fn ref_error_as_sheet_prefix_propagates() {
    let mut model = new_empty_model();
    // A reference into a deleted sheet: the whole expression is #REF!
    model._set("A1", "=#REF!K$48");
    model._set("A2", "=A1*2");
    model.evaluate();

    assert_eq!(model._get_text("A1"), *"#REF!");
    assert_eq!(model._get_text("A2"), *"#REF!");
}
