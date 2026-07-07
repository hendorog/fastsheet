#![allow(clippy::unwrap_used)]

// Excel treats a blank lookup value as the number 0 in MATCH (found via a
// real-world pricing workbook: MATCH(IF(ISTEXT(B18),VALUE(B18),B18),...)
// with a blank B18 must behave like MATCH(0,...)).

use crate::test::util::new_empty_model;

#[test]
fn match_blank_lookup_value_is_zero() {
    let mut model = new_empty_model();
    for (i, v) in ["0", "9", "12"].iter().enumerate() {
        model._set(&format!("{}1", ["M", "N", "O"][i]), v);
    }
    // B18 left blank
    model._set("A1", "=MATCH(B18,M1:O1,1)");
    model._set("A2", "=MATCH(IF(ISTEXT(B18),VALUE(B18),B18),M1:O1,1)");
    model._set("A3", "=MATCH(B18,M1:O1,0)");
    model.evaluate();

    assert_eq!(model._get_text("A1"), *"1");
    assert_eq!(model._get_text("A2"), *"1");
    assert_eq!(model._get_text("A3"), *"1");
}
