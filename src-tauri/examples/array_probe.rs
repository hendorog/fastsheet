// Inventory the ARRAY (array-formula) records in an .xls file:
// spill range, anchor, decoded formula text. Dev diagnostic only.
//
//   cargo run --release --features dev-bins --example array_probe -- ../fixtures/neutral.xls

use calamine::{open_workbook, Reader, Xls};
use fastsheet_lib::xls_biff::{decode_array_formula, scan_xls_shape};
use std::io::Read;

fn col_letter(mut c: i32) -> String {
    let mut s = String::new();
    while c > 0 {
        let rem = ((c - 1) % 26) as u8;
        s.insert(0, (b'A' + rem) as char);
        c = (c - 1) / 26;
    }
    s
}

fn main() {
    let path = std::env::args().nth(1).expect("usage: array_probe <file.xls>");
    let mut bytes = Vec::new();
    std::fs::File::open(&path)
        .expect("open")
        .read_to_end(&mut bytes)
        .expect("read");
    let shape = scan_xls_shape(&bytes);

    let wb: Xls<_> = open_workbook(&path).expect("calamine open");
    let defined_names: Vec<String> = wb.defined_names().iter().map(|(n, _)| n.clone()).collect();

    println!(
        "{} ARRAY ranges, {} anchor rgce entries",
        shape.array_ranges.len(),
        shape.array_formulas.len()
    );
    for (sheet, r1, r2, c1, c2) in &shape.array_ranges {
        let sheet_name = shape
            .biff_sheet_names
            .get(*sheet as usize)
            .cloned()
            .unwrap_or_else(|| format!("sheet{sheet}"));
        let range = format!(
            "{}{}:{}{}",
            col_letter(*c1),
            r1,
            col_letter(*c2),
            r2
        );
        let decoded = shape
            .array_formulas
            .get(&(*sheet, *r1, *c1))
            .and_then(|rgce| {
                decode_array_formula(
                    rgce,
                    &defined_names,
                    &shape.xti_table,
                    &shape.biff_sheet_names,
                    &shape.extern_names,
                )
            })
            .unwrap_or_else(|| "<undecodable>".to_string());
        let rgce_len = shape
            .array_formulas
            .get(&(*sheet, *r1, *c1))
            .map(|b| b.len())
            .unwrap_or(0);
        println!("{sheet_name}!{range}  ({}x{}, rgce {} bytes)  ={decoded}",
            r2 - r1 + 1, c2 - c1 + 1, rgce_len);
    }
}
