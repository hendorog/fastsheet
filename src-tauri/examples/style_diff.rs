// Compare resolved per-cell styles between two xls files (original
// vs round-tripped) through load_xls: fill colors, font colors, font
// attributes, number formats. Quantifies style fidelity losses the
// value-diffing save_probe can't see (e.g. XFEXT true colors
// collapsing to palette approximations).
//
//   cargo run --release --features dev-bins --example style_diff -- a.xls b.xls

use std::collections::HashMap;

fn main() {
    let a_path = std::env::args().nth(1).expect("usage: style_diff <a.xls> <b.xls>");
    let b_path = std::env::args().nth(2).expect("usage: style_diff <a.xls> <b.xls>");
    let (a, _, _, _) = fastsheet_lib::load_xls(&a_path).expect("load a");
    let (b, _, _, _) = fastsheet_lib::load_xls(&b_path).expect("load b");

    let mut total = 0u64;
    let mut diffs: HashMap<&'static str, u64> = HashMap::new();
    let mut samples: Vec<String> = Vec::new();
    for (s, ws) in a.workbook.worksheets.iter().enumerate() {
        for (row, cols) in &ws.sheet_data {
            for col in cols.keys() {
                total += 1;
                let sa = a.get_style_for_cell(s as u32, *row, *col).ok();
                let sb = b.get_style_for_cell(s as u32, *row, *col).ok();
                let (Some(sa), Some(sb)) = (sa, sb) else {
                    *diffs.entry("missing").or_default() += 1;
                    continue;
                };
                let checks: [(&'static str, bool); 6] = [
                    ("fill", sa.fill != sb.fill),
                    ("font-color", sa.font.color != sb.font.color),
                    ("font-face", (sa.font.name.clone(), sa.font.sz, sa.font.b, sa.font.i, sa.font.u)
                        != (sb.font.name.clone(), sb.font.sz, sb.font.b, sb.font.i, sb.font.u)),
                    ("numfmt", sa.num_fmt != sb.num_fmt),
                    ("border", sa.border != sb.border),
                    ("alignment", sa.alignment != sb.alignment),
                ];
                for (kind, is_diff) in checks {
                    if is_diff {
                        *diffs.entry(kind).or_default() += 1;
                        if samples.len() < 12 {
                            let detail = match kind {
                                "fill" => format!("a={:?} b={:?}", sa.fill, sb.fill),
                                "border" => format!("a={:?} b={:?}", sa.border, sb.border),
                                _ => String::new(),
                            };
                            samples.push(format!(
                                "{}!r{}c{} {kind} {detail}",
                                ws.name, row, col
                            ));
                        }
                    }
                }
            }
        }
    }
    println!("cells={total}");
    let mut keys: Vec<_> = diffs.iter().collect();
    keys.sort();
    for (k, v) in keys {
        println!("  {k}: {v}");
    }
    if diffs.is_empty() {
        println!("  styles identical");
    }
    for s in samples {
        println!("  sample: {s}");
    }
}
