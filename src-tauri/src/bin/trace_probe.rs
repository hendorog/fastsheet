//! trace_probe — load an .xls/.xlsx and dump the dependency tree
//! for a given cell. Validates the trace module without firing up
//! the GUI.
//!
//! Usage: trace_probe <file> <sheet>:<row>:<col>
//! Example: trace_probe Hetairos_BLUE.xls 0:4:7

use fastsheet_lib::{load_xls, load_xlsx_with_fallback};

fn main() {
    let args: Vec<String> = std::env::args().skip(1).collect();
    if args.len() < 2 {
        eprintln!("usage: trace_probe <file> <sheet>:<row>:<col>");
        std::process::exit(2);
    }
    let path = &args[0];
    let target = &args[1];
    let parts: Vec<&str> = target.split(':').collect();
    if parts.len() != 3 {
        eprintln!("target must be sheet:row:col (e.g. 0:4:7)");
        std::process::exit(2);
    }
    let s: u32 = parts[0].parse().expect("sheet idx");
    let r: i32 = parts[1].parse().expect("row");
    let c: i32 = parts[2].parse().expect("col");

    let is_xls = std::path::Path::new(path)
        .extension()
        .and_then(|x| x.to_str())
        .map(|e| e.eq_ignore_ascii_case("xls"))
        .unwrap_or(false);
    let mut model = if is_xls {
        load_xls(path).expect("load_xls").0
    } else {
        load_xlsx_with_fallback(path).expect("load_xlsx")
    };
    model.evaluate();

    let tree = fastsheet_lib::trace::trace(&model, s, r, c, None);
    print_tree(&tree, 0);
}

fn print_tree(n: &fastsheet_lib::trace::TraceNode, depth: usize) {
    let pad = "  ".repeat(depth);
    let marker = if n.is_error {
        " ⚠"
    } else if n.cycle {
        " ↺"
    } else if n.truncated {
        " …"
    } else {
        ""
    };
    let f = n
        .formula
        .as_ref()
        .map(|s| format!("  formula={s:?}"))
        .unwrap_or_default();
    let note = n
        .note
        .as_ref()
        .map(|s| format!("  ({s})"))
        .unwrap_or_default();
    println!(
        "{pad}[{}] {} = {:?}{}{}{}",
        n.kind, n.address, n.value, marker, f, note
    );
    for d in &n.deps {
        print_tree(d, depth + 1);
    }
}
