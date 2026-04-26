//! fastsheet — pseudo-spill UDFs.
//!
//! These mirror the semantics of a set of Apache-POI-based Java UDFs
//! (`{Unique,Sort,SortBlank,Transpose,Filter}Udf.java`). Originally written
//! so .xls files (which have no real array formulas) could fake spilling:
//! each cell in the spill range carries
//! the same `=MY*(...)` call with an absolute anchor reference; the function
//! computes its own (cell - anchor) offset and returns the matching matrix entry.
//!
//! In fastsheet the source workbook is xlsx, so the load preprocessor
//! (`expand_my_array_formulas` in xlsx_load.rs) walks every
//! `<f t="array" ref="X1:Yn">MY*(...)</f>` element and rewrites it as one
//! plain `<f>` per cell across the ref range. IronCalc then evaluates each
//! cell independently, calling our function with that cell's
//! `CellReferenceIndex` and the same args / anchor, so each cell picks the
//! right entry from the would-be spill.

use std::collections::BTreeMap;

use crate::calc_result::CalcResult;
use crate::expressions::parser::Node;
use crate::expressions::token::{Error, OpCompare, OpProduct};
use crate::expressions::types::CellReferenceIndex;
use crate::model::Model;

/// Internal canonical form for a single cell's value, simpler to sort/dedup
/// than CalcResult and friendly to the typed-key encoding in MYUNIQUE.
#[derive(Clone, Debug)]
enum Cv {
    Empty,
    Number(f64),
    String(String),
    Boolean(bool),
    Error(Error),
}

impl Cv {
    fn from_calc(c: CalcResult) -> Self {
        match c {
            CalcResult::Number(n) => Cv::Number(n),
            CalcResult::String(s) => Cv::String(s),
            CalcResult::Boolean(b) => Cv::Boolean(b),
            CalcResult::EmptyCell | CalcResult::EmptyArg => Cv::Empty,
            CalcResult::Error { error, .. } => Cv::Error(error),
            // Range/Array don't appear from evaluate_cell on a single cell;
            // treat defensively as empty.
            _ => Cv::Empty,
        }
    }

    fn into_calc(self, origin: CellReferenceIndex) -> CalcResult {
        match self {
            Cv::Empty => CalcResult::String(String::new()),
            Cv::Number(n) => CalcResult::Number(n),
            Cv::String(s) => CalcResult::String(s),
            Cv::Boolean(b) => CalcResult::Boolean(b),
            Cv::Error(error) => CalcResult::Error {
                error,
                origin,
                message: String::new(),
            },
        }
    }

    /// Treat blank, empty string, or whitespace-only string as "blank"
    /// (the Java SortBlankUdf used the same predicate).
    fn is_blankish(&self) -> bool {
        match self {
            Cv::Empty => true,
            Cv::String(s) => s.trim().is_empty(),
            Cv::Error(_) => true,
            _ => false,
        }
    }

    fn as_truthy(&self) -> bool {
        match self {
            Cv::Boolean(b) => *b,
            Cv::Number(n) => n.abs() > 1e-12,
            Cv::String(s) => {
                let t = s.trim();
                t.eq_ignore_ascii_case("TRUE") || t == "1"
            }
            _ => false,
        }
    }

    fn as_string(&self) -> String {
        match self {
            Cv::String(s) => s.clone(),
            Cv::Number(n) => format!("{n}"),
            Cv::Boolean(b) => if *b { "TRUE" } else { "FALSE" }.to_string(),
            _ => String::new(),
        }
    }

    /// Cross-type equality used by FILTER masks. Numbers and numeric strings
    /// match (Excel's coercion rules); empties match each other and only
    /// each other.
    fn loose_eq(&self, other: &Cv) -> bool {
        match (self, other) {
            (Cv::Number(a), Cv::Number(b)) => (a - b).abs() < f64::EPSILON,
            (Cv::String(a), Cv::String(b)) => a.eq_ignore_ascii_case(b),
            (Cv::Boolean(a), Cv::Boolean(b)) => a == b,
            (Cv::Empty, Cv::Empty) => true,
            (Cv::Number(a), Cv::String(b)) | (Cv::String(b), Cv::Number(a)) => {
                b.trim().parse::<f64>().map(|x| (x - a).abs() < f64::EPSILON).unwrap_or(false)
            }
            _ => false,
        }
    }
}

/// Read an integer arg (used for the trailing offset pair injected by
/// `replicate_my_array_formulas`). The Java original passed an absolute
/// cell reference for the anchor and computed `cell - anchor`; we instead
/// inject the offsets as literal numbers so each replicated cell has no
/// dependency on the anchor cell (which in our case IS the calling cell —
/// IronCalc would mark every spill cell as #CIRC! otherwise).
fn read_int_arg(
    model: &mut Model,
    arg: &Node,
    origin: CellReferenceIndex,
) -> Result<i32, CalcResult> {
    let v = model.evaluate_node_in_context(arg, origin);
    match Cv::from_calc(v) {
        Cv::Number(n) => Ok(n.round() as i32),
        Cv::Error(error) => Err(CalcResult::new_error(
            error,
            origin,
            "MY*: invalid offset arg".to_string(),
        )),
        _ => Err(CalcResult::new_error(
            Error::VALUE,
            origin,
            "MY*: offset arg must be numeric".to_string(),
        )),
    }
}

/// Read the trailing two args as (dr, dc).
fn offsets_of(
    model: &mut Model,
    args: &[Node],
    origin: CellReferenceIndex,
) -> Result<(i32, i32), CalcResult> {
    let n = args.len();
    if n < 2 {
        return Err(CalcResult::new_args_number_error(origin));
    }
    let dr = read_int_arg(model, &args[n - 2], origin)?;
    let dc = read_int_arg(model, &args[n - 1], origin)?;
    Ok((dr, dc))
}

/// Read a 2D matrix from a Range result. For scalar/1-cell results, returns
/// a 1×1 matrix; for errors, propagates.
fn read_matrix(
    model: &mut Model,
    arg: &Node,
    origin: CellReferenceIndex,
) -> Result<Vec<Vec<Cv>>, CalcResult> {
    let result = model.evaluate_node_in_context(arg, origin);
    matrix_from_calc(model, result, origin)
}

fn matrix_from_calc(
    model: &mut Model,
    result: CalcResult,
    origin: CellReferenceIndex,
) -> Result<Vec<Vec<Cv>>, CalcResult> {
    match result {
        CalcResult::Range { left, right } => {
            let h = (right.row - left.row + 1).max(0) as usize;
            let w = (right.column - left.column + 1).max(0) as usize;
            let mut out = Vec::with_capacity(h);
            for r in 0..h {
                let mut row = Vec::with_capacity(w);
                for c in 0..w {
                    let v = model.evaluate_cell(CellReferenceIndex {
                        sheet: left.sheet,
                        row: left.row + r as i32,
                        column: left.column + c as i32,
                    });
                    row.push(Cv::from_calc(v));
                }
                out.push(row);
            }
            Ok(out)
        }
        CalcResult::Error { error, message, .. } => Err(CalcResult::Error {
            error,
            origin,
            message,
        }),
        scalar => Ok(vec![vec![Cv::from_calc(scalar)]]),
    }
}

fn pick_or_blank(matrix: &[Vec<Cv>], r: i32, c: i32) -> Cv {
    if r < 0 || c < 0 {
        return Cv::Empty;
    }
    let (r, c) = (r as usize, c as usize);
    if r >= matrix.len() {
        return Cv::Empty;
    }
    let row = &matrix[r];
    if c >= row.len() {
        return Cv::Empty;
    }
    row[c].clone()
}

fn transpose(m: Vec<Vec<Cv>>) -> Vec<Vec<Cv>> {
    if m.is_empty() {
        return m;
    }
    let h = m.len();
    let w = m.iter().map(|r| r.len()).max().unwrap_or(0);
    let mut out: Vec<Vec<Cv>> = (0..w)
        .map(|_| Vec::with_capacity(h))
        .collect();
    for row in m {
        for (c, v) in row.into_iter().enumerate() {
            out[c].push(v);
        }
    }
    out
}

// ---------- typed-key encoding for MYUNIQUE (mirrors Java UniqueUdf) ----------

const SEP: char = '\u{0001}';

fn row_key(row: &[Cv]) -> String {
    let mut s = String::new();
    for (i, v) in row.iter().enumerate() {
        if i > 0 {
            s.push(SEP);
        }
        match v {
            Cv::Empty => s.push_str("X:"),
            Cv::Number(n) => {
                s.push_str("N:");
                s.push_str(&format!("{:016x}", n.to_bits()));
            }
            Cv::String(t) => {
                s.push_str("S:");
                // No need for base64 for our purposes — SEP is U+0001, which
                // we assume doesn't appear in real cell text.
                s.push_str(t);
            }
            Cv::Boolean(b) => {
                s.push_str(if *b { "B:1" } else { "B:0" });
            }
            Cv::Error(_) => s.push_str("E:"),
        }
    }
    s
}

// ---------- FILTER mask walker ----------

/// Strip wrapping ImplicitIntersection nodes — for our FILTER mask args we
/// always want the array-shaped underlying expression, never a single
/// intersected scalar.
fn unwrap_ii(node: &Node) -> &Node {
    let mut n = node;
    while let Node::ImplicitIntersection { child, .. } = n {
        n = child;
    }
    n
}

/// Evaluate a mask expression to a flat boolean vector aligned to the
/// data-row count. Recognises:
///   * a range/defined-name resolving to a Range — read element values as truthy
///   * RANGE = SCALAR / RANGE <> SCALAR  (loose comparison)
///   * SCALAR = RANGE / SCALAR <> RANGE  (mirror)
///   * (mask) * (mask) — element-wise AND
///   * (mask) + (mask) — element-wise OR
///   * a single scalar — broadcast to length `expected_len`
fn evaluate_mask(
    model: &mut Model,
    node: &Node,
    expected_len: usize,
    origin: CellReferenceIndex,
) -> Result<Vec<bool>, CalcResult> {
    let node = unwrap_ii(node);

    // (A) * (B) — AND
    if let Node::OpProductKind { kind: OpProduct::Times, left, right } = node {
        let a = evaluate_mask(model, left, expected_len, origin)?;
        let b = evaluate_mask(model, right, expected_len, origin)?;
        return Ok(a
            .into_iter()
            .zip(b.into_iter())
            .map(|(x, y)| x && y)
            .collect());
    }

    // A = B / A <> B
    if let Node::CompareKind { kind, left, right } = node {
        if matches!(kind, OpCompare::Equal | OpCompare::NonEqual) {
            let l = model.evaluate_node_in_context(left, origin);
            let r = model.evaluate_node_in_context(right, origin);
            let mask = compare_to_mask(model, l, r, origin)?;
            return Ok(if matches!(kind, OpCompare::NonEqual) {
                mask.into_iter().map(|b| !b).collect()
            } else {
                mask
            });
        }
    }

    // Bare range / defined name — truthy on each element.
    let result = model.evaluate_node_in_context(node, origin);
    if let CalcResult::Range { left, right } = result {
        let mut out = Vec::new();
        for r in left.row..=right.row {
            for c in left.column..=right.column {
                let v = Cv::from_calc(model.evaluate_cell(CellReferenceIndex {
                    sheet: left.sheet,
                    row: r,
                    column: c,
                }));
                out.push(v.as_truthy());
            }
        }
        return Ok(out);
    }
    if result.is_error() {
        if let CalcResult::Error { error, message, .. } = result {
            return Err(CalcResult::Error { error, origin, message });
        }
    }
    // Scalar — broadcast.
    let scalar = Cv::from_calc(result).as_truthy();
    Ok(vec![scalar; expected_len])
}

/// Helper used by `=` / `<>`: compare two CalcResult sides where typically
/// one is a Range and the other is a scalar.
fn compare_to_mask(
    model: &mut Model,
    l: CalcResult,
    r: CalcResult,
    origin: CellReferenceIndex,
) -> Result<Vec<bool>, CalcResult> {
    let mat_l = matrix_from_calc(model, l, origin)?;
    let mat_r = matrix_from_calc(model, r, origin)?;
    let flat_l: Vec<Cv> = mat_l.into_iter().flatten().collect();
    let flat_r: Vec<Cv> = mat_r.into_iter().flatten().collect();
    let n = flat_l.len().max(flat_r.len());
    let mut out = Vec::with_capacity(n);
    for i in 0..n {
        let a = if flat_l.len() == 1 { &flat_l[0] } else { &flat_l[i.min(flat_l.len() - 1)] };
        let b = if flat_r.len() == 1 { &flat_r[0] } else { &flat_r[i.min(flat_r.len() - 1)] };
        out.push(a.loose_eq(b));
    }
    Ok(out)
}

// ---------- comparator for MYSORT / MYSORTBLANK ----------

fn cmp_mixed(a: &Cv, b: &Cv) -> std::cmp::Ordering {
    use std::cmp::Ordering::*;
    let an = matches!(a, Cv::Number(_));
    let bn = matches!(b, Cv::Number(_));
    if an && bn {
        if let (Cv::Number(x), Cv::Number(y)) = (a, b) {
            return x.partial_cmp(y).unwrap_or(Equal);
        }
    }
    if an {
        return Less;
    }
    if bn {
        return Greater;
    }
    let abool = matches!(a, Cv::Boolean(_));
    let bbool = matches!(b, Cv::Boolean(_));
    if abool && bbool {
        if let (Cv::Boolean(x), Cv::Boolean(y)) = (a, b) {
            return x.cmp(y);
        }
    }
    a.as_string()
        .to_lowercase()
        .cmp(&b.as_string().to_lowercase())
}

fn cmp_with_blanks_last(a: &Cv, b: &Cv, sort_order: i32) -> std::cmp::Ordering {
    use std::cmp::Ordering::*;
    let ab = a.is_blankish();
    let bb = b.is_blankish();
    if ab && bb {
        return Equal;
    }
    if ab {
        return Greater;
    }
    if bb {
        return Less;
    }
    let cmp = cmp_mixed(a, b);
    if sort_order < 0 {
        cmp.reverse()
    } else {
        cmp
    }
}

// =========== fn_my_* implementations ===========

impl<'a> Model<'a> {
    pub(crate) fn fn_my_transpose(&mut self, args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        if args.len() < 3 {
            return CalcResult::new_args_number_error(cell);
        }
        let (dr, dc) = match offsets_of(self, args, cell) {
            Ok(o) => o,
            Err(e) => return e,
        };
        let data = match read_matrix(self, &args[0], cell) {
            Ok(m) => m,
            Err(e) => return e,
        };
        let t = transpose(data);
        pick_or_blank(&t, dr, dc).into_calc(cell)
    }

    pub(crate) fn fn_my_unique(&mut self, args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        if args.len() < 3 {
            return CalcResult::new_args_number_error(cell);
        }
        let (dr, dc) = match offsets_of(self, args, cell) {
            Ok(o) => o,
            Err(e) => return e,
        };
        let data = match read_matrix(self, &args[0], cell) {
            Ok(m) => m,
            Err(e) => return e,
        };
        // Optional middle args: byCol (after data, before offsets), exactlyOnce.
        // Total layout: data, [byCol], [exactlyOnce], dr, dc.
        let middle = &args[1..args.len() - 2];
        let by_col = middle
            .first()
            .map(|a| Cv::from_calc(self.evaluate_node_in_context(a, cell)).as_truthy())
            .unwrap_or(false);
        let exactly_once = middle
            .get(1)
            .map(|a| Cv::from_calc(self.evaluate_node_in_context(a, cell)).as_truthy())
            .unwrap_or(false);

        let oriented = if by_col { transpose(data) } else { data };
        // dedup row-wise.
        let mut counts: BTreeMap<String, usize> = BTreeMap::new();
        let mut order: Vec<String> = Vec::new();
        let mut row_for: BTreeMap<String, Vec<Cv>> = BTreeMap::new();
        for row in &oriented {
            let k = row_key(row);
            let n = counts.entry(k.clone()).or_insert(0);
            *n += 1;
            if *n == 1 {
                order.push(k.clone());
                row_for.insert(k, row.clone());
            }
        }
        let mut out: Vec<Vec<Cv>> = Vec::new();
        for k in order {
            if exactly_once && counts.get(&k).copied().unwrap_or(0) != 1 {
                continue;
            }
            if let Some(r) = row_for.remove(&k) {
                out.push(r);
            }
        }
        let result = if by_col { transpose(out) } else { out };
        pick_or_blank(&result, dr, dc).into_calc(cell)
    }

    pub(crate) fn fn_my_sort(&mut self, args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        if args.len() < 3 {
            return CalcResult::new_args_number_error(cell);
        }
        let (dr, dc) = match offsets_of(self, args, cell) {
            Ok(o) => o,
            Err(e) => return e,
        };
        let data = match read_matrix(self, &args[0], cell) {
            Ok(m) => m,
            Err(e) => return e,
        };
        // Layout: data, [sortIndex], [sortOrder], [byCol], dr, dc.
        let middle = &args[1..args.len() - 2];
        let sort_index = middle
            .first()
            .map(|a| match Cv::from_calc(self.evaluate_node_in_context(a, cell)) {
                Cv::Number(n) => n as i32,
                _ => 1,
            })
            .unwrap_or(1);
        let sort_order = middle
            .get(1)
            .map(|a| match Cv::from_calc(self.evaluate_node_in_context(a, cell)) {
                Cv::Number(n) => n as i32,
                _ => 1,
            })
            .unwrap_or(1);
        let by_col = middle
            .get(2)
            .map(|a| Cv::from_calc(self.evaluate_node_in_context(a, cell)).as_truthy())
            .unwrap_or(false);

        let mut oriented = if by_col { transpose(data) } else { data };
        let key = (sort_index.max(1) - 1) as usize;
        oriented.sort_by(|r1, r2| {
            let a = r1.get(key).cloned().unwrap_or(Cv::Empty);
            let b = r2.get(key).cloned().unwrap_or(Cv::Empty);
            let cmp = cmp_mixed(&a, &b);
            if sort_order < 0 {
                cmp.reverse()
            } else {
                cmp
            }
        });
        let result = if by_col { transpose(oriented) } else { oriented };
        pick_or_blank(&result, dr, dc).into_calc(cell)
    }

    pub(crate) fn fn_my_sort_blank(&mut self, args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        if args.len() < 3 {
            return CalcResult::new_args_number_error(cell);
        }
        let (dr, dc) = match offsets_of(self, args, cell) {
            Ok(o) => o,
            Err(e) => return e,
        };
        let data = match read_matrix(self, &args[0], cell) {
            Ok(m) => m,
            Err(e) => return e,
        };
        // Layout: data, [sortIndex], [sortOrder], [byCol], [skipBlanks], dr, dc.
        let middle = &args[1..args.len() - 2];
        let sort_index = middle
            .first()
            .map(|a| match Cv::from_calc(self.evaluate_node_in_context(a, cell)) {
                Cv::Number(n) => (n as i32).max(1),
                _ => 1,
            })
            .unwrap_or(1);
        let sort_order = middle
            .get(1)
            .map(|a| match Cv::from_calc(self.evaluate_node_in_context(a, cell)) {
                Cv::Number(n) => n as i32,
                _ => 1,
            })
            .unwrap_or(1);
        let by_col = middle
            .get(2)
            .map(|a| Cv::from_calc(self.evaluate_node_in_context(a, cell)).as_truthy())
            .unwrap_or(false);
        let skip_blanks = middle
            .get(3)
            .map(|a| Cv::from_calc(self.evaluate_node_in_context(a, cell)).as_truthy())
            .unwrap_or(true);

        let mut oriented = if by_col { transpose(data) } else { data };
        if skip_blanks {
            oriented.retain(|row| row.iter().any(|v| !v.is_blankish()));
        }
        let width = oriented.iter().map(|r| r.len()).max().unwrap_or(1);
        let key = (sort_index - 1).min(width.saturating_sub(1) as i32) as usize;
        oriented.sort_by(|r1, r2| {
            let a = r1.get(key).cloned().unwrap_or(Cv::Empty);
            let b = r2.get(key).cloned().unwrap_or(Cv::Empty);
            cmp_with_blanks_last(&a, &b, sort_order)
        });
        let result = if by_col { transpose(oriented) } else { oriented };
        pick_or_blank(&result, dr, dc).into_calc(cell)
    }

    pub(crate) fn fn_my_filter(&mut self, args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        // Layout: data, mask, [if_empty], dr, dc.
        if args.len() < 4 {
            return CalcResult::new_args_number_error(cell);
        }
        let (dr, dc) = match offsets_of(self, args, cell) {
            Ok(o) => o,
            Err(e) => return e,
        };
        let data = match read_matrix(self, &args[0], cell) {
            Ok(m) => m,
            Err(e) => return e,
        };
        let middle = &args[1..args.len() - 2]; // mask, [if_empty]
        let mask_node = match middle.first() {
            Some(n) => n,
            None => return CalcResult::new_args_number_error(cell),
        };
        let if_empty = middle
            .get(1)
            .map(|a| Cv::from_calc(self.evaluate_node_in_context(a, cell)));

        let height = data.len();
        let mask = match evaluate_mask(self, mask_node, height, cell) {
            Ok(m) => m,
            Err(e) => return e,
        };

        let kept: Vec<Vec<Cv>> = if mask.len() == height {
            data.into_iter()
                .zip(mask.into_iter())
                .filter_map(|(row, keep)| if keep { Some(row) } else { None })
                .collect()
        } else {
            // Java fallback: row-major flatten; keep row when ANY element's
            // mask is true. Used when mask shape doesn't match height.
            let width = data.first().map(|r| r.len()).unwrap_or(0).max(1);
            data.into_iter()
                .enumerate()
                .filter_map(|(r, row)| {
                    let any = (0..width)
                        .map(|c| r * width + c)
                        .any(|i| mask.get(i).copied().unwrap_or(false));
                    if any { Some(row) } else { None }
                })
                .collect()
        };

        if kept.is_empty() {
            if dr == 0 && dc == 0 {
                if let Some(v) = if_empty {
                    return v.into_calc(cell);
                }
            }
            return Cv::Empty.into_calc(cell);
        }
        pick_or_blank(&kept, dr, dc).into_calc(cell)
    }
}
