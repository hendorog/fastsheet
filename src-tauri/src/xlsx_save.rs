use std::collections::{BTreeMap, HashMap, HashSet};

use ironcalc::base::types::{Col, Row};

use crate::state::LoadedFile;
use crate::util::{col_letter_i, parse_attr_val};

/// Per-sheet layout state captured at save time so save_preserving can
/// project it back into the xlsx XML alongside the dirty cells. Without
/// this, drag-resized widths / heights, /Worksheet/Column/Hide,
/// /Worksheet/Titles, etc. mutate the in-memory IronCalc model but are
/// silently dropped on reload.
pub(crate) struct SheetLayoutSnapshot {
    pub cols: Vec<Col>,
    pub hidden_cols: HashSet<i32>,
    pub rows: Vec<Row>,
    pub frozen_rows: i32,
    pub frozen_cols: i32,
}

/// One cell edit, classified into the kind of XML we'll write.
#[derive(Debug)]
enum CellWrite {
    /// Numeric literal. Use Excel's general format, no `t` attr.
    Number(f64),
    /// Formula like `=SUM(A1:A5)`. We strip the leading `=` and write
    /// `<f>...</f>` (no cached value — Excel recalcs on open).
    Formula(String),
    /// Inline string. Preserves shared-string table by NOT touching it.
    InlineString(String),
    /// Clear the cell entirely (delete the `<c>` element).
    Empty,
}

fn classify_input(s: &str) -> CellWrite {
    if s.is_empty() {
        return CellWrite::Empty;
    }
    if let Some(stripped) = s.strip_prefix('=') {
        return CellWrite::Formula(stripped.to_string());
    }
    if let Ok(n) = s.parse::<f64>() {
        return CellWrite::Number(n);
    }
    CellWrite::InlineString(s.to_string())
}

fn xml_escape(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    for c in s.chars() {
        match c {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' => out.push_str("&quot;"),
            '\'' => out.push_str("&apos;"),
            _ => out.push(c),
        }
    }
    out
}

/// Render a `<c>` element. Empty values render as the empty string so the
/// caller's range-replacement deletes the cell entirely.
fn build_cell_xml(addr: &str, style: Option<&str>, value: &CellWrite) -> String {
    let style_attr = style
        .map(|s| format!(r#" s="{s}""#))
        .unwrap_or_default();
    match value {
        CellWrite::Empty => String::new(),
        CellWrite::Number(n) => {
            // Format without scientific notation for typical values; trim
            // trailing zero on integers ("42" not "42.0").
            let v = if n.fract() == 0.0 && n.abs() < 1e15 {
                format!("{}", *n as i64)
            } else {
                format!("{}", n)
            };
            format!(r#"<c r="{addr}"{style_attr}><v>{v}</v></c>"#)
        }
        CellWrite::Formula(f) => {
            format!(
                r#"<c r="{addr}"{style_attr}><f>{}</f></c>"#,
                xml_escape(f)
            )
        }
        CellWrite::InlineString(s) => {
            format!(
                r#"<c r="{addr}"{style_attr} t="inlineStr"><is><t xml:space="preserve">{}</t></is></c>"#,
                xml_escape(s)
            )
        }
    }
}

/// Locate `<c r="ADDR" ...>...</c>` (or self-closing form) in worksheet XML.
/// Returns (start_offset, end_offset, opening_tag_text) where the offsets
/// span the entire `<c>` element.
fn find_cell_in_xml<'a>(xml: &'a str, addr: &str) -> Option<(usize, usize, &'a str)> {
    // Quoted address so "A1" doesn't match "A10".
    let needle = format!(r#"<c r="{addr}""#);
    let start = xml.find(&needle)?;
    let bytes = xml.as_bytes();
    let mut tag_end = start + needle.len();
    while tag_end < bytes.len() && bytes[tag_end] != b'>' {
        tag_end += 1;
    }
    if tag_end >= bytes.len() {
        return None;
    }
    let opening = &xml[start..=tag_end];
    if tag_end > 0 && bytes[tag_end - 1] == b'/' {
        return Some((start, tag_end + 1, opening));
    }
    let after = tag_end + 1;
    let close = xml[after..].find("</c>")?;
    Some((start, after + close + 4, opening))
}

/// Find the insertion point inside an existing `<row r="N">…</row>` —
/// returns the offset of the `</row>` close tag, where new cells should be
/// inserted (sort order isn't required by the spec).
fn find_row_close_in_xml(xml: &str, row: i32) -> Option<usize> {
    let needle = format!(r#"<row r="{row}""#);
    let start = xml.find(&needle)?;
    let bytes = xml.as_bytes();
    let mut tag_end = start + needle.len();
    while tag_end < bytes.len() && bytes[tag_end] != b'>' {
        tag_end += 1;
    }
    if tag_end >= bytes.len() {
        return None;
    }
    if tag_end > 0 && bytes[tag_end - 1] == b'/' {
        // Self-closing — empty row. Treat as no row exists for our purposes.
        return None;
    }
    let after = tag_end + 1;
    xml[after..].find("</row>").map(|p| after + p)
}

fn extract_style_attr(opening_tag: &str) -> Option<String> {
    let pos = opening_tag.find(" s=\"")?;
    let after = &opening_tag[pos + 4..];
    let end = after.find('"')?;
    Some(after[..end].to_string())
}

/// Patch a worksheet XML string with the given dirty cells.
/// `dirty` maps (row, col) → user input string (1-indexed addresses).
fn patch_sheet_xml(xml: &str, dirty: &HashMap<(i32, i32), String>) -> String {
    if dirty.is_empty() {
        return xml.to_string();
    }
    // Group by row in ascending order so any new <row> we synthesise stays
    // sorted (Excel tolerates out-of-order rows but stays cleaner).
    let mut by_row: BTreeMap<i32, Vec<(i32, &String)>> = BTreeMap::new();
    for ((r, c), v) in dirty {
        by_row.entry(*r).or_default().push((*c, v));
    }
    for cells in by_row.values_mut() {
        cells.sort_by_key(|(c, _)| *c);
    }

    let mut edits: Vec<(usize, usize, String)> = Vec::new();
    for (row, cells) in &by_row {
        let row_close = find_row_close_in_xml(xml, *row);
        match row_close {
            Some(insert_at) => {
                for (col, value) in cells {
                    let addr = format!("{}{}", col_letter_i(*col), row);
                    let cw = classify_input(value);
                    if let Some((s, e, opening)) = find_cell_in_xml(xml, &addr) {
                        let style = extract_style_attr(opening);
                        edits.push((s, e, build_cell_xml(&addr, style.as_deref(), &cw)));
                    } else if !matches!(cw, CellWrite::Empty) {
                        edits.push((insert_at, insert_at, build_cell_xml(&addr, None, &cw)));
                    }
                }
            }
            None => {
                // No such row yet — synthesise it before </sheetData>.
                let mut row_xml = format!(r#"<row r="{row}">"#);
                let mut any = false;
                for (col, value) in cells {
                    let cw = classify_input(value);
                    if matches!(cw, CellWrite::Empty) {
                        continue;
                    }
                    let addr = format!("{}{}", col_letter_i(*col), row);
                    row_xml.push_str(&build_cell_xml(&addr, None, &cw));
                    any = true;
                }
                row_xml.push_str("</row>");
                if any {
                    if let Some(sd) = xml.find("</sheetData>") {
                        edits.push((sd, sd, row_xml));
                    }
                }
            }
        }
    }

    edits.sort_by(|a, b| b.0.cmp(&a.0));
    let mut out = xml.to_string();
    for (s, e, repl) in edits {
        out.replace_range(s..e, &repl);
    }
    out
}

/// Parse `xl/workbook.xml` + `xl/_rels/workbook.xml.rels` from a loaded
/// xlsx zip and return `[sheet_idx → zip entry path]` ordered to match
/// IronCalc's sheet enumeration (which follows `<sheets>` order).
pub(crate) fn extract_sheet_paths(zip_bytes: &[u8]) -> Result<Vec<String>, String> {
    use std::io::{Cursor, Read};
    use zip::ZipArchive;
    let mut zin = ZipArchive::new(Cursor::new(zip_bytes)).map_err(|e| e.to_string())?;
    let mut workbook_xml = String::new();
    zin.by_name("xl/workbook.xml")
        .map_err(|e| e.to_string())?
        .read_to_string(&mut workbook_xml)
        .map_err(|e| e.to_string())?;
    let mut rels_xml = String::new();
    zin.by_name("xl/_rels/workbook.xml.rels")
        .map_err(|e| e.to_string())?
        .read_to_string(&mut rels_xml)
        .map_err(|e| e.to_string())?;

    // Walk `<sheet ... r:id="rIdN"/>` in workbook.xml (order = display order).
    let mut rids: Vec<String> = Vec::new();
    let mut pos = 0;
    while let Some(rel) = workbook_xml[pos..].find("<sheet ") {
        let abs = pos + rel + 7;
        let close = workbook_xml[abs..]
            .find('>')
            .map(|p| abs + p)
            .unwrap_or(workbook_xml.len());
        let tag = &workbook_xml[abs..close];
        if let Some(id) = parse_attr_val(tag, "r:id").or_else(|| parse_attr_val(tag, "r:Id")) {
            rids.push(id);
        }
        pos = close;
    }

    // rId → Target from the rels file.
    let mut rid_map: HashMap<String, String> = HashMap::new();
    let mut pos = 0;
    while let Some(rel) = rels_xml[pos..].find("<Relationship") {
        let abs = pos + rel;
        let close = rels_xml[abs..]
            .find('>')
            .map(|p| abs + p)
            .unwrap_or(rels_xml.len());
        let tag = &rels_xml[abs..close];
        let id = parse_attr_val(tag, "Id");
        let tgt = parse_attr_val(tag, "Target");
        if let (Some(id), Some(t)) = (id, tgt) {
            rid_map.insert(id, t);
        }
        pos = close;
    }

    let mut out = Vec::with_capacity(rids.len());
    for rid in rids {
        if let Some(t) = rid_map.get(&rid) {
            // Targets in workbook.xml.rels are relative to xl/, but some
            // writers prefix `/` for absolute. Normalise.
            let path = if let Some(stripped) = t.strip_prefix('/') {
                stripped.to_string()
            } else {
                format!("xl/{}", t)
            };
            out.push(path);
        }
    }
    Ok(out)
}

/// Save by patching the original xlsx bytes in place. With no dirty cells
/// AND no layout changes we short-circuit and write the original bytes
/// verbatim — Excel is picky about subtle zip metadata (extra fields,
/// compression hints, central directory ordering) and even a no-op
/// rewrite can trigger its repair dialog. Otherwise we walk the zip
/// preserving each entry's original compression method so we change as
/// little as possible.
pub(crate) fn save_preserving(
    loaded: &LoadedFile,
    dirty: &HashMap<(u32, i32, i32), String>,
    layouts: &HashMap<u32, SheetLayoutSnapshot>,
    target_path: &str,
) -> Result<(), String> {
    use std::io::{Cursor, Read, Write};
    use zip::write::FileOptions;
    use zip::{ZipArchive, ZipWriter};

    if dirty.is_empty() && layouts.is_empty() {
        let _ = std::fs::remove_file(target_path);
        std::fs::write(target_path, &loaded.bytes).map_err(|e| e.to_string())?;
        return Ok(());
    }

    let mut by_sheet: HashMap<u32, HashMap<(i32, i32), String>> = HashMap::new();
    for ((s, r, c), v) in dirty {
        by_sheet.entry(*s).or_default().insert((*r, *c), v.clone());
    }

    let mut zin = ZipArchive::new(Cursor::new(&loaded.bytes)).map_err(|e| e.to_string())?;
    let mut buf = Cursor::new(Vec::<u8>::new());
    {
        let mut zout = ZipWriter::new(&mut buf);
        for i in 0..zin.len() {
            let mut entry = zin.by_index(i).map_err(|e| e.to_string())?;
            let name = entry.name().to_string();
            let entry_method = entry.compression();
            let mut content = Vec::new();
            entry.read_to_end(&mut content).map_err(|e| e.to_string())?;

            if let Some(sheet_idx) = loaded
                .sheet_paths
                .iter()
                .position(|p| p == &name)
                .map(|p| p as u32)
            {
                let sheet_dirty = by_sheet.get(&sheet_idx);
                let layout = layouts.get(&sheet_idx);
                if sheet_dirty.is_some() || layout.is_some() {
                    if let Ok(s) = std::str::from_utf8(&content) {
                        let mut patched = s.to_string();
                        if let Some(d) = sheet_dirty {
                            patched = patch_sheet_xml(&patched, d);
                        }
                        if let Some(l) = layout {
                            patched = patch_rows_in_xml(&patched, &l.rows);
                            patched = patch_cols_in_xml(&patched, &l.cols, &l.hidden_cols);
                            patched = patch_sheet_view_in_xml(&patched, l.frozen_rows, l.frozen_cols);
                        }
                        content = patched.into_bytes();
                    }
                }
            }

            let opts = FileOptions::default().compression_method(entry_method);
            zout.start_file(name, opts).map_err(|e| e.to_string())?;
            zout.write_all(&content).map_err(|e| e.to_string())?;
        }
        zout.finish().map_err(|e| e.to_string())?;
    }
    let new_bytes = buf.into_inner();
    let _ = std::fs::remove_file(target_path);
    std::fs::write(target_path, &new_bytes).map_err(|e| e.to_string())?;
    Ok(())
}

/// Rewrite the `<cols>` block to reflect IronCalc's worksheet.cols (widths
/// + custom-flag) merged with our hidden_cols side-channel. Adjacent cols
/// with identical attrs collapse into a single range entry, mirroring how
/// xlsx writers (and Excel itself) emit the section.
fn patch_cols_in_xml(xml: &str, cols: &[Col], hidden_cols: &HashSet<i32>) -> String {
    let block = render_cols_block(cols, hidden_cols);
    if let Some(start) = xml.find("<cols>") {
        if let Some(end_rel) = xml[start..].find("</cols>") {
            let end = start + end_rel + "</cols>".len();
            let mut out = xml.to_string();
            out.replace_range(start..end, &block);
            return out;
        }
    }
    if let Some(start) = xml.find("<cols/>") {
        let mut out = xml.to_string();
        out.replace_range(start..start + "<cols/>".len(), &block);
        return out;
    }
    // No existing section. Insert before <sheetData>.
    if !block.is_empty() {
        if let Some(sd) = xml.find("<sheetData") {
            let mut out = xml.to_string();
            out.insert_str(sd, &block);
            return out;
        }
    }
    xml.to_string()
}

fn render_cols_block(cols: &[Col], hidden_cols: &HashSet<i32>) -> String {
    // Default IronCalc width in chars for hidden-only cols that have no
    // explicit width. Ratio matches DEFAULT_COLUMN_WIDTH / COLUMN_WIDTH_FACTOR.
    const DEFAULT_W: f64 = 125.0 / 12.0;
    let mut expanded: BTreeMap<i32, (f64, bool, Option<i32>)> = BTreeMap::new();
    for c in cols {
        for i in c.min..=c.max {
            expanded.insert(i, (c.width, c.custom_width, c.style));
        }
    }
    for &i in hidden_cols {
        expanded.entry(i).or_insert((DEFAULT_W, false, None));
    }
    if expanded.is_empty() {
        return String::new();
    }

    let mut out = String::from("<cols>");
    let mut iter = expanded.into_iter().peekable();
    while let Some((min_col, attrs)) = iter.next() {
        let mut max_col = min_col;
        let hidden = hidden_cols.contains(&min_col);
        // Coalesce adjacent identical columns into a single range entry.
        while let Some(&(next, next_attrs)) = iter.peek() {
            if next != max_col + 1 || next_attrs != attrs {
                break;
            }
            if hidden_cols.contains(&next) != hidden {
                break;
            }
            max_col = next;
            iter.next();
        }
        emit_col_entry(&mut out, min_col, max_col, attrs, hidden);
    }
    out.push_str("</cols>");
    out
}

fn emit_col_entry(
    out: &mut String,
    min: i32,
    max: i32,
    attrs: (f64, bool, Option<i32>),
    hidden: bool,
) {
    let (width, custom, style) = attrs;
    out.push_str(&format!(r#"<col min="{}" max="{}" width="{}""#, min, max, width));
    if custom {
        out.push_str(r#" customWidth="1""#);
    }
    if hidden {
        out.push_str(r#" hidden="1""#);
    }
    if let Some(s) = style {
        out.push_str(&format!(r#" style="{}""#, s));
    }
    out.push_str("/>");
}

/// Rewrite the open-tag attrs of every `<row r="N">` mentioned in `rows`.
/// Worksheet.rows only stores rows with non-default attributes (height,
/// hidden, custom_format, ...); rows that don't appear keep the xlsx's
/// existing tags untouched. New rows that exist in the snapshot but not
/// in the XML at all are skipped (their attributes only matter once
/// the row gets a cell, which goes through patch_sheet_xml).
fn patch_rows_in_xml(xml: &str, rows: &[Row]) -> String {
    let mut out = xml.to_string();
    for r in rows {
        let needle = format!(r#"<row r="{}""#, r.r);
        let Some(start) = out.find(&needle) else { continue };
        let bytes = out.as_bytes();
        let mut tag_end = start + needle.len();
        let self_closing;
        loop {
            if tag_end >= bytes.len() {
                self_closing = false;
                break;
            }
            if bytes[tag_end] == b'>' {
                self_closing = tag_end > 0 && bytes[tag_end - 1] == b'/';
                break;
            }
            tag_end += 1;
        }
        if tag_end >= bytes.len() {
            continue;
        }
        let new_tag = render_row_open_tag(r, self_closing);
        out.replace_range(start..=tag_end, &new_tag);
    }
    out
}

fn render_row_open_tag(r: &Row, self_closing: bool) -> String {
    // Matches Excel's typical attr order. We always emit `r=...`; height
    // and hidden are conditional. customFormat / customHeight follow when
    // the corresponding attribute is present.
    let mut s = format!(r#"<row r="{}""#, r.r);
    if r.custom_height || r.hidden {
        // `Row::height` is already in points (xlsx native units) — IronCalc's
        // loader stores `ht="N"` directly into the field without scaling.
        // `get_row_height` later multiplies by ROW_HEIGHT_FACTOR (=2) for
        // its API consumers, but we're writing the raw stored value back to
        // disk, so no scaling here. (A previous version multiplied by 2
        // anyway, which doubled every row's height on each save+reload.)
        s.push_str(&format!(r#" ht="{}""#, r.height));
        if r.custom_height {
            s.push_str(r#" customHeight="1""#);
        }
    }
    if r.hidden {
        s.push_str(r#" hidden="1""#);
    }
    if r.custom_format {
        s.push_str(r#" s="{}" customFormat="1""#);
    }
    s.push_str(if self_closing { "/>" } else { ">" });
    s
}

/// Add or update the `<pane .../>` child inside `<sheetView>` to reflect
/// frozen rows/cols. Removes the pane element entirely when both are 0.
fn patch_sheet_view_in_xml(xml: &str, frozen_rows: i32, frozen_cols: i32) -> String {
    let Some(sv_start) = xml.find("<sheetView") else { return xml.to_string() };
    // End of the sheetView container — either self-closing or </sheetView>.
    let after_open = match xml[sv_start..].find('>') {
        Some(p) => sv_start + p + 1,
        None => return xml.to_string(),
    };
    let opening = &xml[sv_start..after_open];
    let self_closing = opening.ends_with("/>");
    let sv_end = if self_closing {
        after_open
    } else {
        match xml[after_open..].find("</sheetView>") {
            Some(p) => after_open + p + "</sheetView>".len(),
            None => return xml.to_string(),
        }
    };

    let want_pane = frozen_rows > 0 || frozen_cols > 0;

    if !want_pane {
        // Drop any existing <pane .../> from inside the sheetView body.
        let body = &xml[after_open..sv_end];
        if let Some(p_start) = body.find("<pane") {
            if let Some(p_end_rel) = body[p_start..].find("/>") {
                let abs_start = after_open + p_start;
                let abs_end = after_open + p_start + p_end_rel + 2;
                let mut out = xml.to_string();
                out.replace_range(abs_start..abs_end, "");
                return out;
            }
        }
        return xml.to_string();
    }

    let pane = render_pane(frozen_rows, frozen_cols);
    // If the sheetView is self-closing, expand it so we can add a child.
    if self_closing {
        let new_opening = format!("{}>{}</sheetView>", &opening[..opening.len() - 2], pane);
        let mut out = xml.to_string();
        out.replace_range(sv_start..sv_end, &new_opening);
        return out;
    }
    // Replace any existing <pane .../>; otherwise insert just after the
    // opening tag.
    let body = &xml[after_open..sv_end - "</sheetView>".len()];
    if let Some(p_start) = body.find("<pane") {
        if let Some(p_end_rel) = body[p_start..].find("/>") {
            let abs_start = after_open + p_start;
            let abs_end = after_open + p_start + p_end_rel + 2;
            let mut out = xml.to_string();
            out.replace_range(abs_start..abs_end, &pane);
            return out;
        }
    }
    let mut out = xml.to_string();
    out.insert_str(after_open, &pane);
    out
}

fn render_pane(frozen_rows: i32, frozen_cols: i32) -> String {
    let top_left = format!("{}{}", col_letter_i(frozen_cols + 1), frozen_rows + 1);
    let active = match (frozen_rows > 0, frozen_cols > 0) {
        (true, true) => "bottomRight",
        (true, false) => "bottomLeft",
        (false, true) => "topRight",
        (false, false) => "topLeft",
    };
    let mut s = String::from("<pane");
    if frozen_cols > 0 {
        s.push_str(&format!(r#" xSplit="{}""#, frozen_cols));
    }
    if frozen_rows > 0 {
        s.push_str(&format!(r#" ySplit="{}""#, frozen_rows));
    }
    s.push_str(&format!(
        r#" topLeftCell="{}" activePane="{}" state="frozen"/>"#,
        top_left, active
    ));
    s
}
