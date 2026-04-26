use ironcalc::base::Model;
use ironcalc::import::load_from_xlsx_bytes;

use crate::xlsx_save::extract_sheet_paths;

/// Detect a custom `<colors><indexedColors>` palette in `xl/styles.xml`
/// and rewrite every `<fgColor|bgColor|color indexed="N"/>` to its literal
/// RGB equivalent. IronCalc 0.7.x ignores the custom palette, so indexed
/// colours come back as the legacy-default RGB even when the file
/// overrides them (common in spreadsheets originally authored in older
/// Excel versions). Passes through untouched if there's no custom palette.
fn preprocess_xlsx_custom_palette(input: &[u8]) -> Result<Vec<u8>, String> {
    use std::io::{Cursor, Read, Write};
    use zip::write::FileOptions;
    use zip::{ZipArchive, ZipWriter};

    let mut styles_xml = String::new();
    {
        let mut zin = ZipArchive::new(Cursor::new(input)).map_err(|e| e.to_string())?;
        match zin.by_name("xl/styles.xml") {
            Ok(mut e) => e.read_to_string(&mut styles_xml).map_err(|e| e.to_string())?,
            Err(_) => return Ok(input.to_vec()),
        };
    }

    // Locate the <indexedColors>…</indexedColors> section and parse each
    // <rgbColor rgb="AARRGGBB"/> into the palette array, 0-indexed.
    let palette: Vec<String> = if let Some(start) = styles_xml.find("<indexedColors>") {
        let section_start = start + "<indexedColors>".len();
        let section_end = styles_xml[section_start..]
            .find("</indexedColors>")
            .map(|p| section_start + p)
            .unwrap_or(styles_xml.len());
        let section = &styles_xml[section_start..section_end];
        let mut result = Vec::new();
        let needle = "<rgbColor rgb=\"";
        let mut pos = 0;
        while let Some(rel) = section[pos..].find(needle) {
            let abs = pos + rel + needle.len();
            if let Some(close_rel) = section[abs..].find('"') {
                let rgb = &section[abs..abs + close_rel];
                if rgb.len() == 8 {
                    result.push(format!("#{}", &rgb[2..]));
                } else if rgb.len() == 6 {
                    result.push(format!("#{}", rgb));
                }
                pos = abs + close_rel;
            } else {
                break;
            }
        }
        result
    } else {
        return Ok(input.to_vec());
    };

    if palette.is_empty() {
        return Ok(input.to_vec());
    }

    // Substitute `<fgColor|bgColor|color indexed="N"/>` → `<tag rgb="FF<hex>"/>`.
    // Only the bare self-closing form is handled; xlsx writers almost always
    // emit colour nodes this way inside `<fill>` / `<font>` entries.
    for (idx, col) in palette.iter().enumerate() {
        let hex = col.trim_start_matches('#');
        for tag in ["fgColor", "bgColor", "color"] {
            let pat = format!("<{tag} indexed=\"{idx}\"/>");
            let rep = format!("<{tag} rgb=\"FF{hex}\"/>");
            if styles_xml.contains(&pat) {
                styles_xml = styles_xml.replace(&pat, &rep);
            }
        }
    }

    let mut zin = ZipArchive::new(Cursor::new(input)).map_err(|e| e.to_string())?;
    let mut buf = Cursor::new(Vec::<u8>::new());
    {
        let mut zout = ZipWriter::new(&mut buf);
        for i in 0..zin.len() {
            let mut entry = zin.by_index(i).map_err(|e| e.to_string())?;
            let name = entry.name().to_string();
            let method = entry.compression();
            let mut content = Vec::new();
            entry.read_to_end(&mut content).map_err(|e| e.to_string())?;
            if name == "xl/styles.xml" {
                content = styles_xml.clone().into_bytes();
            }
            let opts = FileOptions::default().compression_method(method);
            zout.start_file(name, opts).map_err(|e| e.to_string())?;
            zout.write_all(&content).map_err(|e| e.to_string())?;
        }
        zout.finish().map_err(|e| e.to_string())?;
    }
    Ok(buf.into_inner())
}

/// Strip CSE-array-formula markers from a worksheet XML so IronCalc's
/// loader (which bails on `t="array"` non-dynamic-array formulas) accepts
/// the file. The formula stays on the parent cell as a normal formula —
/// the spilled daughter cells lose their array semantics but their cached
/// values (already in the .xlsx) come through. Same trick for `t="dataTable"`.
fn strip_array_markers_from_sheet_xml(xml: &str) -> String {
    // Attribute order varies in the wild; cover both `t="array" ref="…"` and
    // `ref="…" t="array"`. We only touch attributes adjacent to the array
    // marker so other `ref` attrs (e.g. on `<dimension>`) are untouched.
    let patterns: [(&str, &str); 6] = [
        (" t=\"array\" ref=\"", " ref=\""),
        (" ref=\"", " ref=\""),
        (" t=\"array\"", ""),
        (" t=\"dataTable\"", ""),
        (" aca=\"1\"", ""),
        (" cm=\"\"", ""),
    ];
    let mut out = xml.to_string();
    for (from, to) in patterns {
        if from == to {
            continue;
        }
        out = out.replace(from, to);
    }
    out
}

/// Read an .xlsx file, rewrite its worksheet XMLs to strip array-formula
/// markers, and return a fresh in-memory .xlsx as bytes that IronCalc can
/// load. Files without array formulas are passed through with byte
/// equivalence (modulo the zip writer's compression choices).
fn preprocess_xlsx_bytes_for_ironcalc(input: &[u8]) -> Result<Vec<u8>, String> {
    use std::io::{Cursor, Read, Write};
    use zip::write::FileOptions;
    use zip::{CompressionMethod, ZipArchive, ZipWriter};

    let mut zin = ZipArchive::new(Cursor::new(input)).map_err(|e| e.to_string())?;
    let mut buf = Cursor::new(Vec::<u8>::new());
    {
        let mut zout = ZipWriter::new(&mut buf);
        let opts = FileOptions::default().compression_method(CompressionMethod::Deflated);
        for i in 0..zin.len() {
            let mut entry = zin.by_index(i).map_err(|e| e.to_string())?;
            let name = entry.name().to_string();
            let mut content = Vec::new();
            entry.read_to_end(&mut content).map_err(|e| e.to_string())?;
            if name.starts_with("xl/worksheets/sheet") && name.ends_with(".xml") {
                if let Ok(s) = std::str::from_utf8(&content) {
                    let rewritten = strip_array_markers_from_sheet_xml(s);
                    content = rewritten.into_bytes();
                }
            }
            zout.start_file(name, opts).map_err(|e| e.to_string())?;
            zout.write_all(&content).map_err(|e| e.to_string())?;
        }
        zout.finish().map_err(|e| e.to_string())?;
    }
    Ok(buf.into_inner())
}

/// Walk the original xlsx for `<f t="array" ref="X:Y">MY*(...)</f>` entries
/// and overwrite every cell in the ref range with the same formula via
/// `Model::set_user_input`. POI's xls trick is "every cell in the spill area
/// shares the formula"; the array-marker stripper that runs during load
/// keeps only the parent cell's formula, so without this pass the daughter
/// cells stay as POI's stale `<v>` literals (or are missing entirely).
///
/// Only formulas whose name starts with `MY` (case-insensitive) are
/// expanded — other array formulas (which IronCalc 0.7.1 doesn't really
/// support anyway) get the existing strip-only treatment.
///
/// Must be called before `model.evaluate()` so the new formulas participate
/// in the dependency analysis. Returns the number of cells written.
pub fn replicate_my_array_formulas(
    model: &mut Model<'static>,
    original_bytes: &[u8],
) -> Result<usize, String> {
    let sheet_paths = extract_sheet_paths(original_bytes)?;
    let mut written = 0;
    for (sheet_idx, sheet_path) in sheet_paths.iter().enumerate() {
        let xml = match read_zip_entry(original_bytes, sheet_path) {
            Some(s) => s,
            None => continue,
        };
        for entry in scan_my_array_formulas(&xml) {
            let cells = cells_in_range(&entry.ref_range);
            // Top-left of the spill range = anchor's (row, col).
            let (anchor_row, anchor_col) = match cells.first() {
                Some(c) => *c,
                None => continue,
            };
            // Strip the trailing anchor cell reference and replace it with
            // two integer offsets `dr, dc`. Without this, every replicated
            // cell would reference itself (anchor=this cell), and IronCalc's
            // dependency analyzer would short-circuit each one to #CIRC!.
            let head_only = match strip_trailing_anchor(&entry.formula) {
                Some(s) => s,
                None => continue,
            };
            for (row, col) in cells {
                let dr = row - anchor_row;
                let dc = col - anchor_col;
                let formula = format!("={head_only}, {dr}, {dc})");
                model
                    .set_user_input(sheet_idx as u32, row, col, formula)
                    .map_err(|e| format!("set_user_input {row},{col}: {e}"))?;
                written += 1;
            }
        }
    }
    Ok(written)
}

/// Slice off the trailing anchor argument and the closing `)`.
/// `MYUNIQUE(temp7, B22)` → `MYUNIQUE(temp7`. Caller appends `, dr, dc)`.
/// Returns None on malformed input.
fn strip_trailing_anchor(formula: &str) -> Option<String> {
    let open = formula.find('(')?;
    let bytes = formula.as_bytes();
    let mut depth: i32 = 0;
    let mut close: Option<usize> = None;
    let mut last_comma: Option<usize> = None;
    let mut in_str = false;
    for i in (open + 1)..bytes.len() {
        let c = bytes[i] as char;
        if c == '"' {
            in_str = !in_str;
            continue;
        }
        if in_str {
            continue;
        }
        match c {
            '(' => depth += 1,
            ')' if depth == 0 => {
                close = Some(i);
                break;
            }
            ')' => depth -= 1,
            ',' if depth == 0 => last_comma = Some(i),
            _ => {}
        }
    }
    let _close = close?;
    // Need at least one top-level comma — MY* always has (data, ..., anchor).
    let cut = last_comma?;
    Some(formula[..cut].to_string())
}

#[derive(Debug)]
struct MyArrayEntry {
    /// Inclusive ref range, e.g. "B22:B31".
    ref_range: String,
    /// Formula text, no leading `=`, XML-unescaped.
    formula: String,
}

fn read_zip_entry(zip_bytes: &[u8], name: &str) -> Option<String> {
    use std::io::{Cursor, Read};
    use zip::ZipArchive;
    let mut zin = ZipArchive::new(Cursor::new(zip_bytes)).ok()?;
    let mut entry = zin.by_name(name).ok()?;
    let mut s = String::new();
    entry.read_to_string(&mut s).ok()?;
    Some(s)
}

fn scan_my_array_formulas(xml: &str) -> Vec<MyArrayEntry> {
    let mut out = Vec::new();
    let mut pos = 0;
    while let Some(rel) = xml[pos..].find("<f ") {
        let abs = pos + rel;
        // Find the end of the opening tag.
        let tag_end = match xml[abs..].find('>') {
            Some(p) => abs + p,
            None => break,
        };
        let opening = &xml[abs..=tag_end];
        // Self-closing `<f t="shared" si="0"/>` has no body — skip without
        // searching for `</f>` (which would otherwise jump to the next
        // unrelated formula and silently consume entries in between).
        if opening.ends_with("/>") {
            pos = tag_end + 1;
            continue;
        }
        let close = match xml[tag_end + 1..].find("</f>") {
            Some(p) => tag_end + 1 + p,
            None => {
                pos = tag_end + 1;
                continue;
            }
        };
        let body = &xml[tag_end + 1..close];
        // Look for t="array" + ref="..." attributes on the opening tag.
        if let (Some(_), Some(ref_range)) = (
            attr_value(opening, "t").filter(|v| v == "array"),
            attr_value(opening, "ref"),
        ) {
            let trimmed = body.trim_start();
            // Match MY*( prefix to keep this pass cheap.
            let upper_prefix: String = trimmed
                .chars()
                .take_while(|c| c.is_ascii_alphanumeric())
                .collect::<String>()
                .to_uppercase();
            if upper_prefix.starts_with("MY") {
                out.push(MyArrayEntry {
                    ref_range,
                    formula: xml_unescape(body),
                });
            }
        }
        pos = close + 4;
    }
    out
}

fn attr_value(tag: &str, name: &str) -> Option<String> {
    let needle = format!("{name}=\"");
    let i = tag.find(&needle)?;
    let after = &tag[i + needle.len()..];
    let end = after.find('"')?;
    Some(after[..end].to_string())
}

fn xml_unescape(s: &str) -> String {
    s.replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&quot;", "\"")
        .replace("&apos;", "'")
        .replace("&amp;", "&")
}

/// Parse an inclusive A1:B10 range into 1-based (row, col) tuples.
fn cells_in_range(range: &str) -> Vec<(i32, i32)> {
    let (lo, hi) = match range.split_once(':') {
        Some((l, h)) => (l, h),
        None => (range, range),
    };
    let (r1, c1) = match parse_a1(lo) {
        Some(v) => v,
        None => return Vec::new(),
    };
    let (r2, c2) = match parse_a1(hi) {
        Some(v) => v,
        None => return Vec::new(),
    };
    let (r_lo, r_hi) = (r1.min(r2), r1.max(r2));
    let (c_lo, c_hi) = (c1.min(c2), c1.max(c2));
    let mut out = Vec::with_capacity(((r_hi - r_lo + 1) * (c_hi - c_lo + 1)) as usize);
    for r in r_lo..=r_hi {
        for c in c_lo..=c_hi {
            out.push((r, c));
        }
    }
    out
}

/// Parse "AB12" → (12, 28). 1-based. Strips `$`. Returns None if malformed.
fn parse_a1(s: &str) -> Option<(i32, i32)> {
    let s = s.replace('$', "");
    let mut chars = s.chars();
    let mut col = 0i32;
    let mut row_str = String::new();
    let mut have_letters = false;
    for c in chars.by_ref() {
        if c.is_ascii_alphabetic() {
            col = col * 26 + (c.to_ascii_uppercase() as i32 - 'A' as i32 + 1);
            have_letters = true;
        } else {
            row_str.push(c);
            break;
        }
    }
    for c in chars {
        row_str.push(c);
    }
    if !have_letters {
        return None;
    }
    let row: i32 = row_str.parse().ok()?;
    Some((row, col))
}

/// Load an .xlsx into a Model. Always applies the custom-palette rewrite
/// (so indexed colours overridden by the file's own palette come through
/// correctly), then tries IronCalc's loader. Falls back to stripping
/// array/dataTable formula markers on load failure. Returns the loaded
/// (but un-evaluated) Model. Public so the probe binary can use the same
/// path the GUI does.
pub fn load_xlsx_with_fallback(path: &str) -> Result<Model<'static>, String> {
    let name = std::path::Path::new(path)
        .file_stem()
        .and_then(|s| s.to_str())
        .unwrap_or("workbook")
        .to_string();
    let bytes = std::fs::read(path).map_err(|e| e.to_string())?;
    let palette_fixed = preprocess_xlsx_custom_palette(&bytes).unwrap_or_else(|_| bytes.clone());
    match load_from_xlsx_bytes(&palette_fixed, &name, "en", "UTC") {
        Ok(wb) => ironcalc::base::Model::from_workbook(wb, "en")
            .map_err(|e| format!("from_workbook failed: {e}")),
        Err(e) => {
            let msg = e.to_string();
            if !(msg.contains("array formulas") || msg.contains("data table formulas")) {
                return Err(msg);
            }
            let cleaned = preprocess_xlsx_bytes_for_ironcalc(&palette_fixed)?;
            let wb = load_from_xlsx_bytes(&cleaned, &name, "en", "UTC")
                .map_err(|e2| format!("preprocessed load failed: {e2} (original error: {msg})"))?;
            ironcalc::base::Model::from_workbook(wb, "en")
                .map_err(|e| format!("from_workbook failed: {e}"))
        }
    }
}
