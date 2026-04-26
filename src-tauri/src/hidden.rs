use tauri::State;

use crate::state::AppState;
use crate::util::parse_attr_val;

/// Scrape hidden column ranges out of the raw worksheet XML. IronCalc's
/// `Col` struct lacks a `hidden` field, so this information is otherwise
/// lost on load. Each returned `(min, max)` pair inclusive is hidden.
pub fn extract_hidden_col_ranges(zip_bytes: &[u8], sheet_path: &str) -> Vec<(i32, i32)> {
    use std::io::{Cursor, Read};
    use zip::ZipArchive;
    let mut zin = match ZipArchive::new(Cursor::new(zip_bytes)) {
        Ok(z) => z,
        Err(_) => return Vec::new(),
    };
    let mut xml = String::new();
    if zin
        .by_name(sheet_path)
        .ok()
        .and_then(|mut e| e.read_to_string(&mut xml).ok())
        .is_none()
    {
        return Vec::new();
    }
    // Look only inside the <cols>…</cols> section to avoid matching `hidden`
    // attrs elsewhere in the XML.
    let cols_start = match xml.find("<cols>") {
        Some(i) => i + "<cols>".len(),
        None => return Vec::new(),
    };
    let cols_end = match xml[cols_start..].find("</cols>") {
        Some(i) => cols_start + i,
        None => return Vec::new(),
    };
    let cols_xml = &xml[cols_start..cols_end];
    let mut out = Vec::new();
    let mut pos = 0;
    while let Some(rel) = cols_xml[pos..].find("<col ") {
        let abs = pos + rel;
        let close = cols_xml[abs..]
            .find('>')
            .map(|p| abs + p)
            .unwrap_or(cols_xml.len());
        let tag = &cols_xml[abs..close];
        if tag.contains("hidden=\"1\"") {
            let min = parse_attr_val(tag, "min")
                .and_then(|v| v.parse::<i32>().ok())
                .unwrap_or(0);
            let max = parse_attr_val(tag, "max")
                .and_then(|v| v.parse::<i32>().ok())
                .unwrap_or(0);
            if min > 0 && max >= min {
                out.push((min, max));
            }
        }
        pos = close;
    }
    out
}

/// Diagnostic: report the sheet_path mapping and hidden-col ranges the
/// backend has extracted. Called from the frontend status bar on demand.
#[tauri::command]
pub(crate) fn debug_hidden_cols(sheet: u32, state: State<'_, AppState>) -> Result<String, String> {
    let loaded = state.loaded.lock().unwrap();
    let l = loaded.as_ref().ok_or("no loaded file")?;
    let path = l
        .sheet_paths
        .get(sheet as usize)
        .cloned()
        .unwrap_or_default();
    let ranges = extract_hidden_col_ranges(&l.bytes, &path);
    Ok(format!("sheet {sheet} path {path:?} hidden {ranges:?}"))
}
