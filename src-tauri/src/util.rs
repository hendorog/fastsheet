pub(crate) fn col_letter(mut col: u32) -> String {
    let mut out = String::new();
    while col > 0 {
        let r = ((col - 1) % 26) as u8;
        out.insert(0, (b'A' + r) as char);
        col = (col - 1) / 26;
    }
    out
}

pub(crate) fn col_letter_i(col: i32) -> String {
    if col < 1 {
        return String::new();
    }
    col_letter(col as u32)
}

pub(crate) fn parse_attr_val(tag: &str, name: &str) -> Option<String> {
    let needle = format!("{name}=\"");
    let start = tag.find(&needle)?;
    let after = &tag[start + needle.len()..];
    let end = after.find('"')?;
    Some(after[..end].to_string())
}
