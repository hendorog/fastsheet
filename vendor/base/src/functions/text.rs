use crate::{
    calc_result::CalcResult,
    constants::{LAST_COLUMN, LAST_ROW},
    expressions::{parser::Node, token::Error, types::CellReferenceIndex},
    formatter::format::{format_number, parse_formatted_number},
    model::Model,
    number_format::{to_excel_precision_str, to_precision},
};

use super::{
    text_util::{substitute, text_after, text_before, Case},
    util::from_wildcard_to_regex,
};

/// Win-1252 mapping for the C1 control range (0x80..=0x9F), shared by
/// CHAR and CODE. Five codepoints in this range are undefined in
/// Win-1252 — Excel's CHAR returns #VALUE! for them.
const WIN1252_C1: [Option<char>; 32] = [
    Some('€'), None,      Some('‚'), Some('ƒ'), Some('„'), Some('…'), Some('†'), Some('‡'),
    Some('ˆ'), Some('‰'), Some('Š'), Some('‹'), Some('Œ'), None,      Some('Ž'), None,
    None,      Some('\u{2018}'), Some('\u{2019}'), Some('\u{201C}'), Some('\u{201D}'), Some('•'), Some('–'), Some('—'),
    Some('˜'), Some('™'), Some('š'), Some('›'), Some('œ'), None,      Some('ž'), Some('Ÿ'),
];

/// Finds the first instance of 'search_for' in text starting at char index start
fn find(search_for: &str, text: &str, start: usize) -> Option<i32> {
    let ch = text.chars();
    let mut byte_index = 0;
    for (char_index, c) in ch.enumerate() {
        if char_index + 1 >= start && text[byte_index..].starts_with(search_for) {
            return Some((char_index + 1) as i32);
        }
        byte_index += c.len_utf8();
    }
    None
}

/// You can use the wildcard characters — the question mark (?) and asterisk (*) — in the find_text argument.
/// * A question mark matches any single character.
/// * An asterisk matches any sequence of characters.
/// * If you want to find an actual question mark or asterisk, type a tilde (~) before the character.
fn search(search_for: &str, text: &str, start: usize) -> Option<i32> {
    let re = match from_wildcard_to_regex(search_for, false) {
        Ok(r) => r,
        Err(_) => return None,
    };

    let ch = text.chars();
    let mut byte_index = 0;
    for (char_index, c) in ch.enumerate() {
        if char_index + 1 >= start {
            if let Some(m) = re.find(&text[byte_index..]) {
                return Some((text[0..(m.start() + byte_index)].chars().count() as i32) + 1);
            } else {
                return None;
            }
        }
        byte_index += c.len_utf8();
    }
    None
}

impl<'a> Model<'a> {
    pub(crate) fn fn_concat(&mut self, args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        let mut result = "".to_string();
        for arg in args {
            match self.evaluate_node_in_context(arg, cell) {
                CalcResult::String(value) => result = format!("{result}{value}"),
                CalcResult::Number(value) => result.push_str(&to_excel_precision_str(value)),
                CalcResult::EmptyCell | CalcResult::EmptyArg => {}
                CalcResult::Boolean(value) => {
                    if value {
                        result = format!("{result}TRUE");
                    } else {
                        result = format!("{result}FALSE");
                    }
                }
                error @ CalcResult::Error { .. } => return error,
                CalcResult::Range { left, right } => {
                    if left.sheet != right.sheet {
                        return CalcResult::new_error(
                            Error::VALUE,
                            cell,
                            "Ranges are in different sheets".to_string(),
                        );
                    }
                    for row in left.row..(right.row + 1) {
                        for column in left.column..(right.column + 1) {
                            match self.evaluate_cell(CellReferenceIndex {
                                sheet: left.sheet,
                                row,
                                column,
                            }) {
                                CalcResult::String(value) => {
                                    result = format!("{result}{value}");
                                }
                                CalcResult::Number(value) => result.push_str(&to_excel_precision_str(value)),
                                CalcResult::Boolean(value) => {
                                    if value {
                                        result = format!("{result}TRUE");
                                    } else {
                                        result = format!("{result}FALSE");
                                    }
                                }
                                error @ CalcResult::Error { .. } => return error,
                                CalcResult::EmptyCell | CalcResult::EmptyArg => {}
                                CalcResult::Range { .. } => {}
                                CalcResult::Array(_) => {
                                    return CalcResult::Error {
                                        error: Error::NIMPL,
                                        origin: cell,
                                        message: "Arrays not supported yet".to_string(),
                                    }
                                }
                            }
                        }
                    }
                }
                CalcResult::Array(_) => {
                    return CalcResult::Error {
                        error: Error::NIMPL,
                        origin: cell,
                        message: "Arrays not supported yet".to_string(),
                    }
                }
            };
        }
        CalcResult::String(result)
    }
    pub(crate) fn fn_text(&mut self, args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        if args.len() == 2 {
            let value = match self.evaluate_node_in_context(&args[0], cell) {
                CalcResult::Number(f) => f,
                CalcResult::String(s) => {
                    return CalcResult::String(s);
                }
                CalcResult::Boolean(b) => {
                    return CalcResult::Boolean(b);
                }
                error @ CalcResult::Error { .. } => return error,
                CalcResult::Range { .. } => {
                    // Implicit Intersection not implemented
                    return CalcResult::Error {
                        error: Error::NIMPL,
                        origin: cell,
                        message: "Implicit Intersection not implemented".to_string(),
                    };
                }
                CalcResult::EmptyCell | CalcResult::EmptyArg => 0.0,
                CalcResult::Array(_) => {
                    return CalcResult::Error {
                        error: Error::NIMPL,
                        origin: cell,
                        message: "Arrays not supported yet".to_string(),
                    }
                }
            };
            let format_code = match self.get_string(&args[1], cell) {
                Ok(s) => s,
                Err(s) => return s,
            };
            let d = format_number(value, &format_code, self.locale);
            if let Some(_e) = d.error {
                return CalcResult::Error {
                    error: Error::VALUE,
                    origin: cell,
                    message: "Invalid format code".to_string(),
                };
            }
            CalcResult::String(d.text)
        } else {
            CalcResult::new_args_number_error(cell)
        }
    }

    /// FIND(find_text, within_text, [start_num])
    ///  * FIND and FINDB are case sensitive and don't allow wildcard characters.
    ///  * If find_text is "" (empty text), FIND matches the first character in the search string (that is, the character numbered start_num or 1).
    ///  * Find_text cannot contain any wildcard characters.
    ///  * If find_text does not appear in within_text, FIND and FINDB return the #VALUE! error value.
    ///  * If start_num is not greater than zero, FIND and FINDB return the #VALUE! error value.
    ///  * If start_num is greater than the length of within_text, FIND and FINDB return the #VALUE! error value.
    ///    NB: FINDB is not implemented. It is the same as FIND function unless locale is a DBCS (Double Byte Character Set)
    pub(crate) fn fn_find(&mut self, args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        if args.len() < 2 || args.len() > 3 {
            return CalcResult::new_args_number_error(cell);
        }
        let find_text = match self.get_string(&args[0], cell) {
            Ok(s) => s,
            Err(s) => return s,
        };
        let within_text = match self.get_string(&args[1], cell) {
            Ok(s) => s,
            Err(s) => return s,
        };
        let start_num = if args.len() == 3 {
            match self.get_number(&args[2], cell) {
                Ok(s) => s.floor(),
                Err(s) => return s,
            }
        } else {
            1.0
        };

        if start_num < 1.0 {
            return CalcResult::Error {
                error: Error::VALUE,
                origin: cell,
                message: "Start num must be >= 1".to_string(),
            };
        }
        let start_num = start_num as usize;

        if start_num > within_text.chars().count() {
            return CalcResult::Error {
                error: Error::VALUE,
                origin: cell,
                message: "Start num greater than length".to_string(),
            };
        }
        if let Some(s) = find(&find_text, &within_text, start_num) {
            CalcResult::Number(s as f64)
        } else {
            CalcResult::Error {
                error: Error::VALUE,
                origin: cell,
                message: "Text not found".to_string(),
            }
        }
    }

    /// Same API as FIND but:
    ///  * Allows wildcards
    ///  * It is case insensitive
    ///    SEARCH(find_text, within_text, [start_num])
    pub(crate) fn fn_search(&mut self, args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        if args.len() < 2 || args.len() > 3 {
            return CalcResult::new_args_number_error(cell);
        }
        let find_text = match self.get_string(&args[0], cell) {
            Ok(s) => s,
            Err(s) => return s,
        };
        let within_text = match self.get_string(&args[1], cell) {
            Ok(s) => s,
            Err(s) => return s,
        };
        let start_num = if args.len() == 3 {
            match self.get_number(&args[2], cell) {
                Ok(s) => s.floor(),
                Err(s) => return s,
            }
        } else {
            1.0
        };

        if start_num < 1.0 {
            return CalcResult::Error {
                error: Error::VALUE,
                origin: cell,
                message: "Start num must be >= 1".to_string(),
            };
        }
        let start_num = start_num as usize;

        if start_num > within_text.chars().count() {
            return CalcResult::Error {
                error: Error::VALUE,
                origin: cell,
                message: "Start num greater than length".to_string(),
            };
        }
        // SEARCH is case insensitive
        if let Some(s) = search(
            &find_text.to_lowercase(),
            &within_text.to_lowercase(),
            start_num,
        ) {
            CalcResult::Number(s as f64)
        } else {
            CalcResult::Error {
                error: Error::VALUE,
                origin: cell,
                message: "Text not found".to_string(),
            }
        }
    }

    // LEN, LEFT, RIGHT, MID, LOWER, UPPER, TRIM
    pub(crate) fn fn_len(&mut self, args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        if args.len() == 1 {
            let result = self.evaluate_node_in_context(&args[0], cell);
            let s = match result {
                CalcResult::Number(v) => to_excel_precision_str(v),
                CalcResult::String(v) => v,
                CalcResult::Boolean(b) => {
                    if b {
                        "TRUE".to_string()
                    } else {
                        "FALSE".to_string()
                    }
                }
                error @ CalcResult::Error { .. } => return error,
                CalcResult::Range { .. } => {
                    // Implicit Intersection not implemented
                    return CalcResult::Error {
                        error: Error::NIMPL,
                        origin: cell,
                        message: "Implicit Intersection not implemented".to_string(),
                    };
                }
                CalcResult::EmptyCell | CalcResult::EmptyArg => "".to_string(),
                CalcResult::Array(_) => {
                    return CalcResult::Error {
                        error: Error::NIMPL,
                        origin: cell,
                        message: "Arrays not supported yet".to_string(),
                    }
                }
            };
            return CalcResult::Number(s.chars().count() as f64);
        }
        CalcResult::new_args_number_error(cell)
    }

    pub(crate) fn fn_trim(&mut self, args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        if args.len() == 1 {
            let s = match self.evaluate_node_in_context(&args[0], cell) {
                CalcResult::Number(v) => to_excel_precision_str(v),
                CalcResult::String(v) => v,
                CalcResult::Boolean(b) => {
                    if b {
                        "TRUE".to_string()
                    } else {
                        "FALSE".to_string()
                    }
                }
                error @ CalcResult::Error { .. } => return error,
                CalcResult::Range { .. } => {
                    // Implicit Intersection not implemented
                    return CalcResult::Error {
                        error: Error::NIMPL,
                        origin: cell,
                        message: "Implicit Intersection not implemented".to_string(),
                    };
                }
                CalcResult::EmptyCell | CalcResult::EmptyArg => "".to_string(),
                CalcResult::Array(_) => {
                    return CalcResult::Error {
                        error: Error::NIMPL,
                        origin: cell,
                        message: "Arrays not supported yet".to_string(),
                    }
                }
            };
            // Excel TRIM collapses interior runs of spaces (0x20 only) to a
            // single space as well as stripping the ends; other whitespace
            // (tabs, NBSP) passes through untouched. Fastsheet vendor patch —
            // upstream only stripped the ends.
            let collapsed = s
                .split(' ')
                .filter(|part| !part.is_empty())
                .collect::<Vec<_>>()
                .join(" ");
            return CalcResult::String(collapsed);
        }
        CalcResult::new_args_number_error(cell)
    }

    pub(crate) fn fn_lower(&mut self, args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        if args.len() == 1 {
            let s = match self.evaluate_node_in_context(&args[0], cell) {
                CalcResult::Number(v) => to_excel_precision_str(v),
                CalcResult::String(v) => v,
                CalcResult::Boolean(b) => {
                    if b {
                        "TRUE".to_string()
                    } else {
                        "FALSE".to_string()
                    }
                }
                error @ CalcResult::Error { .. } => return error,
                CalcResult::Range { .. } => {
                    // Implicit Intersection not implemented
                    return CalcResult::Error {
                        error: Error::NIMPL,
                        origin: cell,
                        message: "Implicit Intersection not implemented".to_string(),
                    };
                }
                CalcResult::EmptyCell | CalcResult::EmptyArg => "".to_string(),
                CalcResult::Array(_) => {
                    return CalcResult::Error {
                        error: Error::NIMPL,
                        origin: cell,
                        message: "Arrays not supported yet".to_string(),
                    }
                }
            };
            return CalcResult::String(s.to_lowercase());
        }
        CalcResult::new_args_number_error(cell)
    }

    /// CHAR(n) — Excel-compatible character lookup. n must be 1..=255.
    /// 0x80..=0x9F follow the Win-1252 mapping Excel uses (€, smart quotes,
    /// em-dash, etc.); everything else is the Unicode codepoint with the
    /// matching numeric value (so CHAR(9)=TAB, CHAR(13)=CR, CHAR(65)='A').
    pub(crate) fn fn_char(&mut self, args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        if args.len() != 1 {
            return CalcResult::new_args_number_error(cell);
        }
        let n = match self.get_number(&args[0], cell) {
            Ok(f) => f.trunc() as i32,
            Err(e) => return e,
        };
        if !(1..=255).contains(&n) {
            return CalcResult::new_error(
                Error::VALUE,
                cell,
                "CHAR: argument must be in 1..=255".to_string(),
            );
        }
        let ch = if (0x80..=0x9F).contains(&n) {
            match WIN1252_C1[(n - 0x80) as usize] {
                Some(c) => c,
                None => {
                    return CalcResult::new_error(
                        Error::VALUE,
                        cell,
                        format!("CHAR: {n} is undefined in Win-1252"),
                    );
                }
            }
        } else {
            // 1..=127 and 160..=255 — direct Unicode codepoint.
            char::from_u32(n as u32).unwrap_or('\u{FFFD}')
        };
        CalcResult::String(ch.to_string())
    }

    // CODE(text) — Win-1252 code of the first character (inverse of
    // CHAR). Characters outside the ANSI codepage return 63 ('?'),
    // matching Excel. Fastsheet addition — missing from upstream 0.7.1.
    pub(crate) fn fn_code(&mut self, args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        if args.len() != 1 {
            return CalcResult::new_args_number_error(cell);
        }
        let s = match self.get_string(&args[0], cell) {
            Ok(s) => s,
            Err(e) => return e,
        };
        let Some(c) = s.chars().next() else {
            return CalcResult::new_error(Error::VALUE, cell, "CODE of empty text".to_string());
        };
        let cp = c as u32;
        let code = if cp <= 0x7F || (0xA0..=0xFF).contains(&cp) {
            cp
        } else if let Some(pos) = WIN1252_C1.iter().position(|&e| e == Some(c)) {
            0x80 + pos as u32
        } else {
            63
        };
        CalcResult::Number(code as f64)
    }

    // PROPER(text) — capitalize every letter that follows a non-letter,
    // lowercase the rest. Excel counts ANY non-letter as a word break,
    // so "2-cent's worth" → "2-Cent'S Worth". Fastsheet addition.
    pub(crate) fn fn_proper(&mut self, args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        if args.len() != 1 {
            return CalcResult::new_args_number_error(cell);
        }
        let s = match self.get_string(&args[0], cell) {
            Ok(s) => s,
            Err(e) => return e,
        };
        let mut out = String::with_capacity(s.len());
        let mut prev_is_letter = false;
        for c in s.chars() {
            if c.is_alphabetic() {
                if prev_is_letter {
                    out.extend(c.to_lowercase());
                } else {
                    out.extend(c.to_uppercase());
                }
                prev_is_letter = true;
            } else {
                out.push(c);
                prev_is_letter = false;
            }
        }
        CalcResult::String(out)
    }

    // REPLACE(old_text, start_num, num_chars, new_text) — positional
    // replacement by character index (1-based). start past the end
    // appends; num_chars 0 inserts. Fastsheet addition.
    pub(crate) fn fn_replace(&mut self, args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        if args.len() != 4 {
            return CalcResult::new_args_number_error(cell);
        }
        let old_text = match self.get_string(&args[0], cell) {
            Ok(s) => s,
            Err(e) => return e,
        };
        let start = match self.get_number(&args[1], cell) {
            Ok(f) => f.trunc() as i64,
            Err(e) => return e,
        };
        let num_chars = match self.get_number(&args[2], cell) {
            Ok(f) => f.trunc() as i64,
            Err(e) => return e,
        };
        let new_text = match self.get_string(&args[3], cell) {
            Ok(s) => s,
            Err(e) => return e,
        };
        if start < 1 || num_chars < 0 {
            return CalcResult::new_error(
                Error::VALUE,
                cell,
                "REPLACE: start_num must be >= 1 and num_chars >= 0".to_string(),
            );
        }
        let chars: Vec<char> = old_text.chars().collect();
        let head = ((start - 1) as usize).min(chars.len());
        let tail = head.saturating_add(num_chars as usize).min(chars.len());
        let mut out: String = chars[..head].iter().collect();
        out.push_str(&new_text);
        out.extend(&chars[tail..]);
        CalcResult::String(out)
    }

    // FIXED(number, [decimals = 2], [no_commas = FALSE]) — round to a
    // fixed decimal count and format as text with thousands separators.
    // Negative decimals round to the left of the decimal point.
    // Fastsheet addition.
    pub(crate) fn fn_fixed(&mut self, args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        if args.is_empty() || args.len() > 3 {
            return CalcResult::new_args_number_error(cell);
        }
        let number = match self.get_number(&args[0], cell) {
            Ok(f) => f,
            Err(e) => return e,
        };
        let decimals = if args.len() > 1 {
            match self.get_number(&args[1], cell) {
                Ok(f) => f.trunc() as i32,
                Err(e) => return e,
            }
        } else {
            2
        };
        let no_commas = if args.len() > 2 {
            match self.get_boolean(&args[2], cell) {
                Ok(b) => b,
                Err(e) => return e,
            }
        } else {
            false
        };
        if !(-127..=127).contains(&decimals) {
            return CalcResult::new_error(
                Error::VALUE,
                cell,
                "FIXED: decimals out of range".to_string(),
            );
        }
        let pow = 10f64.powi(decimals);
        let rounded = (number * pow).round() / pow;
        let places = decimals.max(0) as usize;
        let body = format!("{:.*}", places, rounded.abs());
        let (int_part, frac_part) = match body.split_once('.') {
            Some((i, f)) => (i.to_string(), Some(f.to_string())),
            None => (body, None),
        };
        let int_grouped = if no_commas {
            int_part
        } else {
            let digits: Vec<char> = int_part.chars().collect();
            let mut grouped = String::new();
            for (i, c) in digits.iter().enumerate() {
                if i > 0 && (digits.len() - i) % 3 == 0 {
                    grouped.push(',');
                }
                grouped.push(*c);
            }
            grouped
        };
        let sign = if rounded < 0.0 { "-" } else { "" };
        let out = match frac_part {
            Some(f) => format!("{sign}{int_grouped}.{f}"),
            None => format!("{sign}{int_grouped}"),
        };
        CalcResult::String(out)
    }

    // NUMBERVALUE(text, [decimal_sep], [group_sep]) — locale-flexible
    // text-to-number. Whitespace is ignored anywhere, group separators
    // are only valid before the decimal separator, and trailing '%'
    // signs each divide by 100. Empty text is 0. Fastsheet addition.
    pub(crate) fn fn_numbervalue(&mut self, args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        if args.is_empty() || args.len() > 3 {
            return CalcResult::new_args_number_error(cell);
        }
        let text = match self.get_string(&args[0], cell) {
            Ok(s) => s,
            Err(e) => return e,
        };
        let sep_arg = |m: &mut Self, idx: usize, default: char| -> Result<char, CalcResult> {
            if args.len() > idx {
                let s = match m.get_string(&args[idx], cell) {
                    Ok(s) => s,
                    Err(e) => return Err(e),
                };
                match s.chars().next() {
                    Some(c) => Ok(c),
                    None => Err(CalcResult::new_error(
                        Error::VALUE,
                        cell,
                        "NUMBERVALUE: empty separator".to_string(),
                    )),
                }
            } else {
                Ok(default)
            }
        };
        let dec = match sep_arg(self, 1, '.') {
            Ok(c) => c,
            Err(e) => return e,
        };
        let group = match sep_arg(self, 2, ',') {
            Ok(c) => c,
            Err(e) => return e,
        };
        if dec == group {
            return CalcResult::new_error(
                Error::VALUE,
                cell,
                "NUMBERVALUE: separators must differ".to_string(),
            );
        }
        let mut t: String = text.chars().filter(|c| !c.is_whitespace()).collect();
        if t.is_empty() {
            return CalcResult::Number(0.0);
        }
        let mut percents = 0u32;
        while t.ends_with('%') {
            t.pop();
            percents += 1;
        }
        if t.matches(dec).count() > 1 {
            return CalcResult::new_error(
                Error::VALUE,
                cell,
                "NUMBERVALUE: more than one decimal separator".to_string(),
            );
        }
        let (int_raw, frac) = match t.split_once(dec) {
            Some((i, f)) => (i, Some(f)),
            None => (t.as_str(), None),
        };
        if let Some(f) = frac {
            if f.contains(group) {
                return CalcResult::new_error(
                    Error::VALUE,
                    cell,
                    "NUMBERVALUE: group separator after decimal".to_string(),
                );
            }
        }
        let int_clean: String = int_raw.chars().filter(|&c| c != group).collect();
        let candidate = match frac {
            Some(f) => format!("{int_clean}.{f}"),
            None => int_clean,
        };
        match candidate.parse::<f64>() {
            Ok(v) => CalcResult::Number(v / 100f64.powi(percents as i32)),
            Err(_) => CalcResult::new_error(
                Error::VALUE,
                cell,
                format!("NUMBERVALUE: cannot parse '{text}'"),
            ),
        }
    }

    pub(crate) fn fn_unicode(&mut self, args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        if args.len() == 1 {
            let s = match self.evaluate_node_in_context(&args[0], cell) {
                CalcResult::Number(v) => to_excel_precision_str(v),
                CalcResult::String(v) => v,
                CalcResult::Boolean(b) => {
                    if b {
                        "TRUE".to_string()
                    } else {
                        "FALSE".to_string()
                    }
                }
                error @ CalcResult::Error { .. } => return error,
                CalcResult::Range { .. } => {
                    // Implicit Intersection not implemented
                    return CalcResult::Error {
                        error: Error::NIMPL,
                        origin: cell,
                        message: "Implicit Intersection not implemented".to_string(),
                    };
                }
                CalcResult::EmptyCell | CalcResult::EmptyArg => {
                    return CalcResult::Error {
                        error: Error::VALUE,
                        origin: cell,
                        message: "Empty cell".to_string(),
                    }
                }
                CalcResult::Array(_) => {
                    return CalcResult::Error {
                        error: Error::NIMPL,
                        origin: cell,
                        message: "Arrays not supported yet".to_string(),
                    }
                }
            };

            match s.chars().next() {
                Some(c) => {
                    let unicode_number = c as u32;
                    return CalcResult::Number(unicode_number as f64);
                }
                None => {
                    return CalcResult::Error {
                        error: Error::VALUE,
                        origin: cell,
                        message: "Empty cell".to_string(),
                    };
                }
            }
        }
        CalcResult::new_args_number_error(cell)
    }

    pub(crate) fn fn_upper(&mut self, args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        if args.len() == 1 {
            let s = match self.evaluate_node_in_context(&args[0], cell) {
                CalcResult::Number(v) => to_excel_precision_str(v),
                CalcResult::String(v) => v,
                CalcResult::Boolean(b) => {
                    if b {
                        "TRUE".to_string()
                    } else {
                        "FALSE".to_string()
                    }
                }
                error @ CalcResult::Error { .. } => return error,
                CalcResult::Range { .. } => {
                    // Implicit Intersection not implemented
                    return CalcResult::Error {
                        error: Error::NIMPL,
                        origin: cell,
                        message: "Implicit Intersection not implemented".to_string(),
                    };
                }
                CalcResult::EmptyCell | CalcResult::EmptyArg => "".to_string(),
                CalcResult::Array(_) => {
                    return CalcResult::Error {
                        error: Error::NIMPL,
                        origin: cell,
                        message: "Arrays not supported yet".to_string(),
                    }
                }
            };
            return CalcResult::String(s.to_uppercase());
        }
        CalcResult::new_args_number_error(cell)
    }

    pub(crate) fn fn_left(&mut self, args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        if args.len() > 2 || args.is_empty() {
            return CalcResult::new_args_number_error(cell);
        }
        let s = match self.evaluate_node_in_context(&args[0], cell) {
            CalcResult::Number(v) => to_excel_precision_str(v),
            CalcResult::String(v) => v,
            CalcResult::Boolean(b) => {
                if b {
                    "TRUE".to_string()
                } else {
                    "FALSE".to_string()
                }
            }
            error @ CalcResult::Error { .. } => return error,
            CalcResult::Range { .. } => {
                // Implicit Intersection not implemented
                return CalcResult::Error {
                    error: Error::NIMPL,
                    origin: cell,
                    message: "Implicit Intersection not implemented".to_string(),
                };
            }
            CalcResult::EmptyCell | CalcResult::EmptyArg => "".to_string(),
            CalcResult::Array(_) => {
                return CalcResult::Error {
                    error: Error::NIMPL,
                    origin: cell,
                    message: "Arrays not supported yet".to_string(),
                }
            }
        };
        let num_chars = if args.len() == 2 {
            // Excel coerces numeric text and booleans here
            match self.get_number(&args[1], cell) {
                Ok(v) => {
                    if v < 0.0 {
                        return CalcResult::Error {
                            error: Error::VALUE,
                            origin: cell,
                            message: "Number must be >= 0".to_string(),
                        };
                    }
                    v.floor() as usize
                }
                Err(s) => return s,
            }
        } else {
            1
        };
        let mut result = "".to_string();
        for (index, ch) in s.chars().enumerate() {
            if index >= num_chars {
                break;
            }
            result.push(ch);
        }
        CalcResult::String(result)
    }

    pub(crate) fn fn_right(&mut self, args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        if args.len() > 2 || args.is_empty() {
            return CalcResult::new_args_number_error(cell);
        }
        let s = match self.evaluate_node_in_context(&args[0], cell) {
            CalcResult::Number(v) => to_excel_precision_str(v),
            CalcResult::String(v) => v,
            CalcResult::Boolean(b) => {
                if b {
                    "TRUE".to_string()
                } else {
                    "FALSE".to_string()
                }
            }
            error @ CalcResult::Error { .. } => return error,
            CalcResult::Range { .. } => {
                // Implicit Intersection not implemented
                return CalcResult::Error {
                    error: Error::NIMPL,
                    origin: cell,
                    message: "Implicit Intersection not implemented".to_string(),
                };
            }
            CalcResult::EmptyCell | CalcResult::EmptyArg => "".to_string(),
            CalcResult::Array(_) => {
                return CalcResult::Error {
                    error: Error::NIMPL,
                    origin: cell,
                    message: "Arrays not supported yet".to_string(),
                }
            }
        };
        let num_chars = if args.len() == 2 {
            // Excel coerces numeric text and booleans here
            match self.get_number(&args[1], cell) {
                Ok(v) => {
                    if v < 0.0 {
                        return CalcResult::Error {
                            error: Error::VALUE,
                            origin: cell,
                            message: "Number must be >= 0".to_string(),
                        };
                    }
                    v.floor() as usize
                }
                Err(s) => return s,
            }
        } else {
            1
        };
        let mut result = "".to_string();
        for (index, ch) in s.chars().rev().enumerate() {
            if index >= num_chars {
                break;
            }
            result.push(ch);
        }
        CalcResult::String(result.chars().rev().collect::<String>())
    }

    pub(crate) fn fn_mid(&mut self, args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        if args.len() != 3 {
            return CalcResult::new_args_number_error(cell);
        }
        let s = match self.evaluate_node_in_context(&args[0], cell) {
            CalcResult::Number(v) => to_excel_precision_str(v),
            CalcResult::String(v) => v,
            CalcResult::Boolean(b) => {
                if b {
                    "TRUE".to_string()
                } else {
                    "FALSE".to_string()
                }
            }
            error @ CalcResult::Error { .. } => return error,
            CalcResult::Range { .. } => {
                // Implicit Intersection not implemented
                return CalcResult::Error {
                    error: Error::NIMPL,
                    origin: cell,
                    message: "Implicit Intersection not implemented".to_string(),
                };
            }
            CalcResult::EmptyCell | CalcResult::EmptyArg => "".to_string(),
            CalcResult::Array(_) => {
                return CalcResult::Error {
                    error: Error::NIMPL,
                    origin: cell,
                    message: "Arrays not supported yet".to_string(),
                }
            }
        };
        // Excel coerces numeric text and booleans for both count arguments
        let start_num = match self.get_number(&args[1], cell) {
            Ok(v) => {
                if v < 1.0 {
                    return CalcResult::Error {
                        error: Error::VALUE,
                        origin: cell,
                        message: "Number must be >= 1".to_string(),
                    };
                }
                v.floor() as usize
            }
            Err(s) => return s,
        };
        let num_chars = match self.get_number(&args[2], cell) {
            Ok(v) => {
                if v < 0.0 {
                    return CalcResult::Error {
                        error: Error::VALUE,
                        origin: cell,
                        message: "Number must be >= 0".to_string(),
                    };
                }
                v.floor() as usize
            }
            Err(s) => return s,
        };
        let mut result = "".to_string();
        let mut count: usize = 0;
        for (index, ch) in s.chars().enumerate() {
            if count >= num_chars {
                break;
            }
            if index + 1 >= start_num {
                result.push(ch);
                count += 1;
            }
        }
        CalcResult::String(result)
    }

    // REPT(text, number_times)
    pub(crate) fn fn_rept(&mut self, args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        if args.len() != 2 {
            return CalcResult::new_args_number_error(cell);
        }
        let text = match self.get_string(&args[0], cell) {
            Ok(s) => s,
            Err(error) => return error,
        };
        let number_times = match self.get_number(&args[1], cell) {
            Ok(f) => f.floor() as i32,
            Err(s) => return s,
        };
        let text_len = text.len() as i32;

        // We normally don't follow Excel's sometimes archaic size's restrictions
        // But this might be a security issue
        if text_len * number_times > 32767 {
            return CalcResult::Error {
                error: Error::VALUE,
                origin: cell,
                message: "number times too high".to_string(),
            };
        }
        if number_times < 0 {
            return CalcResult::Error {
                error: Error::VALUE,
                origin: cell,
                message: "number times too high".to_string(),
            };
        }
        if number_times == 0 {
            return CalcResult::String("".to_string());
        }
        CalcResult::String(text.repeat(number_times as usize))
    }

    // TEXTAFTER(text, delimiter, [instance_num], [match_mode], [match_end], [if_not_found])
    pub(crate) fn fn_textafter(&mut self, args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        let arg_count = args.len();
        if !(2..=6).contains(&arg_count) {
            return CalcResult::new_args_number_error(cell);
        }
        let text = match self.get_string(&args[0], cell) {
            Ok(s) => s,
            Err(error) => return error,
        };
        let delimiter = match self.get_string(&args[1], cell) {
            Ok(s) => s,
            Err(error) => return error,
        };
        let instance_num = if arg_count > 2 {
            match self.get_number(&args[2], cell) {
                Ok(f) => f.floor() as i32,
                Err(s) => return s,
            }
        } else {
            1
        };
        let match_mode = if arg_count > 3 {
            match self.get_number(&args[3], cell) {
                Ok(f) => {
                    if f == 0.0 {
                        Case::Sensitive
                    } else {
                        Case::Insensitive
                    }
                }
                Err(s) => return s,
            }
        } else {
            Case::Sensitive
        };

        let match_end = if arg_count > 4 {
            match self.get_number(&args[4], cell) {
                Ok(f) => f,
                Err(s) => return s,
            }
        } else {
            // disabled by default
            // the delimiter is specified in the formula
            0.0
        };
        if instance_num == 0 {
            return CalcResult::Error {
                error: Error::VALUE,
                origin: cell,
                message: "instance_num must be <> 0".to_string(),
            };
        }
        if delimiter.len() > text.len() {
            // so this is fun(!)
            // if the function was provided with two arguments is a #VALUE!
            // if it had more is a #N/A (irrespective of their values)
            if arg_count > 2 {
                return CalcResult::Error {
                    error: Error::VALUE,
                    origin: cell,
                    message: "The delimiter is longer than the text is trying to match".to_string(),
                };
            } else {
                return CalcResult::Error {
                    error: Error::NA,
                    origin: cell,
                    message: "The delimiter is longer than the text is trying to match".to_string(),
                };
            }
        }
        if match_end != 0.0 && match_end != 1.0 {
            return CalcResult::Error {
                error: Error::VALUE,
                origin: cell,
                message: "argument must be 0 or 1".to_string(),
            };
        };
        match text_after(&text, &delimiter, instance_num, match_mode) {
            Some(s) => CalcResult::String(s),
            None => {
                if match_end == 1.0 {
                    if instance_num == 1 {
                        return CalcResult::String("".to_string());
                    } else if instance_num == -1 {
                        return CalcResult::String(text);
                    }
                }
                if arg_count == 6 {
                    // An empty cell is converted to empty string (not 0)
                    match self.evaluate_node_in_context(&args[5], cell) {
                        CalcResult::EmptyCell => CalcResult::String("".to_string()),
                        result => result,
                    }
                } else {
                    CalcResult::Error {
                        error: Error::NA,
                        origin: cell,
                        message: "Value not found".to_string(),
                    }
                }
            }
        }
    }

    pub(crate) fn fn_textbefore(&mut self, args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        let arg_count = args.len();
        if !(2..=6).contains(&arg_count) {
            return CalcResult::new_args_number_error(cell);
        }
        let text = match self.get_string(&args[0], cell) {
            Ok(s) => s,
            Err(error) => return error,
        };
        let delimiter = match self.get_string(&args[1], cell) {
            Ok(s) => s,
            Err(error) => return error,
        };
        let instance_num = if arg_count > 2 {
            match self.get_number(&args[2], cell) {
                Ok(f) => f.floor() as i32,
                Err(s) => return s,
            }
        } else {
            1
        };
        let match_mode = if arg_count > 3 {
            match self.get_number(&args[3], cell) {
                Ok(f) => {
                    if f == 0.0 {
                        Case::Sensitive
                    } else {
                        Case::Insensitive
                    }
                }
                Err(s) => return s,
            }
        } else {
            Case::Sensitive
        };

        let match_end = if arg_count > 4 {
            match self.get_number(&args[4], cell) {
                Ok(f) => f,
                Err(s) => return s,
            }
        } else {
            // disabled by default
            // the delimiter is specified in the formula
            0.0
        };
        if instance_num == 0 {
            return CalcResult::Error {
                error: Error::VALUE,
                origin: cell,
                message: "instance_num must be <> 0".to_string(),
            };
        }
        if delimiter.len() > text.len() {
            // so this is fun(!)
            // if the function was provided with two arguments is a #VALUE!
            // if it had more is a #N/A (irrespective of their values)
            if arg_count > 2 {
                return CalcResult::Error {
                    error: Error::VALUE,
                    origin: cell,
                    message: "The delimiter is longer than the text is trying to match".to_string(),
                };
            } else {
                return CalcResult::Error {
                    error: Error::NA,
                    origin: cell,
                    message: "The delimiter is longer than the text is trying to match".to_string(),
                };
            }
        }
        if match_end != 0.0 && match_end != 1.0 {
            return CalcResult::Error {
                error: Error::VALUE,
                origin: cell,
                message: "argument must be 0 or 1".to_string(),
            };
        };
        match text_before(&text, &delimiter, instance_num, match_mode) {
            Some(s) => CalcResult::String(s),
            None => {
                if match_end == 1.0 {
                    if instance_num == -1 {
                        return CalcResult::String("".to_string());
                    } else if instance_num == 1 {
                        return CalcResult::String(text);
                    }
                }
                if arg_count == 6 {
                    // An empty cell is converted to empty string (not 0)
                    match self.evaluate_node_in_context(&args[5], cell) {
                        CalcResult::EmptyCell => CalcResult::String("".to_string()),
                        result => result,
                    }
                } else {
                    CalcResult::Error {
                        error: Error::NA,
                        origin: cell,
                        message: "Value not found".to_string(),
                    }
                }
            }
        }
    }

    // TEXTJOIN(delimiter, ignore_empty, text1, [text2], …)
    pub(crate) fn fn_textjoin(&mut self, args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        let arg_count = args.len();
        if arg_count < 3 {
            return CalcResult::new_args_number_error(cell);
        }
        let delimiter = match self.get_string(&args[0], cell) {
            Ok(s) => s,
            Err(error) => return error,
        };
        let ignore_empty = match self.get_boolean(&args[1], cell) {
            Ok(b) => b,
            Err(error) => return error,
        };
        let mut values = Vec::new();
        for arg in &args[2..] {
            match self.evaluate_node_in_context(arg, cell) {
                CalcResult::Number(value) => values.push(to_excel_precision_str(value)),
                CalcResult::Range { left, right } => {
                    if left.sheet != right.sheet {
                        return CalcResult::new_error(
                            Error::VALUE,
                            cell,
                            "Ranges are in different sheets".to_string(),
                        );
                    }
                    let row1 = left.row;
                    let mut row2 = right.row;
                    let column1 = left.column;
                    let mut column2 = right.column;
                    if row1 == 1 && row2 == LAST_ROW {
                        row2 = match self.workbook.worksheet(left.sheet) {
                            Ok(s) => s.dimension().max_row,
                            Err(_) => {
                                return CalcResult::new_error(
                                    Error::ERROR,
                                    cell,
                                    format!("Invalid worksheet index: '{}'", left.sheet),
                                );
                            }
                        };
                    }
                    if column1 == 1 && column2 == LAST_COLUMN {
                        column2 = match self.workbook.worksheet(left.sheet) {
                            Ok(s) => s.dimension().max_column,
                            Err(_) => {
                                return CalcResult::new_error(
                                    Error::ERROR,
                                    cell,
                                    format!("Invalid worksheet index: '{}'", left.sheet),
                                );
                            }
                        };
                    }
                    for row in row1..row2 + 1 {
                        for column in column1..(column2 + 1) {
                            match self.evaluate_cell(CellReferenceIndex {
                                sheet: left.sheet,
                                row,
                                column,
                            }) {
                                CalcResult::Number(value) => {
                                    values.push(to_excel_precision_str(value));
                                }
                                CalcResult::String(value) => {
                                    // ignore_empty skips empty TEXT too, not
                                    // just empty cells (Excel semantics;
                                    // fastsheet vendor patch).
                                    if !(ignore_empty && value.is_empty()) {
                                        values.push(value)
                                    }
                                }
                                CalcResult::Boolean(value) => {
                                    if value {
                                        values.push("TRUE".to_string())
                                    } else {
                                        values.push("FALSE".to_string())
                                    }
                                }
                                CalcResult::EmptyCell => {
                                    if !ignore_empty {
                                        values.push("".to_string())
                                    }
                                }
                                error @ CalcResult::Error { .. } => return error,
                                CalcResult::EmptyArg | CalcResult::Range { .. } => {}
                                CalcResult::Array(_) => {
                                    return CalcResult::Error {
                                        error: Error::NIMPL,
                                        origin: cell,
                                        message: "Arrays not supported yet".to_string(),
                                    }
                                }
                            }
                        }
                    }
                }
                error @ CalcResult::Error { .. } => return error,
                CalcResult::String(value) => {
                    // ignore_empty skips empty TEXT arguments too (Excel
                    // semantics; fastsheet vendor patch).
                    if !(ignore_empty && value.is_empty()) {
                        values.push(value)
                    }
                }
                CalcResult::Boolean(value) => {
                    if value {
                        values.push("TRUE".to_string())
                    } else {
                        values.push("FALSE".to_string())
                    }
                }
                CalcResult::EmptyCell => {
                    if !ignore_empty {
                        values.push("".to_string())
                    }
                }
                CalcResult::EmptyArg => {}
                CalcResult::Array(_) => {
                    return CalcResult::Error {
                        error: Error::NIMPL,
                        origin: cell,
                        message: "Arrays not supported yet".to_string(),
                    }
                }
            };
        }
        let result = values.join(&delimiter);
        CalcResult::String(result)
    }

    // SUBSTITUTE(text, old_text, new_text, [instance_num])
    pub(crate) fn fn_substitute(&mut self, args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        let arg_count = args.len();
        if !(2..=4).contains(&arg_count) {
            return CalcResult::new_args_number_error(cell);
        }
        let text = match self.get_string(&args[0], cell) {
            Ok(s) => s,
            Err(error) => return error,
        };
        let old_text = match self.get_string(&args[1], cell) {
            Ok(s) => s,
            Err(error) => return error,
        };
        let new_text = match self.get_string(&args[2], cell) {
            Ok(s) => s,
            Err(error) => return error,
        };
        let instance_num = if arg_count > 3 {
            match self.get_number(&args[3], cell) {
                Ok(f) => Some(f.floor() as i32),
                Err(s) => return s,
            }
        } else {
            // means every instance is replaced
            None
        };
        if let Some(num) = instance_num {
            if num < 1 {
                return CalcResult::Error {
                    error: Error::VALUE,
                    origin: cell,
                    message: "Invalid value".to_string(),
                };
            }
            if old_text.is_empty() {
                return CalcResult::String(text);
            }
            CalcResult::String(substitute(&text, &old_text, &new_text, num))
        } else {
            if old_text.is_empty() {
                return CalcResult::String(text);
            }
            CalcResult::String(text.replace(&old_text, &new_text))
        }
    }
    pub(crate) fn fn_concatenate(&mut self, args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        let arg_count = args.len();
        if arg_count == 0 {
            return CalcResult::new_args_number_error(cell);
        }
        let mut text_array = Vec::new();
        for arg in args {
            let text = match self.get_string(arg, cell) {
                Ok(s) => s,
                Err(error) => return error,
            };
            text_array.push(text)
        }
        CalcResult::String(text_array.join(""))
    }

    pub(crate) fn fn_exact(&mut self, args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        if args.len() != 2 {
            return CalcResult::new_args_number_error(cell);
        }
        let result1 = &self.evaluate_node_in_context(&args[0], cell);
        let result2 = &self.evaluate_node_in_context(&args[1], cell);
        // FIXME: Implicit intersection
        if let (CalcResult::Number(number1), CalcResult::Number(number2)) = (result1, result2) {
            // In Excel two numbers are the same if they are the same up to 15 digits.
            CalcResult::Boolean(to_precision(*number1, 15) == to_precision(*number2, 15))
        } else {
            let string1 = match self.cast_to_string(result1.clone(), cell) {
                Ok(s) => s,
                Err(error) => return error,
            };
            let string2 = match self.cast_to_string(result2.clone(), cell) {
                Ok(s) => s,
                Err(error) => return error,
            };
            CalcResult::Boolean(string1 == string2)
        }
    }
    // VALUE(text)
    pub(crate) fn fn_value(&mut self, args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        if args.len() != 1 {
            return CalcResult::new_args_number_error(cell);
        }
        match self.evaluate_node_in_context(&args[0], cell) {
            CalcResult::String(text) => {
                let currencies = vec!["$", "€"];
                if let Ok((value, _)) = parse_formatted_number(&text, &currencies, self.locale) {
                    return CalcResult::Number(value);
                };
                CalcResult::Error {
                    error: Error::VALUE,
                    origin: cell,
                    message: "Invalid number".to_string(),
                }
            }
            CalcResult::Number(f) => CalcResult::Number(f),
            CalcResult::Boolean(_) => CalcResult::Error {
                error: Error::VALUE,
                origin: cell,
                message: "Invalid number".to_string(),
            },
            error @ CalcResult::Error { .. } => error,
            CalcResult::Range { .. } => {
                // TODO Implicit Intersection
                CalcResult::Error {
                    error: Error::VALUE,
                    origin: cell,
                    message: "Invalid number".to_string(),
                }
            }
            CalcResult::EmptyCell | CalcResult::EmptyArg => CalcResult::Number(0.0),
            CalcResult::Array(_) => CalcResult::Error {
                error: Error::NIMPL,
                origin: cell,
                message: "Arrays not supported yet".to_string(),
            },
        }
    }

    pub(crate) fn fn_t(&mut self, args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        if args.len() != 1 {
            return CalcResult::new_args_number_error(cell);
        }
        // FIXME: Implicit intersection
        let result = self.evaluate_node_in_context(&args[0], cell);
        match result {
            CalcResult::String(_) => result,
            error @ CalcResult::Error { .. } => error,
            _ => CalcResult::String("".to_string()),
        }
    }

    // VALUETOTEXT(value)
    pub(crate) fn fn_valuetotext(&mut self, args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        if args.len() != 1 {
            return CalcResult::new_args_number_error(cell);
        }
        let text = match self.evaluate_node_in_context(&args[0], cell) {
            // Unlike the `&` operator, VALUETOTEXT renders numbers as plain
            // decimal in Excel (VALUETOTEXT(7.123456E-9) is
            // "0.000000007123456", not "7.123456E-09")
            CalcResult::Number(v) => format!("{}", to_precision(v, 15)),
            other => match self.cast_to_string(other, cell) {
                Ok(s) => s,
                Err(error) => match error {
                    CalcResult::Error { error, .. } => error.to_string(),
                    _ => "".to_string(),
                },
            },
        };
        CalcResult::String(text)
    }
}
