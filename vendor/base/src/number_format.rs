use crate::{
    formatter::{self, format::Formatted},
    locale::get_locale,
    types::NumFmt,
};

const DEFAULT_NUM_FMTS: &[&str] = &[
    "general",
    "0",
    "0.00",
    "#,##0",
    "#,##0.00",
    "$#,##0; \\ - $#,##0",
    "$#,##0; [Red] \\ - $#,##0",
    "$#,##0.00; \\ - $#,##0.00",
    "$#,##0.00; [Red] \\ - $#,##0.00",
    "0%",
    "0.00%",
    "0.00E + 00",
    "#?/?",
    "#?? / ??",
    "mm-dd-yy",
    "d-mmm-yy",
    "d-mmm",
    "mmm-yy",
    "h:mm AM / PM",
    "h:mm:ss AM / PM",
    "h:mm",
    "h:mm:ss",
    "m / d / yy h:mm",
    "#,##0;()#,##0)",
    "#,##0; [Red]()#,##0)",
    "#,##0.00;()#,##0.00)",
    "#,##0.00; [Red]()#,##0.00)",
    "_()$”*#,##0.00 _); _()$”* \\()#,##0.00\\); _()$”* - ?? _); _()@_)",
    "mm:ss",
    "[h]:mm:ss",
    "mmss .0",
    "##0.0E + 0",
    "@",
    "[$ -404] e / m / d ",
    "m / d / yy",
    "[$ -404] e / m / d",
    "[$ -404] e / / d",
    "[$ -404] e / m / d",
    "t0",
    "t0.00",
    "t#,##0",
    "t#,##0.00",
    "t0%",
    "t0.00 %",
    "t#?/?",
];

pub fn get_default_num_fmt_id(num_fmt: &str) -> Option<i32> {
    for (index, default_num_fmt) in DEFAULT_NUM_FMTS.iter().enumerate() {
        if default_num_fmt == &num_fmt {
            return Some(index as i32);
        };
    }
    None
}

pub fn get_num_fmt(num_fmt_id: i32, num_fmts: &[NumFmt]) -> String {
    // Check if it defined
    for num_fmt in num_fmts {
        if num_fmt.num_fmt_id == num_fmt_id {
            return num_fmt.format_code.clone();
        }
    }
    // Return one of the default ones
    if num_fmt_id < DEFAULT_NUM_FMTS.len() as i32 {
        return DEFAULT_NUM_FMTS[num_fmt_id as usize].to_string();
    }
    // Return general
    DEFAULT_NUM_FMTS[0].to_string()
}

pub fn get_new_num_fmt_index(num_fmts: &[NumFmt]) -> i32 {
    let mut index = DEFAULT_NUM_FMTS.len() as i32;
    let mut found = true;
    while found {
        found = false;
        for num_fmt in num_fmts {
            if num_fmt.num_fmt_id == index {
                found = true;
                index += 1;
                break;
            }
        }
    }
    index
}

pub fn to_precision(value: f64, precision: usize) -> f64 {
    if value.is_infinite() || value.is_nan() {
        return value;
    }
    to_precision_str(value, precision)
        .parse::<f64>()
        .unwrap_or({
            // TODO: do this in a way that does not require a possible error
            0.0
        })
}

/// It rounds a `f64` with `p` significant figures:
/// ```
///     use ironcalc_base::number_format;
///     assert_eq!(number_format::to_precision(0.1+0.2, 15), 0.3);
///     assert_eq!(number_format::to_excel_precision_str(0.1+0.2), "0.3");
/// ```
/// This intends to be equivalent to the js: `${parseFloat(value.toPrecision(precision)})`
/// See ([ecma](https://tc39.es/ecma262/#sec-number.prototype.toprecision)).
pub fn to_excel_precision_str(value: f64) -> String {
    let text = to_precision_str(value, 15);
    if !value.is_finite() || value == 0.0 {
        return text;
    }
    // ryu picks scientific notation by shortness (from 1e16 up, sporadically
    // below 1e-4). Excel renders plain decimal for 1e-4 <= |v| < 1e21 and
    // uppercase scientific with a signed two-digit exponent outside that.
    let (sign, body) = match text.strip_prefix('-') {
        Some(rest) => ("-", rest),
        None => ("", text.as_str()),
    };
    // Extract the significant digits and the decimal exponent of the first one
    let (digits, exp) = match body.split_once(['e', 'E']) {
        Some((mantissa, exp_str)) => {
            let Ok(exp) = exp_str.parse::<i32>() else {
                return text;
            };
            (mantissa.replace('.', ""), exp)
        }
        None => match body.split_once('.') {
            Some((int_part, frac_part)) => {
                if int_part == "0" {
                    let zeros = frac_part.len() - frac_part.trim_start_matches('0').len();
                    (frac_part[zeros..].to_string(), -(zeros as i32) - 1)
                } else {
                    (format!("{int_part}{frac_part}"), int_part.len() as i32 - 1)
                }
            }
            None => (body.to_string(), body.len() as i32 - 1),
        },
    };
    let digits = digits.trim_end_matches('0');
    let digits = if digits.is_empty() { "0" } else { digits };
    if (-4..21).contains(&exp) {
        let mut out = String::from(sign);
        if exp >= 0 {
            let e = exp as usize;
            if digits.len() > e + 1 {
                out.push_str(&digits[..=e]);
                out.push('.');
                out.push_str(&digits[e + 1..]);
            } else {
                out.push_str(digits);
                out.push_str(&"0".repeat(e + 1 - digits.len()));
            }
        } else {
            out.push_str("0.");
            out.push_str(&"0".repeat((-exp - 1) as usize));
            out.push_str(digits);
        }
        out
    } else {
        let mut out = format!("{sign}{}", &digits[..1]);
        if digits.len() > 1 {
            out.push('.');
            out.push_str(&digits[1..]);
        }
        out.push_str(if exp < 0 { "E-" } else { "E+" });
        out.push_str(&format!("{:02}", exp.abs()));
        out
    }
}

pub fn to_excel_precision(value: f64, precision: usize) -> f64 {
    if !value.is_finite() {
        return value;
    }

    let s = format!("{:.*e}", precision.saturating_sub(1), value);
    s.parse::<f64>().unwrap_or(value)
}

pub fn to_precision_str(value: f64, precision: usize) -> String {
    if !value.is_finite() {
        if value.is_infinite() {
            return "inf".to_string();
        } else {
            return "NaN".to_string();
        }
    }

    let s = format!("{:.*e}", precision.saturating_sub(1), value);
    let parsed = s.parse::<f64>().unwrap_or(value);

    // I would love to use the std library. There is not a speed concern here
    // problem is it doesn't do the right thing
    // Also ryu is my favorite _modern_ algorithm
    let mut buffer = ryu::Buffer::new();
    let text = buffer.format(parsed);
    // The above algorithm converts 2 to 2.0 regrettably
    if let Some(stripped) = text.strip_suffix(".0") {
        return stripped.to_string();
    }
    text.to_string()
}

pub fn format_number(value: f64, format_code: &str, locale: &str) -> Formatted {
    let locale = match get_locale(locale) {
        Ok(l) => l,
        Err(_) => {
            return Formatted {
                text: "#ERROR!".to_owned(),
                color: None,
                error: Some("Invalid locale".to_string()),
            }
        }
    };
    formatter::format::format_number(value, format_code, locale)
}
