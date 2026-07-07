/// Parse Roman (classic or Excel variants) → number
pub fn from_roman(s: &str) -> Result<u32, String> {
    if s.is_empty() {
        return Err("empty numeral".into());
    }
    fn val(c: char) -> Option<u32> {
        Some(match c {
            'I' => 1,
            'V' => 5,
            'X' => 10,
            'L' => 50,
            'C' => 100,
            'D' => 500,
            'M' => 1000,
            _ => return None,
        })
    }

    // Accept the union of subtractive pairs used by the tables above (Excel-compatible).
    fn allowed_subtractive(a: char, b: char) -> bool {
        matches!(
            (a, b),
            // classic:
            ('I','V')|('I','X')|('X','L')|('X','C')|('C','D')|('C','M')
            // Excel forms:
            |('V','L')|('L','D')|('L','M') // VL, LD, LM
            |('X','D')|('X','M')          // XD, XM
            |('V','M')                    // VM
            |('I','L')|('I','C')|('I','D')|('I','M') // IL, IC, ID, IM
            |('V','D')|('V','C') // VD, VC
        )
    }

    let chars: Vec<char> = s.chars().map(|c| c.to_ascii_uppercase()).collect();
    let mut total = 0u32;
    let mut i = 0usize;

    // Repetition rules similar to classic Romans:
    // V, L, D cannot repeat; I, X, C, M max 3 in a row.
    let mut last_char: Option<char> = None;
    let mut run_len = 0usize;

    while i < chars.len() {
        let c = chars[i];
        let v = val(c).ok_or_else(|| format!("invalid character '{c}'"))?;

        if Some(c) == last_char {
            run_len += 1;
            match c {
                'V' | 'L' | 'D' => return Err(format!("invalid repetition of '{c}'")),
                _ if run_len >= 3 => return Err(format!("invalid repetition of '{c}'")),
                _ => {}
            }
        } else {
            last_char = Some(c);
            run_len = 0;
        }

        if i + 1 < chars.len() {
            let c2 = chars[i + 1];
            let v2 = val(c2).ok_or_else(|| format!("invalid character '{c2}'"))?;
            if v < v2 {
                if !allowed_subtractive(c, c2) {
                    return Err(format!("invalid subtractive pair '{c}{c2}'"));
                }
                // Disallow stacked subtractives like IIV, XXL:
                if run_len > 0 {
                    return Err(format!("malformed numeral near position {i}"));
                }
                total += v2 - v;
                i += 2;
                last_char = None;
                run_len = 0;
                continue;
            }
        }

        total += v;
        i += 1;
    }
    Ok(total)
}

/// Classic Roman (strict) encoder used as a base for all forms.
fn to_roman(mut n: u32) -> Result<String, String> {
    if !(1..=3999).contains(&n) {
        return Err("value out of range (must be 1..=3999)".into());
    }

    const MAP: &[(u32, &str)] = &[
        (1000, "M"),
        (900, "CM"),
        (500, "D"),
        (400, "CD"),
        (100, "C"),
        (90, "XC"),
        (50, "L"),
        (40, "XL"),
        (10, "X"),
        (9, "IX"),
        (5, "V"),
        (4, "IV"),
        (1, "I"),
    ];

    let mut out = String::with_capacity(15);
    for &(val, sym) in MAP {
        while n >= val {
            out.push_str(sym);
            n -= val;
        }
        if n == 0 {
            break;
        }
    }
    Ok(out)
}

/// Excel/Google Sheets compatible ROMAN(number, [form]) encoder.
/// `form`: 0..=4 (0=Classic, 4=Simplified).
pub fn to_roman_with_form(n: u32, form: i32) -> Result<String, String> {
    let mut s = to_roman(n)?;
    if form == 0 {
        return Ok(s);
    }
    if !(0..=4).contains(&form) {
        return Err("form must be between 0 and 4".into());
    }

    // Base rules (apply for all f >= 1)
    let base_rules: &[(&str, &str)] = &[
        // C(D|M)XC -> L$1XL
        ("CDXC", "LDXL"),
        ("CMXC", "LMXL"),
        // C(D|M)L -> L$1
        ("CDL", "LD"),
        ("CML", "LM"),
        // X(L|C)IX -> V$1IV
        ("XLIX", "VLIV"),
        ("XCIX", "VCIV"),
        // X(L|C)V -> V$1
        ("XLV", "VL"),
        ("XCV", "VC"),
    ];

    // Level 2 extra rules
    let lvl2_rules: &[(&str, &str)] = &[
        // V(L|C)IV -> I$1
        ("VLIV", "IL"),
        ("VCIV", "IC"),
        // L(D|M)XL -> X$1
        ("LDXL", "XD"),
        ("LMXL", "XM"),
        // L(D|M)VL -> X$1V
        ("LDVL", "XDV"),
        ("LMVL", "XMV"),
        // L(D|M)IL -> X$1IX
        ("LDIL", "XDIX"),
        ("LMIL", "XMIX"),
    ];

    // Level 3 extra rules
    let lvl3_rules: &[(&str, &str)] = &[
        // X(D|M)V -> V$1
        ("XDV", "VD"),
        ("XMV", "VM"),
        // X(D|M)IX -> V$1IV
        ("XDIX", "VDIV"),
        ("XMIX", "VMIV"),
    ];

    // Level 4 extra rules
    let lvl4_rules: &[(&str, &str)] = &[
        // V(D|M)IV -> I$1
        ("VDIV", "ID"),
        ("VMIV", "IM"),
    ];

    // Helper to apply a batch of (from -> to) globally, in order.
    fn apply_rules(mut t: String, rules: &[(&str, &str)]) -> String {
        for (from, to) in rules {
            if t.contains(from) {
                t = t.replace(from, to);
            }
        }
        t
    }

    s = apply_rules(s, base_rules);
    if form >= 2 {
        s = apply_rules(s, lvl2_rules);
    }
    if form >= 3 {
        s = apply_rules(s, lvl3_rules);
    }
    if form >= 4 {
        s = apply_rules(s, lvl4_rules);
    }
    Ok(s)
}

/// Rounding directions for [`excel_round`].
pub(crate) enum RoundMode {
    /// ROUND: ties round away from zero
    HalfAwayFromZero,
    /// ROUNDUP: away from zero
    Up,
    /// ROUNDDOWN: toward zero
    Down,
}

/// Round `value` at `digits` decimal places the way Excel does: on the
/// 15-significant-digit *decimal* representation, not on the binary value.
/// Binary scaling gets classic cases wrong (`2.675 * 100` is `267.4999…`,
/// so `ROUND(2.675, 2)` would give 2.67 instead of Excel's 2.68).
pub(crate) fn excel_round(value: f64, digits: f64, mode: RoundMode) -> f64 {
    if value == 0.0 || !value.is_finite() {
        return value;
    }
    let digits = digits.clamp(-400.0, 400.0) as i64;
    // Exact 15-significant-digit decimal view: m holds all 15 digits, e is
    // the decimal exponent ("2.67500000000000e0" -> m = 2675…0, e = 0).
    let s = format!("{:.14e}", value.abs());
    let (m, e) = match s.split_once('e') {
        Some((mantissa, exp)) => match (mantissa.replace('.', "").parse::<i64>(), exp.parse::<i64>())
        {
            (Ok(m), Ok(e)) => (m, e),
            _ => return value,
        },
        None => return value,
    };
    // Number of leading decimal digits kept when rounding at 10^-digits
    let kept = e + digits + 1;
    if kept >= 15 {
        // The rounding unit is below the 15-digit resolution: nothing to drop
        return crate::number_format::to_precision(value, 15);
    }
    let q: i64 = if kept <= 0 {
        // |value| is smaller than the rounding unit
        match mode {
            RoundMode::Down => 0,
            RoundMode::Up => 1,
            RoundMode::HalfAwayFromZero => {
                i64::from(kept == 0 && m >= 5 * 10_i64.pow(14))
            }
        }
    } else {
        let div = 10_i64.pow((15 - kept) as u32);
        let (q, r) = (m / div, m % div);
        match mode {
            RoundMode::Down => q,
            RoundMode::Up => q + i64::from(r > 0),
            RoundMode::HalfAwayFromZero => q + i64::from(r >= div / 2),
        }
    };
    if q == 0 {
        return 0.0;
    }
    // Reconstruct decimally so we land on the nearest double to the exact
    // decimal result (multiplying by powers of ten would reintroduce error).
    let sign = if value < 0.0 { "-" } else { "" };
    format!("{sign}{q}e{}", -digits)
        .parse::<f64>()
        .unwrap_or(value)
}
