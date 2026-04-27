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

/// Process-start instant. Captured the first time `app_start_instant`
/// is called (typically the very top of `lib::run`). Used by phases
/// further down the boot path to compute "time since process start"
/// for the profile log.
pub(crate) fn app_start_instant() -> std::time::Instant {
    use std::sync::OnceLock;
    static START: OnceLock<std::time::Instant> = OnceLock::new();
    *START.get_or_init(std::time::Instant::now)
}

/// Append a profile line to the profile log file when
/// FASTSHEET_PROFILE_LOAD is set. Targets a log file because the
/// Tauri GUI binary is built with subsystem=windows on Windows and
/// has no stderr; writing to stderr there discards the output.
///
/// Default log path: `fastsheet_profile.log` next to the running
/// binary (whatever `std::env::current_exe()` resolves to). Falls
/// back to the OS temp dir if current_exe fails. Override the path
/// entirely with `FASTSHEET_PROFILE_LOG`. Best-effort; failures are
/// silent.
pub(crate) fn profile_log(line: &str) {
    if std::env::var("FASTSHEET_PROFILE_LOAD").is_err() {
        return;
    }
    let path = std::env::var("FASTSHEET_PROFILE_LOG").unwrap_or_else(|_| {
        std::env::current_exe()
            .ok()
            .and_then(|p| p.parent().map(|d| d.join("fastsheet_profile.log")))
            .unwrap_or_else(|| std::env::temp_dir().join("fastsheet_profile.log"))
            .to_string_lossy()
            .into_owned()
    });
    if let Ok(mut f) = std::fs::OpenOptions::new().create(true).append(true).open(&path) {
        use std::io::Write;
        let _ = writeln!(f, "{line}");
    }
    // Also try stderr — visible when launched from a console build.
    eprintln!("{line}");
}
