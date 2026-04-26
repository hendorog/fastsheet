/// WebKit2GTK's DMABuf renderer + accelerated compositing don't work under
/// WSLg — the window stays unmapped silently. webkit reads these env vars
/// during library init (before main runs), so setting them with set_var
/// inside main is too late. Re-exec self with the env vars in place.
#[cfg(target_os = "linux")]
pub(crate) fn apply_wsl_webkit_workaround() {
    use std::os::unix::process::CommandExt;
    if std::env::var_os("FASTSHEET_WSL_FIXED").is_some() {
        return;
    }
    let on_wsl = std::env::var_os("WSL_DISTRO_NAME").is_some()
        || std::fs::read_to_string("/proc/version")
            .map(|s| s.contains("microsoft") || s.contains("WSL"))
            .unwrap_or(false);
    if !on_wsl {
        return;
    }
    let exe = match std::env::current_exe() {
        Ok(p) => p,
        Err(_) => return,
    };
    let args: Vec<String> = std::env::args().skip(1).collect();
    let err = std::process::Command::new(exe)
        .args(args)
        .env("FASTSHEET_WSL_FIXED", "1")
        .env("WEBKIT_DISABLE_DMABUF_RENDERER", "1")
        .env("WEBKIT_DISABLE_COMPOSITING_MODE", "1")
        .env("GDK_BACKEND", "x11")
        .exec();
    eprintln!("fastsheet: re-exec failed: {err}");
}

#[cfg(not(target_os = "linux"))]
pub(crate) fn apply_wsl_webkit_workaround() {}
