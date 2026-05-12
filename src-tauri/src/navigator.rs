use std::path::{Path, PathBuf};

use serde::Serialize;

#[derive(Serialize)]
pub(crate) struct DirEntry {
    name: String,
    is_dir: bool,
    /// Unix epoch seconds for last modification, or None if not available.
    modified: Option<i64>,
    /// File size in bytes for files; None for directories.
    size: Option<u64>,
}

#[derive(Serialize)]
pub(crate) struct DirListing {
    /// The canonical absolute directory listed.
    dir: String,
    /// Path to the parent directory (if any). `..` selection should jump here.
    parent: Option<String>,
    /// Directories first (alphabetical), then spreadsheet files
    /// (.xlsx / .xls, alphabetical).
    entries: Vec<DirEntry>,
}

pub(crate) fn home_dir() -> Option<PathBuf> {
    #[cfg(windows)]
    {
        if let Ok(p) = std::env::var("USERPROFILE") {
            return Some(PathBuf::from(p));
        }
    }
    if let Ok(p) = std::env::var("HOME") {
        return Some(PathBuf::from(p));
    }
    None
}

/// Resolve a user-typed path expression against an optional `cwd`.
/// Handles `~`, `~/...`, drive letters (`c:`), absolute and relative paths.
/// Returns the resolved (but not necessarily canonicalised) path.
fn resolve_input(input: &str, cwd: Option<&Path>) -> Result<PathBuf, String> {
    let s = input.trim();
    if s.is_empty() {
        return cwd.map(PathBuf::from).ok_or_else(|| "empty path".into());
    }
    // Drive letter alone → drive root. Works on Windows; harmless elsewhere
    // (the resulting path will not exist on Linux and list_dir will error).
    if s.len() == 2
        && s.chars().nth(1) == Some(':')
        && s.chars().next().unwrap().is_ascii_alphabetic()
    {
        return Ok(PathBuf::from(format!("{}\\", s)));
    }
    // Home expansion
    if s == "~" {
        return home_dir().ok_or_else(|| "no home dir".into());
    }
    if let Some(rest) = s.strip_prefix("~/").or_else(|| s.strip_prefix("~\\")) {
        let home = home_dir().ok_or_else(|| "no home dir".to_string())?;
        return Ok(home.join(rest));
    }
    let p = PathBuf::from(s);
    if p.is_absolute() {
        return Ok(p);
    }
    Ok(cwd.map(|c| c.join(&p)).unwrap_or(p))
}

#[tauri::command]
pub(crate) fn start_dir() -> Result<String, String> {
    // Prefer the user's home over current_dir — cmd.exe's UNC fallback
    // leaves CWD at C:\Windows\System32 which is a useless place to start.
    if let Some(h) = home_dir() {
        if h.is_dir() {
            return Ok(h.to_string_lossy().into_owned());
        }
    }
    let p = std::env::current_dir().map_err(|e| e.to_string())?;
    Ok(p.to_string_lossy().into_owned())
}

#[tauri::command]
pub(crate) fn home_dir_path() -> Result<String, String> {
    home_dir()
        .map(|p| p.to_string_lossy().into_owned())
        .ok_or_else(|| "no home directory".into())
}

/// Enumerate WSL distros from the Windows registry — doesn't depend on
/// wsl.exe or the LxssManager service being responsive. Every distro
/// registers itself under `HKCU\Software\Microsoft\Windows\CurrentVersion\Lxss\<guid>`
/// with a `DistributionName` REG_SZ value.
#[cfg(windows)]
fn list_wsl_distros() -> Vec<String> {
    use std::process::Command;
    let out = match Command::new("reg.exe")
        .args([
            "query",
            r"HKCU\Software\Microsoft\Windows\CurrentVersion\Lxss",
            "/s",
            "/v",
            "DistributionName",
        ])
        .output()
    {
        Ok(o) => o,
        Err(_) => return Vec::new(),
    };
    let stdout = String::from_utf8_lossy(&out.stdout);
    let mut distros: Vec<String> = stdout
        .lines()
        .filter_map(|line| {
            // Lines look like:  "    DistributionName    REG_SZ    Ubuntu"
            let rest = line.trim_start().strip_prefix("DistributionName")?;
            let rest = rest.trim_start().strip_prefix("REG_SZ")?;
            let name = rest.trim();
            if name.is_empty() {
                None
            } else {
                Some(name.to_string())
            }
        })
        .collect();
    distros.sort_by(|a, b| a.to_lowercase().cmp(&b.to_lowercase()));
    distros.dedup();
    distros
}

#[cfg(not(windows))]
fn list_wsl_distros() -> Vec<String> {
    Vec::new()
}

/// Recognise UNC roots that Rust's `read_dir` can't enumerate directly —
/// `\\wsl.localhost\` and `\\wsl$\` (the WSL share roots), or bare `\\`.
fn is_wsl_unc_root(path: &Path) -> bool {
    let s = path.to_string_lossy().to_lowercase();
    matches!(
        s.as_str(),
        r"\\wsl.localhost" | r"\\wsl.localhost\" | r"\\wsl$" | r"\\wsl$\" | r"\\"
    )
}

/// File-kind hint controlling which extensions `list_dir` surfaces. Workbook
/// pickers (Open / Save As / Compare) show .xlsx + .xls; text pickers (the /D
/// Import flow) show .csv + .tsv + .txt. Unknown values fall back to workbook.
fn allowed_extensions(kind: Option<&str>) -> &'static [&'static str] {
    match kind {
        Some("text") => &["csv", "tsv", "txt"],
        _ => &["xlsx", "xls"],
    }
}

#[tauri::command]
pub(crate) fn list_dir(
    path: String,
    cwd: Option<String>,
    kind: Option<String>,
) -> Result<DirListing, String> {
    let cwd_path = cwd.as_deref().map(Path::new);
    let resolved = resolve_input(&path, cwd_path)?;
    let exts = allowed_extensions(kind.as_deref());
    // Special case: server-root UNC paths like \\wsl.localhost\ — Rust's
    // read_dir can't enumerate these, so we synthesise the listing from
    // wsl.exe instead.
    if is_wsl_unc_root(&resolved) {
        let distros = list_wsl_distros();
        let entries: Vec<DirEntry> = distros
            .into_iter()
            .map(|name| DirEntry {
                name,
                is_dir: true,
                modified: None,
                size: None,
            })
            .collect();
        return Ok(DirListing {
            dir: r"\\wsl.localhost\".to_string(),
            parent: None,
            entries,
        });
    }
    // Try canonicalise so symlinks and `..` normalise. UNC paths and some
    // network shares fail canonicalize but read_dir works on them — fall
    // back to the raw resolved path in that case.
    let canonical = std::fs::canonicalize(&resolved).unwrap_or_else(|_| resolved.clone());
    // Try read_dir first — its error is more informative than is_dir().
    let read = match std::fs::read_dir(&canonical) {
        Ok(r) => r,
        Err(e) => {
            return Err(format!("cannot read {}: {}", canonical.display(), e));
        }
    };
    let mut entries = Vec::new();
    for entry in read {
        let entry = match entry {
            Ok(e) => e,
            Err(_) => continue,
        };
        let name = entry.file_name().to_string_lossy().into_owned();
        if name.starts_with('.') {
            continue; // hide dotfiles for now
        }
        let metadata = match entry.metadata() {
            Ok(m) => m,
            Err(_) => continue,
        };
        let is_dir = metadata.is_dir();
        // Show all directories; for files, filter by the requested kind.
        // Workbook pickers (default) accept .xlsx / .xls. Text pickers
        // (/D Import) accept .csv / .tsv / .txt. Extensions matched
        // case-insensitively.
        if !is_dir {
            let lower = name.to_lowercase();
            let ok = exts.iter().any(|e| {
                lower.len() > e.len() + 1
                    && lower.ends_with(e)
                    && lower.as_bytes()[lower.len() - e.len() - 1] == b'.'
            });
            if !ok {
                continue;
            }
        }
        let modified = metadata
            .modified()
            .ok()
            .and_then(|t| t.duration_since(std::time::UNIX_EPOCH).ok())
            .map(|d| d.as_secs() as i64);
        let size = if is_dir { None } else { Some(metadata.len()) };
        entries.push(DirEntry {
            name,
            is_dir,
            modified,
            size,
        });
    }
    entries.sort_by(|a, b| {
        b.is_dir
            .cmp(&a.is_dir)
            .then_with(|| a.name.to_lowercase().cmp(&b.name.to_lowercase()))
    });
    let parent = canonical
        .parent()
        .map(|p| p.to_string_lossy().into_owned());
    Ok(DirListing {
        dir: canonical.to_string_lossy().into_owned(),
        parent,
        entries,
    })
}
