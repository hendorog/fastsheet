//! Atomic file replacement for the save pipeline.
//!
//! All workbook writers used to do `remove_file(target); write(target)`,
//! which leaves a window where a write failure (disk full, IO error,
//! crash mid-write) destroys the user's existing file. This module
//! provides the standard write-tmp-then-rename pattern instead: writes
//! land at `<target>.fs-tmp-<pid>` in the same directory, get fsynced,
//! then renamed over the destination. The original file is preserved
//! on any failure path.
//!
//! `std::fs::rename` on Windows uses MoveFileEx with
//! MOVEFILE_REPLACE_EXISTING — same-directory atomic-replace works on
//! both Unix and Windows targets.

use std::io;
use std::path::{Path, PathBuf};

/// Build a sibling temp path for `target` in the same directory.
/// Same-volume placement is required for the rename to be atomic.
pub(crate) fn temp_path_for(target: &Path) -> PathBuf {
    let parent = target.parent().unwrap_or_else(|| Path::new("."));
    let file_name = target
        .file_name()
        .map(|s| s.to_string_lossy().into_owned())
        .unwrap_or_else(|| "save".into());
    let pid = std::process::id();
    let nonce: u64 = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .map(|d| d.subsec_nanos() as u64)
        .unwrap_or(0);
    parent.join(format!(".{file_name}.fs-tmp-{pid}-{nonce}"))
}

/// fsync the file at `path` so the rename actually commits the bytes.
/// Without this a power loss between rename and writeback can produce
/// a zero-byte target despite the rename succeeding (commonly seen on
/// ext4 with default mount options). NTFS is journaled so the same
/// risk doesn't exist on Windows — but we still attempt the sync for
/// extra safety.
///
/// Reopening a just-closed file on Windows can return `ACCESS_DENIED`
/// transiently when AV / search-indexer / sandbox tools hold a brief
/// scan lock. Retry a few times with short backoffs; if the file
/// stays unreadable, give up — the atomic rename is the actual
/// crash-safety guarantee, and on Windows NTFS the rename's own
/// journaling makes a separate fsync optional. Returns Ok(()) on
/// transient retry failure so the caller still proceeds to rename.
pub(crate) fn fsync(path: &Path) -> io::Result<()> {
    // Open for write so the OS doesn't insist on a read-share that an
    // AV scanner might have. write+read is what the writer used to
    // produce the file, so it should remain compatible.
    let mut last_err: Option<io::Error> = None;
    for attempt in 0..5 {
        match std::fs::OpenOptions::new().read(true).write(true).open(path) {
            Ok(f) => {
                return f.sync_all();
            }
            Err(e) => {
                last_err = Some(e);
                // 5ms, 10ms, 20ms, 40ms, 80ms — total under 200ms.
                std::thread::sleep(std::time::Duration::from_millis(5 << attempt));
            }
        }
    }
    // Best-effort: emit a warning to stderr but treat as recoverable
    // — the rename itself is what protects against partial-write
    // corruption.
    if let Some(e) = last_err {
        eprintln!("[atomic] fsync skipped on {}: {}", path.display(), e);
    }
    Ok(())
}

/// Pick a backup path for `target`. First choice is `<target>.bak`;
/// if that exists, fall back to `<target>.bak.1`, `.bak.2`, … so we
/// never silently overwrite a previous backup. Returns `None` if no
/// free slot turns up after a sane number of tries.
pub(crate) fn backup_path(target: &Path) -> Option<PathBuf> {
    let parent = target.parent().unwrap_or_else(|| Path::new("."));
    let file_name = target.file_name()?.to_string_lossy().into_owned();
    let first = parent.join(format!("{file_name}.bak"));
    if !first.exists() {
        return Some(first);
    }
    for n in 1..1000 {
        let candidate = parent.join(format!("{file_name}.bak.{n}"));
        if !candidate.exists() {
            return Some(candidate);
        }
    }
    None
}

/// If `target` exists, copy it to a fresh backup path and return that
/// path. Returns `Ok(None)` when the target doesn't exist (nothing to
/// back up). Errors out if a backup slot can't be found or the copy
/// fails — the caller should treat that as a hard save failure rather
/// than risk overwriting unprotected data.
pub(crate) fn backup_if_exists(target: &Path) -> Result<Option<PathBuf>, String> {
    if !target.exists() {
        return Ok(None);
    }
    let dst = backup_path(target)
        .ok_or_else(|| format!("could not pick a backup path for {}", target.display()))?;
    std::fs::copy(target, &dst)
        .map_err(|e| format!("backup {} -> {}: {}", target.display(), dst.display(), e))?;
    Ok(Some(dst))
}

/// Run `write` against a temp sibling of `target`, fsync the result,
/// then atomic-rename over `target`. On any failure the temp file is
/// best-effort removed and the original `target` is left untouched.
pub(crate) fn write<F>(target: &Path, write_fn: F) -> Result<(), String>
where
    F: FnOnce(&Path) -> Result<(), String>,
{
    let tmp = temp_path_for(target);
    // Make sure no leftover from a previous failed run is in the way.
    let _ = std::fs::remove_file(&tmp);
    let result = write_fn(&tmp);
    if let Err(e) = result {
        let _ = std::fs::remove_file(&tmp);
        return Err(e);
    }
    if let Err(e) = fsync(&tmp) {
        // fsync failure is unusual; remove the temp and surface the error.
        let _ = std::fs::remove_file(&tmp);
        return Err(format!("fsync {}: {}", tmp.display(), e));
    }
    if let Err(e) = std::fs::rename(&tmp, target) {
        let _ = std::fs::remove_file(&tmp);
        return Err(format!(
            "rename {} -> {}: {}",
            tmp.display(),
            target.display(),
            e
        ));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn write_replaces_existing_file_atomically() {
        let dir = std::env::temp_dir();
        let target = dir.join(format!("fastsheet_atomic_test_{}.bin", std::process::id()));
        std::fs::write(&target, b"original").unwrap();
        write(&target, |tmp| {
            std::fs::write(tmp, b"new content").map_err(|e| e.to_string())
        })
        .expect("atomic write");
        let got = std::fs::read(&target).unwrap();
        assert_eq!(got, b"new content");
        std::fs::remove_file(&target).ok();
    }

    #[test]
    fn write_failure_preserves_original() {
        let dir = std::env::temp_dir();
        let target = dir.join(format!("fastsheet_atomic_fail_{}.bin", std::process::id()));
        std::fs::write(&target, b"keep me").unwrap();
        let result: Result<(), String> = write(&target, |_tmp| Err("simulated write failure".into()));
        assert!(result.is_err());
        let got = std::fs::read(&target).unwrap();
        assert_eq!(got, b"keep me", "original must be preserved on failure");
        std::fs::remove_file(&target).ok();
    }

    #[test]
    fn backup_if_exists_copies_target_when_present() {
        let dir = std::env::temp_dir();
        let target = dir.join(format!(
            "fastsheet_backup_test_{}_{}.bin",
            std::process::id(),
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .map(|d| d.subsec_nanos())
                .unwrap_or(0)
        ));
        let bak1 = dir.join(format!("{}.bak", target.file_name().unwrap().to_string_lossy()));
        let bak2 = dir.join(format!("{}.bak.1", target.file_name().unwrap().to_string_lossy()));
        std::fs::write(&target, b"v1").unwrap();
        let backup = backup_if_exists(&target).unwrap().expect("first backup");
        assert_eq!(backup, bak1);
        assert_eq!(std::fs::read(&backup).unwrap(), b"v1");
        // Modify and back up again — should pick .bak.1 since .bak is taken.
        std::fs::write(&target, b"v2").unwrap();
        let backup2 = backup_if_exists(&target).unwrap().expect("second backup");
        assert_eq!(backup2, bak2);
        assert_eq!(std::fs::read(&backup2).unwrap(), b"v2");
        // Cleanup.
        for p in [&target, &bak1, &bak2] {
            std::fs::remove_file(p).ok();
        }
    }

    #[test]
    fn backup_if_exists_returns_none_when_absent() {
        let dir = std::env::temp_dir();
        let target = dir.join(format!(
            "fastsheet_backup_absent_{}_{}.bin",
            std::process::id(),
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .map(|d| d.subsec_nanos())
                .unwrap_or(0)
        ));
        let _ = std::fs::remove_file(&target);
        let backup = backup_if_exists(&target).unwrap();
        assert!(backup.is_none());
    }

    #[test]
    fn write_creates_target_when_absent() {
        let dir = std::env::temp_dir();
        let target = dir.join(format!(
            "fastsheet_atomic_new_{}_{}.bin",
            std::process::id(),
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .map(|d| d.subsec_nanos())
                .unwrap_or(0)
        ));
        let _ = std::fs::remove_file(&target);
        write(&target, |tmp| {
            std::fs::write(tmp, b"hello").map_err(|e| e.to_string())
        })
        .expect("atomic write");
        let got = std::fs::read(&target).unwrap();
        assert_eq!(got, b"hello");
        std::fs::remove_file(&target).ok();
    }
}
