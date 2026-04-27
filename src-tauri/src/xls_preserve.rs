//! Capture-and-replay for .xls structured-storage subtrees that the
//! greenfield BIFF writer doesn't model — currently just VBA / macro
//! storages.
//!
//! Why this works without offset arithmetic: VBA lives in its own
//! storage tree (`/_VBA_PROJECT_CUR`, optionally `/VBA`, `/Macros`,
//! `/_VBA_PROJECT`) inside the OLE2 compound file. The streams there
//! reference each other (via PROJECTSTREAM / dir / module dependency
//! chains) but those references are **internal to the VBA tree** —
//! they're indexes into other VBA streams, not byte offsets into the
//! workbook BIFF stream. So the bytes can be lifted from the source
//! and dropped into any target CFB without rewriting.
//!
//! The user's case for this: they ship .xls templates with
//! Apache-POI-style pseudo-array UDFs (MROUND, MYUNIQUE, etc.) that
//! Excel needs to evaluate when the file is opened in Excel. fastsheet
//! itself emulates those functions natively (see
//! vendor/base/src/functions/fastsheet_udfs.rs) and never executes the
//! VBA — it just needs to round-trip the macro bytes so Excel users
//! still see working formulas.
//!
//! Anything outside the VBA storages (charts, drawings, pivots,
//! comments, conditional formatting, OLE links) is NOT preserved by
//! this module — those features have BIFF-record-level cross-
//! references with the workbook stream that a fresh write would have
//! reshuffled. Full xls preservation is a bigger lift; tracked as
//! pending #1.
//!
//! The four storage paths covered:
//!   * `/_VBA_PROJECT_CUR`  — Excel 97+ canonical VBA root
//!   * `/_VBA_PROJECT`      — Excel 5/95 form
//!   * `/VBA`               — alternate location some writers use
//!   * `/Macros`            — Excel 4 macro sheets

use std::io::{Cursor, Read};
use std::path::PathBuf;

/// Snapshot of one entry from the source CFB's VBA storage tree.
/// `is_storage` separates the two CFB entry kinds — a storage is a
/// directory; a stream is a leaf with bytes. `data` is `Some` only
/// for streams.
#[derive(Debug, Clone)]
pub struct PreservedEntry {
    pub path: PathBuf,
    pub is_storage: bool,
    pub data: Option<Vec<u8>>,
}

/// Bundle of preserved entries from one source workbook. Iterating in
/// order recreates the storage tree (parents before children — the
/// extractor walks in CFB preorder).
#[derive(Debug, Clone, Default)]
pub struct PreservedXlsData {
    pub entries: Vec<PreservedEntry>,
}

impl PreservedXlsData {
    pub fn is_empty(&self) -> bool {
        self.entries.is_empty()
    }
}

/// Storage paths we preserve from the source. We cover the four known
/// VBA homes; anything else is silently ignored. Order doesn't matter —
/// each subtree is captured independently.
const VBA_STORAGE_ROOTS: &[&str] = &[
    "/_VBA_PROJECT_CUR",
    "/_VBA_PROJECT",
    "/VBA",
    "/Macros",
];

/// Walk the source CFB for VBA-related storage subtrees and capture
/// every storage + stream within them. Returns an empty bundle if the
/// file isn't a valid CFB or has no VBA storages — never errors,
/// because preservation is a best-effort extra on top of the normal
/// load path.
pub fn extract(bytes: &[u8]) -> PreservedXlsData {
    let cursor = Cursor::new(bytes.to_vec());
    let mut cfb = match cfb::CompoundFile::open(cursor) {
        Ok(c) => c,
        Err(_) => return PreservedXlsData::default(),
    };
    let mut entries = Vec::new();
    for root in VBA_STORAGE_ROOTS {
        if !cfb.exists(root) || !cfb.is_storage(root) {
            continue;
        }
        // Collect entry metadata first so the walk iterator (which
        // borrows the cfb immutably) is dropped before we grab a
        // mutable borrow to open each stream below. walk_storage
        // gives preorder traversal — the storage entry itself first,
        // then descendants. That order matches what create_storage_all
        // + create_stream replay expects.
        let metas: Vec<(PathBuf, bool)> = match cfb.walk_storage(root) {
            Ok(w) => w
                .map(|e| (e.path().to_path_buf(), e.is_storage()))
                .collect(),
            Err(_) => continue,
        };
        for (path, is_storage) in metas {
            if is_storage {
                entries.push(PreservedEntry {
                    path,
                    is_storage: true,
                    data: None,
                });
            } else {
                let mut data = Vec::new();
                let mut stream = match cfb.open_stream(&path) {
                    Ok(s) => s,
                    Err(_) => continue,
                };
                if stream.read_to_end(&mut data).is_err() {
                    continue;
                }
                entries.push(PreservedEntry {
                    path,
                    is_storage: false,
                    data: Some(data),
                });
            }
        }
    }
    PreservedXlsData { entries }
}

/// Replay the captured entries into a target CFB. Called by the xls
/// writer after the `/Workbook` stream is in place. Best-effort: a
/// failed inject doesn't fail the save — the workbook itself is
/// already written; we just lose macros for that save.
pub fn inject<F: std::io::Read + std::io::Write + std::io::Seek>(
    cfb: &mut cfb::CompoundFile<F>,
    preserved: &PreservedXlsData,
) {
    if preserved.is_empty() {
        return;
    }
    for entry in &preserved.entries {
        if entry.is_storage {
            // create_storage_all is idempotent on existing storages,
            // so it's safe to call for every storage entry even when
            // a parent was already created via an earlier child.
            let _ = cfb.create_storage_all(&entry.path);
        } else if let Some(data) = &entry.data {
            // Ensure parent storage exists. Won't normally fire — the
            // walk_storage iterator emits parents before children — but
            // it's free insurance against unusual source layouts.
            if let Some(parent) = entry.path.parent() {
                if !parent.as_os_str().is_empty() {
                    let _ = cfb.create_storage_all(parent);
                }
            }
            let mut stream = match cfb.create_stream(&entry.path) {
                Ok(s) => s,
                Err(_) => continue,
            };
            use std::io::Write;
            let _ = stream.write_all(data);
        }
    }
}
