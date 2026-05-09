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

/// One BIFF record in its raw form. The writer can copy these straight
/// into the output stream when we want to passthrough features the
/// greenfield writer doesn't model (drawings, data validation,
/// AutoFilter, sheet protection, conditional formatting, print
/// settings, page-layout-view, theme/XFEXT/STYLEEXT, etc.).
#[derive(Debug, Clone)]
pub struct RawRecord {
    pub opcode: u16,
    pub data: Vec<u8>,
}

/// Records of one substream from the source `/Workbook`. `bof_dt` is
/// the BOF dt field — 0x05 globals, 0x10 worksheet, 0x20 chart, 0x40
/// macro sheet. `records` includes everything between BOF (inclusive)
/// and EOF (inclusive). `index_in_source` is the original substream
/// index in source order; needed because BOUNDSHEET8 entries in
/// globals point at substream lbPlyPos, and we need to know which
/// source substream corresponds to which IronCalc sheet during the
/// splice phase.
#[derive(Debug, Clone)]
pub struct SubstreamSnapshot {
    pub bof_dt: u16,
    pub records: Vec<RawRecord>,
    pub index_in_source: usize,
}

/// Bundle of preserved entries from one source workbook. Iterating
/// `entries` in order recreates the VBA storage tree (parents before
/// children — the extractor walks in CFB preorder).
/// `workbook_substreams` is the per-substream record list of the
/// source `/Workbook`, used by the writer's preservation splice path.
#[derive(Debug, Clone, Default)]
pub struct PreservedXlsData {
    pub entries: Vec<PreservedEntry>,
    pub workbook_substreams: Vec<SubstreamSnapshot>,
}

impl PreservedXlsData {
    pub fn is_empty(&self) -> bool {
        self.entries.is_empty() && self.workbook_substreams.is_empty()
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
    // Pull the workbook stream and split it into per-substream record
    // lists. Used by xls_save's preservation splice to copy through
    // features the greenfield writer doesn't model. Empty when the
    // file isn't a valid xls.
    let workbook_substreams = if cfb.exists("/Workbook") {
        let mut wb_bytes = Vec::new();
        match cfb.open_stream("/Workbook").and_then(|mut s| s.read_to_end(&mut wb_bytes).map(|_| ())) {
            Ok(()) => parse_substreams(&wb_bytes),
            Err(_) => Vec::new(),
        }
    } else {
        Vec::new()
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
    PreservedXlsData { entries, workbook_substreams }
}

/// Walk the workbook stream byte-by-byte, splitting it at each BOF
/// (0x0809) into a separate substream snapshot. Returns an empty
/// vector if the stream is malformed; the caller treats absence of
/// preserved records as "no passthrough available," which falls back
/// to the greenfield writer's own emission.
fn parse_substreams(bytes: &[u8]) -> Vec<SubstreamSnapshot> {
    const R_BOF: u16 = 0x0809;
    const R_EOF: u16 = 0x000A;
    let mut out = Vec::new();
    let mut i = 0usize;
    let mut current: Option<SubstreamSnapshot> = None;
    while i + 4 <= bytes.len() {
        let opcode = u16::from_le_bytes([bytes[i], bytes[i + 1]]);
        let size = u16::from_le_bytes([bytes[i + 2], bytes[i + 3]]) as usize;
        if i + 4 + size > bytes.len() {
            break;
        }
        let data = bytes[i + 4..i + 4 + size].to_vec();
        i += 4 + size;
        if opcode == R_BOF {
            // Close previous substream if it lacks an EOF (malformed
            // but tolerable) and start a fresh one.
            if let Some(prev) = current.take() {
                out.push(prev);
            }
            // BOF data layout: vers(2) + dt(2) + ... — `dt` is what
            // tells us whether this is globals/worksheet/chart/macro.
            let dt = if size >= 4 {
                u16::from_le_bytes([data[2], data[3]])
            } else {
                0
            };
            current = Some(SubstreamSnapshot {
                bof_dt: dt,
                records: vec![RawRecord { opcode, data }],
                index_in_source: out.len(),
            });
            continue;
        }
        if let Some(snap) = current.as_mut() {
            snap.records.push(RawRecord { opcode, data });
            if opcode == R_EOF {
                out.push(current.take().unwrap());
            }
        }
        // Records before any BOF (shouldn't happen in a valid file)
        // are silently dropped.
    }
    if let Some(prev) = current {
        out.push(prev);
    }
    out
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
