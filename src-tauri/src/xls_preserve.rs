//! Capture-and-replay for .xls structured-storage subtrees that the
//! greenfield BIFF writer doesn't model — VBA / macro storages plus
//! the top-level OLE property streams (CompObj, SummaryInformation,
//! DocumentSummaryInformation).
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

/// One array-entered (CSE) formula range from the source file. The
/// loader replicates the anchor formula across the spill range as
/// per-cell `MY*(..., dr, dc)` formulas so IronCalc can evaluate
/// them; this snapshot lets the writer reverse that expansion and
/// re-emit the original FORMULA(PtgExp) + ARRAY record structure.
/// Without it, Excel/POI receive per-cell calls whose trailing args
/// are numbers where the VBA/POI UDF expects an anchor reference,
/// and every spill cell except the anchor evaluates wrong.
#[derive(Debug, Clone)]
pub struct PreservedArrayFormula {
    /// Sheet name at load time; the writer resolves it back to a
    /// model sheet index case-insensitively (sheet indices can shift
    /// if sheets are added/removed between load and save).
    pub sheet_name: String,
    /// 1-based inclusive spill rectangle; (r1, c1) is the anchor.
    pub r1: i32,
    pub r2: i32,
    pub c1: i32,
    pub c2: i32,
    /// The ARRAY record's rgce, verbatim from the source. PtgName
    /// inames index the SOURCE file's NAME table and must be
    /// remapped against the writer's rebuilt tables before replay.
    pub rgce: Vec<u8>,
    /// Canonical per-cell formula text (row-major over the range,
    /// via `Model::get_cell_formula`) captured right after load-time
    /// replication. At save the writer re-reads the same cells; any
    /// mismatch means the user edited inside the range and the whole
    /// range falls back to per-cell expansion.
    pub cell_formulas: Vec<Option<String>>,
}

/// One source NAME (Lbl) record the loader could NOT import into the
/// IronCalc model — builtin names (_FilterDatabase = AutoFilter
/// ranges, Print_Titles = print repeat rows/cols), `_xlfn.*` /
/// `_xleta.*` stubs, names IronCalc's identifier grammar rejects
/// (e.g. containing `?`), and #REF! names. The writer replays these
/// as appended Lbl records with the rgce remapped into its rebuilt
/// NAME / EXTERNSHEET tables; dropping them silently degrades
/// AutoFilter, print setup and any later reference to the name.
#[derive(Debug, Clone)]
pub struct PreservedName {
    /// Raw name string; a single control char for builtin names.
    pub name: String,
    pub grbit: u16,
    pub ch_key: u8,
    /// Scope sheet resolved from the Lbl itab at load time; None =
    /// workbook-global. Stored as a name so sheet insertion/removal
    /// between load and save can't silently retarget the scope.
    pub sheet_name: Option<String>,
    /// Source rgce verbatim; PtgName inames / Ptg*3d ixtis index the
    /// SOURCE tables and are remapped at save.
    pub rgce: Vec<u8>,
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
    /// Array-formula ranges the loader replicated (see
    /// `PreservedArrayFormula`). Populated by `load_xls`, not by
    /// `extract` — the data comes from the BIFF shape scan plus the
    /// post-replication model, neither of which `extract` sees.
    pub array_formulas: Vec<PreservedArrayFormula>,
    /// The source workbook's defined names in BIFF NAME-record order.
    /// PtgName inames in any preserved rgce (arrays, error cells,
    /// preserved names) are 1-based indices into this list.
    pub source_defined_names: Vec<String>,
    /// Source NAME records the loader couldn't import (see
    /// `PreservedName`). Populated by `load_xls`.
    pub preserved_names: Vec<PreservedName>,
    /// Per source EXTERNSHEET xti index: the (first, last) sheet
    /// names for internal-supbook single/multi-sheet entries, `None`
    /// for add-in / external-book / deleted-sheet entries. Lets the
    /// writer remap a preserved rgce's Ptg*3d ixti into its own
    /// rebuilt table by sheet name.
    pub source_xti_sheets: Vec<Option<(String, String)>>,
}

impl PreservedXlsData {
    pub fn is_empty(&self) -> bool {
        self.entries.is_empty()
            && self.workbook_substreams.is_empty()
            && self.array_formulas.is_empty()
            && self.preserved_names.is_empty()
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

/// Top-level OLE property streams preserved verbatim. These are
/// self-contained (no references into the workbook stream):
/// SummaryInformation / DocumentSummaryInformation carry document
/// metadata (author, title, company, custom properties); CompObj
/// identifies the document class for OLE embedding. Dropping them
/// loses metadata and shell integration (2026-08-14 report, Issue 3).
const PROPERTY_STREAMS: &[&str] = &[
    "/\u{1}CompObj",
    "/\u{5}SummaryInformation",
    "/\u{5}DocumentSummaryInformation",
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
    for path in PROPERTY_STREAMS {
        if !cfb.exists(path) || cfb.is_storage(path) {
            continue;
        }
        let mut data = Vec::new();
        let ok = cfb
            .open_stream(path)
            .and_then(|mut s| s.read_to_end(&mut data).map(|_| ()))
            .is_ok();
        if ok {
            entries.push(PreservedEntry {
                path: PathBuf::from(path),
                is_storage: false,
                data: Some(data),
            });
        }
    }
    PreservedXlsData {
        entries,
        workbook_substreams,
        array_formulas: Vec::new(),
        source_defined_names: Vec::new(),
        preserved_names: Vec::new(),
        source_xti_sheets: Vec::new(),
    }
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
    // Depth tracker. BOF/EOF nest in the workbook stream because chart
    // substreams CAN be embedded inside worksheet substreams (per
    // [MS-XLS] §2.1.7.20.1), in which case the inner BOF/EOF are
    // records of the parent — not substream boundaries. Only BOFs at
    // depth 0 start a new top-level substream; only EOFs that close
    // the outermost open substream end one.
    let mut depth = 0u32;
    while i + 4 <= bytes.len() {
        let opcode = u16::from_le_bytes([bytes[i], bytes[i + 1]]);
        let size = u16::from_le_bytes([bytes[i + 2], bytes[i + 3]]) as usize;
        if i + 4 + size > bytes.len() {
            break;
        }
        let data = bytes[i + 4..i + 4 + size].to_vec();
        i += 4 + size;
        if opcode == R_BOF {
            if depth == 0 {
                // Top-level BOF: start a new substream.
                if let Some(prev) = current.take() {
                    out.push(prev);
                }
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
                depth = 1;
                continue;
            }
            // Embedded BOF — keep it as a record of the parent.
            depth += 1;
        } else if opcode == R_EOF {
            // Treat EOF symmetrically: only the outermost EOF closes
            // the substream.
            if depth > 1 {
                depth -= 1;
                if let Some(snap) = current.as_mut() {
                    snap.records.push(RawRecord { opcode, data });
                }
                continue;
            }
            if depth == 1 {
                if let Some(snap) = current.as_mut() {
                    snap.records.push(RawRecord { opcode, data });
                }
                if let Some(prev) = current.take() {
                    out.push(prev);
                }
                depth = 0;
                continue;
            }
            // EOF at depth 0 — orphan; drop.
            continue;
        }
        if let Some(snap) = current.as_mut() {
            snap.records.push(RawRecord { opcode, data });
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
