//! Passthrough helpers for the .xls writer's preservation path.
//!
//! When the source workbook was loaded through `xls_load`, we kept a
//! per-substream snapshot of every BIFF record. The greenfield writer
//! re-emits the records IronCalc's model knows about (cells, fonts,
//! XFs, formats, BOUNDSHEET8s, etc.); this module identifies the
//! records that the writer DOES NOT model — drawings, data
//! validation, AutoFilter, sheet protection, conditional formatting,
//! print settings, page-layout-view, etc. — and supplies them to the
//! writer as raw bytes for verbatim copy-through.
//!
//! The splice happens at three positions per sheet substream:
//!
//! 1. **Pre-DIMENSIONS zone** — page-setup records that appear
//!    between WSBOOL and COLINFO/DIMENSIONS in the source: classic
//!    HEADER/FOOTER, HCENTER, VCENTER, margin records, SETUP, PLS,
//!    page breaks, sheet-protection PROTECT/PASSWORD, etc.
//!
//! 2. **Post-cells / pre-WINDOW2 zone** — records that appear
//!    between the last cell and WINDOW2 in source: MERGECELLS,
//!    CONDFMT/CF, HLINK, DVAL/DV, PHONETICINFO, FEAT, etc.
//!
//! 3. **Post-WINDOW2 zone** — drawings (MSODRAWING/OBJ/TXO +
//!    CONTINUE), SHEETPROTECTION (FRT), RANGEPROTECTION, FRT
//!    HEADERFOOTER, PLV, FORCEFULLCALCULATION, FILTERMODE/
//!    AUTOFILTERINFO/AUTOFILTER, etc. WINDOW2/PANE/SELECTION
//!    themselves are filtered out — the writer emits its own from
//!    the IronCalc model so frozen-pane state stays consistent
//!    with model edits.
//!
//! Globals get their own simpler passthrough — we collect any
//! "trailing" records the writer doesn't emit (MSODRAWINGGROUP,
//! THEME, XFEXT, STYLEEXT, DXF, TABLESTYLES, etc.) and append them
//! before EOF.

use crate::xls_preserve::{RawRecord, SubstreamSnapshot};
use std::collections::HashMap;

// Per-sheet record opcodes the writer owns. We strip these from
// passthrough zones because the writer re-emits them from the
// IronCalc model. Anything not in this set passes through verbatim.
//
// BOF/EOF are intentionally NOT in this set: at the substream
// boundary the zone-splitting logic excludes them (zones start at
// index 1 and stop before the trailing EOF), but EMBEDDED BOF/EOF
// pairs — chart substreams nested inside a worksheet substream per
// [MS-XLS] §2.1.7.20.1 — must pass through verbatim, otherwise the
// embedded chart loses its framing on save.
const OWNED_PER_SHEET: &[u16] = &[
    // Calculation / display headers
    0x000D, // CALCMODE
    0x000C, // CALCCOUNT
    0x000F, // REFMODE
    0x0011, // ITERATION
    0x0010, // DELTA
    0x005F, // SAVERECALC
    0x002A, // PRINTHEADERS
    0x002B, // PRINTGRIDLINES
    0x0082, // GRIDSET
    0x0080, // GUTS
    0x0225, // DEFAULTROWHEIGHT
    0x0081, // WSBOOL
    // Column info
    0x0055, // DEFCOLWIDTH
    0x007D, // COLINFO
    // Dimensions / cell table
    0x0200, // DIMENSIONS
    0x020B, // INDEX
    0x00D7, // DBCELL
    0x0208, // ROW
    // Cell records
    0x0201, // BLANK
    0x0203, // NUMBER
    0x027E, // RK
    0x00BD, // MULRK
    0x00BE, // MULBLANK
    0x00FD, // LABELSST
    0x0204, // LABEL
    0x0006, // FORMULA
    0x0207, // STRING
    0x04BC, // SHRFMLA
    0x0221, // ARRAY
    0x0205, // BOOLERR
    0x0036, // TABLEOP
    0x0037, // TABLEOP2
    // View / freeze
    0x023E, // WINDOW2
    0x0041, // PANE
    0x001D, // SELECTION
    // Merged ranges — the writer re-emits MERGECELLS from the model's
    // merge list (populated on load, mutated by set_sheet_merges), so
    // the source record would be stale after a /Move or /Range Merge.
    0x00E5, // MERGECELLS
];

// Globals records the writer owns. Anything else in the source
// globals substream (THEME, XFEXT, STYLEEXT, DXF, TABLESTYLES,
// XFCRC, MSODRAWINGGROUP, FORCEFULLCALCULATION, EXCEL9FILE,
// RECALCID, FNGROUPCOUNT, TABID, etc.) gets passed through.
const OWNED_GLOBALS: &[u16] = &[
    0x0809, // BOF
    0x000A, // EOF
    0x00E1, // INTERFACEHDR
    0x00E2, // INTERFACEEND
    0x00C1, // MMS
    0x005C, // WRITEACCESS
    0x0042, // CODEPAGE
    0x0161, // DSF
    0x01B7, // REFRESHALL
    0x00DA, // BOOKBOOL
    0x0031, // FONT
    0x041E, // FORMAT
    0x00E0, // XF
    0x0293, // STYLE
    0x0092, // PALETTE
    0x0160, // USESELFS
    0x008D, // HIDEOBJ
    0x0022, // DATEMODE
    0x000E, // PRECISION
    0x0040, // BACKUP
    0x003D, // WINDOW1
    0x0085, // BOUNDSHEET8
    0x008C, // COUNTRY
    0x01AE, // SUPBOOK
    0x0017, // EXTERNSHEET
    0x0023, // EXTERNNAME (BIFF5 form, what calamine reads)
    0x0223, // EXTERNNAME (BIFF8 form — handle either)
    0x0018, // NAME / Lbl (BIFF5 form)
    0x0218, // NAME / Lbl (BIFF8 form)
    0x00FC, // SST
    0x00FF, // EXTSST
    // XF-table-coupled FRT records. The writer rebuilds the XF table
    // from the model, so the source's versions are stale: XFCRC holds
    // a count + CRC of the XF records (guaranteed mismatch → Excel
    // fails Office File Validation, opens the file in Protected View
    // or with "File error: data may have been lost"), and XFEXT /
    // STYLEEXT reference XF / STYLE slots by index into the SOURCE
    // tables. The model already carries the resolved XFEXT colors
    // (applied at load), so dropping these loses nothing.
    0x087C, // XFCRC
    0x087D, // XFEXT
    0x0892, // STYLEEXT
    // External-workbook cell caches. XCT/CRN belong to an external-
    // file SUPBOOK block; the writer only emits the internal and
    // add-in SUPBOOKs, so passed-through copies would be orphaned —
    // a grammar violation Office File Validation rejects.
    0x0059, // XCT
    0x005A, // CRN
];

const R_BOUNDSHEET8: u16 = 0x0085;
const R_DIMENSIONS: u16 = 0x0200;
const R_WINDOW2: u16 = 0x023E;
const R_SELECTION_LOCAL: u16 = 0x001D;
const R_EOF: u16 = 0x000A;
const R_CONTINUE: u16 = 0x003C;

/// Identify cell-table records (cells + ROW) used to find the post-
/// cells zone boundary.
fn is_cell_or_row(opcode: u16) -> bool {
    matches!(
        opcode,
        0x0208 // ROW
            | 0x0201 // BLANK
            | 0x0203 // NUMBER
            | 0x027E // RK
            | 0x00BD // MULRK
            | 0x00BE // MULBLANK
            | 0x00FD // LABELSST
            | 0x0204 // LABEL
            | 0x0006 // FORMULA
            | 0x0207 // STRING
            | 0x04BC // SHRFMLA
            | 0x0221 // ARRAY
            | 0x0205 // BOOLERR
            | 0x00D7 // DBCELL
    )
}

fn is_owned_per_sheet(opcode: u16) -> bool {
    OWNED_PER_SHEET.contains(&opcode)
}

fn is_owned_globals(opcode: u16) -> bool {
    OWNED_GLOBALS.contains(&opcode)
}

/// Per-sheet passthrough zones split out from a source substream.
/// A zone is a slice of raw records (filtered to drop opcodes the
/// writer owns); the writer emits the slice verbatim at the
/// appropriate splice point.
///
/// `pre_dim` lands BEFORE the writer's COLINFO records (i.e., right
/// after WSBOOL). `post_cells_pre_win2` lands AFTER the writer's
/// cells but BEFORE its WINDOW2/PANE. `post_win2` lands AFTER the
/// writer's WINDOW2/PANE but BEFORE EOF.
pub struct SheetZones<'a> {
    pub pre_dim: Vec<&'a RawRecord>,
    pub post_cells_pre_win2: Vec<&'a RawRecord>,
    /// Records between WINDOW2 and the first SELECTION in the source
    /// (PLV, SCL). The window-block grammar is WINDOW2 → PLV → SCL →
    /// PANE → SELECTION; these must land BEFORE the writer's
    /// SELECTION records.
    pub post_win2_pre_selection: Vec<&'a RawRecord>,
    pub post_win2: Vec<&'a RawRecord>,
}

impl<'a> SheetZones<'a> {
    pub fn empty() -> Self {
        Self {
            pre_dim: vec![],
            post_cells_pre_win2: vec![],
            post_win2_pre_selection: vec![],
            post_win2: vec![],
        }
    }
}

/// Split a source sheet substream into the three passthrough zones.
/// CONTINUE records are KEPT in their adjacency to the previous
/// record; they're passthrough-friendly because BIFF readers stitch
/// them based on file order, not content. Owned-record opcodes are
/// dropped.
pub fn split_sheet_zones<'a>(src: &'a SubstreamSnapshot) -> SheetZones<'a> {
    let recs = &src.records;
    if recs.is_empty() {
        return SheetZones::empty();
    }
    // Embedded-substream depth per record: an embedded chart inside a
    // worksheet substream carries records whose opcodes COLLIDE with
    // writer-owned sheet records (the chart's own DIMENSIONS, PROTECT,
    // SCL, …). Zone anchors must only match at depth 0, and the
    // owned-opcode filter must never fire inside an embedded stream —
    // stripping the chart's internal DIMENSIONS corrupted every saved
    // workbook with an embedded chart (Office File Validation failure,
    // 2026-08-14 report follow-up).
    let depths = embedded_depths(recs);
    // Last meaningful position is len-1 (EOF) — exclude it from any
    // zone.
    let end_excl = recs.len().saturating_sub(1);
    let dim_pos = (0..recs.len())
        .find(|&i| depths[i] == 0 && recs[i].opcode == R_DIMENSIONS)
        .unwrap_or(end_excl);
    let last_cell = (0..recs.len())
        .rev()
        .find(|&i| depths[i] == 0 && is_cell_or_row(recs[i].opcode));
    let win2_pos = (0..recs.len())
        .find(|&i| depths[i] == 0 && recs[i].opcode == R_WINDOW2);

    // pre_dim: BOF+1 .. dim_pos. `restore_continue` does the
    // owned-opcode filtering and CONTINUE-chain adjacency in one
    // pass.
    let pre_dim = restore_continue(&recs[1..dim_pos], &depths[1..dim_pos]);

    // mid zone: after last cell+1 .. min(win2_pos, eof)
    let mid_start = last_cell.map(|p| p + 1).unwrap_or(dim_pos);
    let mid_end = win2_pos.unwrap_or(end_excl);
    let mid_end = mid_end.max(mid_start);
    let post_cells_pre_win2 = restore_continue(&recs[mid_start..mid_end], &depths[mid_start..mid_end]);

    // post zone: WINDOW2 onwards .. eof. Skip WINDOW2/PANE/SELECTION
    // themselves (writer re-emits these from model state). Split at
    // the source's first SELECTION: PLV/SCL come between WINDOW2 and
    // SELECTION per the window-block grammar.
    let post_start = win2_pos.unwrap_or(end_excl);
    let sel_pos = (post_start..end_excl)
        .find(|&i| depths[i] == 0 && recs[i].opcode == R_SELECTION_LOCAL)
        .unwrap_or(post_start);
    let post_win2_pre_selection = restore_continue(&recs[post_start..sel_pos], &depths[post_start..sel_pos]);
    let post_win2 = restore_continue(&recs[sel_pos..end_excl], &depths[sel_pos..end_excl]);

    let mut zones = SheetZones { pre_dim, post_cells_pre_win2, post_win2_pre_selection, post_win2 };
    // Diagnostic (dev-only): drop each embedded chart complex — the
    // chart BOF..EOF plus its anchoring MSODRAWING/OBJ pair directly
    // before it — from the mid zone.
    if std::env::var("FASTSHEET_XLS_DROP_CHART_COMPLEX").is_ok() {
        zones.post_cells_pre_win2 = drop_chart_complexes(zones.post_cells_pre_win2);
    }
    // Diagnostic bisection hook (dev-only): drop whole zone classes
    // via FASTSHEET_XLS_DROP_SHEET_ZONES containing p / m / w.
    if let Ok(v) = std::env::var("FASTSHEET_XLS_DROP_SHEET_ZONES") {
        if v.contains('p') { zones.pre_dim.clear(); }
        if v.contains('m') { zones.post_cells_pre_win2.clear(); }
        if v.contains('w') { zones.post_win2.clear(); zones.post_win2_pre_selection.clear(); }
    }
    zones
}

/// Remove every embedded chart complex from a zone record list: the
/// embedded BOF..EOF stream plus the MSODRAWING (and OBJ/CONTINUE)
/// records that anchor it immediately before.
fn drop_chart_complexes<'a>(recs: Vec<&'a RawRecord>) -> Vec<&'a RawRecord> {
    const R_BOF: u16 = 0x0809;
    const R_MSODRAWING: u16 = 0x00EC;
    let mut out: Vec<&'a RawRecord> = Vec::with_capacity(recs.len());
    let mut i = 0usize;
    while i < recs.len() {
        if recs[i].opcode == R_BOF {
            // Walk back over the anchor: OBJ / CONTINUE back to and
            // including the nearest MSODRAWING.
            while let Some(last) = out.last() {
                let op = last.opcode;
                out.pop();
                if op == R_MSODRAWING {
                    break;
                }
            }
            // Skip the embedded stream through its matching EOF.
            let mut depth = 1u32;
            i += 1;
            while i < recs.len() && depth > 0 {
                if recs[i].opcode == R_BOF {
                    depth += 1;
                } else if recs[i].opcode == R_EOF {
                    depth -= 1;
                }
                i += 1;
            }
            continue;
        }
        out.push(recs[i]);
        i += 1;
    }
    out
}

/// Depth of each record relative to embedded substreams: 0 for the
/// worksheet's own records, >0 inside an embedded (chart) BOF..EOF.
/// Index 0 (the sheet's own BOF) is depth 0; an embedded BOF and its
/// matching EOF are both tagged with the embedded depth so the whole
/// chart stream travels as one verbatim unit.
fn embedded_depths(recs: &[RawRecord]) -> Vec<u32> {
    const R_BOF: u16 = 0x0809;
    let mut out = Vec::with_capacity(recs.len());
    let mut depth = 0u32;
    for (i, r) in recs.iter().enumerate() {
        if i == 0 {
            out.push(0); // the sheet's own BOF
            continue;
        }
        if r.opcode == R_BOF {
            depth += 1;
            out.push(depth);
        } else if r.opcode == R_EOF {
            out.push(depth);
            depth = depth.saturating_sub(1);
        } else {
            out.push(depth);
        }
    }
    out
}

/// Filter out owned record opcodes while preserving CONTINUE
/// adjacency: a CONTINUE record belongs to the previous record in
/// file order. If we drop a parent (because it's owned), drop its
/// CONTINUE chain too. If we keep a parent, keep its CONTINUE chain.
/// Records inside an embedded substream (depth > 0) pass through
/// verbatim — their opcodes belong to the chart, not the sheet.
fn restore_continue<'a>(slice: &'a [RawRecord], depths: &[u32]) -> Vec<&'a RawRecord> {
    // Diagnostic bisection hook, mirroring the globals one: extra
    // depth-0 opcodes to drop, plus the token "emb" to drop embedded
    // substreams. Dev-only.
    let (extra_drop, drop_embedded) = match std::env::var("FASTSHEET_XLS_SHEET_PT_DROP") {
        Ok(v) => {
            let drop_emb = v.contains("emb");
            let set: std::collections::HashSet<u16> = v
                .split(',')
                .filter_map(|s| u16::from_str_radix(s.trim().trim_start_matches("0x"), 16).ok())
                .collect();
            (set, drop_emb)
        }
        Err(_) => (std::collections::HashSet::new(), false),
    };
    let mut out = Vec::with_capacity(slice.len());
    let mut last_kept = false;
    for (rec, &depth) in slice.iter().zip(depths) {
        if depth > 0 {
            if !drop_embedded {
                out.push(rec);
                last_kept = true;
            } else {
                last_kept = false;
            }
            continue;
        }
        if extra_drop.contains(&rec.opcode) {
            last_kept = false;
            continue;
        }
        if rec.opcode == R_CONTINUE {
            if last_kept {
                out.push(rec);
            }
            // last_kept stays true so further CONTINUEs in the chain
            // also pass through.
        } else if !is_owned_per_sheet(rec.opcode) {
            out.push(rec);
            last_kept = true;
        } else {
            last_kept = false;
        }
    }
    out
}

/// Globals records the writer should pass through verbatim. Skips
/// owned opcodes and EOF; the caller appends these before the
/// writer's own R_EOF in globals.
/// Globals passthrough records, grouped by placement anchor.
///
/// Office File Validation checks the globals substream against the
/// [MS-XLS] grammar, and record ORDER is part of it — appending every
/// passthrough record at the tail (the old behavior) put position-
/// sensitive records (TABID, EXCEL9FILE, the protection block,
/// MSODRAWINGGROUP, …) in grammar-illegal spots, and Excel opened the
/// file in Protected View / with "File error: data may have been
/// lost" (2026-08-14 report follow-up). To reproduce the source's
/// grammar-valid relative order, each passthrough record is keyed by
/// the NEXT writer-owned opcode after it in the source; the writer
/// flushes each group just before it emits that anchor. Records whose
/// anchor the writer never emits fall through to the trailing flush
/// before EOF.
pub struct GlobalsZones<'a> {
    /// (anchor opcode, records) in source order. Anchor R_EOF holds
    /// the trailing group.
    zones: Vec<(u16, Vec<&'a RawRecord>)>,
}

impl<'a> GlobalsZones<'a> {
    pub fn empty() -> Self {
        GlobalsZones { zones: Vec::new() }
    }
    /// Records anchored at `anchor`, removing them from the map.
    pub fn take(&mut self, anchor: u16) -> Vec<&'a RawRecord> {
        let mut out = Vec::new();
        for (a, recs) in self.zones.iter_mut() {
            if *a == anchor {
                out.append(recs);
            }
        }
        out
    }
    /// Everything not yet taken, in source order. Call before EOF.
    pub fn drain_remaining(&mut self) -> Vec<&'a RawRecord> {
        let mut out = Vec::new();
        for (_, recs) in self.zones.iter_mut() {
            out.append(recs);
        }
        out
    }
}

pub fn globals_passthrough_zones<'a>(src: &'a SubstreamSnapshot) -> GlobalsZones<'a> {
    // Diagnostic bisection hook: FASTSHEET_XLS_PT_DROP holds comma-
    // separated hex opcodes to additionally filter from the globals
    // passthrough. Dev-only; unset in normal operation.
    let extra_drop: std::collections::HashSet<u16> = std::env::var("FASTSHEET_XLS_PT_DROP")
        .map(|v| {
            v.split(',')
                .filter_map(|s| u16::from_str_radix(s.trim().trim_start_matches("0x"), 16).ok())
                .collect()
        })
        .unwrap_or_default();
    let mut zones: Vec<(u16, Vec<&'a RawRecord>)> = Vec::new();
    let mut pending: Vec<&'a RawRecord> = Vec::new();
    let mut last_kept = false;
    for rec in &src.records {
        if rec.opcode == R_EOF {
            continue;
        }
        if rec.opcode == R_CONTINUE {
            // A CONTINUE travels with its parent: pend it only when
            // the parent was a kept passthrough record.
            if last_kept {
                pending.push(rec);
            }
            continue;
        }
        if is_owned_globals(rec.opcode) {
            // Owned anchor: everything pending is placed just before
            // the writer's emission of this opcode.
            last_kept = false;
            if !pending.is_empty() {
                zones.push((rec.opcode, std::mem::take(&mut pending)));
            }
        } else if extra_drop.contains(&rec.opcode) {
            last_kept = false;
        } else {
            pending.push(rec);
            last_kept = true;
        }
    }
    if !pending.is_empty() {
        zones.push((R_EOF, pending));
    }
    GlobalsZones { zones }
}

/// Build a `sheet name (lowercased) → source-substream-index` map by
/// walking BOUNDSHEET8 records in the globals substream. Source
/// substream order matches BOUNDSHEET8 order: substream `k+1` (skip
/// globals at index 0) corresponds to `BOUNDSHEET8[k]`. The writer
/// looks up by IronCalc worksheet name so chart/macro substreams
/// (which IronCalc doesn't carry) don't shift the indexing.
pub fn build_sheet_substream_index(globals: &SubstreamSnapshot) -> HashMap<String, usize> {
    let mut out = HashMap::new();
    let mut sub_idx = 1usize;
    for rec in &globals.records {
        if rec.opcode == R_BOUNDSHEET8 {
            if let Some(name) = parse_boundsheet_name(&rec.data) {
                out.insert(name.to_lowercase(), sub_idx);
            }
            sub_idx += 1;
        }
    }
    out
}

/// One non-worksheet substream found in the source — chart or macro
/// sheets that IronCalc doesn't carry. The writer emits the
/// substream's records verbatim after the worksheet substreams,
/// with a corresponding BOUNDSHEET8 entry appended in globals.
pub struct ExtraSubstream<'a> {
    pub name: String,
    pub hs_state: u8,
    pub dt: u8,
    pub records: &'a [RawRecord],
}

/// Walk the source's globals + workbook substreams to collect any
/// substream whose BOUNDSHEET8 entry has dt != 0 (worksheet).
/// Currently that's chart sheets (dt=2) and macro sheets (dt=1, dt=6
/// for VB modules — but VB modules have no substream of their own).
/// Returns slices that borrow from `substreams`, so the caller keeps
/// the snapshot alive while emitting.
pub fn extract_extra_substreams<'a>(
    substreams: &'a [SubstreamSnapshot],
) -> Vec<ExtraSubstream<'a>> {
    let Some(globals) = substreams.iter().find(|s| s.bof_dt == 0x0005) else {
        return Vec::new();
    };
    let mut out = Vec::new();
    let mut sub_idx = 1usize; // index of next non-globals substream
    for rec in &globals.records {
        if rec.opcode == R_BOUNDSHEET8 {
            let parsed = parse_boundsheet_full(&rec.data);
            if let Some((name, hs_state, dt)) = parsed {
                if dt != 0 && sub_idx < substreams.len() {
                    out.push(ExtraSubstream {
                        name,
                        hs_state,
                        dt,
                        records: &substreams[sub_idx].records,
                    });
                }
            }
            sub_idx += 1;
        }
    }
    out
}

/// Parse BOUNDSHEET8: lbPlyPos (4) + hsState (1) + dt (1) + name. The
/// name is a ShortXLUnicodeString. Returns (name, hsState, dt) or
/// None if malformed.
fn parse_boundsheet_full(data: &[u8]) -> Option<(String, u8, u8)> {
    if data.len() < 8 {
        return None;
    }
    let hs_state = data[4];
    let dt = data[5];
    let name = parse_boundsheet_name(data)?;
    Some((name, hs_state, dt))
}

/// BOUNDSHEET8 layout: lbPlyPos (u32), hsState (u8), dt (u8), then a
/// ShortXLUnicodeString — cch (u8), grbit (u8), bytes. grbit bit 0 =
/// 1 means UTF-16; otherwise 1-byte ASCII. Returns None on malformed
/// input.
fn parse_boundsheet_name(data: &[u8]) -> Option<String> {
    if data.len() < 8 {
        return None;
    }
    let cch = data[6] as usize;
    let grbit = data[7];
    let high_byte = (grbit & 0x01) != 0;
    if high_byte {
        if data.len() < 8 + 2 * cch {
            return None;
        }
        let mut chars = Vec::with_capacity(cch);
        for i in 0..cch {
            let b0 = data[8 + 2 * i] as u16;
            let b1 = data[8 + 2 * i + 1] as u16;
            chars.push(b0 | (b1 << 8));
        }
        String::from_utf16(&chars).ok()
    } else {
        if data.len() < 8 + cch {
            return None;
        }
        Some(String::from_utf8_lossy(&data[8..8 + cch]).into_owned())
    }
}
