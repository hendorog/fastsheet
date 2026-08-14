//! BIFF8 (.xls) writer.
//!
//! Mirrors what `xls_load.rs` + `xls_biff.rs` reads. We re-emit the
//! workbook as an OLE2 compound file containing a single `/Workbook`
//! stream of BIFF8 records.
//!
//! Phase 1 (this commit): byte-level scaffolding only. Primitives for
//! record emission, BIFF8 string encoding, CONTINUE-splitting for
//! oversized payloads, and a minimal "valid empty xls" producer used by
//! the unit tests. Workbook-globals (SST/XF/FONT/etc.), cell records,
//! and formula encoding are added in later commits — see tasks
//! #2 / #3 / #4 in the task list.
//!
//! References for record layouts:
//! * [MS-XLS]: Excel (.xls) Binary File Format
//!   https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/
//! * OpenOffice.org's Excel File Format documentation (BIFF5/8).

use std::collections::{BTreeMap, HashMap, HashSet};
use std::io::Write;
use std::path::Path;

// ---------------------------------------------------------------------------
// BIFF8 record types (mirror of xls_biff.rs constants — kept duplicated so
// the writer doesn't import from the reader and vice versa).
// ---------------------------------------------------------------------------

const R_BOF: u16 = 0x0809;
const R_EOF: u16 = 0x000A;
const R_BOUNDSHEET8: u16 = 0x0085;
const R_DIMENSIONS: u16 = 0x0200;
const R_WINDOW2: u16 = 0x023E;
const R_PANE: u16 = 0x0041;
const R_SELECTION: u16 = 0x001D;
const R_MERGECELLS: u16 = 0x00E5;
const R_CONTINUE: u16 = 0x003C;
const R_CODEPAGE: u16 = 0x0042;
const R_DSF: u16 = 0x0161;
const R_WINDOW1: u16 = 0x003D;
const R_DATEMODE: u16 = 0x0022;
const R_PRECISION: u16 = 0x000E;
const R_BACKUP: u16 = 0x0040;
const R_HIDEOBJ: u16 = 0x008D;
const R_FONT: u16 = 0x0031;
const R_FORMAT: u16 = 0x041E;
const R_XF: u16 = 0x00E0;
const R_XFCRC: u16 = 0x087C;
const R_XFEXT: u16 = 0x087D;
const R_STYLE: u16 = 0x0293;
const R_USESELFS: u16 = 0x0160;
const R_COUNTRY: u16 = 0x008C;
const R_BOOKBOOL: u16 = 0x00DA;
const R_SST: u16 = 0x00FC;
const R_EXTSST: u16 = 0x00FF;
const R_INTERFACEHDR: u16 = 0x00E1;
const R_INTERFACEEND: u16 = 0x00E2;
const R_MMS: u16 = 0x00C1;
const R_WRITEACCESS: u16 = 0x005C;
const R_REFRESHALL: u16 = 0x01B7;
const R_CALCMODE: u16 = 0x000D;
const R_CALCCOUNT: u16 = 0x000C;
const R_REFMODE: u16 = 0x000F;
const R_ITERATION: u16 = 0x0011;
const R_DELTA: u16 = 0x0010;
const R_SAVERECALC: u16 = 0x005F;
const R_PRINTHEADERS: u16 = 0x002A;
const R_PRINTGRIDLINES: u16 = 0x002B;
const R_GRIDSET: u16 = 0x0082;
const R_GUTS: u16 = 0x0080;
const R_DEFAULTROWHEIGHT: u16 = 0x0225;
const R_WSBOOL: u16 = 0x0081;
const R_DEFCOLWIDTH: u16 = 0x0055;
const R_BLANK: u16 = 0x0201;
const R_NUMBER: u16 = 0x0203;
const R_LABELSST: u16 = 0x00FD;
const R_BOOLERR: u16 = 0x0205;
const R_FORMULA: u16 = 0x0006;
const R_STRING: u16 = 0x0207;
const R_ARRAY: u16 = 0x0221;
const R_ROW: u16 = 0x0208;
const R_COLINFO: u16 = 0x007D;
const R_SUPBOOK: u16 = 0x01AE;
const R_EXTERNSHEET: u16 = 0x0017;
// Match calamine's read of NAME (0x0018) — pick the BIFF5-style
// EXTERNNAME (0x0023) for symmetry. xls_biff.rs reads both 0x0223
// and 0x0023; calamine ignores both (we only need our reader to
// resolve them).
const R_EXTERNNAME: u16 = 0x0023;
// NAME record: spec ([MS-XLS] 2.4.187 Lbl) lists 0x0218 as the BIFF8
// opcode, but in practice every reader (calamine, Excel itself, POI)
// expects the BIFF5 0x0018 form — calamine xls.rs:361 only matches
// that exact value. Use 0x0018 to be widely compatible.
const R_LBL: u16 = 0x0018;

// BOF substream types.
const DT_GLOBALS: u16 = 0x0005;
const DT_WORKSHEET: u16 = 0x0010;

// Maximum record body length before a CONTINUE record is required.
// [MS-XLS] §2.1.7.20.1 — record max size is 8224 bytes.
const MAX_RECORD_BODY: usize = 8224;

// ---------------------------------------------------------------------------
// Primitive byte-level emitter.
// ---------------------------------------------------------------------------

/// Builds BIFF8 records into a contiguous byte buffer. Records are
/// emitted by value-typed helpers (`u16`, `u32`, etc.); record framing
/// (the 4-byte type+len header and CONTINUE splitting) is handled by
/// `write_record`.
#[derive(Default)]
pub(crate) struct BiffWriter {
    buf: Vec<u8>,
}

impl BiffWriter {
    pub(crate) fn new() -> Self {
        Self { buf: Vec::with_capacity(4096) }
    }

    pub(crate) fn into_bytes(self) -> Vec<u8> {
        self.buf
    }

    /// Current byte offset in the stream — used for INDEX / DBCELL /
    /// BOUNDSHEET8 stream-position fixups.
    pub(crate) fn pos(&self) -> u32 {
        self.buf.len() as u32
    }

    /// Patch a previously-written u32 in place. Used to fill in
    /// stream-position fields (e.g. BOUNDSHEET8.lbPlyPos) once the
    /// referenced record has been emitted.
    pub(crate) fn patch_u32(&mut self, offset: usize, value: u32) {
        let bytes = value.to_le_bytes();
        self.buf[offset..offset + 4].copy_from_slice(&bytes);
    }

    /// Emit one BIFF record. If the body exceeds `MAX_RECORD_BODY`,
    /// split into CONTINUE records of the same opcode-style framing
    /// (subsequent fragments use `R_CONTINUE` as their type).
    ///
    /// Note: a few records (notably SST) have non-trivial CONTINUE
    /// rules where the per-fragment payload prefix must be repeated.
    /// Those records call `write_record_raw` instead and split their
    /// own bodies. For simple "fits in one record" payloads this
    /// helper is the right tool.
    pub(crate) fn write_record(&mut self, rec_type: u16, body: &[u8]) {
        if body.len() <= MAX_RECORD_BODY {
            self.write_record_raw(rec_type, body);
            return;
        }
        // Generic split: first fragment under rec_type, rest under
        // CONTINUE. Suitable for records whose CONTINUE semantics
        // are "just keep emitting bytes" (most non-SST records).
        let (first, rest) = body.split_at(MAX_RECORD_BODY);
        self.write_record_raw(rec_type, first);
        let mut remaining = rest;
        while !remaining.is_empty() {
            let chunk_len = remaining.len().min(MAX_RECORD_BODY);
            self.write_record_raw(R_CONTINUE, &remaining[..chunk_len]);
            remaining = &remaining[chunk_len..];
        }
    }

    /// Write a single record header + body without any CONTINUE
    /// handling. Caller guarantees `body.len() <= MAX_RECORD_BODY`.
    pub(crate) fn write_record_raw(&mut self, rec_type: u16, body: &[u8]) {
        debug_assert!(
            body.len() <= MAX_RECORD_BODY,
            "BIFF record body {} exceeds max {}",
            body.len(),
            MAX_RECORD_BODY
        );
        self.buf.extend_from_slice(&rec_type.to_le_bytes());
        self.buf.extend_from_slice(&(body.len() as u16).to_le_bytes());
        self.buf.extend_from_slice(body);
    }
}

// ---------------------------------------------------------------------------
// Body-builder helpers. Most callers want to assemble a small Vec<u8> of
// little-endian primitives, then hand it to `write_record`. These ride
// on a plain Vec so they compose in the obvious way.
// ---------------------------------------------------------------------------

pub(crate) trait BiffBody {
    fn put_u8(&mut self, v: u8);
    fn put_u16(&mut self, v: u16);
    fn put_u32(&mut self, v: u32);
    fn put_f64(&mut self, v: f64);
    /// XLUnicodeString-style: u16 char count + u8 flag + chars.
    /// Compressed (1 byte/char) when all codepoints fit in U+0000..=U+00FF;
    /// otherwise uncompressed (2 bytes/char).
    fn put_xl_unicode_string(&mut self, s: &str);
    /// ShortXLUnicodeString: u8 char count + u8 flag + chars. Used by
    /// BOUNDSHEET8 (sheet name) and a handful of other records.
    fn put_short_xl_unicode_string(&mut self, s: &str);
}

impl BiffBody for Vec<u8> {
    fn put_u8(&mut self, v: u8) { self.push(v); }
    fn put_u16(&mut self, v: u16) { self.extend_from_slice(&v.to_le_bytes()); }
    fn put_u32(&mut self, v: u32) { self.extend_from_slice(&v.to_le_bytes()); }
    fn put_f64(&mut self, v: f64) { self.extend_from_slice(&v.to_le_bytes()); }

    fn put_xl_unicode_string(&mut self, s: &str) {
        let chars: Vec<u16> = s.encode_utf16().collect();
        let high_byte = chars.iter().any(|&c| c > 0xFF);
        // [MS-XLS] 2.5.295 — `cch` is u16, then 1-byte flag (bit 0 = high-byte).
        // chars count is the UTF-16 code unit count, capped at u16::MAX.
        let cch = chars.len().min(u16::MAX as usize) as u16;
        self.put_u16(cch);
        self.put_u8(if high_byte { 0x01 } else { 0x00 });
        if high_byte {
            for c in chars.iter().take(cch as usize) {
                self.put_u16(*c);
            }
        } else {
            for c in chars.iter().take(cch as usize) {
                self.put_u8(*c as u8);
            }
        }
    }

    fn put_short_xl_unicode_string(&mut self, s: &str) {
        // [MS-XLS] 2.5.240 — like XLUnicodeString but `cch` is u8.
        let chars: Vec<u16> = s.encode_utf16().collect();
        let high_byte = chars.iter().any(|&c| c > 0xFF);
        let cch = chars.len().min(u8::MAX as usize) as u8;
        self.put_u8(cch);
        self.put_u8(if high_byte { 0x01 } else { 0x00 });
        if high_byte {
            for c in chars.iter().take(cch as usize) {
                self.put_u16(*c);
            }
        } else {
            for c in chars.iter().take(cch as usize) {
                self.put_u8(*c as u8);
            }
        }
    }
}

// ---------------------------------------------------------------------------
// Common record builders.
// ---------------------------------------------------------------------------

/// BOF — Beginning of File / substream marker. [MS-XLS] 2.4.21.
///
/// The version fields must declare a MODERN Excel writer. The old
/// values (rupBuild 0x0DBB, rupYear 0x07CC = 1996, bfh 0x41) said
/// "last saved by an Excel-97-era build", and Excel gates its trust
/// in post-2003 FRT records on them — with the old BOF, Excel
/// silently discarded our XFEXT true-color records even with a valid
/// XFCRC (2026-08-14). These constants mirror what current Excel
/// writes in real files (neutral.xls: rupBuild 20314, rupYear 1997,
/// verLastXLSaved in bfh = 0x8xxxx band, verLowestBiff 6).
fn build_bof(substream: u16) -> Vec<u8> {
    let mut body = Vec::with_capacity(16);
    body.put_u16(0x0600); // vers = BIFF8
    body.put_u16(substream); // dt = substream type
    body.put_u16(0x4F5A);    // rupBuild (20314, modern Excel)
    body.put_u16(0x07CD);    // rupYear (1997 — fixed for BIFF8)
    body.put_u32(0x000380D9); // bfh: history flags + verLastXLSaved
    body.put_u32(0x00000806); // sfo: verLowestBiff=6 + reserved
    body
}

fn build_eof() -> Vec<u8> { Vec::new() }

/// DIMENSIONS — [MS-XLS] 2.4.91. rwMic / rwMac / colMic / colMac / reserved.
/// COLINFO record (0x007D) [MS-XLS] 2.4.39. Body (12 bytes):
///   colFirst u16, colLast u16, colWidth u16, ixfe u16, grbit u16,
///   reserved u16. colWidth is in 1/256ths of the default font's
///   '0' character width. grbit bit 0 = hidden.
///
/// IronCalc stores Col.width as PLAIN CHARACTERS (the
/// `set_column_width` API divides its argument by COLUMN_WIDTH_FACTOR
/// before storing — see vendor/base/src/worksheet.rs:420). Convert
/// chars → 1/256ths by multiplying by 256.
fn build_colinfo_range(
    min_col_1based: i32,
    max_col_1based: i32,
    width_chars: f64,
    ixfe: u16,
    hidden: bool,
) -> Vec<u8> {
    let mut body = Vec::with_capacity(12);
    let col_first = (min_col_1based - 1).max(0) as u16;
    let col_last = (max_col_1based - 1).max(0) as u16;
    let biff_units = (width_chars * 256.0).round();
    let col_width = biff_units.max(0.0).min(u16::MAX as f64) as u16;
    body.put_u16(col_first);
    body.put_u16(col_last);
    body.put_u16(col_width);
    body.put_u16(ixfe);
    body.put_u16(if hidden { 0x0001 } else { 0 });
    body.put_u16(0); // reserved
    body
}

/// Build the full set of COLINFO bodies for a worksheet, layering the
/// AppState `hidden_cols` side-channel on top of `worksheet.cols`. Cols
/// with the same (width, ixfe, hidden) coalesce into a single range
/// entry so output mirrors what Excel itself would emit.
///
/// Hidden cols not covered by an explicit IronCalc Col entry get the
/// Excel default width (8.43 chars) so unhide restores a reasonable
/// width — IronCalc's `Col` has no hidden field, so a "user just
/// hid this column" gesture in our UI never updates `worksheet.cols`.
fn build_colinfo_records(
    ws: &ironcalc::base::types::Worksheet,
    hidden_cols: &HashSet<i32>,
) -> Vec<Vec<u8>> {
    const DEFAULT_HIDDEN_WIDTH_CHARS: f64 = 8.43;

    let mut per_col: BTreeMap<i32, (u64, u16, bool)> = BTreeMap::new();
    for col in &ws.cols {
        let ixfe = col.style.unwrap_or(15) as u16;
        // f64 isn't Eq, so quantize widths to the BIFF units we'll emit.
        let w_units = (col.width * 256.0).round().max(0.0).min(u16::MAX as f64) as u64;
        for i in col.min..=col.max {
            per_col.insert(i, (w_units, ixfe, false));
        }
    }
    let default_units = (DEFAULT_HIDDEN_WIDTH_CHARS * 256.0).round() as u64;
    for &c in hidden_cols {
        per_col
            .entry(c)
            .and_modify(|e| e.2 = true)
            .or_insert((default_units, 15u16, true));
    }
    if per_col.is_empty() {
        return Vec::new();
    }

    let mut out = Vec::new();
    let mut iter = per_col.into_iter().peekable();
    while let Some((min_col, attrs)) = iter.next() {
        let mut max_col = min_col;
        while let Some(&(next, next_attrs)) = iter.peek() {
            if next != max_col + 1 || next_attrs != attrs {
                break;
            }
            max_col = next;
            iter.next();
        }
        let (w_units, ixfe, hidden) = attrs;
        let width_chars = (w_units as f64) / 256.0;
        out.push(build_colinfo_range(min_col, max_col, width_chars, ixfe, hidden));
    }
    out
}

/// ROW record (0x0208) [MS-XLS] 2.4.220. Body (16 bytes):
///   rw u16, colMic u16, colMac u16, miyRw u16, reserved1 u16,
///   reserved2 u16, grbit u16, ixfe u16. miyRw stores height in
///   twips (1/20 point) in the bottom 15 bits, with bit 15 = "custom
///   height" flag.
///
/// IronCalc stores Row.height as PLAIN POINTS (the `set_row_height`
/// API divides its argument by ROW_HEIGHT_FACTOR before storing —
/// see vendor/base/src/worksheet.rs:353). Convert points → twips
/// by multiplying by 20.
fn build_row_record(row: &ironcalc::base::types::Row, col_mic: u16, col_mac: u16) -> Vec<u8> {
    let mut body = Vec::with_capacity(16);
    body.put_u16((row.r - 1).max(0) as u16);
    body.put_u16(col_mic);
    body.put_u16(col_mac);
    let twips = (row.height * 20.0).round().max(0.0).min(0x7FFF as f64) as u16;
    let custom = if row.custom_height { 0x8000 } else { 0 };
    body.put_u16(twips | custom);
    body.put_u16(0); // reserved1
    body.put_u16(0); // reserved2
    let grbit: u16 = if row.hidden { 0x0020 } else { 0 };
    body.put_u16(grbit);
    body.put_u16(row.s.max(0).min(u16::MAX as i32) as u16);
    body
}

fn build_dimensions(rw_mic: u32, rw_mac: u32, col_mic: u16, col_mac: u16) -> Vec<u8> {
    let mut body = Vec::with_capacity(14);
    body.put_u32(rw_mic);
    body.put_u32(rw_mac); // exclusive — first row past last used row
    body.put_u16(col_mic);
    body.put_u16(col_mac); // exclusive — first col past last used col
    body.put_u16(0);
    body
}

/// WINDOW2 — [MS-XLS] 2.4.358. Selection-state record. When the
/// worksheet has frozen panes, set fFrozen (bit 3, 0x0008) and
/// fFrozenNoSplit (bit 8, 0x0100) — Excel pairs the two; without
/// fFrozenNoSplit the sheet renders as a draggable split instead of
/// frozen titles even with PANE present.
fn build_window2(frozen_rows: i32, frozen_columns: i32, is_active: bool) -> Vec<u8> {
    let mut body = Vec::with_capacity(18);
    // Base: fDspGrid=1, fDspRwCol=1, fDspZeros=1, fDefaultHdr=1,
    //       fDspGuts=1.
    let mut grbit: u16 = 0x00B6;
    // fSelected (0x0200) + fPaged (0x0400) mark the ACTIVE sheet
    // only. Setting them on every sheet (the old behavior) made
    // Excel open the workbook with all sheets selected — "Group"
    // mode in the title bar, where an edit hits every sheet at once.
    if is_active {
        grbit |= 0x0600;
    }
    if frozen_rows > 0 || frozen_columns > 0 {
        grbit |= 0x0008; // fFrozen
        grbit |= 0x0100; // fFrozenNoSplit
    }
    body.put_u16(grbit);
    body.put_u16(0); // top row
    body.put_u16(0); // left col
    body.put_u32(0x00000040); // icvHdr (default)
    body.put_u16(0); // pagebreak preview zoom
    body.put_u16(0); // normal view zoom
    body.put_u32(0); // reserved
    body
}

/// PANE — [MS-XLS] 2.4.213. Emitted only when the worksheet has frozen
/// panes (i.e. WINDOW2's fFrozen is set). For frozen panes:
///   x       = number of columns in the left (frozen) pane
///   y       = number of rows in the top (frozen) pane
///   rwTop   = top visible row in the bottom pane (= y when frozen)
///   colLeft = left visible col in the right pane (= x when frozen)
///   pnnAct  = active pane (0 = lower-right, 2 = lower-left). When
///             only rows are frozen, no vertical split exists so the
///             data area is the lower-left pane (pnnAct=2). Otherwise
///             the data area is the lower-right (pnnAct=0).
/// ifnt field of chart records that reference the globals FONT table:
/// FONTX (0x1026, body = a single u16 ifnt) and FBI (0x1060, ifnt at
/// body offset 8). [MS-XLS] 2.4.126 / 2.4.111.
fn zone_record_ifnt(opcode: u16, data: &[u8]) -> Option<u16> {
    match opcode {
        0x1026 if data.len() >= 2 => Some(u16::from_le_bytes([data[0], data[1]])),
        0x1060 if data.len() >= 10 => Some(u16::from_le_bytes([data[8], data[9]])),
        _ => None,
    }
}

/// Emit a passthrough record, rewriting chart font references through
/// the source→ours ifnt remap built in build_workbook_stream.
fn write_zone_record(
    w: &mut BiffWriter,
    rec: &crate::xls_preserve::RawRecord,
    font_remap: &HashMap<u16, u16>,
) {
    if let Some(ifnt) = zone_record_ifnt(rec.opcode, &rec.data) {
        if let Some(&new_ifnt) = font_remap.get(&ifnt) {
            let mut d = rec.data.clone();
            let off = if rec.opcode == 0x1026 { 0 } else { 8 };
            d[off..off + 2].copy_from_slice(&new_ifnt.to_le_bytes());
            w.write_record(rec.opcode, &d);
            return;
        }
    }
    w.write_record(rec.opcode, &rec.data);
}

/// Panes needing a SELECTION record, in Excel's emission order:
/// top-left (3) first, then top-right (1) / bottom-left (2) as the
/// freeze axes require, with the data-area pane last (matching
/// PANE.pnnAcc so the active selection lands in the scrollable area).
fn selection_panes(frozen_rows: i32, frozen_columns: i32) -> Vec<u8> {
    match (frozen_rows > 0, frozen_columns > 0) {
        (false, false) => vec![3],
        (true, false) => vec![3, 2],
        (false, true) => vec![3, 1],
        (true, true) => vec![3, 1, 2, 0],
    }
}

/// SELECTION — [MS-XLS] 2.4.248. pnn(1) + rwAct(2) + colAct(2) +
/// irefAct(2) + cref(2) + one RefU (rwFirst u16, rwLast u16,
/// colFirst u8, colLast u8): a default A1 selection in the pane.
fn build_selection(pnn: u8) -> Vec<u8> {
    let mut body = Vec::with_capacity(15);
    body.put_u8(pnn);
    body.put_u16(0); // rwAct
    body.put_u16(0); // colAct
    body.put_u16(0); // irefAct
    body.put_u16(1); // cref
    body.put_u16(0); // rwFirst
    body.put_u16(0); // rwLast
    body.put_u8(0);  // colFirst
    body.put_u8(0);  // colLast
    body
}

fn build_pane(frozen_rows: i32, frozen_columns: i32) -> Vec<u8> {
    let cols = frozen_columns.max(0).min(u16::MAX as i32) as u16;
    let rows = frozen_rows.max(0).min(u16::MAX as i32) as u16;
    // pnnAcc is a 2-byte field — the record must be exactly 10 bytes.
    // A 1-byte emit here produced files Apache POI rejects outright
    // ("Not enough data (1) to read requested (2) bytes" in PaneRecord).
    let pnn_acc: u16 = if cols == 0 { 2 } else { 0 };
    let mut body = Vec::with_capacity(10);
    body.put_u16(cols);  // x
    body.put_u16(rows);  // y
    body.put_u16(rows);  // rwTop
    body.put_u16(cols);  // colLeft
    body.put_u16(pnn_acc);
    body
}

/// Parse a plain A1 range like "C3:D3" (no sheet, no $) into 1-based
/// (row1, col1, row2, col2). Merge-list entries are stored in exactly
/// this shape by both load paths.
fn parse_merge_range(range: &str) -> Option<(u16, u16, u16, u16)> {
    fn parse_cell(s: &str) -> Option<(u16, u16)> {
        let mut col: u32 = 0;
        let mut row: u32 = 0;
        let mut in_digits = false;
        for c in s.trim().trim_start_matches('$').chars() {
            if c == '$' {
                continue;
            }
            if c.is_ascii_alphabetic() && !in_digits {
                col = col * 26 + (c.to_ascii_uppercase() as u32 - 'A' as u32 + 1);
            } else if c.is_ascii_digit() {
                in_digits = true;
                row = row * 10 + (c as u32 - '0' as u32);
            } else {
                return None;
            }
        }
        if col == 0 || row == 0 || col > 0x4000 || row > 0x10000 {
            return None;
        }
        Some((row as u16, col as u16))
    }
    let (a, b) = match range.split_once(':') {
        Some((a, b)) => (a, b),
        None => (range, range),
    };
    let (r1, c1) = parse_cell(a)?;
    let (r2, c2) = parse_cell(b)?;
    Some((r1.min(r2), c1.min(c2), r1.max(r2), c1.max(c2)))
}

/// MERGECELLS — [MS-XLS] 2.4.168. cmcs (u16) then REF8 per range:
/// rwFirst / rwLast / colFirst / colLast, all u16 and 0-based. At most
/// 1026 REF8s fit a record; callers chunk. Ranges that don't parse or
/// exceed BIFF's 65 536 rows are skipped.
fn build_mergecells(ranges: &[&str]) -> Vec<u8> {
    let refs: Vec<(u16, u16, u16, u16)> = ranges
        .iter()
        .filter_map(|r| parse_merge_range(r))
        .collect();
    let mut body = Vec::with_capacity(2 + refs.len() * 8);
    body.put_u16(refs.len() as u16);
    for (r1, c1, r2, c2) in refs {
        body.put_u16(r1 - 1);
        body.put_u16(r2 - 1);
        body.put_u16(c1 - 1);
        body.put_u16(c2 - 1);
    }
    body
}

/// BOUNDSHEET8 — [MS-XLS] 2.4.28. lbPlyPos (u32, file-offset of the
/// sheet's BOF — patched in after the substream is emitted), grbit (u16,
/// hidden/very-hidden flags), then a ShortXLUnicodeString of the name.
/// Returns the body and the offset of the lbPlyPos field within it (so
/// the caller can later patch the value).
fn build_boundsheet8(name: &str) -> (Vec<u8>, usize) {
    build_boundsheet8_full(name, 0, 0)
}

/// Variant that accepts the visibility flag and substream type — used
/// by the preservation path to emit BOUNDSHEET8 entries for chart /
/// macro substreams from the source.
fn build_boundsheet8_full(name: &str, hs_state: u8, dt: u8) -> (Vec<u8>, usize) {
    let mut body = Vec::with_capacity(8 + name.len() + 2);
    let lb_ply_pos_offset = 0;
    body.put_u32(0); // lbPlyPos placeholder
    body.put_u8(hs_state);
    body.put_u8(dt);
    body.put_short_xl_unicode_string(name);
    (body, lb_ply_pos_offset)
}

// ---------------------------------------------------------------------------
// Globals-substream record builders. The set + ordering below mirrors
// what HSSF writes for a "blank" workbook — Excel accepts smaller
// subsets but rejects unexpected orderings, so it's safer to follow the
// canonical sequence.
// ---------------------------------------------------------------------------

fn build_codepage() -> Vec<u8> {
    let mut body = Vec::with_capacity(2);
    // 1200 = UTF-16 LE; 1252 = ANSI Western. We can use either since
    // every string we emit carries its own high-byte flag, but 1200
    // matches what modern Excel emits.
    body.put_u16(1200);
    body
}

fn build_dsf() -> Vec<u8> {
    let mut body = Vec::with_capacity(2);
    body.put_u16(0); // single-stream file
    body
}

fn build_window1() -> Vec<u8> {
    let mut body = Vec::with_capacity(18);
    body.put_u16(360);   // xWn (twips)
    body.put_u16(270);   // yWn
    body.put_u16(14940); // dxWn
    body.put_u16(9150);  // dyWn
    // grbit: fHidden=0, fIconic=0, fDspHScroll=1, fDspVScroll=1,
    //        fBotAdornment=1, fDspFmlaBar=1, fDspStatusBar=1
    body.put_u16(0x0038);
    body.put_u16(0); // itabCur — active sheet
    body.put_u16(0); // itabFirst — first selected sheet
    body.put_u16(1); // ctabSel — number of selected sheets
    body.put_u16(600); // wTabRatio (default 60%)
    body
}

fn build_datemode_1900() -> Vec<u8> {
    let mut body = Vec::with_capacity(2);
    body.put_u16(0); // 0 = 1900 base, 1 = 1904 base
    body
}

fn build_precision_full() -> Vec<u8> {
    let mut body = Vec::with_capacity(2);
    body.put_u16(1); // full precision (not "as displayed")
    body
}

fn build_backup_off() -> Vec<u8> {
    let mut body = Vec::with_capacity(2);
    body.put_u16(0);
    body
}

fn build_hideobj_show_all() -> Vec<u8> {
    let mut body = Vec::with_capacity(2);
    body.put_u16(0);
    body
}

fn build_useselfs_yes() -> Vec<u8> {
    let mut body = Vec::with_capacity(2);
    body.put_u16(1);
    body
}

fn build_country() -> Vec<u8> {
    let mut body = Vec::with_capacity(4);
    body.put_u16(1); // user country (US)
    body.put_u16(1); // language country
    body
}

fn build_bookbool() -> Vec<u8> {
    let mut body = Vec::with_capacity(2);
    body.put_u16(0);
    body
}

fn build_refreshall() -> Vec<u8> {
    let mut body = Vec::with_capacity(2);
    body.put_u16(0);
    body
}

fn build_interfacehdr() -> Vec<u8> {
    let mut body = Vec::with_capacity(2);
    body.put_u16(1200); // codepage hint
    body
}

fn build_mms() -> Vec<u8> {
    let mut body = Vec::with_capacity(2);
    body.put_u16(0);
    body
}

fn build_writeaccess() -> Vec<u8> {
    // [MS-XLS] 2.4.351 — userName: an XLUnicodeString padded to 112
    // bytes total. Use a generic owner name; no PII.
    let mut body = Vec::with_capacity(112);
    body.put_xl_unicode_string("fastsheet");
    while body.len() < 112 {
        body.push(0x20); // ASCII space pad
    }
    body
}

// Per-sheet calc/print/view records that Excel expects in every sheet
// substream's prologue.
fn build_calcmode_auto() -> Vec<u8> { let mut b = Vec::new(); b.put_u16(1); b }
fn build_calccount() -> Vec<u8> { let mut b = Vec::new(); b.put_u16(0x0064); b } // 100 iters max
fn build_refmode_a1() -> Vec<u8> { let mut b = Vec::new(); b.put_u16(1); b }
fn build_iteration_off() -> Vec<u8> { let mut b = Vec::new(); b.put_u16(0); b }
fn build_delta() -> Vec<u8> {
    let mut b = Vec::new();
    b.put_f64(0.001); // standard 0.001 convergence delta
    b
}
fn build_saverecalc() -> Vec<u8> { let mut b = Vec::new(); b.put_u16(1); b }
fn build_printheaders_off() -> Vec<u8> { let mut b = Vec::new(); b.put_u16(0); b }
fn build_printgridlines_off() -> Vec<u8> { let mut b = Vec::new(); b.put_u16(0); b }
fn build_gridset() -> Vec<u8> { let mut b = Vec::new(); b.put_u16(1); b }
fn build_guts() -> Vec<u8> {
    let mut b = Vec::new();
    b.put_u16(0); // dxRwGut
    b.put_u16(0); // dxColGut
    b.put_u16(0); // iLevelRwMac
    b.put_u16(0); // iLevelColMac
    b
}
fn build_defaultrowheight() -> Vec<u8> {
    let mut b = Vec::new();
    b.put_u16(0); // grbit (fUnsynced=0, fDyZero=0, fExAsc=0, fExDsc=0)
    b.put_u16(255); // height in twips (= 12.75 points = standard)
    b
}
fn build_wsbool() -> Vec<u8> {
    let mut b = Vec::new();
    // grbit: fShowAutoBreaks=1, fOutline=1, fSyncHoriz/Vert=0,
    //        fAltExprEval=0, fAltFormulaEntry=0, fDspGuts=1,
    //        fSyncHoriz=0, fSyncVert=0, fTransitionEval=0,
    //        fAltExprEval=0, fSummaryRowsBelow=1, fSummaryColsRight=1
    b.put_u16(0x04C1);
    b
}

// ---------------------------------------------------------------------------
// FONT / FORMAT / XF / PALETTE — model-driven style tables.
//
// IronCalc stores styles as separate vectors: workbook.styles.{fonts,
// fills, borders, num_fmts, cell_xfs}, each cell carrying a `s: i32`
// index into cell_xfs. We translate each into the matching BIFF8
// records.
//
// BIFF8 reserves XF indices 0..15 as "style" XFs (with fStyle=1) and
// uses 16+ as "cell" XFs that user cells reference via ixfe. So we
// emit 15 default style XFs first, then one cell XF per IronCalc
// cell_xf, and store the mapping (ironcalc index → biff index) so
// `emit_sheet_cells` can translate `cell.s` correctly.
//
// FONT-index-4 skip: XF.ifnt values skip 4 ({0,1,2,3,5,...}) but the
// stream contains NO placeholder record — record k carries ifnt k for
// k < 4 and ifnt k+1 from the fifth record on ([MS-XLS] 2.4.122).
// Our reader mirrors this by inserting a phantom Vec entry at index 4.
// ---------------------------------------------------------------------------

const PALETTE_LAST_USER_INDEX: u16 = 63;
const ICV_AUTO: u16 = 0x7FFF;
/// Custom-palette assignments start at icv=24 — past the 16 primary
/// colors (icv 8..23: black/white/RGB/CMY/grays/maroons/teals) so
/// fonts referring to those default-named colors via OUR icvs still
/// resolve correctly even if the source workbook had a custom palette
/// that we round-tripped only partially.
const PALETTE_FIRST_CUSTOM_SLOT: u16 = 24;

/// BIFF8 default palette for icv 8..63 — verbatim from
/// xls_biff.rs::seed_default_palette so reading and writing agree.
/// Used when emitting PALETTE records to fill unallocated slots
/// instead of zeroing them (which would over-write standard colors
/// like red=10 with black, breaking fonts that reference those icvs).
const BIFF_DEFAULT_PALETTE: &[(u16, [u8; 3])] = &[
    (8,  [0x00, 0x00, 0x00]), (9,  [0xFF, 0xFF, 0xFF]),
    (10, [0xFF, 0x00, 0x00]), (11, [0x00, 0xFF, 0x00]),
    (12, [0x00, 0x00, 0xFF]), (13, [0xFF, 0xFF, 0x00]),
    (14, [0xFF, 0x00, 0xFF]), (15, [0x00, 0xFF, 0xFF]),
    (16, [0x80, 0x00, 0x00]), (17, [0x00, 0x80, 0x00]),
    (18, [0x00, 0x00, 0x80]), (19, [0x80, 0x80, 0x00]),
    (20, [0x80, 0x00, 0x80]), (21, [0x00, 0x80, 0x80]),
    (22, [0xC0, 0xC0, 0xC0]), (23, [0x80, 0x80, 0x80]),
    (24, [0x99, 0x99, 0xFF]), (25, [0x99, 0x33, 0x66]),
    (26, [0xFF, 0xFF, 0xCC]), (27, [0xCC, 0xFF, 0xFF]),
    (28, [0x66, 0x00, 0x66]), (29, [0xFF, 0x80, 0x80]),
    (30, [0x00, 0x66, 0xCC]), (31, [0xCC, 0xCC, 0xFF]),
    (32, [0x00, 0x00, 0x80]), (33, [0xFF, 0x00, 0xFF]),
    (34, [0xFF, 0xFF, 0x00]), (35, [0x00, 0xFF, 0xFF]),
    (36, [0x80, 0x00, 0x80]), (37, [0x80, 0x00, 0x00]),
    (38, [0x00, 0x80, 0x80]), (39, [0x00, 0x00, 0xFF]),
    (40, [0x00, 0xCC, 0xFF]), (41, [0xCC, 0xFF, 0xFF]),
    (42, [0xCC, 0xFF, 0xCC]), (43, [0xFF, 0xFF, 0x99]),
    (44, [0x99, 0xCC, 0xFF]), (45, [0xFF, 0x99, 0xCC]),
    (46, [0xCC, 0x99, 0xFF]), (47, [0xFF, 0xCC, 0x99]),
    (48, [0x33, 0x66, 0xFF]), (49, [0x33, 0xCC, 0xCC]),
    (50, [0x99, 0xCC, 0x00]), (51, [0xFF, 0xCC, 0x00]),
    (52, [0xFF, 0x99, 0x00]), (53, [0xFF, 0x66, 0x00]),
    (54, [0x66, 0x66, 0x99]), (55, [0x96, 0x96, 0x96]),
    (56, [0x00, 0x33, 0x66]), (57, [0x33, 0x99, 0x66]),
    (58, [0x00, 0x33, 0x00]), (59, [0x33, 0x33, 0x00]),
    (60, [0x99, 0x33, 0x00]), (61, [0x99, 0x33, 0x66]),
    (62, [0x33, 0x33, 0x99]), (63, [0x33, 0x33, 0x33]),
];

/// Look up an RGB color in the BIFF default palette. Returns the icv
/// of the first matching slot, or None for non-matches. Used so
/// fonts with standard colors (black, red, etc.) can use built-in
/// icv values rather than custom-palette overrides.
fn match_default_palette_icv(rgb: [u8; 3]) -> Option<u16> {
    BIFF_DEFAULT_PALETTE
        .iter()
        .find_map(|&(icv, c)| if c == rgb { Some(icv) } else { None })
}

/// MsoCRC32 ([MS-OSHARED] 2.4.3): the checksum Excel stores in XFCRC
/// (0x087C) over the concatenated data of all XF records, and checks
/// on load to decide whether the file's XFEXT records are trustworthy
/// (a writer that rebuilt the XF table without updating the crc has
/// stale extensions). The cache is spec-literal — including the odd
/// 16-bit mask — and verified against Excel's own XFCRC in
/// neutral.xls (0xB4C2D16A over 2,111 XF records, 2026-08-14).
fn mso_crc32(data: &[u8], init: u32) -> u32 {
    let mut cache = [0u32; 256];
    for (i, slot) in cache.iter_mut().enumerate() {
        let mut v = (i as u32) << 24;
        for _ in 0..8 {
            v = if v & 0x8000_0000 != 0 {
                (v << 1) ^ 0xAF
            } else {
                v << 1
            };
        }
        *slot = v & 0xFFFF;
    }
    let mut crc = init;
    for &b in data {
        let idx = ((crc >> 24) as u8 ^ b) as usize;
        crc = (crc << 8) ^ cache[idx];
    }
    crc
}

/// XFCRC record body ([MS-XLS] 2.4.354): frtHeader(12) + reserved(2)
/// + cxfs(2) + the MsoCRC32 of all XF record bodies.
fn build_xfcrc_record(xfs: &[Vec<u8>]) -> Vec<u8> {
    let mut blob = Vec::new();
    for xf in xfs {
        blob.extend_from_slice(xf);
    }
    let crc = mso_crc32(&blob, 0);
    let mut body = Vec::with_capacity(20);
    body.put_u16(0x087C); // frtHeader.rt
    body.put_u16(0);      // frtHeader.grbitFrt
    body.extend_from_slice(&[0u8; 8]); // frtHeader.reserved
    body.put_u16(0);      // reserved1
    body.put_u16(xfs.len() as u16); // cxfs
    body.put_u32(crc);
    body
}

/// XFEXT record body ([MS-XLS] 2.4.355): frtHeader(12) + reserved1(2)
/// + ixfe(2) + reserved2(2) + cexts(2) + ExtProps. Each ExtProp is
/// extType(2) + cb(2)=20 + FullColorExt(xclrType=2 RGB, tint=0,
/// LongRGBA value, 8 unused).
fn build_xfext_record(ixfe: u16, exts: &[(u16, [u8; 3])]) -> Vec<u8> {
    let mut body = Vec::with_capacity(20 + exts.len() * 20);
    body.put_u16(0x087D);
    body.put_u16(0);
    body.extend_from_slice(&[0u8; 8]);
    body.put_u16(0);
    body.put_u16(ixfe);
    body.put_u16(0);
    body.put_u16(exts.len() as u16);
    for &(ext_type, rgb) in exts {
        body.put_u16(ext_type);
        body.put_u16(20); // cb: 4 header + 16 FullColorExt
        body.put_u16(2);  // xclrType = RGB
        body.put_u16(0);  // nTintShade
        body.put_u8(rgb[0]);
        body.put_u8(rgb[1]);
        body.put_u8(rgb[2]);
        body.put_u8(0xFF); // LongRGBA alpha
        body.extend_from_slice(&[0u8; 8]); // unused
    }
    body
}

/// Pre-computed BIFF style records derived from the IronCalc model.
struct StyleTables {
    /// FONT record bodies in file-emission order. Length is at least 5
    /// (BIFF8 minimum: indices 0..4 with 4 as a placeholder).
    fonts: Vec<Vec<u8>>,
    /// FORMAT records to emit (custom only, ifmt >= 164).
    formats: Vec<Vec<u8>>,
    /// XF record bodies in file-emission order. First 15 are style XFs
    /// (indices 0..14), then `default_cell_xf_idx` is index 15, then
    /// one entry per `model.styles.cell_xfs`.
    xfs: Vec<Vec<u8>>,
    /// Default cell XF index — used as a fallback when a cell's
    /// `s` field is out of range.
    default_cell_xf_idx: u16,
    /// Maps IronCalc cell_xf index → BIFF XF index. Length matches
    /// `model.styles.cell_xfs.len()`.
    cell_xf_index: Vec<u16>,
    /// Custom palette: (icv, rgb) entries that override the BIFF
    /// default palette at specific slots. Slots not in this list
    /// keep their default colors. Empty ⇒ no PALETTE record needed.
    palette: Vec<(u16, [u8; 3])>,
    /// True-color extensions per BIFF XF index, for colors the
    /// 56-slot palette could only approximate: (ixfe, [(extType,
    /// rgb)]). Emitted as XFCRC + XFEXT records so modern Excel (and
    /// our own loader) render the exact colors while legacy readers
    /// fall back to the nearest-palette icv baked into the XF —
    /// the same two-tier scheme Excel itself writes.
    xfext: Vec<(u16, Vec<(u16, [u8; 3])>)>,
}

/// Map a model font's array index to its BIFF8 file font index, applying
/// the index-4 skip convention.
fn model_font_to_biff_ifnt(model_idx: u32) -> u16 {
    if model_idx < 4 {
        model_idx as u16
    } else {
        (model_idx + 1) as u16
    }
}

/// Parse a `#RRGGBB` (or `#AARRGGBB`) hex color string. Returns None
/// for malformed input or `None` (no color set).
pub(crate) fn parse_hex_color(s: &str) -> Option<[u8; 3]> {
    let s = s.trim_start_matches('#');
    let bytes = match s.len() {
        6 => s,
        8 => &s[2..], // ignore alpha prefix
        _ => return None,
    };
    let r = u8::from_str_radix(&bytes[0..2], 16).ok()?;
    let g = u8::from_str_radix(&bytes[2..4], 16).ok()?;
    let b = u8::from_str_radix(&bytes[4..6], 16).ok()?;
    Some([r, g, b])
}

/// Mutable state for palette assignment across a workbook walk.
/// Custom slots avoid colliding with default-palette icvs that any
/// font / fill in the workbook actually uses — otherwise our PALETTE
/// record's overlay would silently break those default-icv references
/// (e.g. light cyan icv=27 getting clobbered by the 4th custom color
/// allocated from icv=24 onward).
#[derive(Default)]
struct PaletteState {
    custom: Vec<(u16, [u8; 3])>,
    /// BIFF default icvs that some color matched via
    /// match_default_palette_icv. Custom slot allocation skips these.
    used_defaults: std::collections::HashSet<u16>,
}

/// Resolve an OOXML hex color to a BIFF8 icv. Strategy:
///   1. None / unparseable → ICV_AUTO (font uses system foreground)
///   2. Color matches a BIFF default palette slot → use that icv
///      directly, mark it as "in use" so subsequent custom slot
///      allocations don't collide
///   3. Already in our custom-overrides → reuse its slot
///   4. Otherwise → scan icv 24..63 for the first slot not in
///      used_defaults and not already custom; assign there
///   5. No free slots → fall back to ICV_AUTO
///
/// Pre-fix the writer matched colors against an empty-then-grow list
/// starting at icv=8, so the very first non-default font color
/// shoved BLACK out of icv=8 — which broke every font that referenced
/// "black" via the standard icv.
///
/// Pre-fix #2 the writer allocated custom slots sequentially from
/// icv=24, which clobbered icv 27 (light cyan default) once the 4th
/// custom color was registered. User reported cyan-background cells
/// turning dark blue. Now custom slots avoid any used default icv.
fn palette_assign(state: &mut PaletteState, color: Option<&str>) -> u16 {
    let Some(s) = color else { return ICV_AUTO };
    let Some(rgb) = parse_hex_color(s) else { return ICV_AUTO };
    if let Some(icv) = match_default_palette_icv(rgb) {
        state.used_defaults.insert(icv);
        return icv;
    }
    if let Some(&(icv, _)) = state.custom.iter().find(|(_, c)| *c == rgb) {
        return icv;
    }
    let mut candidate = PALETTE_FIRST_CUSTOM_SLOT;
    while candidate <= PALETTE_LAST_USER_INDEX {
        let in_used_defaults = state.used_defaults.contains(&candidate);
        let in_custom = state.custom.iter().any(|(i, _)| *i == candidate);
        if !in_used_defaults && !in_custom {
            state.custom.push((candidate, rgb));
            return candidate;
        }
        candidate += 1;
    }
    // Palette exhausted (a workbook with more unique colors than the
    // 56 slots — neutral.xls has 66 fills alone). Falling back to
    // ICV_AUTO here made solid fills render as NO fill (auto
    // foreground on a solid pattern reads back as unset): 6,136 cells
    // lost their fill on the 2026-08-14 fidelity audit. Do what Excel
    // itself does when quantizing to the legacy palette: nearest
    // color by squared RGB distance over every assigned slot. Exact
    // colors ride the XFEXT records on top.
    nearest_palette_icv(state, rgb)
}

/// Nearest already-assigned palette slot (defaults + custom
/// overrides) by squared RGB distance.
fn nearest_palette_icv(state: &PaletteState, rgb: [u8; 3]) -> u16 {
    let dist = |a: [u8; 3]| -> i64 {
        let dr = a[0] as i64 - rgb[0] as i64;
        let dg = a[1] as i64 - rgb[1] as i64;
        let db = a[2] as i64 - rgb[2] as i64;
        dr * dr + dg * dg + db * db
    };
    // Custom slots override the default color at their icv, so a
    // default entry whose slot was reassigned no longer exists.
    BIFF_DEFAULT_PALETTE
        .iter()
        .filter(|(icv, _)| !state.custom.iter().any(|(ci, _)| ci == icv))
        .map(|&(icv, c)| (icv, dist(c)))
        .chain(state.custom.iter().map(|&(icv, c)| (icv, dist(c))))
        .min_by_key(|&(_, d)| d)
        .map(|(icv, _)| icv)
        .unwrap_or(ICV_AUTO)
}

/// The color currently occupying a palette slot: custom override
/// first, then the BIFF default. None for auto / out-of-range icvs.
fn palette_rgb(state: &PaletteState, icv: u16) -> Option<[u8; 3]> {
    if let Some(&(_, rgb)) = state.custom.iter().find(|(i, _)| *i == icv) {
        return Some(rgb);
    }
    BIFF_DEFAULT_PALETTE
        .iter()
        .find_map(|&(i, c)| if i == icv { Some(c) } else { None })
}

/// Build a BIFF FONT record from an IronCalc Font.
fn build_font_record(
    font: &ironcalc::base::types::Font,
    palette: &mut PaletteState,
) -> Vec<u8> {
    let mut body = Vec::with_capacity(32);
    body.put_u16((font.sz * 20).max(1) as u16); // dyHeight (twips)
    let mut grbit: u16 = 0;
    if font.i { grbit |= 0x0002; }
    if font.strike { grbit |= 0x0008; }
    body.put_u16(grbit);
    body.put_u16(palette_assign(palette, font.color.as_deref()));
    body.put_u16(if font.b { 700 } else { 400 });
    body.put_u16(0); // sss (super/subscript)
    body.put_u8(if font.u_double { 2 } else if font.u { 1 } else { 0 }); // uls
    body.put_u8(font.family.max(0).min(255) as u8); // bFamily
    body.put_u8(0); // bCharSet
    body.put_u8(0); // reserved
    body.put_short_xl_unicode_string(&font.name);
    body
}

/// Build a placeholder FONT record (used at index 4 to satisfy the
/// BIFF8 skip convention). Duplicates the body of the workbook's
/// default font.
fn build_font_placeholder() -> Vec<u8> {
    let mut body = Vec::with_capacity(32);
    body.put_u16(220); body.put_u16(0); body.put_u16(ICV_AUTO);
    body.put_u16(400); body.put_u16(0); body.put_u8(0);
    body.put_u8(0); body.put_u8(0); body.put_u8(0);
    body.put_short_xl_unicode_string("Calibri");
    body
}

/// IronCalc's built-in number-format table, shared from the vendored
/// crate (made `pub` by the fastsheet patch that also fixed its
/// mangled entries) so id resolution here can never drift from what
/// the model itself resolves.
const IRONCALC_DEFAULT_NUM_FMTS: &[&str] =
    ironcalc::base::number_format::DEFAULT_NUM_FMTS;

/// Look up the format-code string for an IronCalc num_fmt_id. Mirrors
/// `ironcalc::base::number_format::get_num_fmt`: registered entries in
/// `num_fmts` win, then the IronCalc default table, then "general".
fn ironcalc_format_code_for_id(num_fmts: &[ironcalc::base::types::NumFmt], id: i32) -> String {
    for nf in num_fmts {
        if nf.num_fmt_id == id {
            return nf.format_code.clone();
        }
    }
    if id >= 0 && (id as usize) < IRONCALC_DEFAULT_NUM_FMTS.len() {
        return IRONCALC_DEFAULT_NUM_FMTS[id as usize].to_string();
    }
    "general".to_string()
}

/// Built-in BIFF format index for a given format string, or None if
/// it needs to be a custom FORMAT record. Mirror of xls_biff.rs's
/// `builtin_format` function — same numeric mapping.
fn builtin_fmt_idx(fmt: &str) -> Option<u16> {
    let s = fmt.trim();
    if s.eq_ignore_ascii_case("general") || s.is_empty() {
        return Some(0);
    }
    Some(match s {
        "0" => 1,
        "0.00" => 2,
        "#,##0" => 3,
        "#,##0.00" => 4,
        "0%" => 9,
        "0.00%" => 10,
        "0.00E+00" => 11,
        "# ?/?" => 12,
        "# ??/??" => 13,
        "m/d/yyyy" => 14,
        "d-mmm-yy" => 15,
        "d-mmm" => 16,
        "mmm-yy" => 17,
        "h:mm AM/PM" => 18,
        "h:mm:ss AM/PM" => 19,
        "h:mm" => 20,
        "h:mm:ss" => 21,
        "m/d/yyyy h:mm" => 22,
        "#,##0_);(#,##0)" => 37,
        "#,##0_);[Red](#,##0)" => 38,
        "#,##0.00_);(#,##0.00)" => 39,
        "#,##0.00_);[Red](#,##0.00)" => 40,
        "mm:ss" => 45,
        "[h]:mm:ss" => 46,
        "mm:ss.0" => 47,
        "##0.0E+0" => 48,
        "@" => 49,
        _ => return None,
    })
}

/// Map an IronCalc BorderStyle to its BIFF8 dg* code (0=none, 1=thin, etc.).
fn border_style_to_biff(style: &ironcalc::base::types::BorderStyle) -> u8 {
    use ironcalc::base::types::BorderStyle;
    match style {
        BorderStyle::Thin => 1,
        BorderStyle::Medium => 2,
        BorderStyle::Thick => 5,
        BorderStyle::Double => 6,
        BorderStyle::Dotted => 4,
        BorderStyle::SlantDashDot => 13,
        BorderStyle::MediumDashed => 8,
        BorderStyle::MediumDashDotDot => 12,
        BorderStyle::MediumDashDot => 10,
    }
}

/// Convert IronCalc Fill.pattern_type to its BIFF8 fill-pattern code.
fn fill_pattern_to_biff(pattern: &str) -> u16 {
    match pattern {
        "none" => 0,
        "solid" => 1,
        "darkGray" => 3,
        "mediumGray" => 4,
        "lightGray" => 5,
        "gray125" => 17,
        "gray0625" => 18,
        _ => 0,
    }
}

/// Build a BIFF XF record body (20 bytes) from the resolved attributes.
#[allow(clippy::too_many_arguments)]
fn build_xf_record(
    ifnt: u16,
    ifmt: u16,
    is_style: bool,
    ixf_parent: u16,
    h_align: u8,
    v_align: u8,
    wrap: bool,
    border_left: u8,
    border_right: u8,
    border_top: u8,
    border_bottom: u8,
    border_icvs: (u16, u16, u16, u16),
    fill_pattern: u16,
    fill_fg_icv: u16,
    fill_bg_icv: u16,
) -> Vec<u8> {
    let mut body = Vec::with_capacity(20);
    body.put_u16(ifnt);
    body.put_u16(ifmt);
    let f_style: u16 = if is_style { 0b100 } else { 0 };
    let locked: u16 = 0x0001;
    let cell_options = locked | f_style | (ixf_parent << 4);
    body.put_u16(cell_options);
    let align_byte = (h_align & 0x07)
        | (if wrap { 0x08 } else { 0 })
        | ((v_align & 0x07) << 4);
    body.put_u8(align_byte);
    body.put_u8(0); // rotation
    body.put_u8(0); // indent / shrink / merge / reading order
    // fAtr* used-attribute bits: 0xFC sets all six (number format,
    // font, alignment, border, fill, protection). Both style XFs and
    // cell XFs declare ownership across the board so the cell's
    // explicit values win over the parent style. Empirically this is
    // what real-world XLS templates do, and the round-trip tests fail
    // if we leave the bits clear for cell XFs.
    body.put_u8(0xFC);
    // Border styles: nibble per side.
    let border_styles_u16 = (border_left as u16 & 0x0F)
        | ((border_right as u16 & 0x0F) << 4)
        | ((border_top as u16 & 0x0F) << 8)
        | ((border_bottom as u16 & 0x0F) << 12);
    body.put_u16(border_styles_u16);
    let (icv_left, icv_right, icv_top, icv_bottom) = border_icvs;
    body.put_u16((icv_left & 0x7F) | ((icv_right & 0x7F) << 7));
    // u32 at offset 14: icvTop bits 0-6, icvBottom 7-13, icvDiag 14-20,
    // dgDiag 21-24, fHasXFExt 25, fls (fill pattern) 26-31.
    let fill_pattern_bits = ((fill_pattern as u32) & 0x3F) << 26
        | (icv_top as u32 & 0x7F)
        | ((icv_bottom as u32 & 0x7F) << 7);
    body.put_u32(fill_pattern_bits);
    // u16 at offset 18: icvFore + icvBack.
    let pattern_colors = (fill_fg_icv & 0x7F) | ((fill_bg_icv & 0x7F) << 7);
    body.put_u16(pattern_colors);
    body
}

/// Build the full StyleTables for the given model.
fn build_style_tables(model: &ironcalc::base::Model) -> StyleTables {
    let styles = &model.workbook.styles;
    let mut palette = PaletteState::default();
    // Pre-scan all colors to populate used_defaults BEFORE allocating
    // any custom slots. Without this, a color that gets a custom slot
    // (e.g. icv=24) early can later find that icv ALSO matches the
    // BIFF default palette for some other font's color — that font
    // then references icv=24 expecting one color but the PALETTE
    // record overrides it with the custom one. User-reported as
    // light cyan cells turning dark blue.
    for f in &styles.fonts {
        let _ = palette_assign(&mut palette, f.color.as_deref());
    }
    for fill in &styles.fills {
        let _ = palette_assign(&mut palette, fill.fg_color.as_deref());
        let _ = palette_assign(&mut palette, fill.bg_color.as_deref());
    }
    // Reset the custom assignments — we only wanted used_defaults
    // populated. The actual assignment happens below as fonts/XFs
    // are emitted. used_defaults persists across the reset so custom
    // slot allocation knows which default icvs to avoid.
    palette.custom.clear();

    // ---- FONT records ----
    // Emit at minimum 5 fonts so the BIFF8 ifnt-skip-4 convention is
    // satisfied. If the model has fewer fonts, pad with the workbook
    // default. Index 4 is always a placeholder.
    let mut fonts: Vec<Vec<u8>> = Vec::with_capacity(styles.fonts.len() + 1);
    let default_font_record = if let Some(f0) = styles.fonts.first() {
        build_font_record(f0, &mut palette)
    } else {
        build_font_placeholder()
    };
    // First 4 file slots: indices 0..3 — populate from model fonts 0..3
    // where available, padding with the default font otherwise.
    for i in 0..4 {
        match styles.fonts.get(i) {
            Some(f) => fonts.push(build_font_record(f, &mut palette)),
            None => fonts.push(default_font_record.clone()),
        }
    }
    // File slots 4+: model fonts 4+, emitted directly — there is NO
    // placeholder record in the stream. Per [MS-XLS] 2.4.122 the ifnt
    // numbering simply skips 4: stream record k has ifnt k for k < 4
    // and ifnt k+1 from the fifth record on (model_font_to_biff_ifnt
    // applies the same shift). Emitting a physical placeholder here
    // made Excel resolve every font from ifnt 5 up one record early.
    for f in styles.fonts.iter().skip(4) {
        fonts.push(build_font_record(f, &mut palette));
    }
    // BIFF8 readers expect at least 5 FONT records; pad with a
    // placeholder (never referenced — it would be ifnt 5).
    if fonts.len() < 5 {
        fonts.push(build_font_placeholder());
    }

    // ---- FORMAT records + IronCalc-id → BIFF-ifmt mapping ----
    // IronCalc assigns custom num_fmt_ids starting at 46 (length of its
    // DEFAULT_NUM_FMTS table). BIFF8 reserves 0..49 for built-ins (and
    // 47 in particular is "mm:ss.0", a time format). So passing
    // IronCalc's num_fmt_id straight through to BIFF as `ifmt` causes
    // collisions when their id ranges overlap but their format-string
    // assignments diverge (most notably IronCalc 28-32 vs BIFF 45-49).
    //
    // Strategy: for every distinct num_fmt_id referenced by any cell
    // XF or registered in num_fmts, look up its format string (via the
    // model's num_fmts table or IronCalc's own DEFAULT_NUM_FMTS for
    // built-in ids), then resolve to a BIFF ifmt:
    //   - If the format string matches a known BIFF built-in: that idx
    //   - Else: assign a fresh BIFF custom slot (164+) and emit a
    //     FORMAT record so readers can resolve it on load
    let mut formats: Vec<Vec<u8>> = Vec::new();
    let mut fmt_id_map: std::collections::HashMap<i32, u16> =
        std::collections::HashMap::new();
    let mut next_custom_ifmt: u16 = 164;
    let mut distinct_ids: Vec<i32> = Vec::new();
    {
        use std::collections::HashSet;
        let mut seen: HashSet<i32> = HashSet::new();
        for cxf in &styles.cell_xfs {
            if seen.insert(cxf.num_fmt_id) { distinct_ids.push(cxf.num_fmt_id); }
        }
        for nf in &styles.num_fmts {
            if seen.insert(nf.num_fmt_id) { distinct_ids.push(nf.num_fmt_id); }
        }
    }
    for &id in &distinct_ids {
        let code = ironcalc_format_code_for_id(&styles.num_fmts, id);
        let ifmt = if let Some(idx) = builtin_fmt_idx(&code) {
            idx
        } else {
            let i = next_custom_ifmt;
            next_custom_ifmt = next_custom_ifmt.saturating_add(1);
            let mut body = Vec::new();
            body.put_u16(i);
            body.put_xl_unicode_string(&code);
            formats.push(body);
            i
        };
        fmt_id_map.insert(id, ifmt);
    }

    // ---- XF records ----
    // Emit 15 style XFs (indices 0..14) all referencing font 0 / fmt 0.
    // Then one default cell XF at index 15 (parent = style 0). Then one
    // cell XF per model cell_xf at indices 16+.
    let mut xfs: Vec<Vec<u8>> = Vec::with_capacity(15 + 1 + styles.cell_xfs.len());
    for _ in 0..15 {
        xfs.push(build_xf_record(0, 0, true, 0xFFF, 0, 2, false, 0, 0, 0, 0, (0, 0, 0, 0), 0, 0, 0));
    }
    let default_cell_xf_idx = 15u16;
    xfs.push(build_xf_record(0, 0, false, 0, 0, 2, false, 0, 0, 0, 0, (0, 0, 0, 0), 0, 0, 0));

    let mut cell_xf_index: Vec<u16> = Vec::with_capacity(styles.cell_xfs.len());
    let mut xfext: Vec<(u16, Vec<(u16, [u8; 3])>)> = Vec::new();
    for cxf in &styles.cell_xfs {
        // Resolve font.
        let ifnt = if cxf.font_id >= 0 {
            model_font_to_biff_ifnt(cxf.font_id as u32)
        } else { 0 };

        // Resolve format via the IronCalc-id → BIFF-ifmt map built
        // above. Falls back to general (0) when the id isn't known.
        let ifmt = resolve_fmt_idx(&fmt_id_map, cxf.num_fmt_id);

        // Resolve fill.
        let (fill_pattern, fill_fg_icv, fill_bg_icv) = if cxf.fill_id >= 0 {
            if let Some(fill) = styles.fills.get(cxf.fill_id as usize) {
                let pat = fill_pattern_to_biff(&fill.pattern_type);
                let fg = palette_assign(&mut palette, fill.fg_color.as_deref());
                let bg = palette_assign(&mut palette, fill.bg_color.as_deref());
                (pat, fg, bg)
            } else { (0, 0, 0) }
        } else { (0, 0, 0) };

        // Resolve borders — style code + color icv per edge. Explicit
        // colors go through the palette; edges without one use icv 64
        // ("auto", renders black).
        let (bl, br, bt, bb, icv_l, icv_r, icv_t, icv_b) = if cxf.border_id >= 0 {
            if let Some(bd) = styles.borders.get(cxf.border_id as usize) {
                let mut edge = |item: &Option<ironcalc::base::types::BorderItem>| match item {
                    Some(b) => (
                        border_style_to_biff(&b.style),
                        match b.color.as_deref() {
                            Some(hex) => palette_assign(&mut palette, Some(hex)),
                            None => 64,
                        },
                    ),
                    None => (0u8, 0u16),
                };
                let (l, il) = edge(&bd.left);
                let (r, ir) = edge(&bd.right);
                let (t, it) = edge(&bd.top);
                let (bot, ib) = edge(&bd.bottom);
                (l, r, t, bot, il, ir, it, ib)
            } else { (0, 0, 0, 0, 0, 0, 0, 0) }
        } else { (0, 0, 0, 0, 0, 0, 0, 0) };

        // Resolve alignment.
        let (h_align, v_align, wrap) = if let Some(al) = &cxf.alignment {
            use ironcalc::base::types::{HorizontalAlignment as H, VerticalAlignment as V};
            let h = match al.horizontal {
                H::General => 0, H::Left => 1, H::Center => 2, H::Right => 3,
                H::Fill => 4, H::Justify => 5, H::CenterContinuous => 6, H::Distributed => 7,
            };
            let v = match al.vertical {
                V::Top => 0, V::Center => 1, V::Bottom => 2, V::Justify => 3, V::Distributed => 4,
            };
            (h, v, al.wrap_text)
        } else { (0, 2, false) };

        xfs.push(build_xf_record(
            ifnt, ifmt, false, default_cell_xf_idx,
            h_align, v_align, wrap,
            bl, br, bt, bb,
            (icv_l, icv_r, icv_t, icv_b),
            fill_pattern, fill_fg_icv, fill_bg_icv,
        ));
        cell_xf_index.push((xfs.len() - 1) as u16);

        // XFEXT capture: any color whose palette slot doesn't hold the
        // exact RGB (nearest-match fallback, or auto because the color
        // never got a slot) gets a FullColorExt carrying the true
        // value. extType codes per [MS-XLS] 2.5.108 / our scanner:
        // 0x04 fill fg, 0x05 fill bg, 0x07-0x0A border t/b/l/r,
        // 0x0D font color.
        let mut exts: Vec<(u16, [u8; 3])> = Vec::new();
        fn ext_check(
            palette: &PaletteState,
            exts: &mut Vec<(u16, [u8; 3])>,
            ext_type: u16,
            color: Option<&str>,
            icv: u16,
        ) {
            let Some(hex) = color else { return };
            let Some(rgb) = parse_hex_color(hex) else { return };
            if palette_rgb(palette, icv) != Some(rgb) {
                exts.push((ext_type, rgb));
            }
        }
        if cxf.fill_id >= 0 {
            if let Some(fill) = styles.fills.get(cxf.fill_id as usize) {
                ext_check(&palette, &mut exts, 0x0004, fill.fg_color.as_deref(), fill_fg_icv);
                ext_check(&palette, &mut exts, 0x0005, fill.bg_color.as_deref(), fill_bg_icv);
            }
        }
        if cxf.border_id >= 0 {
            if let Some(bd) = styles.borders.get(cxf.border_id as usize) {
                let color = |it: &Option<ironcalc::base::types::BorderItem>| {
                    it.as_ref().and_then(|b| b.color.clone())
                };
                ext_check(&palette, &mut exts, 0x0007, color(&bd.top).as_deref(), icv_t);
                ext_check(&palette, &mut exts, 0x0008, color(&bd.bottom).as_deref(), icv_b);
                ext_check(&palette, &mut exts, 0x0009, color(&bd.left).as_deref(), icv_l);
                ext_check(&palette, &mut exts, 0x000A, color(&bd.right).as_deref(), icv_r);
            }
        }
        if cxf.font_id >= 0 {
            if let Some(font) = styles.fonts.get(cxf.font_id as usize) {
                if let Some(hex) = font.color.as_deref() {
                    // Re-derive the icv the FONT record got — the
                    // palette state is stable by now, so assignment
                    // is a pure lookup.
                    let icv = palette_assign(&mut palette, Some(hex));
                    ext_check(&palette, &mut exts, 0x000D, Some(hex), icv);
                }
            }
        }
        if !exts.is_empty() {
            // fHasXFExt — bit 25 of the u32 at XF offset 14; Excel
            // ignores XFEXT records whose XF doesn't opt in.
            let body = xfs.last_mut().unwrap();
            body[17] |= 0x02;
            xfext.push(((xfs.len() - 1) as u16, exts));
        }
    }

    StyleTables {
        fonts,
        formats,
        xfs,
        default_cell_xf_idx,
        cell_xf_index,
        palette: palette.custom,
        xfext,
    }
}

/// Resolve an IronCalc num_fmt_id to a BIFF ifmt via the pre-built
/// remap. Anything not present in the map falls back to general (0)
/// — IronCalc shouldn't reference an unregistered num_fmt_id, but
/// being defensive avoids crashes on malformed models.
fn resolve_fmt_idx(map: &std::collections::HashMap<i32, u16>, num_fmt_id: i32) -> u16 {
    if num_fmt_id < 0 { return 0; }
    if let Some(&idx) = map.get(&num_fmt_id) { return idx; }
    // Not in the map ⇒ fall back to general. (We never trust raw
    // IronCalc IDs to align with BIFF built-ins — see fmt_id_map
    // construction in build_style_tables.)
    0
}

/// STYLE record — register a built-in style mapping XF index → built-in
/// style ID. [MS-XLS] 2.4.269.
fn build_style_builtin(ixf: u16, builtin_id: u8, level: u8) -> Vec<u8> {
    let mut body = Vec::with_capacity(4);
    // ixfe: bits 0-11 = XF index, bit 15 = fBuiltIn
    body.put_u16(ixf | 0x8000);
    body.put_u8(builtin_id); // 0=Normal, 3=Comma, 4=Currency, 5=Percent, 6=Comma[0], 7=Currency[0]
    body.put_u8(level);
    body
}

// FORMAT records ([MS-XLS] 2.4.126) are emitted inline inside
// `build_style_tables` next to where the BIFF ifmt is allocated for
// each custom IronCalc num_fmt — see the loop near `formats.push(...)`.
// Putting the body assembly there keeps allocation and emission in
// the same place; an extracted helper would have been a third
// indirection nobody else calls.

// ---------------------------------------------------------------------------
// SST (Shared String Table). [MS-XLS] 2.4.265.
//
// The SST collects all unique strings that LABELSST cells reference by
// index. Each entry is encoded as an XLUnicodeRichExtendedString —
// effectively the same as XLUnicodeString for plain text (no rich /
// phonetic / extended bits). Records >8224 bytes split into CONTINUE
// chunks; a CONTINUE break that lands mid-string requires repeating
// the high-byte flag at the start of the continuation. For simplicity
// we split only at string boundaries here — the cost is occasional
// trailing waste in a CONTINUE record. Real-world SSTs almost never
// hit the 8224-byte limit per entry.
// ---------------------------------------------------------------------------

/// String collector → SST builder. Walks every cell in every sheet of
/// the model, registers each distinct string, and returns:
///   - a vector of strings in SST order (model.workbook.shared_strings
///     prefix, then any inline CellFormulaString values appended)
///   - a HashMap mapping each string to its SST index
///
/// We append CellFormulaString values to the SST so they can be emitted
/// as LABELSST cells instead of FORMULA + STRING follow-ups. Until the
/// real ptg encoder lands (task #4), our placeholder rgce uses PtgStr
/// which has a u8 cch cap (255 chars). Long string values (e.g.
/// 500+ char paragraph-style cells) get truncated through that path.
/// LABELSST has no length cap and round-trips full content.
pub(crate) fn build_sst_table(model: &ironcalc::base::Model) -> (Vec<String>, std::collections::HashMap<String, u32>) {
    use ironcalc::base::types::Cell;
    use std::collections::HashMap;
    let mut table: Vec<String> = model.workbook.shared_strings.clone();
    let mut idx: HashMap<String, u32> = HashMap::with_capacity(table.len() + 16);
    for (i, s) in table.iter().enumerate() {
        idx.entry(s.clone()).or_insert(i as u32);
    }
    // Walk every CellFormulaString and add its `v` to the SST if not
    // already present. Empty strings are skipped (those use the
    // FORMULA-Blank tag, no SST entry needed).
    for ws in &model.workbook.worksheets {
        for cols in ws.sheet_data.values() {
            for cell in cols.values() {
                if let Cell::CellFormulaString { v, .. } = cell {
                    if v.is_empty() { continue; }
                    if !idx.contains_key(v) {
                        idx.insert(v.clone(), table.len() as u32);
                        table.push(v.clone());
                    }
                }
            }
        }
    }
    (table, idx)
}

/// Emit SST + EXTSST records into the workbook stream. `total_refs` is
/// the count of LABELSST references the workbook will produce — for
/// model-driven save it equals the count of `Cell::SharedString` across
/// all sheets. EXTSST is an optional acceleration index; we emit a
/// minimal stub since correctness doesn't depend on it.
fn emit_sst(w: &mut BiffWriter, table: &[String], total_refs: u32) {
    // SST header: cstTotal (u32) + cstUnique (u32), then the strings.
    let mut header = Vec::with_capacity(8);
    header.put_u32(total_refs);
    header.put_u32(table.len() as u32);

    // We split at string boundaries: pack header + strings into one
    // record; if it overflows MAX_RECORD_BODY, finalise the current
    // record and start a CONTINUE for the next batch. This is simpler
    // than mid-string splitting and matches what most BIFF writers do
    // when they can.
    let mut body = header;
    let mut first_record = true;
    for s in table {
        let mut entry = Vec::with_capacity(3 + s.len() * 2);
        entry.put_xl_unicode_string(s);
        if body.len() + entry.len() > MAX_RECORD_BODY {
            // Flush current record.
            if first_record {
                w.write_record_raw(R_SST, &body);
                first_record = false;
            } else {
                w.write_record_raw(R_CONTINUE, &body);
            }
            body = Vec::with_capacity(entry.len());
        }
        body.extend_from_slice(&entry);
    }
    // Flush the final record (always emit, even if empty — the SST
    // header alone is valid).
    if first_record {
        w.write_record_raw(R_SST, &body);
    } else {
        w.write_record_raw(R_CONTINUE, &body);
    }

    // EXTSST: dsst (u16, strings per group) + rgisstinf (8 bytes per
    // group, all zero for our stub). One group is the minimum.
    let mut extsst = Vec::with_capacity(2 + 8);
    extsst.put_u16(8); // 8 strings per bucket
    extsst.put_u32(0); // ib — first-string stream offset (stub)
    extsst.put_u16(0); // cb — first-string within-record offset (stub)
    extsst.put_u16(0); // reserved
    w.write_record(R_EXTSST, &extsst);
}

// ---------------------------------------------------------------------------
// Cell record builders.
//
// Each cell record's first 6 bytes are identical: rw (u16), col (u16),
// ixfe (u16). Then the type-specific payload follows. Until phase 2.5
// ports per-cell XFs, every cell uses ixfe=15 (the default cell XF
// emitted in the globals substream).
// ---------------------------------------------------------------------------

fn put_cell_header(body: &mut Vec<u8>, rw: u32, col: u32, ixfe: u16) {
    body.put_u16(rw as u16);
    body.put_u16(col as u16);
    body.put_u16(ixfe);
}

/// BLANK record (0x0201) — empty cell with formatting only.
fn build_blank(rw: u32, col: u32, ixfe: u16) -> Vec<u8> {
    let mut body = Vec::with_capacity(6);
    put_cell_header(&mut body, rw, col, ixfe);
    body
}

/// NUMBER record (0x0203) — IEEE 754 double-precision value.
fn build_number(rw: u32, col: u32, ixfe: u16, value: f64) -> Vec<u8> {
    let mut body = Vec::with_capacity(14);
    put_cell_header(&mut body, rw, col, ixfe);
    body.put_f64(value);
    body
}

/// LABELSST record (0x00FD) — string cell pointing at the SST entry
/// at `sst_index`.
fn build_labelsst(rw: u32, col: u32, ixfe: u16, sst_index: u32) -> Vec<u8> {
    let mut body = Vec::with_capacity(10);
    put_cell_header(&mut body, rw, col, ixfe);
    body.put_u32(sst_index);
    body
}

/// BOOLERR record (0x0205) — boolean or error value.
/// `is_error == false` ⇒ bool, `byte_value` is 0/1.
/// `is_error == true`  ⇒ error code, `byte_value` is the BIFF error
/// number (0x07=DIV/0, 0x0F=VALUE, 0x17=REF, 0x1D=NAME, 0x24=NUM, 0x2A=NA).
fn build_boolerr(rw: u32, col: u32, ixfe: u16, byte_value: u8, is_error: bool) -> Vec<u8> {
    let mut body = Vec::with_capacity(8);
    put_cell_header(&mut body, rw, col, ixfe);
    body.put_u8(byte_value);
    body.put_u8(if is_error { 1 } else { 0 });
    body
}

/// Map an IronCalc Error enum variant to its BIFF error code.
fn ironcalc_error_to_biff(e: &ironcalc::base::expressions::token::Error) -> u8 {
    use ironcalc::base::expressions::token::Error;
    match e {
        Error::NULL => 0x00,
        Error::DIV => 0x07,
        Error::VALUE => 0x0F,
        Error::REF => 0x17,
        Error::NAME => 0x1D,
        Error::NUM => 0x24,
        Error::NA => 0x2A,
        // IronCalc has error variants BIFF doesn't natively
        // represent. Map to the closest BIFF equivalent:
        //   ERROR  -> #VALUE! (Excel's catch-all for generic
        //             evaluation failure; better than #N/A).
        //   NIMPL  -> #VALUE! (we couldn't implement that fn).
        //   SPILL  -> #VALUE! (no BIFF #SPILL!).
        //   CALC   -> #VALUE! (generic calc failure).
        //   CIRC   -> #VALUE! (Excel writes circular refs as 0,
        //             but as a stable error we pick #VALUE!).
        // The original variant is lost in xls; this is the
        // best-faith mapping, and it eliminates the BUG-01
        // "#ERROR! -> #N/A" round-trip drift.
        Error::ERROR | Error::NIMPL | Error::SPILL | Error::CALC | Error::CIRC => 0x0F,
    }
}

/// Cached-value bytes for the FORMULA record's `val` field
/// ([MS-XLS] 2.5.133 FormulaValue). Returns 8 bytes that the reader
/// interprets via the type-tag in byte 0 + sentinel 0xFFFF in bytes 6-7.
///
/// Tag values per spec (also matches xls_biff.rs decoder + calamine):
///   0 = string in following STRING (0x0207) record
///   1 = boolean (byte[2] = value)
///   2 = error  (byte[2] = error code)
///   3 = blank / empty — no STRING follows; result is empty string
enum FormulaCachedValue {
    Number(f64),
    Bool(bool),
    Error(u8),
    /// Non-empty cached string — caller must follow up with a STRING
    /// record carrying the value.
    StringPending,
    /// Empty / blank result — no STRING record follows. Used when a
    /// formula evaluates to "" so we don't emit a 3-byte STRING that
    /// calamine rejects (parse_string requires ≥4 bytes of body).
    Blank,
}

fn build_formula_cached_bytes(v: &FormulaCachedValue) -> [u8; 8] {
    match v {
        FormulaCachedValue::Number(n) => n.to_le_bytes(),
        FormulaCachedValue::Bool(b) => {
            [0x01, 0x00, if *b { 1 } else { 0 }, 0x00, 0x00, 0x00, 0xFF, 0xFF]
        }
        FormulaCachedValue::Error(e) => {
            [0x02, 0x00, *e, 0x00, 0x00, 0x00, 0xFF, 0xFF]
        }
        FormulaCachedValue::StringPending => {
            [0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xFF, 0xFF]
        }
        FormulaCachedValue::Blank => {
            [0x03, 0x00, 0x00, 0x00, 0x00, 0x00, 0xFF, 0xFF]
        }
    }
}

/// FORMULA record (0x0006). `rgce` is the parse-tree byte sequence
/// (ptgs). For phase 3 the rgce is a placeholder that just pushes the
/// cached value onto the operand stack (so when Excel recalcs, the
/// result matches the cache). Phase 4 replaces this with a real ptg
/// encoding of the formula text.
fn build_formula(
    rw: u32,
    col: u32,
    ixfe: u16,
    cached: &FormulaCachedValue,
    rgce: &[u8],
) -> Vec<u8> {
    build_formula_with_grbit(rw, col, ixfe, cached, rgce, 0)
}

/// Same as `build_formula` but with an explicit grbit. Array-member
/// cells (rgce = PtgExp) carry 0x0020 (fClearErrors), matching what
/// Excel writes for every cell of an array-entered range.
fn build_formula_with_grbit(
    rw: u32,
    col: u32,
    ixfe: u16,
    cached: &FormulaCachedValue,
    rgce: &[u8],
    grbit: u16,
) -> Vec<u8> {
    let mut body = Vec::with_capacity(22 + rgce.len());
    put_cell_header(&mut body, rw, col, ixfe);
    body.extend_from_slice(&build_formula_cached_bytes(cached));
    body.put_u16(grbit);
    body.put_u32(0); // chn — reserved
    body.put_u16(rgce.len() as u16);
    body.extend_from_slice(rgce);
    body
}

/// STRING record (0x0207) — cached string value following a FORMULA
/// whose `val` byte-0 is 0x03. Body is just an XLUnicodeString.
fn build_string(s: &str) -> Vec<u8> {
    let mut body = Vec::with_capacity(3 + s.len() * 2);
    body.put_xl_unicode_string(s);
    body
}

// ---------------------------------------------------------------------------
// Formula ptg encoder. Walks an IronCalc parsed-formula Node tree and
// emits BIFF8 ptg bytes (rgce) so the formula round-trips on save.
//
// Coverage in this commit:
//   - Number / Bool / String / Error literals
//   - Same-sheet PtgRef / PtgArea (with abs/rel flags)
//   - All arithmetic / compare / concat / unary operators
//   - PtgFuncVar for ~120 functions present in the BIFF FTAB
//
// Anything else (3D refs, defined names, INDIRECT/OFFSET dynamic refs,
// post-2007 _xlfn functions like XLOOKUP/IFERROR, complex constructs)
// returns None and the caller falls back to the placeholder rgce. The
// fallback yields the cached value but loses the formula text — same
// as before this commit.
// ---------------------------------------------------------------------------

/// Anchor cell context used to resolve relative refs. IronCalc's
/// parser stores relative refs as OFFSETS from this anchor in
/// `Node::ReferenceKind.{row, column}` (see vendor/base parser/mod.rs
/// — the `row - context.row` adjustment); BIFF8 stores absolute
/// targets, so the encoder adds anchor coords back when emitting
/// PtgRef / PtgArea for relative refs.
#[derive(Clone, Copy)]
struct CellAnchor {
    row: i32,    // 1-based
    column: i32, // 1-based
}

/// EXTERNSHEET / SUPBOOK supporting state for 3D refs. BIFF8 stores
/// foreign-sheet refs as PtgRef3d / PtgArea3d with a `ixti` index
/// into a workbook-level xti table; that table maps each entry to a
/// (supBook, itabFirst, itabLast) tuple. For self-referencing files
/// (the only kind we emit), the supBook is a single SUPBOOK marker
/// record at index 0 representing "this workbook".
///
/// We build the xti table eagerly: walk every parsed formula for
/// every sheet, find foreign-sheet refs, register an entry per
/// distinct (sheet_first, sheet_last) tuple. The encoder then looks
/// up each ref's ixti during emission.
#[derive(Default)]
struct XtiTable {
    /// (itabFirst, itabLast) per ixti index. For single-sheet 3D
    /// refs both equal the target sheet index.
    entries: Vec<(u16, u16)>,
    /// sheet_idx → ixti for single-sheet refs (the common case).
    single_sheet_lookup: std::collections::HashMap<u32, u16>,
    /// ixti of the add-in supbook entry (iSupBook=1, itabs=0xFFFE),
    /// allocated once when the workbook calls any ATP function.
    /// PtgNameX must reference this entry — POI resolves the ptg
    /// through EXTERNSHEET → SUPBOOK and NPEs on a dangling ixti.
    addin_entry: Option<u16>,
}

impl XtiTable {
    fn lookup_or_add_single(&mut self, sheet_idx: u32) -> u16 {
        if let Some(&i) = self.single_sheet_lookup.get(&sheet_idx) {
            return i;
        }
        let i = self.entries.len() as u16;
        self.entries.push((sheet_idx as u16, sheet_idx as u16));
        self.single_sheet_lookup.insert(sheet_idx, i);
        i
    }
    fn lookup_single(&self, sheet_idx: u32) -> Option<u16> {
        self.single_sheet_lookup.get(&sheet_idx).copied()
    }
    /// (first, last) sheet-span entry — used by preserved-rgce
    /// remapping, where a source Ptg*3d may span sheets. The add-in
    /// entry's (0xFFFE, 0xFFFE) marker can't collide with real sheet
    /// indices. Linear scan; xti tables stay small.
    fn lookup_or_add_pair(&mut self, first: u16, last: u16) -> u16 {
        if first == last {
            return self.lookup_or_add_single(first as u32);
        }
        if let Some(i) = self.entries.iter().position(|&e| e == (first, last)) {
            return i as u16;
        }
        let i = self.entries.len() as u16;
        self.entries.push((first, last));
        i
    }
    fn lookup_pair(&self, first: u16, last: u16) -> Option<u16> {
        if first == last {
            return self.lookup_single(first as u32);
        }
        self.entries
            .iter()
            .position(|&e| e == (first, last))
            .map(|i| i as u16)
    }
    fn addin_ixti(&mut self) -> u16 {
        if let Some(i) = self.addin_entry {
            return i;
        }
        let i = self.entries.len() as u16;
        self.entries.push((0xFFFE, 0xFFFE));
        self.addin_entry = Some(i);
        i
    }
}

/// Walk a single Node tree, registering xti entries for every 3D ref
/// inside it. Used both by build_xti_table (for cell formulas) and
/// build_defined_name_table (for defined-name formulas).
fn walk_node_for_xti(
    node: &ironcalc::base::expressions::parser::Node,
    current_sheet: u32,
    xti: &mut XtiTable,
) {
    use ironcalc::base::expressions::parser::Node;
    match node {
        Node::ReferenceKind { sheet_index, .. } | Node::RangeKind { sheet_index, .. } => {
            if *sheet_index != current_sheet {
                xti.lookup_or_add_single(*sheet_index);
            }
        }
        Node::OpSumKind { left, right, .. }
        | Node::OpProductKind { left, right, .. }
        | Node::OpPowerKind { left, right }
        | Node::OpConcatenateKind { left, right }
        | Node::CompareKind { left, right, .. }
        | Node::OpRangeKind { left, right } => {
            walk_node_for_xti(left, current_sheet, xti);
            walk_node_for_xti(right, current_sheet, xti);
        }
        Node::UnaryKind { right, .. } => walk_node_for_xti(right, current_sheet, xti),
        Node::FunctionKind { args, .. } => {
            for a in args {
                walk_node_for_xti(a, current_sheet, xti);
            }
        }
        _ => {}
    }
}

/// Walk every parsed formula in every sheet and register an xti
/// entry for each distinct foreign-sheet referenced.
fn build_xti_table(model: &ironcalc::base::Model) -> XtiTable {
    let mut xti = XtiTable::default();
    for (sheet_idx, sheet_formulas) in model.parsed_formulas.iter().enumerate() {
        for node in sheet_formulas {
            walk_node_for_xti(node, sheet_idx as u32, &mut xti);
        }
    }
    xti
}

/// Defined-name table: registered names with their encoded rgce
/// bytes plus a name → ILBL lookup so the formula encoder can
/// resolve `Node::DefinedNameKind` to a PtgName ptg.
///
/// ILBL is 1-based and matches the order in which NAME (Lbl)
/// records are emitted in the workbook globals substream.
#[derive(Default)]
struct DefinedNameTable {
    /// (name, sheet_id_for_local_or_None_for_global, rgce_bytes)
    /// Order is the emission order of the NAME records.
    entries: Vec<DefinedNameEntry>,
    /// Lookup keyed by (lowercase_name, sheet_id). Excel treats
    /// defined-name lookup as case-insensitive.
    lookup: std::collections::HashMap<(String, Option<u32>), u16>,
}

#[derive(Default)]
struct DefinedNameEntry {
    name: String,
    sheet_id: Option<u32>,
    rgce: Vec<u8>,
}

impl DefinedNameTable {
    fn ilbl_for(&self, name: &str, scope: Option<u32>) -> Option<u16> {
        let key = (name.to_lowercase(), scope);
        if let Some(&i) = self.lookup.get(&key) {
            return Some(i);
        }
        // Fall back to a global name with the same string when
        // looking up from a sheet-scoped context (Excel resolution).
        let key_global = (name.to_lowercase(), None);
        self.lookup.get(&key_global).copied()
    }
}

/// Walk model.workbook.defined_names, parse each one's formula via
/// IronCalc's parser, register any 3D refs in xti, encode the rgce,
/// and stash the result for emission later. Names whose formula
/// can't be parsed or encoded get an empty rgce — their NAME record
/// still emits so PtgName references don't dangle, but readers will
/// resolve them to #REF!. That's a strictly better fallback than
/// silently dropping the name and breaking every formula that
/// references it.
fn build_defined_name_table(
    model: &ironcalc::base::Model,
    xti: &mut XtiTable,
    extern_names: &ExternNameTable,
) -> DefinedNameTable {
    use ironcalc::base::expressions::parser::new_parser_english;
    use ironcalc::base::expressions::types::CellReferenceRC;
    let worksheet_names: Vec<String> = model
        .workbook
        .worksheets
        .iter()
        .map(|w| w.name.clone())
        .collect();
    let mut parser = new_parser_english(
        worksheet_names.clone(),
        Vec::new(),
        std::collections::HashMap::new(),
    );
    let mut table = DefinedNameTable::default();
    for dn in &model.workbook.defined_names {
        let formula = dn.formula.strip_prefix('=').unwrap_or(&dn.formula);
        // dn.sheet_id is the worksheet's ID field, NOT its index
        // (new_defined_name converts the caller's index via
        // ws.sheet_id). Resolve back to the index — emitting the raw
        // ID as if it were an index pointed sheet-scoped names one
        // sheet too far (IDs start at 1 in typical workbooks).
        let scope_index: Option<u32> = dn.sheet_id.and_then(|id| {
            model
                .workbook
                .worksheets
                .iter()
                .position(|w| w.sheet_id == id)
                .map(|i| i as u32)
        });
        let context = CellReferenceRC {
            sheet: worksheet_names
                .get(scope_index.unwrap_or(0) as usize)
                .cloned()
                .unwrap_or_default(),
            column: 1,
            row: 1,
        };
        let node = parser.parse(formula, &context);
        // NAME records' rgce must always emit 3D refs (PtgRef3d /
        // PtgArea3d) — the name has no enclosing-sheet context the
        // way a cell formula does, so sheet qualifiers need to be
        // explicit in the rgce. Use u32::MAX as a sentinel sheet
        // index so the encoder's `sheet_index == current_sheet`
        // check never fires; every ref takes the 3D branch.
        const NAME_SHEET_SENTINEL: u32 = u32::MAX;
        walk_node_for_xti(&node, NAME_SHEET_SENTINEL, xti);
        let anchor = CellAnchor { row: 1, column: 1 };
        let rgce = encode_formula_rgce_with_names(
            &node,
            NAME_SHEET_SENTINEL,
            anchor,
            xti,
            &DefinedNameTable::default(),
            extern_names,
            CLS_REF,
        )
        .unwrap_or_default();
        let ilbl = (table.entries.len() + 1) as u16;
        table.lookup.insert((dn.name.to_lowercase(), dn.sheet_id), ilbl);
        table.entries.push(DefinedNameEntry {
            name: dn.name.clone(),
            sheet_id: scope_index,
            rgce,
        });
    }
    table
}

/// Routing table for function names with no BIFF8 iftab. Two routes,
/// mirroring what Excel itself writes ([MS-XLS] + observed output —
/// 2026-08-14 corruption report):
///
///  - **Add-in route** (genuine Analysis ToolPak functions — CONVERT,
///    MROUND, …): PtgNameX pointing at an EXTERNNAME record, wrapped
///    in PtgFuncVar(iftab=255 "User"). On reload calamine renders
///    `User([PtgNameX], args…)`; xls_load's patch_ptg_name_x rewrites
///    `[PtgNameX]` → `_xlfn.<name>` and unwrap_user_xlfn collapses the
///    wrapper to `<name>(…)`.
///  - **NAME route** (everything else): a hidden internal NAME record
///    referenced by PtgName, wrapped in the same PtgFuncVar(255).
///    Post-2003 builtins get the `_xlfn.<name>` future-function form
///    (Excel maps it back to the modern builtin); other names (the MY*
///    UDFs, genuine VBA UDFs) emit under their own name with
///    fFunc|fProc set, which binds to a same-workbook VBA function
///    when the VBA project is present. Routing these through
///    EXTERNNAME instead pointed Excel at a nonexistent add-in:
///    `#NAME?` on every affected cell even with the VBA intact.
#[derive(Default)]
struct ExternNameTable {
    /// Add-in-route names in registration order — PtgNameX nameindex
    /// is 1-based into this list.
    names: Vec<String>,
    /// lowercase_name → 1-based nameindex.
    lookup: std::collections::HashMap<String, u16>,
    /// NAME-route entries: (name as emitted in the Lbl record —
    /// `_xlfn.CONCAT` / `MYUNIQUE` — , is_xlfn). Emitted as Lbl
    /// records directly after the model's defined names.
    name_fns: Vec<(String, bool)>,
    /// canonical lowercase function name → 0-based index into name_fns.
    name_fn_lookup: std::collections::HashMap<String, usize>,
    /// Count of model defined names. Lbl records for name_fns follow
    /// them, so PtgName ilbl = ilbl_base + index + 1 (1-based).
    ilbl_base: u16,
    /// EXTERNSHEET xti index of the add-in supbook entry — the ixti
    /// PtgNameX must carry. Set by build_workbook_stream when `names`
    /// is non-empty.
    addin_ixti: u16,
}

impl ExternNameTable {
    fn lookup_or_add(&mut self, name: &str) -> u16 {
        let key = name.to_lowercase();
        if let Some(&idx) = self.lookup.get(&key) {
            return idx;
        }
        let idx = (self.names.len() + 1) as u16;
        self.names.push(name.to_string());
        self.lookup.insert(key, idx);
        idx
    }
    fn get(&self, name: &str) -> Option<u16> {
        self.lookup.get(&name.to_lowercase()).copied()
    }
    fn add_name_fn(&mut self, name: &str) {
        let key = name.to_lowercase();
        if self.name_fn_lookup.contains_key(&key) {
            return;
        }
        let (emitted, is_xlfn) = match xlfn_route_for_name(name) {
            Some(x) => (x, true),
            None => (name.to_string(), false),
        };
        self.name_fn_lookup.insert(key, self.name_fns.len());
        self.name_fns.push((emitted, is_xlfn));
    }
    fn name_fn_ilbl(&self, name: &str) -> Option<u16> {
        let idx = *self.name_fn_lookup.get(&name.to_lowercase())?;
        Some(self.ilbl_base + idx as u16 + 1)
    }
}

/// Walk every parsed formula in the workbook and register each
/// function whose name isn't in the BIFF FTAB: ATP functions on the
/// add-in route, everything else on the NAME route. Names like
/// "MYUNIQUE" appear once in the table even if used by hundreds of
/// cells.
fn build_extern_name_table(model: &ironcalc::base::Model) -> ExternNameTable {
    use ironcalc::base::expressions::parser::Node;
    let lang = ironcalc::base::language::get_default_language();
    let mut table = ExternNameTable::default();
    table.ilbl_base = model.workbook.defined_names.len() as u16;
    fn walk(node: &Node, lang: &ironcalc::base::language::Language, table: &mut ExternNameTable) {
        match node {
            Node::FunctionKind { kind, args } => {
                let name = kind.to_localized_name(lang);
                if ftab_iftab_for_name(&name).is_none() {
                    if is_atp_function(&name) {
                        table.lookup_or_add(&name);
                    } else {
                        table.add_name_fn(&name);
                    }
                }
                for a in args {
                    walk(a, lang, table);
                }
            }
            Node::OpSumKind { left, right, .. }
            | Node::OpProductKind { left, right, .. }
            | Node::OpPowerKind { left, right }
            | Node::OpConcatenateKind { left, right }
            | Node::CompareKind { left, right, .. }
            | Node::OpRangeKind { left, right } => {
                walk(left, lang, table);
                walk(right, lang, table);
            }
            Node::UnaryKind { right, .. } => walk(right, lang, table),
            _ => {}
        }
    }
    for sheet_formulas in &model.parsed_formulas {
        for node in sheet_formulas {
            walk(node, lang, &mut table);
        }
    }
    table
}

/// EXTERNNAME body for a function-name reference. [MS-XLS] 2.4.69.
/// Layout we emit (matches xls_biff.rs reader expectations):
///   options(2) = 0x0000, one(2) = 0x0001, itab(2) = 0x0000,
///   cch(1), grbit(1) = 0x00 for compressed, name chars,
///   cce(2) + rgce — for ATP/add-in fns this is 0x02, 0x00, 0x1C, 0x17
///   (cce=2; PtgErr 0x17 = #VALUE!), the convention real Excel uses
///   for unbound-add-in references.
fn build_externname_record(name: &str) -> Vec<u8> {
    let chars: Vec<u16> = name.encode_utf16().collect();
    let high_byte = chars.iter().any(|&c| c > 0xFF);
    let cch = chars.len().min(255) as u8;
    let mut body = Vec::with_capacity(8 + chars.len() * 2 + 4);
    body.put_u16(0x0000); // options
    body.put_u16(0x0000); // reserved — Excel writes zeros here, not 1
    body.put_u16(0x0000); // itab
    body.put_u8(cch);
    body.put_u8(if high_byte { 0x01 } else { 0x00 });
    if high_byte {
        for c in chars.iter().take(cch as usize) { body.put_u16(*c); }
    } else {
        for c in chars.iter().take(cch as usize) { body.put_u8(*c as u8); }
    }
    // Trailing formula data — unbound add-in placeholder.
    body.put_u16(0x0002); // cce
    body.put_u8(0x1C);    // PtgErr
    body.put_u8(0x17);    // #VALUE!
    body
}

/// Build the body of one function-name Lbl (NAME) record, matching
/// Excel's own encoding byte-for-byte (dumped from an Excel-authored
/// file, 2026-08-14):
///  - future function (`_xlfn.*`): flags 0x038B (fHidden|fFunc|fProc,
///    fGrp=14 "User Defined"), rgce = PtgErr #NAME? (0x1C 0x1D)
///  - VBA UDF / other: flags 0x000E (fFunc|fOB|fProc), empty rgce
fn build_fn_lbl_record(name: &str, is_xlfn: bool) -> Vec<u8> {
    let chars: Vec<u16> = name.encode_utf16().collect();
    let high_byte = chars.iter().any(|&c| c > 0xFF);
    let cch = chars.len().min(255) as u8;
    let rgce: &[u8] = if is_xlfn { &[0x1C, 0x1D] } else { &[] };
    let mut body = Vec::with_capacity(14 + chars.len() * 2 + rgce.len());
    body.put_u16(if is_xlfn { 0x038B } else { 0x000E }); // grbit
    body.put_u8(0);  // chKey
    body.put_u8(cch);
    body.put_u16(rgce.len() as u16); // cce
    body.put_u16(0); // ixals (must be 0 in BIFF8)
    body.put_u16(0); // itab — global scope
    body.put_u8(0); // cchCustMenu
    body.put_u8(0); // cchDescription
    body.put_u8(0); // cchHelpTopic
    body.put_u8(0); // cchStatusText
    body.put_u8(if high_byte { 0x01 } else { 0x00 });
    if high_byte {
        for c in chars.iter().take(cch as usize) { body.put_u16(*c); }
    } else {
        for c in chars.iter().take(cch as usize) { body.put_u8(*c as u8); }
    }
    body.extend_from_slice(rgce);
    body
}

/// Select which preserved source NAME records get replayed and give
/// each its final ilbl. Skips names the writer already emits (a model
/// defined name, or a function NAME on the `_xlfn.*` / VBA-UDF
/// route), and sheet-scoped names whose scope sheet no longer exists.
/// Returns (index into `preserved_names`, Lbl itab) per record, in
/// emission order; ilbls continue the numbering after the model
/// names and the function NAMEs and are registered in the
/// defined-name lookup so preserved rgce referencing them resolves.
fn plan_preserved_names(
    preserved: Option<&crate::xls_preserve::PreservedXlsData>,
    defined_names: &mut DefinedNameTable,
    extern_names: &ExternNameTable,
    sheet_by_name: &HashMap<String, u32>,
) -> Vec<(usize, u16)> {
    let mut plans: Vec<(usize, u16)> = Vec::new();
    let Some(p) = preserved else { return plans };
    let emitted_fn_names: std::collections::HashSet<String> = extern_names
        .name_fns
        .iter()
        .map(|(n, _)| n.to_lowercase())
        .collect();
    let base = defined_names.entries.len() + extern_names.name_fns.len();
    for (idx, pn) in p.preserved_names.iter().enumerate() {
        if pn.name.is_empty() {
            continue;
        }
        let key = pn.name.to_lowercase();
        if defined_names.lookup.contains_key(&(key.clone(), None))
            || emitted_fn_names.contains(&key)
        {
            continue;
        }
        let (itab, scope) = match &pn.sheet_name {
            None => (0u16, None),
            Some(sn) => match sheet_by_name.get(&sn.to_lowercase()) {
                Some(&i) => (i as u16 + 1, Some(i)),
                None => continue,
            },
        };
        let ilbl = (base + plans.len() + 1) as u16;
        defined_names.lookup.insert((key, scope), ilbl);
        plans.push((idx, itab));
    }
    plans
}

/// Lbl record body for a preserved source NAME, replayed with its
/// original flags / chKey and a remapped rgce. Same field layout as
/// `build_lbl_record` / `build_fn_lbl_record` ([MS-XLS] 2.4.187).
fn build_preserved_lbl_record(
    pn: &crate::xls_preserve::PreservedName,
    itab: u16,
    rgce: &[u8],
) -> Vec<u8> {
    let chars: Vec<u16> = pn.name.encode_utf16().collect();
    let cch = chars.len().min(255) as u8;
    let high_byte = chars.iter().any(|&c| c > 0xFF);
    let mut body = Vec::with_capacity(15 + chars.len() * 2 + rgce.len());
    body.put_u16(pn.grbit);
    body.put_u8(pn.ch_key);
    body.put_u8(cch);
    body.put_u16(rgce.len() as u16);
    body.put_u16(0); // ixals (reserved)
    body.put_u16(itab);
    body.put_u8(0); // cchCustMenu
    body.put_u8(0); // cchDescription
    body.put_u8(0); // cchHelpTopic
    body.put_u8(0); // cchStatusText
    body.put_u8(if high_byte { 0x01 } else { 0x00 });
    if high_byte {
        for c in chars.iter().take(cch as usize) {
            body.put_u16(*c);
        }
    } else {
        for c in chars.iter().take(cch as usize) {
            body.put_u8(*c as u8);
        }
    }
    body.extend_from_slice(rgce);
    body
}

/// Build the body of one Lbl (NAME) record. [MS-XLS] 2.4.187.
fn build_lbl_record(entry: &DefinedNameEntry) -> Vec<u8> {
    let chars: Vec<u16> = entry.name.encode_utf16().collect();
    let high_byte = chars.iter().any(|&c| c > 0xFF);
    let cch = chars.len().min(255) as u8;
    let mut body = Vec::with_capacity(14 + chars.len() * 2 + entry.rgce.len());
    body.put_u16(0); // grbit (no special flags)
    body.put_u8(0);  // chKey
    body.put_u8(cch);
    body.put_u16(entry.rgce.len() as u16); // cce
    body.put_u16(0); // ixals (must be 0 in BIFF8)
    body.put_u16(entry.sheet_id.map(|s| s as u16 + 1).unwrap_or(0)); // itab
    body.put_u8(0); // cchCustMenu
    body.put_u8(0); // cchDescription
    body.put_u8(0); // cchHelpTopic
    body.put_u8(0); // cchStatusText
    body.put_u8(if high_byte { 0x01 } else { 0x00 });
    if high_byte {
        for c in chars.iter().take(cch as usize) { body.put_u16(*c); }
    } else {
        for c in chars.iter().take(cch as usize) { body.put_u8(*c as u8); }
    }
    body.extend_from_slice(&entry.rgce);
    body
}

/// Operand-class bits ([MS-XLS] 2.5.198.2), OR'd into the base ptg id
/// of operand tokens (Ref 0x04, Area 0x05, Ref3d 0x1A, Area3d 0x1B,
/// Name 0x03, NameX 0x19, FuncVar 0x02). Excel's recalc engine is
/// strict about them in places: a reference-class PtgName or Ptg*3d
/// where a value is required (formula root, operator operand) makes
/// every recalculated dependent show #VALUE! — verified 2026-08-14 by
/// flipping a single class byte (0x3A→0x5A) in a broken save and
/// watching the cell recover. Ref class is only correct where Excel
/// itself writes it: ref-taking builtin args (per FN_PARAM_CLASSES)
/// and User-call (VBA UDF / _xlfn / add-in) ref args.
const CLS_REF: u8 = 0x20;
const CLS_VAL: u8 = 0x40;

/// Encode an IronCalc Node tree to BIFF8 rgce bytes. `root_cls` is
/// the operand class required by the rgce's context: CLS_VAL for cell
/// formulas (the result is a value), CLS_REF for defined-name
/// definitions (a name is a reference). Returns None when any subtree
/// contains a construct we can't yet encode — the caller then falls
/// back to the placeholder rgce path. We deliberately fail-closed
/// rather than try to half-encode, because a malformed rgce rendered
/// through a BIFF reader produces nonsense formulas (the parser walks
/// ptgs by fixed sizes and mis-aligns on garbage).
fn encode_formula_rgce_with_names(
    node: &ironcalc::base::expressions::parser::Node,
    sheet_idx: u32,
    anchor: CellAnchor,
    xti: &XtiTable,
    defined_names: &DefinedNameTable,
    extern_names: &ExternNameTable,
    root_cls: u8,
) -> Option<Vec<u8>> {
    let mut rgce = Vec::with_capacity(64);
    encode_node(&mut rgce, node, sheet_idx, anchor, xti, defined_names, extern_names, root_cls).ok()?;
    Some(rgce)
}

/// True for Node variants that compile to a multi-ptg expression
/// (binary ops, range concatenation). Used to decide when a child
/// operand needs PtgParen wrapping so text-based renderers like
/// calamine reproduce the original grouping. Without this, formulas
/// like `(A+B+C)/5` lose their parens and re-parse as `A+B+C/5`,
/// changing the result via operator precedence.
fn is_compound_binop(node: &ironcalc::base::expressions::parser::Node) -> bool {
    use ironcalc::base::expressions::parser::Node;
    matches!(
        node,
        Node::OpSumKind { .. }
            | Node::OpProductKind { .. }
            | Node::OpPowerKind { .. }
            | Node::OpConcatenateKind { .. }
            | Node::CompareKind { .. }
            | Node::OpRangeKind { .. }
    )
}

/// Encode `node`, then emit PtgParen if it's a compound op so that
/// text-rendering of the parent expression preserves grouping.
fn encode_node_parenthesized(
    rgce: &mut Vec<u8>,
    node: &ironcalc::base::expressions::parser::Node,
    sheet_idx: u32,
    anchor: CellAnchor,
    xti: &XtiTable,
    defined_names: &DefinedNameTable,
    extern_names: &ExternNameTable,
    cls: u8,
) -> Result<(), ()> {
    encode_node(rgce, node, sheet_idx, anchor, xti, defined_names, extern_names, cls)?;
    if is_compound_binop(node) {
        rgce.put_u8(0x15); // PtgParen — passive; affects rendering only
    }
    Ok(())
}

/// `cls` is the operand class this node's result must carry (CLS_REF
/// or CLS_VAL). It lands on direct operand ptgs (refs, areas, names)
/// and on PtgFuncVar; operator operands always evaluate as values so
/// the recursion resets to CLS_VAL below them, and function args get
/// their class from FN_PARAM_CLASSES (builtins) or CLS_REF (User
/// calls) inside emit_function_call_named.
fn encode_node(
    rgce: &mut Vec<u8>,
    node: &ironcalc::base::expressions::parser::Node,
    sheet_idx: u32,
    anchor: CellAnchor,
    xti: &XtiTable,
    defined_names: &DefinedNameTable,
    extern_names: &ExternNameTable,
    cls: u8,
) -> Result<(), ()> {
    use ironcalc::base::expressions::parser::Node;
    use ironcalc::base::expressions::token::{OpCompare, OpProduct, OpSum, OpUnary};
    match node {
        Node::NumberKind(n) => emit_number_literal(rgce, *n),
        Node::BooleanKind(b) => {
            rgce.put_u8(0x1D); // PtgBool
            rgce.put_u8(if *b { 1 } else { 0 });
        }
        Node::StringKind(s) => emit_string_literal(rgce, s)?,
        Node::ErrorKind(e) => {
            rgce.put_u8(0x1C); // PtgErr
            rgce.put_u8(ironcalc_error_to_biff(e));
        }
        Node::ReferenceKind {
            sheet_index,
            absolute_row,
            absolute_column,
            row,
            column,
            ..
        } => {
            if *sheet_index == sheet_idx {
                emit_ptg_ref(rgce, *row, *column, *absolute_row, *absolute_column, anchor, cls)?;
            } else {
                let ixti = xti.lookup_single(*sheet_index).ok_or(())?;
                emit_ptg_ref3d(rgce, ixti, *row, *column, *absolute_row, *absolute_column, anchor, cls)?;
            }
        }
        Node::RangeKind {
            sheet_index,
            absolute_row1,
            absolute_column1,
            row1,
            column1,
            absolute_row2,
            absolute_column2,
            row2,
            column2,
            ..
        } => {
            if *sheet_index == sheet_idx {
                emit_ptg_area(
                    rgce,
                    *row1, *column1, *absolute_row1, *absolute_column1,
                    *row2, *column2, *absolute_row2, *absolute_column2,
                    anchor, cls,
                )?;
            } else {
                let ixti = xti.lookup_single(*sheet_index).ok_or(())?;
                emit_ptg_area3d(
                    rgce, ixti,
                    *row1, *column1, *absolute_row1, *absolute_column1,
                    *row2, *column2, *absolute_row2, *absolute_column2,
                    anchor, cls,
                )?;
            }
        }
        Node::OpSumKind { kind, left, right } => {
            encode_node_parenthesized(rgce, left, sheet_idx, anchor, xti, defined_names, extern_names, CLS_VAL)?;
            encode_node_parenthesized(rgce, right, sheet_idx, anchor, xti, defined_names, extern_names, CLS_VAL)?;
            rgce.put_u8(match kind {
                OpSum::Add => 0x03,    // PtgAdd
                OpSum::Minus => 0x04,  // PtgSub
            });
        }
        Node::OpProductKind { kind, left, right } => {
            encode_node_parenthesized(rgce, left, sheet_idx, anchor, xti, defined_names, extern_names, CLS_VAL)?;
            encode_node_parenthesized(rgce, right, sheet_idx, anchor, xti, defined_names, extern_names, CLS_VAL)?;
            rgce.put_u8(match kind {
                OpProduct::Times => 0x05,   // PtgMul
                OpProduct::Divide => 0x06,  // PtgDiv
            });
        }
        Node::OpPowerKind { left, right } => {
            encode_node_parenthesized(rgce, left, sheet_idx, anchor, xti, defined_names, extern_names, CLS_VAL)?;
            encode_node_parenthesized(rgce, right, sheet_idx, anchor, xti, defined_names, extern_names, CLS_VAL)?;
            rgce.put_u8(0x07); // PtgPower
        }
        Node::OpConcatenateKind { left, right } => {
            encode_node_parenthesized(rgce, left, sheet_idx, anchor, xti, defined_names, extern_names, CLS_VAL)?;
            encode_node_parenthesized(rgce, right, sheet_idx, anchor, xti, defined_names, extern_names, CLS_VAL)?;
            rgce.put_u8(0x08); // PtgConcat
        }
        Node::CompareKind { kind, left, right } => {
            encode_node_parenthesized(rgce, left, sheet_idx, anchor, xti, defined_names, extern_names, CLS_VAL)?;
            encode_node_parenthesized(rgce, right, sheet_idx, anchor, xti, defined_names, extern_names, CLS_VAL)?;
            rgce.put_u8(match kind {
                OpCompare::LessThan => 0x09,            // PtgLT
                OpCompare::LessOrEqualThan => 0x0A,     // PtgLE
                OpCompare::Equal => 0x0B,               // PtgEQ
                OpCompare::GreaterOrEqualThan => 0x0C,  // PtgGE
                OpCompare::GreaterThan => 0x0D,         // PtgGT
                OpCompare::NonEqual => 0x0E,            // PtgNE
            });
        }
        Node::UnaryKind { kind, right } => {
            encode_node_parenthesized(rgce, right, sheet_idx, anchor, xti, defined_names, extern_names, CLS_VAL)?;
            rgce.put_u8(match kind {
                OpUnary::Minus => 0x13,       // PtgUminus
                OpUnary::Percentage => 0x14,  // PtgPercent
            });
        }
        Node::FunctionKind { kind, args } => {
            // `kind` is IronCalc's internal Function enum value; we
            // can't name the type (private module) but we can call
            // its methods. to_localized_name with the default English
            // language gives us the canonical Excel function name.
            let lang = ironcalc::base::language::get_default_language();
            let name = kind.to_localized_name(lang);
            emit_function_call_named(rgce, &name, args, sheet_idx, anchor, xti, defined_names, extern_names, cls)?;
        }
        Node::DefinedNameKind((name, scope, _formula)) => {
            let ilbl = defined_names.ilbl_for(name, *scope).ok_or(())?;
            rgce.put_u8(0x03 | cls); // PtgName
            // Wire: u16 ilbl + u16 reserved. Calamine reads as one
            // u32 nameindex (LE), so this comes out 1-based.
            rgce.put_u16(ilbl);
            rgce.put_u16(0);
        }
        // Everything else — tables, array literals, parse errors —
        // falls back to the caller's placeholder rgce.
        _ => return Err(()),
    }
    Ok(())
}

/// Pick the most compact ptg for a numeric literal. Integers in
/// 0..=65535 use PtgInt (3 bytes); everything else PtgNum (9 bytes).
fn emit_number_literal(rgce: &mut Vec<u8>, n: f64) {
    if n.is_finite() && n.fract() == 0.0 && n >= 0.0 && n <= u16::MAX as f64 {
        rgce.put_u8(0x1E); // PtgInt
        rgce.put_u16(n as u16);
    } else {
        rgce.put_u8(0x1F); // PtgNum
        rgce.put_f64(n);
    }
}

/// PtgStr — opcode + u8 cch + flag + chars. u8 cch limits the literal
/// to 255 chars; longer strings make the encoder bail (caller falls
/// back to LABELSST routing for cell values that wouldn't fit).
///
/// Important: IronCalc's parser stores `Node::StringKind` with the
/// Excel-escape form preserved — i.e. `="6"" headboard"` parses to
/// StringKind("6\"\" headboard") (13 chars, two `"`s) even though
/// the evaluator unescapes `""` → `"` when computing the cached
/// value (one `"`, 12 chars). PtgStr stores RAW chars in BIFF
/// (no escaping in the wire format), so we must unescape before
/// writing — otherwise calamine on reload reads two quote chars
/// from PtgStr verbatim, then patch_ptg_strings re-doubles each
/// to `""""`, IronCalc unescapes once → two quotes survive in the
/// reloaded value. Each round-trip would add another doubling.
fn emit_string_literal(rgce: &mut Vec<u8>, s: &str) -> Result<(), ()> {
    let unescaped = s.replace("\"\"", "\"");
    let chars: Vec<u16> = unescaped.encode_utf16().collect();
    if chars.len() > 255 {
        return Err(());
    }
    let high_byte = chars.iter().any(|&c| c > 0xFF);
    if high_byte {
        // Calamine 0x26.1 has a read bug for high-byte PtgStr: it
        // advances `rgce[2 + cch..]` regardless of whether each char
        // takes 1 or 2 bytes (xls.rs:1271), so 2-byte-encoded content
        // mis-aligns the rgce stream and subsequent ptgs decode as
        // garbage. Bail rather than emit a corrupt rgce — the caller
        // falls back to placeholder + (for CellFormulaString) LABELSST
        // routing, which round-trips the cached value via the SST in
        // full Unicode without going through PtgStr.
        return Err(());
    }
    rgce.put_u8(0x17); // PtgStr
    rgce.put_u8(chars.len() as u8);
    rgce.put_u8(0x00); // flag = compressed (high_byte=false guaranteed)
    for c in &chars { rgce.put_u8(*c as u8); }
    Ok(())
}

/// Resolve an IronCalc ref (row + abs_flag) to a BIFF-storage 1-based
/// absolute target. IronCalc stores relative refs as offsets from the
/// formula's anchor cell (e.g., "R[9]" with anchor row=26 has Node.row
/// = 9). BIFF stores the absolute target with a flag — convert by
/// adding the anchor.
fn resolve_to_absolute_1based(coord: i32, abs_flag: bool, anchor_coord: i32) -> i32 {
    if abs_flag { coord } else { coord + anchor_coord }
}

/// PtgRef (base 0x04 | operand class). [MS-XLS] 2.5.198 RgceLoc.
/// Resolves IronCalc's relative-ref offsets to absolute targets via
/// the anchor cell, then converts 1-based → 0-based for BIFF storage.
/// Bits 0..13 of col_word = column; bit 14 = fColRel; bit 15 = fRwRel.
fn emit_ptg_ref(
    rgce: &mut Vec<u8>,
    row: i32,
    column: i32,
    absolute_row: bool,
    absolute_column: bool,
    anchor: CellAnchor,
    cls: u8,
) -> Result<(), ()> {
    let abs_row = resolve_to_absolute_1based(row, absolute_row, anchor.row);
    let abs_col = resolve_to_absolute_1based(column, absolute_column, anchor.column);
    let r0 = abs_row.checked_sub(1).ok_or(())? as u32;
    let c0 = abs_col.checked_sub(1).ok_or(())? as u32;
    if r0 > 0xFFFF || c0 > 0x3FFF { return Err(()); }
    let mut col_word = c0 as u16 & 0x3FFF;
    if !absolute_column { col_word |= 0x4000; }
    if !absolute_row    { col_word |= 0x8000; }
    rgce.put_u8(0x04 | cls);
    rgce.put_u16(r0 as u16);
    rgce.put_u16(col_word);
    Ok(())
}

/// PtgArea (base 0x05 | operand class). [MS-XLS] 2.5.198a — pair of RgceLocs.
fn emit_ptg_area(
    rgce: &mut Vec<u8>,
    row1: i32, column1: i32, abs_row1: bool, abs_col1: bool,
    row2: i32, column2: i32, abs_row2: bool, abs_col2: bool,
    anchor: CellAnchor,
    cls: u8,
) -> Result<(), ()> {
    let r1_abs = resolve_to_absolute_1based(row1, abs_row1, anchor.row);
    let c1_abs = resolve_to_absolute_1based(column1, abs_col1, anchor.column);
    let r2_abs = resolve_to_absolute_1based(row2, abs_row2, anchor.row);
    let c2_abs = resolve_to_absolute_1based(column2, abs_col2, anchor.column);
    let r1 = r1_abs.checked_sub(1).ok_or(())? as u32;
    let c1 = c1_abs.checked_sub(1).ok_or(())? as u32;
    let r2 = r2_abs.checked_sub(1).ok_or(())? as u32;
    let c2 = c2_abs.checked_sub(1).ok_or(())? as u32;
    if r1 > 0xFFFF || r2 > 0xFFFF || c1 > 0x3FFF || c2 > 0x3FFF { return Err(()); }
    let pack = |c: u32, abs_c: bool, abs_r: bool| -> u16 {
        let mut w = c as u16 & 0x3FFF;
        if !abs_c { w |= 0x4000; }
        if !abs_r { w |= 0x8000; }
        w
    };
    rgce.put_u8(0x05 | cls);
    rgce.put_u16(r1 as u16);
    rgce.put_u16(r2 as u16);
    rgce.put_u16(pack(c1, abs_col1, abs_row1));
    rgce.put_u16(pack(c2, abs_col2, abs_row2));
    Ok(())
}

/// Emit a function call as PtgFuncVar. We look the function up by its
/// English name (via `to_localized_name`) rather than matching on
/// the Function enum variants — IronCalc's `functions` module isn't
/// publicly re-exported, so the variant identifiers aren't reachable
/// from outside the crate. The string-keyed lookup is morally
/// equivalent and decouples the writer from internal renames.
fn emit_function_call_named(
    rgce: &mut Vec<u8>,
    func_name: &str,
    args: &[ironcalc::base::expressions::parser::Node],
    sheet_idx: u32,
    anchor: CellAnchor,
    xti: &XtiTable,
    defined_names: &DefinedNameTable,
    extern_names: &ExternNameTable,
    cls: u8,
) -> Result<(), ()> {
    if args.len() > 255 { return Err(()); }
    if let Some(iftab) = ftab_iftab_for_name(func_name) {
        // Builtin args carry the operand class Excel's own writer
        // uses, from the per-function table (last class repeats for
        // varargs). Missing table entry → value class, the common
        // case. Getting this wrong breaks recalc in Excel: a
        // ref-class arg where the function wants a value makes the
        // call error, and a value-class arg where it wants a
        // reference (SUM, OFFSET, ...) collapses the range via
        // implicit intersection.
        let params = crate::xls_fn_classes::FN_PARAM_CLASSES
            .get(iftab as usize)
            .copied()
            .flatten();
        for (i, arg) in args.iter().enumerate() {
            let arg_cls = params
                .and_then(|p| p.get(i).or(p.last()))
                .copied()
                .unwrap_or(CLS_VAL);
            encode_node(rgce, arg, sheet_idx, anchor, xti, defined_names, extern_names, arg_cls)?;
        }
        rgce.put_u8(0x02 | cls); // PtgFuncVar
        rgce.put_u8(args.len() as u8);
        rgce.put_u16(iftab);
        return Ok(());
    }
    if args.len().saturating_add(1) > 255 { return Err(()); }
    // NAME route: future functions (_xlfn.*) and VBA UDFs reference an
    // internal NAME record via PtgName; the Lbl records are emitted
    // right after the model's defined names, so ilbl continues their
    // numbering. Same PtgFuncVar(255) wrapper as the add-in route.
    // User-call ref/name args keep reference class — that is what
    // Excel writes for array-entered VBA UDF calls (neutral.xls ARRAY
    // rgces), and it is what delivers a Range object to a VBA
    // `As Range` parameter.
    if let Some(ilbl) = extern_names.name_fn_ilbl(func_name) {
        rgce.put_u8(0x23); // PtgName (fn-name slot is always ref class)
        rgce.put_u16(ilbl);
        rgce.put_u16(0);   // reserved (calamine reads ilbl as u32 LE)
        for arg in args {
            encode_node(rgce, arg, sheet_idx, anchor, xti, defined_names, extern_names, CLS_REF)?;
        }
        rgce.put_u8(0x02 | cls); // PtgFuncVar
        rgce.put_u8((args.len() + 1) as u8); // argc = args + 1 for the name
        rgce.put_u16(0xFF); // iftab = 255 ("User": first arg is the fn)
        return Ok(());
    }
    // Add-in route (genuine ATP functions). PtgNameX appears as the
    // first arg (the "function name"); calamine's PtgFuncVar with
    // iftab=255 ("User") renders the call as `User([PtgNameX], arg1,
    // ...)`. xls_load.rs's patch_ptg_name_x + unwrap_user_xlfn
    // collapse that back to `<name>(arg1, ...)` on reload.
    let nameindex = extern_names.get(func_name).ok_or(())?;
    rgce.put_u8(0x39);    // PtgNameX (fn-name slot is always ref class)
    rgce.put_u16(extern_names.addin_ixti); // xti entry of the add-in supbook
    rgce.put_u16(nameindex);
    rgce.put_u16(0);      // reserved
    for arg in args {
        encode_node(rgce, arg, sheet_idx, anchor, xti, defined_names, extern_names, CLS_REF)?;
    }
    rgce.put_u8(0x02 | cls); // PtgFuncVar
    rgce.put_u8((args.len() + 1) as u8); // argc = args + 1 for the name itself
    rgce.put_u16(0xFF);   // iftab = 255 (FTAB[255] = "User")
    Ok(())
}

/// PtgRef3d (base 0x1A | operand class). [MS-XLS] 2.5.198.97. Layout:
/// ptg + ixti(u16) + rwu(u16) + colu_with_flags(u16). The ixti
/// indexes into the EXTERNSHEET xti table, which in turn maps to a
/// (supBook, sheet_first, sheet_last) tuple. For self-referencing
/// 3D refs we use a single SUPBOOK marker at index 0.
fn emit_ptg_ref3d(
    rgce: &mut Vec<u8>,
    ixti: u16,
    row: i32,
    column: i32,
    absolute_row: bool,
    absolute_column: bool,
    anchor: CellAnchor,
    cls: u8,
) -> Result<(), ()> {
    let abs_row = resolve_to_absolute_1based(row, absolute_row, anchor.row);
    let abs_col = resolve_to_absolute_1based(column, absolute_column, anchor.column);
    let r0 = abs_row.checked_sub(1).ok_or(())? as u32;
    let c0 = abs_col.checked_sub(1).ok_or(())? as u32;
    if r0 > 0xFFFF || c0 > 0x3FFF { return Err(()); }
    let mut col_word = c0 as u16 & 0x3FFF;
    if !absolute_column { col_word |= 0x4000; }
    if !absolute_row    { col_word |= 0x8000; }
    rgce.put_u8(0x1A | cls);
    rgce.put_u16(ixti);
    rgce.put_u16(r0 as u16);
    rgce.put_u16(col_word);
    Ok(())
}

/// PtgArea3d (base 0x1B | operand class). Layout: ptg + ixti(u16) +
/// rwFirst(u16) + rwLast(u16) + colFirst_w_flags(u16) + colLast_w_flags(u16).
fn emit_ptg_area3d(
    rgce: &mut Vec<u8>,
    ixti: u16,
    row1: i32, column1: i32, abs_row1: bool, abs_col1: bool,
    row2: i32, column2: i32, abs_row2: bool, abs_col2: bool,
    anchor: CellAnchor,
    cls: u8,
) -> Result<(), ()> {
    let r1_abs = resolve_to_absolute_1based(row1, abs_row1, anchor.row);
    let c1_abs = resolve_to_absolute_1based(column1, abs_col1, anchor.column);
    let r2_abs = resolve_to_absolute_1based(row2, abs_row2, anchor.row);
    let c2_abs = resolve_to_absolute_1based(column2, abs_col2, anchor.column);
    let r1 = r1_abs.checked_sub(1).ok_or(())? as u32;
    let c1 = c1_abs.checked_sub(1).ok_or(())? as u32;
    let r2 = r2_abs.checked_sub(1).ok_or(())? as u32;
    let c2 = c2_abs.checked_sub(1).ok_or(())? as u32;
    if r1 > 0xFFFF || r2 > 0xFFFF || c1 > 0x3FFF || c2 > 0x3FFF { return Err(()); }
    let pack = |c: u32, abs_c: bool, abs_r: bool| -> u16 {
        let mut w = c as u16 & 0x3FFF;
        if !abs_c { w |= 0x4000; }
        if !abs_r { w |= 0x8000; }
        w
    };
    rgce.put_u8(0x1B | cls);
    rgce.put_u16(ixti);
    rgce.put_u16(r1 as u16);
    rgce.put_u16(r2 as u16);
    rgce.put_u16(pack(c1, abs_col1, abs_row1));
    rgce.put_u16(pack(c2, abs_col2, abs_row2));
    Ok(())
}

/// Map a BIFF function name to its FTAB iftab index. Names are
/// uppercased ASCII (Excel's canonical form). The table is the
/// complete calamine FTAB (`calamine-0.26.1/src/utils.rs::FTAB`,
/// whose indices match [MS-XLS] 2.5.198.17) capped at index 368
/// (BAHTTEXT) — the last function Excel 2003 knows. This cap is
/// load-bearing three ways:
///
///  1. Completeness: the previous hand-curated ~120-name subset sent
///     genuine builtins (SMALL, LARGE, …) down the add-in EXTERNNAME
///     route, which Excel renders as `#NAME?` calls to a nonexistent
///     add-in (2026-08-14 corruption report, Issue 2).
///  2. The 384..=476 range is the Analysis ToolPak set — Excel never
///     writes those as iftabs in .xls; it uses add-in EXTERNNAME
///     references. See `is_atp_function`.
///  3. The 480..=484 tail (IFERROR, COUNTIFS, SUMIFS, AVERAGEIF,
///     AVERAGEIFS) are Excel-2007 functions that Excel encodes as
///     `_xlfn.*` future-function NAMEs in BIFF8, not iftabs. Emitting
///     iftab 480+ round-trips through calamine but is not portable —
///     route them via the _xlfn path instead (`xlfn_route_for_name`).
fn ftab_iftab_for_name(name: &str) -> Option<u16> {
    Some(match name {
        "COUNT" => 0,
        "IF" => 1,
        "ISNA" => 2,
        "ISERROR" => 3,
        "SUM" => 4,
        "AVERAGE" => 5,
        "MIN" => 6,
        "MAX" => 7,
        "ROW" => 8,
        "COLUMN" => 9,
        "NA" => 10,
        "NPV" => 11,
        "STDEV" => 12,
        "DOLLAR" => 13,
        "FIXED" => 14,
        "SIN" => 15,
        "COS" => 16,
        "TAN" => 17,
        "ATAN" => 18,
        "PI" => 19,
        "SQRT" => 20,
        "EXP" => 21,
        "LN" => 22,
        "LOG10" => 23,
        "ABS" => 24,
        "INT" => 25,
        "SIGN" => 26,
        "ROUND" => 27,
        "LOOKUP" => 28,
        "INDEX" => 29,
        "REPT" => 30,
        "MID" => 31,
        "LEN" => 32,
        "VALUE" => 33,
        "TRUE" => 34,
        "FALSE" => 35,
        "AND" => 36,
        "OR" => 37,
        "NOT" => 38,
        "MOD" => 39,
        "DCOUNT" => 40,
        "DSUM" => 41,
        "DAVERAGE" => 42,
        "DMIN" => 43,
        "DMAX" => 44,
        "DSTDEV" => 45,
        "VAR" => 46,
        "DVAR" => 47,
        "TEXT" => 48,
        "LINEST" => 49,
        "TREND" => 50,
        "LOGEST" => 51,
        "GROWTH" => 52,
        "GOTO" => 53,
        "HALT" => 54,
        "RETURN" => 55,
        "PV" => 56,
        "FV" => 57,
        "NPER" => 58,
        "PMT" => 59,
        "RATE" => 60,
        "MIRR" => 61,
        "IRR" => 62,
        "RAND" => 63,
        "MATCH" => 64,
        "DATE" => 65,
        "TIME" => 66,
        "DAY" => 67,
        "MONTH" => 68,
        "YEAR" => 69,
        "WEEKDAY" => 70,
        "HOUR" => 71,
        "MINUTE" => 72,
        "SECOND" => 73,
        "NOW" => 74,
        "AREAS" => 75,
        "ROWS" => 76,
        "COLUMNS" => 77,
        "OFFSET" => 78,
        "ABSREF" => 79,
        "RELREF" => 80,
        "ARGUMENT" => 81,
        "SEARCH" => 82,
        "TRANSPOSE" => 83,
        "ERROR" => 84,
        "STEP" => 85,
        "TYPE" => 86,
        "ECHO" => 87,
        "SET.NAME" => 88,
        "CALLER" => 89,
        "DEREF" => 90,
        "WINDOWS" => 91,
        "SERIES" => 92,
        "DOCUMENTS" => 93,
        "ACTIVE.CELL" => 94,
        "SELECTION" => 95,
        "RESULT" => 96,
        "ATAN2" => 97,
        "ASIN" => 98,
        "ACOS" => 99,
        "CHOOSE" => 100,
        "HLOOKUP" => 101,
        "VLOOKUP" => 102,
        "LINKS" => 103,
        "INPUT" => 104,
        "ISREF" => 105,
        "GET.FORMULA" => 106,
        "GET.NAME" => 107,
        "SET.VALUE" => 108,
        "LOG" => 109,
        "EXEC" => 110,
        "CHAR" => 111,
        "LOWER" => 112,
        "UPPER" => 113,
        "PROPER" => 114,
        "LEFT" => 115,
        "RIGHT" => 116,
        "EXACT" => 117,
        "TRIM" => 118,
        "REPLACE" => 119,
        "SUBSTITUTE" => 120,
        "CODE" => 121,
        "NAMES" => 122,
        "DIRECTORY" => 123,
        "FIND" => 124,
        "CELL" => 125,
        "ISERR" => 126,
        "ISTEXT" => 127,
        "ISNUMBER" => 128,
        "ISBLANK" => 129,
        "T" => 130,
        "N" => 131,
        "FOPEN" => 132,
        "FCLOSE" => 133,
        "FSIZE" => 134,
        "FREADLN" => 135,
        "FREAD" => 136,
        "FWRITELN" => 137,
        "FWRITE" => 138,
        "FPOS" => 139,
        "DATEVALUE" => 140,
        "TIMEVALUE" => 141,
        "SLN" => 142,
        "SYD" => 143,
        "DDB" => 144,
        "GET.DEF" => 145,
        "REFTEXT" => 146,
        "TEXTREF" => 147,
        "INDIRECT" => 148,
        "REGISTER" => 149,
        "CALL" => 150,
        "ADD.BAR" => 151,
        "ADD.MENU" => 152,
        "ADD.COMMAND" => 153,
        "ENABLE.COMMAND" => 154,
        "CHECK.COMMAND" => 155,
        "RENAME.COMMAND" => 156,
        "SHOW.BAR" => 157,
        "DELETE.MENU" => 158,
        "DELETE.COMMAND" => 159,
        "GET.CHART.ITEM" => 160,
        "DIALOG.BOX" => 161,
        "CLEAN" => 162,
        "MDETERM" => 163,
        "MINVERSE" => 164,
        "MMULT" => 165,
        "FILES" => 166,
        "IPMT" => 167,
        "PPMT" => 168,
        "COUNTA" => 169,
        "CANCEL.KEY" => 170,
        "FOR" => 171,
        "WHILE" => 172,
        "BREAK" => 173,
        "NEXT" => 174,
        "INITIATE" => 175,
        "REQUEST" => 176,
        "POKE" => 177,
        "EXECUTE" => 178,
        "TERMINATE" => 179,
        "RESTART" => 180,
        "HELP" => 181,
        "GET.BAR" => 182,
        "PRODUCT" => 183,
        "FACT" => 184,
        "GET.CELL" => 185,
        "GET.WORKSPACE" => 186,
        "GET.WINDOW" => 187,
        "GET.DOCUMENT" => 188,
        "DPRODUCT" => 189,
        "ISNONTEXT" => 190,
        "GET.NOTE" => 191,
        "NOTE" => 192,
        "STDEVP" => 193,
        "VARP" => 194,
        "DSTDEVP" => 195,
        "DVARP" => 196,
        "TRUNC" => 197,
        "ISLOGICAL" => 198,
        "DCOUNTA" => 199,
        "DELETE.BAR" => 200,
        "UNREGISTER" => 201,
        "USDOLLAR" => 204,
        "FINDB" => 205,
        "SEARCHB" => 206,
        "REPLACEB" => 207,
        "LEFTB" => 208,
        "RIGHTB" => 209,
        "MIDB" => 210,
        "LENB" => 211,
        "ROUNDUP" => 212,
        "ROUNDDOWN" => 213,
        "ASC" => 214,
        "DBCS" => 215,
        "RANK" => 216,
        "ADDRESS" => 219,
        "DAYS360" => 220,
        "TODAY" => 221,
        "VDB" => 222,
        "ELSE" => 223,
        "ELSE.IF" => 224,
        "END.IF" => 225,
        "FOR.CELL" => 226,
        "MEDIAN" => 227,
        "SUMPRODUCT" => 228,
        "SINH" => 229,
        "COSH" => 230,
        "TANH" => 231,
        "ASINH" => 232,
        "ACOSH" => 233,
        "ATANH" => 234,
        "DGET" => 235,
        "CREATE.OBJECT" => 236,
        "VOLATILE" => 237,
        "LAST.ERROR" => 238,
        "CUSTOM.UNDO" => 239,
        "CUSTOM.REPEAT" => 240,
        "FORMULA.CONVERT" => 241,
        "GET.LINK.INFO" => 242,
        "TEXT.BOX" => 243,
        "INFO" => 244,
        "GROUP" => 245,
        "GET.OBJECT" => 246,
        "DB" => 247,
        "PAUSE" => 248,
        "RESUME" => 251,
        "FREQUENCY" => 252,
        "ADD.TOOLBAR" => 253,
        "DELETE.TOOLBAR" => 254,
        "RESET.TOOLBAR" => 256,
        "EVALUATE" => 257,
        "GET.TOOLBAR" => 258,
        "GET.TOOL" => 259,
        "SPELLING.CHECK" => 260,
        "ERROR.TYPE" => 261,
        "APP.TITLE" => 262,
        "WINDOW.TITLE" => 263,
        "SAVE.TOOLBAR" => 264,
        "ENABLE.TOOL" => 265,
        "PRESS.TOOL" => 266,
        "REGISTER.ID" => 267,
        "GET.WORKBOOK" => 268,
        "AVEDEV" => 269,
        "BETADIST" => 270,
        "GAMMALN" => 271,
        "BETAINV" => 272,
        "BINOMDIST" => 273,
        "CHIDIST" => 274,
        "CHIINV" => 275,
        "COMBIN" => 276,
        "CONFIDENCE" => 277,
        "CRITBINOM" => 278,
        "EVEN" => 279,
        "EXPONDIST" => 280,
        "FDIST" => 281,
        "FINV" => 282,
        "FISHER" => 283,
        "FISHERINV" => 284,
        "FLOOR" => 285,
        "GAMMADIST" => 286,
        "GAMMAINV" => 287,
        "CEILING" => 288,
        "HYPGEOMDIST" => 289,
        "LOGNORMDIST" => 290,
        "LOGINV" => 291,
        "NEGBINOMDIST" => 292,
        "NORMDIST" => 293,
        "NORMSDIST" => 294,
        "NORMINV" => 295,
        "NORMSINV" => 296,
        "STANDARDIZE" => 297,
        "ODD" => 298,
        "PERMUT" => 299,
        "POISSON" => 300,
        "TDIST" => 301,
        "WEIBULL" => 302,
        "SUMXMY2" => 303,
        "SUMX2MY2" => 304,
        "SUMX2PY2" => 305,
        "CHITEST" => 306,
        "CORREL" => 307,
        "COVAR" => 308,
        "FORECAST" => 309,
        "FTEST" => 310,
        "INTERCEPT" => 311,
        "PEARSON" => 312,
        "RSQ" => 313,
        "STEYX" => 314,
        "SLOPE" => 315,
        "TTEST" => 316,
        "PROB" => 317,
        "DEVSQ" => 318,
        "GEOMEAN" => 319,
        "HARMEAN" => 320,
        "SUMSQ" => 321,
        "KURT" => 322,
        "SKEW" => 323,
        "ZTEST" => 324,
        "LARGE" => 325,
        "SMALL" => 326,
        "QUARTILE" => 327,
        "PERCENTILE" => 328,
        "PERCENTRANK" => 329,
        "MODE" => 330,
        "TRIMMEAN" => 331,
        "TINV" => 332,
        "MOVIE.COMMAND" => 334,
        "GET.MOVIE" => 335,
        "CONCATENATE" => 336,
        "POWER" => 337,
        "PIVOT.ADD.DATA" => 338,
        "GET.PIVOT.TABLE" => 339,
        "GET.PIVOT.FIELD" => 340,
        "GET.PIVOT.ITEM" => 341,
        "RADIANS" => 342,
        "DEGREES" => 343,
        "SUBTOTAL" => 344,
        "SUMIF" => 345,
        "COUNTIF" => 346,
        "COUNTBLANK" => 347,
        "SCENARIO.GET" => 348,
        "OPTIONS.LISTS.GET" => 349,
        "ISPMT" => 350,
        "DATEDIF" => 351,
        "DATESTRING" => 352,
        "NUMBERSTRING" => 353,
        "ROMAN" => 354,
        "OPEN.DIALOG" => 355,
        "SAVE.DIALOG" => 356,
        "VIEW.GET" => 357,
        "GETPIVOTDATA" => 358,
        "HYPERLINK" => 359,
        "PHONETIC" => 360,
        "AVERAGEA" => 361,
        "MAXA" => 362,
        "MINA" => 363,
        "STDEVPA" => 364,
        "VARPA" => 365,
        "STDEVA" => 366,
        "VARA" => 367,
        "BAHTTEXT" => 368,
        _ => return None,
    })
}

/// Classic Analysis ToolPak function set — calamine FTAB indices
/// 384..=476. In BIFF8 Excel encodes calls to these as add-in
/// EXTERNNAME references (PtgNameX), never as iftab builtins;
/// Excel 2003 has no iftab for them and POI follows Excel.
fn is_atp_function(name: &str) -> bool {
    matches!(name,
        "HEX2BIN" | "HEX2DEC" | "HEX2OCT" | "DEC2BIN" | "DEC2HEX" | "DEC2OCT" |
        "OCT2BIN" | "OCT2HEX" | "OCT2DEC" | "BIN2DEC" | "BIN2OCT" | "BIN2HEX" |
        "IMSUB" | "IMDIV" | "IMPOWER" | "IMABS" | "IMSQRT" | "IMLN" |
        "IMLOG2" | "IMLOG10" | "IMSIN" | "IMCOS" | "IMEXP" | "IMARGUMENT" |
        "IMCONJUGATE" | "IMAGINARY" | "IMREAL" | "COMPLEX" | "IMSUM" | "IMPRODUCT" |
        "SERIESSUM" | "FACTDOUBLE" | "SQRTPI" | "QUOTIENT" | "DELTA" | "GESTEP" |
        "ISEVEN" | "ISODD" | "MROUND" | "ERF" | "ERFC" | "BESSELJ" |
        "BESSELK" | "BESSELY" | "BESSELI" | "XIRR" | "XNPV" | "PRICEMAT" |
        "YIELDMAT" | "INTRATE" | "RECEIVED" | "DISC" | "PRICEDISC" | "YIELDDISC" |
        "TBILLEQ" | "TBILLPRICE" | "TBILLYIELD" | "PRICE" | "YIELD" | "DOLLARDE" |
        "DOLLARFR" | "NOMINAL" | "EFFECT" | "CUMPRINC" | "CUMIPMT" | "EDATE" |
        "EOMONTH" | "YEARFRAC" | "COUPDAYBS" | "COUPDAYS" | "COUPDAYSNC" | "COUPNCD" |
        "COUPNUM" | "COUPPCD" | "DURATION" | "MDURATION" | "ODDLPRICE" | "ODDLYIELD" |
        "ODDFPRICE" | "ODDFYIELD" | "RANDBETWEEN" | "WEEKNUM" | "AMORDEGRC" | "AMORLINC" |
        "CONVERT" | "ACCRINT" | "ACCRINTM" | "WORKDAY" | "NETWORKDAYS" | "GCD" |
        "MULTINOMIAL" | "LCM" | "FVSCHEDULE")
}

/// Functions added to Excel after 2003 with no BIFF8 iftab. Excel
/// encodes calls to these in .xls as future functions: a hidden NAME
/// record `_xlfn.<name>` referenced by PtgName, wrapped in
/// PtgFuncVar(iftab=255). Excel maps `_xlfn.*` back to the modern
/// builtin on open and displays the call normally. Returns the full
/// name to emit in the NAME record (the worksheet-scope dynamic-array
/// functions take an extra `_xlws.` segment, per Excel's own output).
fn xlfn_route_for_name(name: &str) -> Option<String> {
    match name {
        // Dynamic-array worksheet functions: _xlfn._xlws. prefix.
        "SORT" | "FILTER" => return Some(format!("_xlfn._xlws.{name}")),
        // Excel 2007
        "IFERROR" | "COUNTIFS" | "SUMIFS" | "AVERAGEIF" | "AVERAGEIFS"
        // Excel 2010 dotted statistics + misc
        | "STDEV.P" | "STDEV.S" | "VAR.P" | "VAR.S" | "MODE.SNGL" | "MODE.MULT"
        | "PERCENTILE.INC" | "PERCENTILE.EXC" | "QUARTILE.INC" | "QUARTILE.EXC"
        | "RANK.EQ" | "RANK.AVG" | "NORM.DIST" | "NORM.INV" | "NORM.S.DIST"
        | "NORM.S.INV" | "LOGNORM.DIST" | "LOGNORM.INV" | "T.DIST" | "T.DIST.2T"
        | "T.DIST.RT" | "T.INV" | "T.INV.2T" | "T.TEST" | "F.DIST" | "F.DIST.RT"
        | "F.INV" | "F.INV.RT" | "F.TEST" | "CHISQ.DIST" | "CHISQ.DIST.RT"
        | "CHISQ.INV" | "CHISQ.INV.RT" | "CHISQ.TEST" | "BINOM.DIST" | "BINOM.INV"
        | "POISSON.DIST" | "EXPON.DIST" | "GAMMA.DIST" | "GAMMA.INV"
        | "GAMMALN.PRECISE" | "BETA.DIST" | "BETA.INV" | "HYPGEOM.DIST"
        | "NEGBINOM.DIST" | "WEIBULL.DIST" | "Z.TEST" | "CONFIDENCE.NORM"
        | "CONFIDENCE.T" | "COVARIANCE.P" | "COVARIANCE.S" | "AGGREGATE"
        | "NETWORKDAYS.INTL" | "WORKDAY.INTL" | "CEILING.PRECISE" | "FLOOR.PRECISE"
        | "ISO.CEILING" | "ERF.PRECISE" | "ERFC.PRECISE"
        // Excel 2013
        | "IFNA" | "SHEET" | "SHEETS" | "FORMULATEXT" | "ISFORMULA" | "ISOWEEKNUM"
        | "DAYS" | "BITAND" | "BITOR" | "BITXOR" | "BITLSHIFT" | "BITRSHIFT"
        | "ARABIC" | "BASE" | "DECIMAL" | "COMBINA" | "PERMUTATIONA" | "GAMMA"
        | "GAUSS" | "PHI" | "SKEW.P" | "FLOOR.MATH" | "CEILING.MATH" | "ACOT"
        | "ACOTH" | "COT" | "COTH" | "CSC" | "CSCH" | "SEC" | "SECH" | "RRI"
        | "PDURATION" | "UNICHAR" | "UNICODE" | "NUMBERVALUE" | "BINOM.DIST.RANGE"
        | "FORECAST.LINEAR"
        // Excel 2016+
        | "IFS" | "SWITCH" | "XOR" | "CONCAT" | "TEXTJOIN" | "MAXIFS" | "MINIFS"
        // Excel 2019 / 365
        | "XLOOKUP" | "XMATCH" | "SEQUENCE" | "RANDARRAY" | "UNIQUE" | "SORTBY"
        | "LET" | "LAMBDA" | "MAP" | "REDUCE" | "SCAN" | "BYROW" | "BYCOL"
        | "MAKEARRAY" | "ISOMITTED" | "TEXTBEFORE" | "TEXTAFTER" | "TEXTSPLIT"
        | "TOROW" | "TOCOL" | "VSTACK" | "HSTACK" | "TAKE" | "DROP"
        | "CHOOSEROWS" | "CHOOSECOLS" | "WRAPROWS" | "WRAPCOLS" | "EXPAND"
        | "VALUETOTEXT" | "ARRAYTOTEXT" => Some(format!("_xlfn.{name}")),
        _ => None,
    }
}

/// Walk a preserved rgce and decide whether it can be replayed
/// verbatim into a rebuilt workbook stream. Any ptg that indexes a
/// workbook-global table (PtgName → NAME ilbl, PtgNameX → EXTERNSHEET
/// ixti + EXTERNNAME, PtgRef3d / PtgArea3d / PtgRefErr3d /
/// PtgAreaErr3d → EXTERNSHEET ixti) makes it non-replayable — the
/// writer renumbers those tables. Unknown ptgs and ptgs with trailing
/// extra data (PtgArray, PtgMemArea) also return false; the caller
/// then falls back to the cached-value placeholder, which was the
/// behavior for all preserved cells before verbatim replay existed.
fn rgce_is_self_contained(rgce: &[u8]) -> bool {
    let mut i = 0usize;
    while i < rgce.len() {
        let ptg = rgce[i];
        let size = match ptg {
            // Table-referencing ptgs — not replayable.
            0x23 | 0x43 | 0x63 => return false, // PtgName
            0x39 | 0x59 | 0x79 => return false, // PtgNameX
            0x3A | 0x5A | 0x7A => return false, // PtgRef3d
            0x3B | 0x5B | 0x7B => return false, // PtgArea3d
            0x3C | 0x5C | 0x7C => return false, // PtgRefErr3d
            0x3D | 0x5D | 0x7D => return false, // PtgAreaErr3d
            // Binary / unary operators, paren, missing arg.
            0x03..=0x16 => 1,
            // PtgStr: cch + grbit + chars.
            0x17 => {
                if i + 3 > rgce.len() { return false; }
                let cch = rgce[i + 1] as usize;
                let wide = rgce[i + 2] & 1 == 1;
                3 + cch * if wide { 2 } else { 1 }
            }
            // PtgAttr: sub-type in byte 1; AttrChoose carries a jump
            // table of (nc+1) u16 offsets.
            0x19 => {
                if i + 4 > rgce.len() { return false; }
                if rgce[i + 1] & 0x04 != 0 {
                    let nc = u16::from_le_bytes([rgce[i + 2], rgce[i + 3]]) as usize;
                    4 + 2 * (nc + 1)
                } else {
                    4
                }
            }
            0x1C | 0x1D => 2,          // PtgErr, PtgBool
            0x1E => 3,                 // PtgInt
            0x1F => 9,                 // PtgNum
            0x21 | 0x41 | 0x61 => 3,   // PtgFunc
            0x22 | 0x42 | 0x62 => 4,   // PtgFuncVar
            0x24 | 0x44 | 0x64 => 5,   // PtgRef
            0x25 | 0x45 | 0x65 => 9,   // PtgArea
            0x29 | 0x49 | 0x69 => 3,   // PtgMemFunc
            0x2A | 0x4A | 0x6A => 5,   // PtgRefErr
            0x2B | 0x4B | 0x6B => 9,   // PtgAreaErr
            0x2C | 0x4C | 0x6C => 5,   // PtgRefN
            0x2D | 0x4D | 0x6D => 9,   // PtgAreaN
            // PtgExp/PtgTbl (shared-formula anchors), PtgArray and
            // PtgMemArea (trailing extra data), anything unknown.
            _ => return false,
        };
        i += size;
    }
    i == rgce.len()
}

/// Size of the ptg at offset `i`, or `None` for ptgs preserved-rgce
/// replay doesn't handle (PtgExp / PtgTbl / PtgArray / PtgMemArea
/// carry trailing data or anchor semantics; anything unknown).
fn ptg_size_at(rgce: &[u8], i: usize) -> Option<usize> {
    Some(match rgce[i] {
        0x03..=0x16 => 1,
        0x17 => {
            if i + 3 > rgce.len() {
                return None;
            }
            let cch = rgce[i + 1] as usize;
            let wide = rgce[i + 2] & 1 == 1;
            3 + cch * if wide { 2 } else { 1 }
        }
        0x19 => {
            if i + 4 > rgce.len() {
                return None;
            }
            if rgce[i + 1] & 0x04 != 0 {
                let nc = u16::from_le_bytes([rgce[i + 2], rgce[i + 3]]) as usize;
                4 + 2 * (nc + 1)
            } else {
                4
            }
        }
        0x1C | 0x1D => 2,        // PtgErr, PtgBool
        0x1E => 3,               // PtgInt
        0x1F => 9,               // PtgNum
        0x21 | 0x41 | 0x61 => 3, // PtgFunc
        0x22 | 0x42 | 0x62 => 4, // PtgFuncVar
        0x23 | 0x43 | 0x63 => 5, // PtgName
        0x24 | 0x44 | 0x64 => 5, // PtgRef
        0x25 | 0x45 | 0x65 => 9, // PtgArea
        0x29 | 0x49 | 0x69 => 3, // PtgMemFunc
        0x2A | 0x4A | 0x6A => 5, // PtgRefErr
        0x2B | 0x4B | 0x6B => 9, // PtgAreaErr
        0x2C | 0x4C | 0x6C => 5, // PtgRefN
        0x2D | 0x4D | 0x6D => 9, // PtgAreaN
        0x39 | 0x59 | 0x79 => 7,  // PtgNameX
        0x3A | 0x5A | 0x7A => 7,  // PtgRef3d
        0x3B | 0x5B | 0x7B => 11, // PtgArea3d
        0x3C | 0x5C | 0x7C => 7,  // PtgRefErr3d
        0x3D | 0x5D | 0x7D => 11, // PtgAreaErr3d
        _ => return None,
    })
}

/// Source-file table context for preserved-rgce replay: the NAME list
/// in source Lbl order (PtgName inames are 1-based into it), the
/// per-source-ixti sheet names, and the current model's sheet-name →
/// index map for resolving those back.
struct SourceRgceCtx<'a> {
    names: &'a [String],
    xti_sheets: &'a [Option<(String, String)>],
    sheet_by_name: &'a HashMap<String, u32>,
}

impl<'a> SourceRgceCtx<'a> {
    /// Resolve a source ixti to the model's (first, last) sheet
    /// indices; None for add-in / external / deleted entries or
    /// sheets that no longer exist.
    fn resolve_ixti(&self, src_ixti: usize) -> Option<(u16, u16)> {
        let (first, last) = self.xti_sheets.get(src_ixti)?.as_ref()?;
        let fi = *self.sheet_by_name.get(&first.to_lowercase())? as u16;
        let li = *self.sheet_by_name.get(&last.to_lowercase())? as u16;
        Some((fi, li))
    }
}

/// Rewrite a preserved (source-file) rgce so its table references
/// point into THIS writer's rebuilt tables: PtgName inames remap by
/// name string, Ptg*3d / Ptg*Err3d ixtis remap by sheet name through
/// the writer's xti table (whose entries must have been registered by
/// `register_preserved_xti` before EXTERNSHEET emission). Everything
/// else passes verbatim — including the source's operand classes,
/// which are Excel-authored and correct by construction. Returns
/// `None` (caller falls back) on PtgNameX (the add-in list isn't
/// name-remappable), unresolvable names/sheets, or unknown ptgs.
fn remap_source_rgce(
    rgce: &[u8],
    ctx: &SourceRgceCtx,
    defined_names: &DefinedNameTable,
    extern_names: &ExternNameTable,
    xti: &XtiTable,
) -> Option<Vec<u8>> {
    let mut out = rgce.to_vec();
    let mut i = 0usize;
    while i < out.len() {
        match out[i] {
            0x23 | 0x43 | 0x63 => {
                if i + 5 > out.len() {
                    return None;
                }
                let iname =
                    u32::from_le_bytes([out[i + 1], out[i + 2], out[i + 3], out[i + 4]]) as usize;
                let name = ctx.names.get(iname.checked_sub(1)?)?;
                let canon = name.strip_prefix("_xlfn.").unwrap_or(name);
                let ilbl = extern_names
                    .name_fn_ilbl(canon)
                    .or_else(|| defined_names.ilbl_for(name, None))?;
                out[i + 1..i + 3].copy_from_slice(&ilbl.to_le_bytes());
                out[i + 3] = 0;
                out[i + 4] = 0;
            }
            0x3A | 0x5A | 0x7A | 0x3B | 0x5B | 0x7B | 0x3C | 0x5C | 0x7C | 0x3D | 0x5D
            | 0x7D => {
                if i + 3 > out.len() {
                    return None;
                }
                let src_ixti = u16::from_le_bytes([out[i + 1], out[i + 2]]) as usize;
                let (fi, li) = ctx.resolve_ixti(src_ixti)?;
                let new_ixti = xti.lookup_pair(fi, li)?;
                out[i + 1..i + 3].copy_from_slice(&new_ixti.to_le_bytes());
            }
            0x39 | 0x59 | 0x79 => return None, // PtgNameX
            _ => {}
        }
        i += ptg_size_at(&out, i)?;
    }
    (i == out.len()).then_some(out)
}

/// Pre-pass before EXTERNSHEET emission: register in the writer's
/// xti table every sheet span referenced by a preserved rgce (array
/// formulas, unimported NAME records, #ERROR!-cell replays), so
/// `remap_source_rgce` finds them at emission time.
fn register_preserved_xti(
    preserved: Option<&crate::xls_preserve::PreservedXlsData>,
    preserved_rgce: Option<&HashMap<(u32, i32, i32), Vec<u8>>>,
    ctx: &SourceRgceCtx,
    xti: &mut XtiTable,
) {
    let mut reg = |rgce: &[u8]| {
        let mut i = 0usize;
        while i < rgce.len() {
            if matches!(
                rgce[i],
                0x3A | 0x5A | 0x7A | 0x3B | 0x5B | 0x7B | 0x3C | 0x5C | 0x7C | 0x3D | 0x5D | 0x7D
            ) && i + 3 <= rgce.len()
            {
                let src_ixti = u16::from_le_bytes([rgce[i + 1], rgce[i + 2]]) as usize;
                if let Some((fi, li)) = ctx.resolve_ixti(src_ixti) {
                    xti.lookup_or_add_pair(fi, li);
                }
            }
            match ptg_size_at(rgce, i) {
                Some(sz) => i += sz,
                None => break,
            }
        }
    };
    if let Some(p) = preserved {
        for pa in &p.array_formulas {
            reg(&pa.rgce);
        }
        for pn in &p.preserved_names {
            reg(&pn.rgce);
        }
    }
    if let Some(m) = preserved_rgce {
        for r in m.values() {
            reg(r);
        }
    }
}

/// Per-sheet plan for folding load-time-replicated array formulas
/// back into their FORMULA(PtgExp) + ARRAY record structure on save.
/// Built from `PreservedXlsData.array_formulas` by
/// `build_array_plans`; only ranges whose cells are byte-for-byte
/// unedited since load make it in — anything touched falls back to
/// the per-cell expansion (which is also the pre-existing behaviour
/// for every array range when no plan exists).
struct SheetArrayPlan {
    /// (row, col) 1-based → the range anchor as 0-based (rw, col),
    /// ready for the PtgExp payload. Includes the anchor cell itself.
    members: HashMap<(i32, i32), (u16, u16)>,
    /// anchor (row, col) 1-based → prebuilt ARRAY record body,
    /// written immediately after the anchor's FORMULA record.
    anchors: HashMap<(i32, i32), Vec<u8>>,
}

fn build_array_plans(
    model: &ironcalc::base::Model,
    preserved: Option<&crate::xls_preserve::PreservedXlsData>,
    ctx: Option<&SourceRgceCtx>,
    defined_names: &DefinedNameTable,
    extern_names: &ExternNameTable,
    xti: &XtiTable,
) -> HashMap<u32, SheetArrayPlan> {
    let mut plans: HashMap<u32, SheetArrayPlan> = HashMap::new();
    let Some(p) = preserved else { return plans };
    for pa in &p.array_formulas {
        // BIFF limits: rows u16, cols u8 in the ARRAY record's range.
        if pa.r1 < 1 || pa.c1 < 1 || pa.r2 > 0x1_0000 || pa.c2 > 256 {
            continue;
        }
        let Some(sheet_idx) = model
            .workbook
            .worksheets
            .iter()
            .position(|w| w.name.eq_ignore_ascii_case(&pa.sheet_name))
        else {
            continue;
        };
        let sheet_idx = sheet_idx as u32;
        // Unedited check: the model's canonical formula text for every
        // cell in the range must equal the snapshot taken right after
        // load-time replication.
        let mut snap = pa.cell_formulas.iter();
        let unedited = (pa.r1..=pa.r2).all(|r| {
            (pa.c1..=pa.c2).all(|c| {
                let actual = model.get_cell_formula(sheet_idx, r, c).ok().flatten();
                snap.next().map(|expected| *expected == actual).unwrap_or(false)
            })
        });
        if !unedited {
            continue;
        }
        let Some(rgce) = ctx
            .and_then(|c| remap_source_rgce(&pa.rgce, c, defined_names, extern_names, xti))
        else {
            continue;
        };
        // ARRAY record body ([MS-XLS] 2.4.4): rwFirst(2) rwLast(2)
        // colFirst(1) colLast(1) flags(2) chn(4) cce(2) rgce.
        let mut body = Vec::with_capacity(14 + rgce.len());
        body.put_u16((pa.r1 - 1) as u16);
        body.put_u16((pa.r2 - 1) as u16);
        body.put_u8((pa.c1 - 1) as u8);
        body.put_u8((pa.c2 - 1) as u8);
        body.put_u16(0); // flags
        body.put_u32(0); // chn — reserved, readers rebuild it
        body.put_u16(rgce.len() as u16);
        body.extend_from_slice(&rgce);
        let plan = plans.entry(sheet_idx).or_insert_with(|| SheetArrayPlan {
            members: HashMap::new(),
            anchors: HashMap::new(),
        });
        let anchor0 = ((pa.r1 - 1) as u16, (pa.c1 - 1) as u16);
        for r in pa.r1..=pa.r2 {
            for c in pa.c1..=pa.c2 {
                plan.members.insert((r, c), anchor0);
            }
        }
        plan.anchors.insert((pa.r1, pa.c1), body);
    }
    plans
}

/// Cached-value classification for an array-member cell. `None` for
/// non-formula cells — can't happen for a validated range (the
/// unedited check requires formula text in every cell), but the
/// caller falls back to normal emission rather than panicking.
fn array_member_cached(
    cell: &ironcalc::base::types::Cell,
) -> Option<(FormulaCachedValue, Option<&str>)> {
    use ironcalc::base::types::Cell;
    match cell {
        Cell::CellFormulaNumber { v, .. } => Some((FormulaCachedValue::Number(*v), None)),
        Cell::CellFormulaBoolean { v, .. } => Some((FormulaCachedValue::Bool(*v), None)),
        Cell::CellFormulaError { ei, .. } => {
            Some((FormulaCachedValue::Error(ironcalc_error_to_biff(ei)), None))
        }
        Cell::CellFormulaString { v, .. } if v.is_empty() => {
            Some((FormulaCachedValue::Blank, None))
        }
        Cell::CellFormulaString { v, .. } => {
            Some((FormulaCachedValue::StringPending, Some(v.as_str())))
        }
        Cell::CellFormula { .. } => Some((FormulaCachedValue::Number(0.0), None)),
        _ => None,
    }
}

/// Try to encode the parsed formula at index `f`; on any failure (no
/// such index, encoder hit an unsupported node), fall back to the
/// placeholder rgce that just pushes the cached value. Anchor is the
/// 1-based (row, col) of the cell containing the formula — used by
/// the encoder to resolve IronCalc's relative-ref offsets back to
/// absolute targets.
fn encode_or_placeholder(
    parsed_formulas: &[ironcalc::base::expressions::parser::Node],
    f: i32,
    sheet_idx: u32,
    anchor: CellAnchor,
    xti: &XtiTable,
    defined_names: &DefinedNameTable,
    extern_names: &ExternNameTable,
    cached: &FormulaCachedValue,
    cached_string: Option<&str>,
) -> Vec<u8> {
    if let Some(node) = parsed_formulas.get(f as usize) {
        if let Some(rgce) = encode_formula_rgce_with_names(node, sheet_idx, anchor, xti, defined_names, extern_names, CLS_VAL) {
            return rgce;
        }
    }
    placeholder_rgce(cached, cached_string)
}

/// Placeholder rgce: emit a single ptg that pushes a constant value.
/// Real formulas come in phase 4. Excel happily evaluates the result
/// of any rgce that leaves exactly one value on the stack.
fn placeholder_rgce(cached: &FormulaCachedValue, cached_string: Option<&str>) -> Vec<u8> {
    match cached {
        FormulaCachedValue::Number(n) => {
            let mut rgce = Vec::with_capacity(9);
            rgce.put_u8(0x1F); // PtgNum
            rgce.put_f64(*n);
            rgce
        }
        FormulaCachedValue::Bool(b) => {
            let mut rgce = Vec::with_capacity(2);
            rgce.put_u8(0x1D); // PtgBool
            rgce.put_u8(if *b { 1 } else { 0 });
            rgce
        }
        FormulaCachedValue::Error(e) => {
            let mut rgce = Vec::with_capacity(2);
            rgce.put_u8(0x1C); // PtgErr
            rgce.put_u8(*e);
            rgce
        }
        FormulaCachedValue::StringPending => {
            // PtgStr (0x17) — opcode + cch (u8) + flag (u8) + chars.
            // Truncate to 255 codepoints; longer cached strings are rare
            // and full handling lands in phase 4.
            let s = cached_string.unwrap_or("");
            let chars: Vec<u16> = s.encode_utf16().collect();
            let high_byte = chars.iter().any(|&c| c > 0xFF);
            let cch = chars.len().min(255) as u8;
            let mut rgce = Vec::with_capacity(3 + chars.len() * 2);
            rgce.put_u8(0x17);
            rgce.put_u8(cch);
            rgce.put_u8(if high_byte { 0x01 } else { 0x00 });
            if high_byte {
                for c in chars.iter().take(cch as usize) { rgce.put_u16(*c); }
            } else {
                for c in chars.iter().take(cch as usize) { rgce.put_u8(*c as u8); }
            }
            rgce
        }
        FormulaCachedValue::Blank => {
            // PtgStr with cch=0 → empty string literal `""`. Calamine
            // emits formula text `""` for this; xls_load parses as
            // `=""` and IronCalc stores the cell as a
            // CellFormulaString { v: "" }. Without this (when we'd
            // emitted PtgMissArg 0x16), calamine produced empty
            // formula text → xls_load skipped the cell → it became
            // EmptyCell. That broke downstream COUNTA / COUNTIF
            // because Excel treats `""`-formula cells differently
            // from truly-empty ones.
            vec![0x17, 0x00, 0x00] // PtgStr opcode + cch=0 + flag=0
        }
    }
}

// ---------------------------------------------------------------------------
// Pre-pass: walk all cells across all sheets to collect (a) the count
// of LABELSST references per sheet (cstTotal in the SST header is a
// workbook-wide sum) and (b) the per-sheet used range (DIMENSIONS).
// IronCalc's `Worksheet.dimension` field exists but is sometimes "" or
// stale; we recompute it ourselves to be safe.
// ---------------------------------------------------------------------------

#[derive(Default, Debug)]
struct SheetExtents {
    /// Inclusive min/max in 1-based coords. `None` ⇒ sheet is empty.
    bounds: Option<(i32, i32, i32, i32)>, // (min_row, max_row, min_col, max_col)
    labelsst_count: u32,
}

fn compute_sheet_extents(ws: &ironcalc::base::types::Worksheet) -> SheetExtents {
    use ironcalc::base::types::Cell;
    let mut out = SheetExtents::default();
    for (row, cols) in ws.sheet_data.iter() {
        for (col, cell) in cols.iter() {
            // Update bounds.
            out.bounds = Some(match out.bounds {
                None => (*row, *row, *col, *col),
                Some((rmin, rmax, cmin, cmax)) => (
                    rmin.min(*row),
                    rmax.max(*row),
                    cmin.min(*col),
                    cmax.max(*col),
                ),
            });
            // SharedString cells emit LABELSST directly. CellFormulaString
            // cells with non-empty values also emit LABELSST (we route
            // them through the SST instead of FORMULA+STRING — see
            // build_sst_table for rationale).
            match cell {
                Cell::SharedString { .. } => out.labelsst_count += 1,
                Cell::CellFormulaString { v, .. } if !v.is_empty() => {
                    out.labelsst_count += 1;
                }
                _ => {}
            }
        }
    }
    out
}

/// Emit cell records for one sheet. Cells are walked in row-major
/// (row asc, col asc) order so the output is deterministic and matches
/// what BIFF readers expect for INDEX / DBCELL acceleration.
///
/// `style_tables.cell_xf_index` resolves IronCalc's per-cell `s` field
/// to the BIFF XF index for the cell record's `ixfe`.
fn emit_sheet_cells(
    w: &mut BiffWriter,
    ws: &ironcalc::base::types::Worksheet,
    sheet_idx: u32,
    shared_strings: &[String],
    sst_index: &std::collections::HashMap<String, u32>,
    parsed_formulas: &[ironcalc::base::expressions::parser::Node],
    xti: &XtiTable,
    defined_names: &DefinedNameTable,
    extern_names: &ExternNameTable,
    style_tables: &StyleTables,
    preserved_rgce: Option<&HashMap<(u32, i32, i32), Vec<u8>>>,
    arrays: Option<&SheetArrayPlan>,
    remap_ctx: Option<&SourceRgceCtx>,
) {
    use ironcalc::base::types::Cell;
    // The cell table is organized in row blocks of up to 32 rows:
    // the block's ROW records, then the block's cell records
    // ([MS-XLS] 2.1.7.20.2). Excel expects a ROW record for EVERY
    // populated row — emitting only the custom-height rows (the old
    // behavior) made Excel show "File error: data may have been
    // lost" on open (2026-08-14 report follow-up). Height-only Row
    // entries with no cells are included too.
    let mut row_union: std::collections::BTreeSet<i32> = ws.sheet_data.keys().copied().collect();
    let row_props: HashMap<i32, &ironcalc::base::types::Row> =
        ws.rows.iter().map(|r| (r.r, r)).collect();
    row_union.extend(row_props.keys().copied());
    let all_rows: Vec<i32> = row_union.into_iter().collect();
    for block in all_rows.chunks(32) {
        for row in block {
            // Per-row column bounds (colMic inclusive, colMac
            // exclusive, 0-based) from the row's own cells.
            let (col_mic, col_mac) = match ws.sheet_data.get(row) {
                Some(cols) if !cols.is_empty() => {
                    let min = cols.keys().min().copied().unwrap_or(1);
                    let max = cols.keys().max().copied().unwrap_or(1);
                    (((min - 1).max(0) as u16), (max.max(0) as u16))
                }
                _ => (0u16, 0u16),
            };
            match row_props.get(row) {
                Some(r) => w.write_record(R_ROW, &build_row_record(r, col_mic, col_mac)),
                None => {
                    // Default-height row: custom_height=false so Excel
                    // recomputes from the font; 12.75pt matches the
                    // DEFAULTROWHEIGHT record (255 twips).
                    let default_row = ironcalc::base::types::Row {
                        r: *row,
                        height: 12.75,
                        custom_format: false,
                        custom_height: false,
                        s: 0,
                        hidden: false,
                    };
                    w.write_record(R_ROW, &build_row_record(&default_row, col_mic, col_mac));
                }
            }
        }
        for row in block {
            let Some(cols_map) = ws.sheet_data.get(row) else { continue };
        let mut cols: Vec<&i32> = cols_map.keys().collect();
        cols.sort();
        for col in cols {
            let cell = &cols_map[col];
            // BIFF rows / cols are 0-based in records, but IronCalc
            // stores them 1-based. Convert here. Keep the 1-based
            // values around for the formula-encoder anchor: relative
            // refs in the parsed tree are offsets from this anchor.
            let rw = (*row - 1).max(0) as u32;
            let cl = (*col - 1).max(0) as u32;
            let anchor = CellAnchor { row: *row, column: *col };
            // Resolve cell.s (IronCalc cell_xf index) → BIFF XF index.
            // Out-of-range falls back to the default cell XF.
            let s = match cell {
                ironcalc::base::types::Cell::EmptyCell { s } => *s,
                ironcalc::base::types::Cell::BooleanCell { s, .. } => *s,
                ironcalc::base::types::Cell::NumberCell { s, .. } => *s,
                ironcalc::base::types::Cell::ErrorCell { s, .. } => *s,
                ironcalc::base::types::Cell::SharedString { s, .. } => *s,
                ironcalc::base::types::Cell::CellFormula { s, .. } => *s,
                ironcalc::base::types::Cell::CellFormulaBoolean { s, .. } => *s,
                ironcalc::base::types::Cell::CellFormulaNumber { s, .. } => *s,
                ironcalc::base::types::Cell::CellFormulaString { s, .. } => *s,
                ironcalc::base::types::Cell::CellFormulaError { s, .. } => *s,
            };
            let ixfe = style_tables
                .cell_xf_index
                .get(s as usize)
                .copied()
                .unwrap_or(style_tables.default_cell_xf_idx);
            // Array-entered cell (validated-unedited range from the
            // source file): emit FORMULA with a PtgExp rgce pointing
            // at the range anchor; the anchor's FORMULA is followed
            // by the ARRAY record carrying the real (remapped) rgce.
            // Record order matches Excel: FORMULA → ARRAY → STRING.
            if let Some((anchor0, (cached, pending_str))) = arrays.and_then(|plan| {
                let anchor0 = plan.members.get(&(*row, *col))?;
                Some((anchor0, array_member_cached(cell)?))
            }) {
                let mut rgce = Vec::with_capacity(5);
                rgce.put_u8(0x01); // PtgExp
                rgce.put_u16(anchor0.0);
                rgce.put_u16(anchor0.1);
                w.write_record(
                    R_FORMULA,
                    &build_formula_with_grbit(rw, cl, ixfe, &cached, &rgce, 0x0020),
                );
                if let Some(body) = arrays.and_then(|plan| plan.anchors.get(&(*row, *col))) {
                    w.write_record(R_ARRAY, body);
                }
                if let Some(v) = pending_str {
                    w.write_record(R_STRING, &build_string(v));
                }
                continue;
            }
            match cell {
                Cell::EmptyCell { s } => {
                    // Only emit BLANK if the cell has non-default styling
                    // (style index != 0). Empty cells with no formatting
                    // contribute nothing and just bloat the file.
                    if *s != 0 {
                        w.write_record(R_BLANK, &build_blank(rw, cl, ixfe));
                    }
                }
                Cell::BooleanCell { v, .. } => {
                    w.write_record(R_BOOLERR, &build_boolerr(rw, cl, ixfe, if *v { 1 } else { 0 }, false));
                }
                Cell::NumberCell { v, .. } => {
                    w.write_record(R_NUMBER, &build_number(rw, cl, ixfe, *v));
                }
                Cell::ErrorCell { ei, .. } => {
                    let code = ironcalc_error_to_biff(ei);
                    w.write_record(R_BOOLERR, &build_boolerr(rw, cl, ixfe, code, true));
                }
                Cell::SharedString { si, .. } => {
                    // si is the 0-based SST index (matches our SST emit).
                    if (*si as usize) < shared_strings.len() {
                        w.write_record(R_LABELSST, &build_labelsst(rw, cl, ixfe, *si as u32));
                    }
                }
                Cell::CellFormula { f, .. } => {
                    let cached = FormulaCachedValue::Number(0.0);
                    let rgce = encode_or_placeholder(parsed_formulas, *f, sheet_idx, anchor, xti, defined_names, extern_names, &cached, None);
                    w.write_record(R_FORMULA, &build_formula(rw, cl, ixfe, &cached, &rgce));
                }
                Cell::CellFormulaBoolean { f, v, .. } => {
                    let cached = FormulaCachedValue::Bool(*v);
                    let rgce = encode_or_placeholder(parsed_formulas, *f, sheet_idx, anchor, xti, defined_names, extern_names, &cached, None);
                    w.write_record(R_FORMULA, &build_formula(rw, cl, ixfe, &cached, &rgce));
                }
                Cell::CellFormulaNumber { f, v, .. } => {
                    let cached = FormulaCachedValue::Number(*v);
                    let rgce = encode_or_placeholder(parsed_formulas, *f, sheet_idx, anchor, xti, defined_names, extern_names, &cached, None);
                    w.write_record(R_FORMULA, &build_formula(rw, cl, ixfe, &cached, &rgce));
                }
                Cell::CellFormulaString { f, v, .. } => {
                    // Try the real formula encoder regardless of
                    // whether the cached value is empty. Pre-fix this
                    // branch only encoded for non-empty v and emitted
                    // a placeholder PtgMissArg formula for empty
                    // values, which dropped the formula identity:
                    // formulas like IF(...="Bolt Rope","",HLOOKUP(...))
                    // turned into EmptyCells on reload, and downstream
                    // IF checks comparing `=0` to that cell flipped
                    // (Excel coerces empty cells to 0 numerically but
                    // treats "" string as ≠ 0). Cascaded as numeric
                    // drift through the cost columns.
                    let mut emitted = false;
                    if let Some(node) = parsed_formulas.get(*f as usize) {
                        if let Some(rgce) = encode_formula_rgce_with_names(
                            node, sheet_idx, anchor, xti, defined_names, extern_names, CLS_VAL,
                        ) {
                            if v.is_empty() {
                                // Empty cached result — Blank tag
                                // (0x03), no STRING record. Real rgce
                                // still carries the formula text so
                                // dependents see the same expression.
                                let cached = FormulaCachedValue::Blank;
                                w.write_record(R_FORMULA, &build_formula(rw, cl, ixfe, &cached, &rgce));
                            } else {
                                let cached = FormulaCachedValue::StringPending;
                                w.write_record(R_FORMULA, &build_formula(rw, cl, ixfe, &cached, &rgce));
                                w.write_record(R_STRING, &build_string(v));
                            }
                            emitted = true;
                        }
                    }
                    if !emitted {
                        // Encoding failed (unsupported node type).
                        // For empty: emit FORMULA with Blank tag +
                        // placeholder rgce — formula identity lost,
                        // value preserved as "". For non-empty: route
                        // through the SST as LABELSST so the cached
                        // value survives in full (PtgStr truncates
                        // strings >255 chars; LABELSST has no cap).
                        if v.is_empty() {
                            let cached = FormulaCachedValue::Blank;
                            let rgce = placeholder_rgce(&cached, None);
                            w.write_record(R_FORMULA, &build_formula(rw, cl, ixfe, &cached, &rgce));
                        } else if let Some(&sst_idx) = sst_index.get(v) {
                            w.write_record(R_LABELSST, &build_labelsst(rw, cl, ixfe, sst_idx));
                        }
                    }
                }
                Cell::CellFormulaError { f, ei, .. } => {
                    let code = ironcalc_error_to_biff(ei);
                    let cached = FormulaCachedValue::Error(code);
                    // BUG-01 round-trip: if we preserved the original
                    // rgce on load (for cells whose formulas IronCalc
                    // couldn't parse), emit those bytes verbatim. The
                    // synthesized placeholder otherwise reloads as
                    // #VALUE! instead of the source's #ERROR!.
                    // Verbatim replay is only valid when the rgce is
                    // self-contained: PtgName / PtgNameX / Ptg*3d
                    // indices point into the SOURCE file's NAME /
                    // EXTERNSHEET tables, which this writer rebuilds
                    // with different numbering — replaying those made
                    // POI throw IndexOutOfBounds on the dangling ixti
                    // (2026-08-14 report follow-up).
                    // Verbatim replay when self-contained; otherwise
                    // remap the source rgce's table references into
                    // this file's rebuilt tables (the xti spans were
                    // registered before EXTERNSHEET emission). Keeps
                    // formula identity for cells IronCalc can't parse
                    // — e.g. the `#REF!#REF!` dead-3D-ref idiom.
                    let key = (sheet_idx, *row, *col);
                    let rgce = preserved_rgce
                        .and_then(|m| m.get(&key))
                        .and_then(|r| {
                            if rgce_is_self_contained(r) {
                                Some(r.clone())
                            } else {
                                remap_ctx.and_then(|c| {
                                    remap_source_rgce(r, c, defined_names, extern_names, xti)
                                })
                            }
                        })
                        .unwrap_or_else(|| {
                            encode_or_placeholder(parsed_formulas, *f, sheet_idx, anchor, xti, defined_names, extern_names, &cached, None)
                        });
                    w.write_record(R_FORMULA, &build_formula(rw, cl, ixfe, &cached, &rgce));
                }
            }
        }
        }
    }
}

// ---------------------------------------------------------------------------
// Top-level entry: build a minimal valid xls and write it as an OLE2
// compound file with a single `/Workbook` stream. The minimal file has
// one empty worksheet — enough to validate the framing and round-trip
// through Excel / our own BIFF reader.
// ---------------------------------------------------------------------------

pub fn save_xls<P: AsRef<Path>>(model: &ironcalc::base::Model, path: P) -> Result<(), String> {
    save_xls_with_preserved(model, path, None, None)
}

/// Save with optional VBA / macro storage preservation. When
/// `preserved` is `Some`, after the workbook stream is written the
/// captured storage subtrees from the source file are replayed into
/// the new CFB. Excel-side macros come back; everything else (charts,
/// pivots, drawings, conditional formatting, OLE links) is still
/// dropped — those would need full BIFF-stream offset preservation
/// (see pending #1).
///
/// `hidden_cols`: per-sheet (0-based sheet index → set of 1-based col
/// indices) hidden-column map. IronCalc's `Col` struct carries no
/// hidden field, so the caller threads this through from the AppState
/// side-channel populated on load.
pub fn save_xls_with_preserved<P: AsRef<Path>>(
    model: &ironcalc::base::Model,
    path: P,
    preserved: Option<&crate::xls_preserve::PreservedXlsData>,
    hidden_cols: Option<&HashMap<u32, HashSet<i32>>>,
) -> Result<(), String> {
    let bytes = build_xls_bytes_with_options(model, hidden_cols, None, preserved);
    write_xls_bytes_with_preserved(path, &bytes, preserved)
}

pub fn build_xls_bytes(
    model: &ironcalc::base::Model,
    hidden_cols: Option<&HashMap<u32, HashSet<i32>>>,
) -> Vec<u8> {
    build_workbook_stream(model, hidden_cols, None, None)
}

/// Save-path entry that also accepts a preserved-rgce map AND the
/// full source-substream snapshot. The rgce map handles `#ERROR!`
/// round-trip (BUG-01); the substream snapshot drives the
/// preservation passthrough that copies drawings, data validation,
/// AutoFilter, sheet protection, conditional formatting, print
/// settings, page-layout-view, theme/XFEXT/etc. through verbatim
/// (greenfield writer doesn't model those).
pub fn build_xls_bytes_with_options(
    model: &ironcalc::base::Model,
    hidden_cols: Option<&HashMap<u32, HashSet<i32>>>,
    preserved_rgce: Option<&HashMap<(u32, i32, i32), Vec<u8>>>,
    preserved: Option<&crate::xls_preserve::PreservedXlsData>,
) -> Vec<u8> {
    build_workbook_stream(model, hidden_cols, preserved_rgce, preserved)
}

pub fn write_xls_bytes_with_preserved<P: AsRef<Path>>(
    path: P,
    workbook_bytes: &[u8],
    preserved: Option<&crate::xls_preserve::PreservedXlsData>,
) -> Result<(), String> {
    write_compound_file(path.as_ref(), workbook_bytes, preserved)
}

/// Build the full workbook-stream byte sequence: globals (header
/// records + style table + SST + BOUNDSHEET8s) followed by one
/// substream per worksheet. Cell records are not yet emitted — phase
/// 3 walks `model.workbook.worksheets[i].sheet_data` and writes
/// BLANK / NUMBER / LABELSST / FORMULA records.
fn build_workbook_stream(
    model: &ironcalc::base::Model,
    hidden_cols: Option<&HashMap<u32, HashSet<i32>>>,
    preserved_rgce: Option<&HashMap<(u32, i32, i32), Vec<u8>>>,
    preserved: Option<&crate::xls_preserve::PreservedXlsData>,
) -> Vec<u8> {
    use crate::xls_save_passthrough as pt;
    // Build sheet-name → source-substream-index map (lowercased keys
    // for case-insensitive lookup, since BIFF is case-insensitive
    // about sheet names). Empty when there's no preserved data, in
    // which case every per-sheet zone lookup falls through to the
    // greenfield emission with no passthrough.
    let globals_substream = preserved.and_then(|p| {
        p.workbook_substreams.iter().find(|s| s.bof_dt == 0x0005)
    });
    let sheet_substream_index: std::collections::HashMap<String, usize> = globals_substream
        .map(pt::build_sheet_substream_index)
        .unwrap_or_default();
    let lookup_sheet_substream = |name: &str| -> Option<&crate::xls_preserve::SubstreamSnapshot> {
        let idx = sheet_substream_index.get(&name.to_lowercase())?;
        preserved?.workbook_substreams.get(*idx)
    };
    let mut w = BiffWriter::new();
    let sheet_names: Vec<&str> = model
        .workbook
        .worksheets
        .iter()
        .map(|s| s.name.as_str())
        .collect();

    // ----- Globals substream -----
    // Passthrough records from the source globals, grouped by the
    // writer-owned record they preceded in the source. Flushed at the
    // matching emission points below so Office File Validation sees
    // them in grammar-valid positions; whatever anchor the writer
    // never reaches drains right before EOF.
    let mut globals_zones = globals_substream
        .map(pt::globals_passthrough_zones)
        .unwrap_or_else(pt::GlobalsZones::empty);
    macro_rules! flush_zone {
        ($w:expr, $anchor:expr) => {
            for rec in globals_zones.take($anchor) {
                $w.write_record(rec.opcode, &rec.data);
            }
        };
    }
    w.write_record(R_BOF, &build_bof(DT_GLOBALS));
    w.write_record(R_INTERFACEHDR, &build_interfacehdr());
    w.write_record(R_MMS, &build_mms());
    w.write_record(R_INTERFACEEND, &[]);
    w.write_record(R_WRITEACCESS, &build_writeaccess());
    w.write_record(R_CODEPAGE, &build_codepage());
    w.write_record(R_DSF, &build_dsf());
    // The early positional block (EXCEL9FILE, TABID, OBPROJ,
    // CODENAME, FNGROUPCOUNT, workbook protection) sits between DSF
    // and WINDOW1 in Excel-authored files.
    flush_zone!(w, R_WINDOW1);
    w.write_record(R_REFRESHALL, &build_refreshall());
    w.write_record(R_BOOKBOOL, &build_bookbool());

    // Build the model-driven style tables once; emit FONT, FORMAT,
    // XF, PALETTE records from them. The same StyleTables drives
    // emit_sheet_cells's per-cell ixfe lookup later in the function.
    let mut style_tables = build_style_tables(model);

    // Per-sheet passthrough zones, computed up front (rather than in
    // the sheet loop) because embedded chart streams reference the
    // globals FONT table by ifnt: FONTX / FBI records carry indices
    // into the SOURCE font table, which this writer rebuilds smaller.
    // Dangling ifnts made Office File Validation reject every saved
    // workbook with an embedded chart (2026-08-14 report follow-up).
    // Carry each referenced source FONT into our table and remap the
    // ifnt bytes at emission (`write_zone_record`).
    let pt_suppress_names: Vec<String> = std::env::var("FASTSHEET_XLS_NO_PT_SHEETS")
        .map(|v| v.split(',').map(|s| s.trim().to_lowercase()).collect())
        .unwrap_or_default();
    let all_zones: Vec<pt::SheetZones> = sheet_names
        .iter()
        .map(|name| {
            let suppressed = pt_suppress_names.iter().any(|s| s == &name.to_lowercase());
            lookup_sheet_substream(name)
                .filter(|_| !suppressed)
                .map(pt::split_sheet_zones)
                .unwrap_or_else(pt::SheetZones::empty)
        })
        .collect();
    let extra_substreams = preserved
        .map(|p| pt::extract_extra_substreams(&p.workbook_substreams))
        .unwrap_or_default();
    let source_fonts: Vec<&crate::xls_preserve::RawRecord> = globals_substream
        .map(|g| g.records.iter().filter(|r| r.opcode == R_FONT).collect())
        .unwrap_or_default();
    let mut wanted_ifnts: std::collections::BTreeSet<u16> = std::collections::BTreeSet::new();
    {
        let mut scan = |recs: &[&crate::xls_preserve::RawRecord]| {
            for rec in recs {
                if let Some(ifnt) = zone_record_ifnt(rec.opcode, &rec.data) {
                    wanted_ifnts.insert(ifnt);
                }
            }
        };
        for z in &all_zones {
            scan(&z.pre_dim);
            scan(&z.post_cells_pre_win2);
            scan(&z.post_win2_pre_selection);
            scan(&z.post_win2);
        }
        for extra in &extra_substreams {
            for rec in extra.records {
                if let Some(ifnt) = zone_record_ifnt(rec.opcode, &rec.data) {
                    wanted_ifnts.insert(ifnt);
                }
            }
        }
    }
    let mut font_remap: HashMap<u16, u16> = HashMap::new();
    for &ifnt in &wanted_ifnts {
        // ifnt ↔ stream index: BIFF skips ifnt 4 with no placeholder
        // record, so stream record k carries ifnt k for k<4 and k+1
        // from the fifth record on.
        let src_idx = if ifnt < 4 { ifnt } else { ifnt - 1 } as usize;
        let Some(body) = source_fonts.get(src_idx) else { continue };
        let our_idx = match style_tables.fonts.iter().position(|f| f == &body.data) {
            Some(i) => i,
            None => {
                style_tables.fonts.push(body.data.clone());
                style_tables.fonts.len() - 1
            }
        };
        let our_ifnt = if our_idx < 4 { our_idx } else { our_idx + 1 } as u16;
        if our_ifnt != ifnt {
            font_remap.insert(ifnt, our_ifnt);
        }
    }

    flush_zone!(w, R_FONT);
    for body in &style_tables.fonts {
        w.write_record(R_FONT, body);
    }
    flush_zone!(w, R_FORMAT);
    for body in &style_tables.formats {
        w.write_record(R_FORMAT, body);
    }
    flush_zone!(w, R_XF);
    for body in &style_tables.xfs {
        w.write_record(R_XF, body);
    }
    // XFCRC + XFEXT: true colors for palette-approximated slots.
    // Excel validates the crc against the XF records above and only
    // honours the XFEXTs when it matches.
    if !style_tables.xfext.is_empty() {
        w.write_record(R_XFCRC, &build_xfcrc_record(&style_tables.xfs));
        for (ixfe, exts) in &style_tables.xfext {
            w.write_record(R_XFEXT, &build_xfext_record(*ixfe, exts));
        }
    }

    // STYLE records — register the 6 standard built-in styles. The
    // first 6 of the 15 style XFs we emitted serve as their backing
    // XF entries.
    flush_zone!(w, R_STYLE);
    w.write_record(R_STYLE, &build_style_builtin(0, 0, 0xFF));   // Normal
    w.write_record(R_STYLE, &build_style_builtin(1, 3, 0xFF));   // Comma
    w.write_record(R_STYLE, &build_style_builtin(2, 6, 0xFF));   // Comma [0]
    w.write_record(R_STYLE, &build_style_builtin(3, 4, 0xFF));   // Currency
    w.write_record(R_STYLE, &build_style_builtin(4, 7, 0xFF));   // Currency [0]
    w.write_record(R_STYLE, &build_style_builtin(5, 5, 0xFF));   // Percent

    // PALETTE record — only emit when we actually need to override
    // a default slot. [MS-XLS] 2.4.188: ccv (u16=56) + 56 LongRGB
    // (R, G, B, reserved) covering icv 8..63. We seed it with the
    // BIFF defaults so unaltered slots keep their standard colors;
    // then we overlay our custom slots on top. Without this seeding,
    // every font referencing red=10, blue=12, etc. via default icvs
    // would render as black (the previous "unused slot" filler).
    flush_zone!(w, 0x0092 /* PALETTE */);
    if !style_tables.palette.is_empty() {
        let mut body = Vec::with_capacity(2 + 56 * 4);
        body.put_u16(56);
        // Seed with BIFF defaults for icv 8..63.
        let mut slots: [[u8; 3]; 56] = [[0, 0, 0]; 56];
        for &(icv, rgb) in BIFF_DEFAULT_PALETTE {
            if (8..=63).contains(&icv) {
                slots[(icv - 8) as usize] = rgb;
            }
        }
        // Overlay custom assignments.
        for &(icv, rgb) in &style_tables.palette {
            if (8..=63).contains(&icv) {
                slots[(icv - 8) as usize] = rgb;
            }
        }
        for rgb in &slots {
            body.put_u8(rgb[0]);
            body.put_u8(rgb[1]);
            body.put_u8(rgb[2]);
            body.put_u8(0); // reserved
        }
        const R_PALETTE_LOCAL: u16 = 0x0092;
        w.write_record(R_PALETTE_LOCAL, &body);
    }

    flush_zone!(w, R_USESELFS);
    w.write_record(R_USESELFS, &build_useselfs_yes());
    flush_zone!(w, R_HIDEOBJ);
    w.write_record(R_HIDEOBJ, &build_hideobj_show_all());
    flush_zone!(w, R_DATEMODE);
    w.write_record(R_DATEMODE, &build_datemode_1900());
    flush_zone!(w, R_PRECISION);
    w.write_record(R_PRECISION, &build_precision_full());
    flush_zone!(w, R_BACKUP);
    w.write_record(R_BACKUP, &build_backup_off());
    // WINDOW1 — required globals record per [MS-XLS] 2.4.346. Missing
    // it triggered Excel's "file format or extension is not valid" /
    // corrupt-file warning on round-trip even though the workbook
    // stream parsed cleanly via calamine + cfb. Carries window
    // geometry, the active sheet index (itabCur), and the selected-
    // sheet count. Static defaults are fine — Excel re-saves with
    // its real geometry on first save anyway.
    w.write_record(R_WINDOW1, &build_window1());

    // BOUNDSHEET8 — one per sheet. Track each lbPlyPos byte offset so
    // we can patch in the real sheet-BOF byte position later.
    flush_zone!(w, R_BOUNDSHEET8);
    let mut boundsheet_patch_offsets: Vec<usize> = Vec::with_capacity(sheet_names.len());
    for name in &sheet_names {
        let (body, field_offset) = build_boundsheet8(name);
        let header_offset = w.pos() as usize;
        w.write_record(R_BOUNDSHEET8, &body);
        boundsheet_patch_offsets.push(header_offset + 4 + field_offset);
    }

    // Chart / macro sheets from the source (extra_substreams was
    // computed with the zones above so its font references joined the
    // ifnt remap). Tab order shifts (charts move to the end) —
    // acceptable cost for preservation.
    let mut extra_patch_offsets: Vec<usize> = Vec::with_capacity(extra_substreams.len());
    for extra in &extra_substreams {
        let (body, field_offset) = build_boundsheet8_full(&extra.name, extra.hs_state, extra.dt);
        let header_offset = w.pos() as usize;
        w.write_record(R_BOUNDSHEET8, &body);
        extra_patch_offsets.push(header_offset + 4 + field_offset);
    }

    flush_zone!(w, R_COUNTRY);
    w.write_record(R_COUNTRY, &build_country());

    // Pre-pass: compute per-sheet extents + LABELSST refcount so the
    // SST header can carry the correct cstTotal.
    let extents: Vec<SheetExtents> = model
        .workbook
        .worksheets
        .iter()
        .map(compute_sheet_extents)
        .collect();
    let total_labelsst_refs: u32 = extents.iter().map(|e| e.labelsst_count).sum();

    // SUPBOOK + EXTERNSHEET + EXTERNNAME + NAME records. Pre-pass
    // the workbook to collect all xti entries (cell formulas +
    // defined names) and the function names that need EXTERNNAME
    // routing (UDFs + post-2007 functions not in the BIFF FTAB).
    let mut xti = build_xti_table(model);
    let mut extern_names = build_extern_name_table(model);
    if !extern_names.names.is_empty() {
        // Reserve the add-in supbook's EXTERNSHEET entry BEFORE the
        // defined-name build below — defined-name rgces may contain
        // ATP calls whose PtgNameX bakes this ixti in.
        extern_names.addin_ixti = xti.addin_ixti();
    }
    let mut defined_names = build_defined_name_table(model, &mut xti, &extern_names);
    // Preserved-rgce replay context: source NAME order + source xti
    // sheet names + the model's sheet-name index. Shared by array
    // re-emission, unimported-NAME replay, and #ERROR!-cell replay.
    let sheet_by_name: HashMap<String, u32> = model
        .workbook
        .worksheets
        .iter()
        .enumerate()
        .map(|(i, ws)| (ws.name.to_lowercase(), i as u32))
        .collect();
    let src_ctx: Option<SourceRgceCtx> = preserved.map(|p| SourceRgceCtx {
        names: &p.source_defined_names,
        xti_sheets: &p.source_xti_sheets,
        sheet_by_name: &sheet_by_name,
    });
    // Unimported source NAME records to append after the writer's own
    // Lbl records (registers their ilbls in the lookup so preserved
    // rgce referencing them can remap).
    let preserved_name_plans =
        plan_preserved_names(preserved, &mut defined_names, &extern_names, &sheet_by_name);
    // Register every sheet span referenced from preserved rgce BEFORE
    // the EXTERNSHEET record is emitted below — remapping at cell/Lbl
    // emission time can only LOOK UP entries.
    if let Some(ctx) = src_ctx.as_ref() {
        register_preserved_xti(preserved, preserved_rgce, ctx, &mut xti);
    }
    // Array-formula re-emission plans: which cell ranges get folded
    // back into FORMULA(PtgExp) + ARRAY records instead of expanded
    // per-cell formulas. Needs both name tables for the rgce remap.
    let array_plans = build_array_plans(
        model,
        preserved,
        src_ctx.as_ref(),
        &defined_names,
        &extern_names,
        &xti,
    );
    if !xti.entries.is_empty() {
        // Record order is load-bearing for POI's LinkTable: each
        // SUPBOOK is immediately followed by ITS EXTERNNAME records,
        // and the single EXTERNSHEET comes after all supbook blocks.
        // (The previous layout — EXTERNNAMEs after EXTERNSHEET,
        // attached to the internal supbook — made POI NPE in
        // resolveNameXText on every ATP call.)
        //
        // SUPBOOK 0: ctab(u16) = workbook sheet count, then 0x0401
        // marker meaning "this workbook is the supporting book".
        flush_zone!(w, R_SUPBOOK);
        let mut body: Vec<u8> = Vec::with_capacity(4);
        body.put_u16(model.workbook.worksheets.len() as u16);
        body.put_u16(0x0401);
        w.write_record(R_SUPBOOK, &body);

        if !extern_names.names.is_empty() {
            // SUPBOOK 1: the add-in supbook ([MS-XLS] 2.4.271 —
            // ctab=1, virtPath marker 0x3A01), holding one EXTERNNAME
            // per ATP function. PtgNameX nameindex is 1-based into
            // this list.
            let mut body: Vec<u8> = Vec::with_capacity(4);
            body.put_u16(0x0001);
            body.put_u16(0x3A01);
            w.write_record(R_SUPBOOK, &body);
            for name in &extern_names.names {
                let body = build_externname_record(name);
                w.write_record(R_EXTERNNAME, &body);
            }
        }

        // EXTERNSHEET: cXTI(u16), then per entry: iSupBook(u16) +
        // itabFirst(u16) + itabLast(u16). Sheet entries point at the
        // internal SUPBOOK (iSupBook=0); the reserved add-in entry
        // points at SUPBOOK 1 with the sheet-independent 0xFFFE tabs.
        let mut body: Vec<u8> = Vec::with_capacity(2 + xti.entries.len() * 6);
        body.put_u16(xti.entries.len() as u16);
        for (i, &(first, last)) in xti.entries.iter().enumerate() {
            let is_addin = xti.addin_entry == Some(i as u16);
            body.put_u16(if is_addin { 1 } else { 0 }); // iSupBook
            body.put_u16(first);
            body.put_u16(last);
        }
        w.write_record(R_EXTERNSHEET, &body);
    }

    // Emit NAME (Lbl) records — one per defined name, in the same
    // order as defined_names.entries (so PtgName.ilbl values resolve
    // correctly on read).
    flush_zone!(w, R_EXTERNSHEET);
    flush_zone!(w, R_LBL);
    for entry in &defined_names.entries {
        let body = build_lbl_record(entry);
        w.write_record(R_LBL, &body);
    }
    // Then the NAME-route function names (_xlfn future functions +
    // VBA UDFs) — their ilbl values continue the defined-name
    // numbering (extern_names.ilbl_base), so order and position here
    // are load-bearing.
    debug_assert_eq!(defined_names.entries.len() as u16, extern_names.ilbl_base);
    for (name, is_xlfn) in &extern_names.name_fns {
        let body = build_fn_lbl_record(name, *is_xlfn);
        w.write_record(R_LBL, &body);
    }
    // Then the preserved source NAME records the loader couldn't
    // import (builtin _FilterDatabase / Print_Titles, _xlfn stubs,
    // grammar-rejected and #REF! names) — replayed with their rgce
    // remapped into this file's tables. A failed remap emits the Lbl
    // with an empty rgce (resolves #REF!) rather than dropping the
    // record: the name keeps existing, which is what AutoFilter /
    // print setup / VBA lookups care about.
    if let Some(p) = preserved {
        for &(idx, itab) in &preserved_name_plans {
            let pn = &p.preserved_names[idx];
            let rgce = src_ctx
                .as_ref()
                .and_then(|c| remap_source_rgce(&pn.rgce, c, &defined_names, &extern_names, &xti))
                .unwrap_or_default();
            w.write_record(R_LBL, &build_preserved_lbl_record(pn, itab, &rgce));
        }
    }

    // SST + EXTSST. The pre-SST zone carries RECALCID, FEAT headers,
    // and MSODRAWINGGROUP (+CONTINUE) — the drawing-objects shared
    // state the per-sheet MSODRAWING records reference.
    flush_zone!(w, 0x00FC /* SST */);
    let (sst_table, sst_idx) = build_sst_table(model);
    emit_sst(&mut w, &sst_table, total_labelsst_refs);

    // Trailing zone (post-EXTSST FRTs: BOOKEXT, THEME,
    // COMPRESSPICTURES, GUIDTYPELIB, …) plus any zone whose anchor
    // the writer never emitted, in source order.
    for rec in globals_zones.drain_remaining() {
        w.write_record(rec.opcode, &rec.data);
    }

    w.write_record(R_EOF, &build_eof());

    // ----- Per-sheet substreams -----
    let empty_formulas: Vec<ironcalc::base::expressions::parser::Node> = Vec::new();
    for (i, name) in sheet_names.iter().enumerate() {
        let ws = &model.workbook.worksheets[i];
        let parsed = model
            .parsed_formulas
            .get(i)
            .map(|v| v.as_slice())
            .unwrap_or(&empty_formulas);
        let bof_pos = w.pos();
        w.patch_u32(boundsheet_patch_offsets[i], bof_pos);
        w.write_record(R_BOF, &build_bof(DT_WORKSHEET));
        w.write_record(R_CALCMODE, &build_calcmode_auto());
        w.write_record(R_CALCCOUNT, &build_calccount());
        w.write_record(R_REFMODE, &build_refmode_a1());
        w.write_record(R_ITERATION, &build_iteration_off());
        w.write_record(R_DELTA, &build_delta());
        w.write_record(R_SAVERECALC, &build_saverecalc());
        w.write_record(R_PRINTHEADERS, &build_printheaders_off());
        w.write_record(R_PRINTGRIDLINES, &build_printgridlines_off());
        w.write_record(R_GRIDSET, &build_gridset());
        w.write_record(R_GUTS, &build_guts());
        w.write_record(R_DEFAULTROWHEIGHT, &build_defaultrowheight());
        w.write_record(R_WSBOOL, &build_wsbool());

        // Preservation: per-sheet zones split out from the source
        // substream of the sheet with the matching name. Empty zones
        // when there's no preserved data — the writer falls back to
        // greenfield emission.
        let zones = &all_zones[i];

        // pre_dim zone: classic page-setup records (HEADER/FOOTER,
        // HCENTER, VCENTER, margins, SETUP, PLS, page breaks),
        // sheet-protection PROTECT/PASSWORD, etc. — landed here per
        // [MS-XLS] §2.1.7.20.1 ordering.
        for rec in &zones.pre_dim {
            write_zone_record(&mut w, rec, &font_remap);
        }

        // DEFCOLWIDTH — Excel always writes it between the page-setup
        // block and COLINFO. 8 chars is the standard default.
        let mut dcw = Vec::with_capacity(2);
        dcw.put_u16(8);
        w.write_record(R_DEFCOLWIDTH, &dcw);

        // COLINFO records — derived from IronCalc Col entries layered
        // with the AppState hidden_cols side-channel. BIFF8 ordering
        // puts COLINFO between WSBOOL and DIMENSIONS.
        let empty_hidden = HashSet::new();
        let sheet_hidden = hidden_cols
            .and_then(|m| m.get(&(i as u32)))
            .unwrap_or(&empty_hidden);
        for body in build_colinfo_records(ws, sheet_hidden) {
            w.write_record(R_COLINFO, &body);
        }

        // DIMENSIONS — 0-based inclusive→exclusive: rwMic / rwMac /
        // colMic / colMac. Excel rejects rwMac=0 in non-empty sheets,
        // so for an empty sheet we emit (0, 0, 0, 0); for populated
        // ones we convert IronCalc's 1-based bounds.
        let (rw_mic, rw_mac, col_mic, col_mac) = match extents[i].bounds {
            None => (0u32, 0u32, 0u16, 0u16),
            Some((rmin, rmax, cmin, cmax)) => (
                (rmin - 1).max(0) as u32,
                rmax as u32, // already exclusive: rmax is the last-used 1-based row
                (cmin - 1).max(0) as u16,
                cmax as u16,
            ),
        };
        w.write_record(R_DIMENSIONS, &build_dimensions(rw_mic, rw_mac, col_mic, col_mac));

        // ROW + cell records are emitted together by emit_sheet_cells
        // in [MS-XLS] row-block layout (≤32 ROW records, then those
        // rows' cells) — every populated row gets a ROW record.
        emit_sheet_cells(&mut w, ws, i as u32, &sst_table, &sst_idx, parsed, &xti, &defined_names, &extern_names, &style_tables, preserved_rgce, array_plans.get(&(i as u32)), src_ctx.as_ref());

        // post_cells_pre_win2 zone: CONDFMT/CF, HLINK, DVAL/DV,
        // PHONETICINFO, FEAT, etc. — these land between the cell
        // table and WINDOW2 per spec. MERGECELLS is filtered out of
        // the passthrough (writer-owned) and re-emitted from the
        // model's merge list below so merge edits survive a save.
        for rec in &zones.post_cells_pre_win2 {
            write_zone_record(&mut w, rec, &font_remap);
        }

        w.write_record(R_WINDOW2, &build_window2(ws.frozen_rows, ws.frozen_columns, i == 0));
        // Window-block grammar: WINDOW2 → PLV/SCL → PANE → SELECTION.
        // The pre-selection slice of the post zone carries the
        // source's PLV / SCL records.
        for rec in &zones.post_win2_pre_selection {
            write_zone_record(&mut w, rec, &font_remap);
        }
        if ws.frozen_rows > 0 || ws.frozen_columns > 0 {
            w.write_record(R_PANE, &build_pane(ws.frozen_rows, ws.frozen_columns));
        }
        // SELECTION — one per visible pane, top-left first, the
        // data-area pane (PANE.pnnAcc) last. Excel emits these after
        // WINDOW2/PANE; omitting them entirely (the old behavior)
        // contributed to Excel's "File error: data may have been
        // lost" repair prompt on open.
        for pnn in selection_panes(ws.frozen_rows, ws.frozen_columns) {
            w.write_record(R_SELECTION, &build_selection(pnn));
        }
        // MERGECELLS goes AFTER WINDOW2/PANE/SELECTION — that's where
        // Excel writes it. Emitting it before WINDOW2 (the old spot)
        // put it inside the sheet's object block on sheets with
        // drawings, which Office File Validation rejects.
        if !ws.merge_cells.is_empty() {
            let ranges: Vec<&str> = ws.merge_cells.iter().map(|s| s.as_str()).collect();
            for chunk in ranges.chunks(1026) {
                w.write_record(R_MERGECELLS, &build_mergecells(chunk));
            }
        }

        // post_win2 zone: drawings (MSODRAWING/OBJ/TXO + CONTINUE),
        // SHEETPROTECTION (FRT), RANGEPROTECTION, FRT HEADERFOOTER,
        // PLV, FORCEFULLCALCULATION, FILTERMODE/AUTOFILTERINFO/
        // AUTOFILTER. WINDOW2/PANE/SELECTION are filtered out
        // because we re-emit them above from the model — keeps
        // frozen-pane state consistent with edits.
        for rec in &zones.post_win2 {
            write_zone_record(&mut w, rec, &font_remap);
        }

        w.write_record(R_EOF, &build_eof());
    }

    // ----- Preserved chart / macro substreams -----
    // IronCalc doesn't represent chart sheets, so they round-tripped
    // through nothing in the model. Emit the raw record stream for
    // each with the BOUNDSHEET8 lbPlyPos patched to the substream's
    // position in the output. Records (BOF, the chart's frames, EOF)
    // are copied verbatim; cross-references inside the chart
    // substream are self-contained.
    for (extra, patch_off) in extra_substreams.iter().zip(extra_patch_offsets.iter()) {
        let bof_pos = w.pos();
        w.patch_u32(*patch_off, bof_pos);
        for rec in extra.records {
            write_zone_record(&mut w, rec, &font_remap);
        }
    }

    w.into_bytes()
}

/// Wrap a workbook-stream byte sequence in an OLE2 compound file at
/// `path`, with a single `/Workbook` stream entry. Overwrites any
/// existing file at the target path.
///
/// Uses CFB version 3 (512-byte sectors) — that's what Excel and every
/// in-the-wild .xls reader expects. cfb-rs's `create()` defaults to
/// version 4 (4096-byte sectors); cfb-rs can re-read those, but
/// calamine's stricter-or-older CFB parser rejects them with "Empty
/// Root directory". (Reproduced 2026-04-25; reading a v4-CFB-wrapped
/// xls via calamine returned that error even though cfb-rs
/// round-tripped the same bytes fine.)
fn write_compound_file(
    path: &Path,
    workbook_bytes: &[u8],
    preserved: Option<&crate::xls_preserve::PreservedXlsData>,
) -> Result<(), String> {
    // Atomic save: build the CFB at a sibling temp path so a mid-write
    // failure (disk full, IO error) leaves the existing file intact.
    crate::atomic::write(path, |tmp| build_cfb_at(tmp, workbook_bytes, preserved))
}

/// Synthesize the CFB-wrapped workbook stream at `path`. Caller is
/// responsible for atomic placement; this just writes the file.
fn build_cfb_at(
    path: &Path,
    workbook_bytes: &[u8],
    preserved: Option<&crate::xls_preserve::PreservedXlsData>,
) -> Result<(), String> {
    let file = std::fs::OpenOptions::new()
        .read(true)
        .write(true)
        .create_new(true)
        .open(path)
        .map_err(|e| format!("xls create file: {e}"))?;
    let mut comp = cfb::CompoundFile::create_with_version(cfb::Version::V3, file)
        .map_err(|e| format!("xls create cfb: {e}"))?;
    {
        let mut stream = comp
            .create_stream("/Workbook")
            .map_err(|e| format!("xls create stream: {e}"))?;
        stream
            .write_all(workbook_bytes)
            .map_err(|e| format!("xls write stream: {e}"))?;
    }
    // Replay preserved VBA / macro storages from the source CFB. Lives
    // in its own self-contained subtree so verbatim copy is safe — no
    // offset cross-references with the workbook stream we just wrote.
    if let Some(p) = preserved {
        crate::xls_preserve::inject(&mut comp, p);
    }
    comp.flush().map_err(|e| format!("xls flush: {e}"))?;
    Ok(())
}

// ---------------------------------------------------------------------------
// Tests — round-trip the minimal stub through cfb, parse the BIFF
// records back, assert the framing matches what we wrote.
// ---------------------------------------------------------------------------

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Read;

    fn read_records(buf: &[u8]) -> Vec<(u16, usize)> {
        let mut out = Vec::new();
        let mut pos = 0;
        while pos + 4 <= buf.len() {
            let rt = u16::from_le_bytes([buf[pos], buf[pos + 1]]);
            let rl = u16::from_le_bytes([buf[pos + 2], buf[pos + 3]]) as usize;
            out.push((rt, rl));
            pos += 4 + rl;
        }
        out
    }

    /// Build a model-derived workbook stream with the given sheet
    /// names — convenience for tests that don't need real cell content.
    fn build_with_sheets(names: &[&str]) -> Vec<u8> {
        let mut model = ironcalc::base::Model::new_empty("test", "en", "UTC", "en").unwrap();
        // Model::new_empty starts with one sheet "Sheet1". Rename / add
        // to match the requested set.
        if let Some(first) = names.first() {
            let _ = model.rename_sheet_by_index(0, first);
        }
        for n in names.iter().skip(1) {
            let _ = model.add_sheet(n);
        }
        build_workbook_stream(&model, None, None, None)
    }

    #[test]
    fn workbook_stream_has_expected_globals_framing() {
        let bytes = build_with_sheets(&["Sheet1"]);
        let recs = read_records(&bytes);
        let types: Vec<u16> = recs.iter().map(|(t, _)| *t).collect();

        // Spot-check ordering of key records up to the first EOF.
        let globals_eof = types.iter().position(|&t| t == R_EOF).unwrap();
        let globals = &types[..=globals_eof];
        // Required header records are present.
        assert!(globals.contains(&R_INTERFACEHDR));
        assert!(globals.contains(&R_CODEPAGE));
        assert!(globals.contains(&R_WRITEACCESS));
        // Style table. With model-driven emit: at least 5 FONTs
        // (BIFF8 skip-index-4 convention) and at least 16 XFs (15
        // style + ≥1 cell). New_empty's default model has 1 cell_xf,
        // bringing the XF total to 17.
        assert!(globals.iter().filter(|&&t| t == R_FONT).count() >= 5);
        assert!(globals.iter().filter(|&&t| t == R_XF).count() >= 16);
        assert_eq!(globals.iter().filter(|&&t| t == R_STYLE).count(), 6);
        // Sheet registration + SST.
        assert_eq!(globals.iter().filter(|&&t| t == R_BOUNDSHEET8).count(), 1);
        assert!(globals.contains(&R_SST));
        assert!(globals.contains(&R_EXTSST));
        // WINDOW1 — required globals record. Missing it triggered
        // Excel's corrupt-file warning on round-trip.
        assert!(globals.contains(&R_WINDOW1), "globals must include WINDOW1");

        // After globals EOF: sheet substream BOF / DIMENSIONS / WINDOW2 /
        // EOF must be present in order.
        let sheet = &types[globals_eof + 1..];
        assert_eq!(sheet[0], R_BOF);
        assert_eq!(*sheet.last().unwrap(), R_EOF);
        assert!(sheet.contains(&R_DIMENSIONS));
        assert!(sheet.contains(&R_WINDOW2));
    }

    #[test]
    fn boundsheet_lbplypos_points_at_sheet_bof() {
        let sheet_names = ["Sheet1", "Sheet2"];
        let bytes = build_with_sheets(&sheet_names);

        // Walk records to find each sheet BOF's file offset (BOFs that
        // appear AFTER the first EOF; the first BOF is the globals one).
        let mut sheet_bof_offsets: Vec<u32> = Vec::new();
        let mut pos = 0u32;
        let mut seen_globals_eof = false;
        while (pos as usize) + 4 <= bytes.len() {
            let p = pos as usize;
            let rt = u16::from_le_bytes([bytes[p], bytes[p + 1]]);
            let rl = u16::from_le_bytes([bytes[p + 2], bytes[p + 3]]) as u32;
            if rt == R_EOF && !seen_globals_eof {
                seen_globals_eof = true;
            } else if rt == R_BOF && seen_globals_eof {
                sheet_bof_offsets.push(pos);
            }
            pos += 4 + rl;
        }
        assert_eq!(sheet_bof_offsets.len(), sheet_names.len());

        // Walk BOUNDSHEET8 records in the globals substream and read
        // lbPlyPos out of each — must equal the sheet BOF offsets above.
        let mut bs_pos = 0u32;
        let mut bs_idx = 0;
        while (bs_pos as usize) + 4 <= bytes.len() {
            let p = bs_pos as usize;
            let rt = u16::from_le_bytes([bytes[p], bytes[p + 1]]);
            let rl = u16::from_le_bytes([bytes[p + 2], bytes[p + 3]]) as u32;
            if rt == R_BOUNDSHEET8 {
                let body = &bytes[p + 4..p + 4 + rl as usize];
                let lb_ply_pos = u32::from_le_bytes([body[0], body[1], body[2], body[3]]);
                assert_eq!(
                    lb_ply_pos,
                    sheet_bof_offsets[bs_idx],
                    "BOUNDSHEET8[{}].lbPlyPos must equal sheet BOF offset",
                    bs_idx
                );
                bs_idx += 1;
            } else if rt == R_EOF {
                break;
            }
            bs_pos += 4 + rl;
        }
        assert_eq!(bs_idx, sheet_names.len());
    }

    #[test]
    fn save_xls_writes_readable_compound_file() {
        let model = ironcalc::base::Model::new_empty("test", "en", "UTC", "en").unwrap();
        let dir = std::env::temp_dir();
        let path = dir.join("fastsheet_xls_save_smoke.xls");
        save_xls(&model, &path).expect("save_xls");
        assert!(path.exists());

        // Re-open via cfb and read the workbook stream back; verify the
        // critical records are present in the right substreams.
        let mut comp = cfb::open(&path).expect("cfb open");
        assert!(comp.exists("/Workbook"));
        let mut stream = comp.open_stream("/Workbook").expect("open stream");
        let mut buf = Vec::new();
        stream.read_to_end(&mut buf).expect("read stream");
        let recs = read_records(&buf);
        let types: Vec<u16> = recs.iter().map(|(t, _)| *t).collect();
        // First record must be a BOF (globals).
        assert_eq!(types[0], R_BOF);
        // Style table emitted.
        assert!(types.contains(&R_FONT));
        assert!(types.contains(&R_XF));
        assert!(types.contains(&R_STYLE));
        // Sheet registration emitted.
        assert!(types.contains(&R_BOUNDSHEET8));
        // SST + EXTSST present.
        assert!(types.contains(&R_SST));
        assert!(types.contains(&R_EXTSST));
        std::fs::remove_file(&path).ok();
    }

    #[test]
    fn sst_table_clones_model_shared_strings() {
        let mut model = ironcalc::base::Model::new_empty("t", "en", "UTC", "en").unwrap();
        // Inject a shared-string cell so the model populates its SST.
        model.set_user_input(0, 1, 1, "Hello".to_string()).unwrap();
        model.set_user_input(0, 2, 1, "World".to_string()).unwrap();
        let (table, idx) = build_sst_table(&model);
        assert!(table.iter().any(|s| s == "Hello"));
        assert!(table.iter().any(|s| s == "World"));
        assert!(idx.contains_key("Hello"));
    }

    #[test]
    fn cell_records_emitted_for_populated_sheet() {
        let mut model = ironcalc::base::Model::new_empty("t", "en", "UTC", "en").unwrap();
        // Mix of types: number, string, formula→number, blank-with-style.
        model.set_user_input(0, 1, 1, "42".to_string()).unwrap();
        model.set_user_input(0, 1, 2, "hello".to_string()).unwrap();
        model.set_user_input(0, 2, 1, "=A1*2".to_string()).unwrap();
        // Force evaluation so the formula gets a cached numeric value.
        model.evaluate();

        let bytes = build_workbook_stream(&model, None, None, None);
        let recs = read_records(&bytes);
        let types: Vec<u16> = recs.iter().map(|(t, _)| *t).collect();

        // We should now see at least one NUMBER, LABELSST, FORMULA in
        // the sheet substream (i.e. after the globals EOF).
        let globals_eof = types.iter().position(|&t| t == R_EOF).unwrap();
        let sheet = &types[globals_eof + 1..];
        assert!(sheet.contains(&R_NUMBER), "expected NUMBER cell record");
        assert!(sheet.contains(&R_LABELSST), "expected LABELSST cell record");
        assert!(sheet.contains(&R_FORMULA), "expected FORMULA cell record");
    }

    #[test]
    fn dimensions_reflects_used_range() {
        let mut model = ironcalc::base::Model::new_empty("t", "en", "UTC", "en").unwrap();
        // Cells at (3, 2) and (5, 7) → rwMic=2, rwMac=5 (1-based last row),
        // colMic=1, colMac=7 (1-based last col).
        model.set_user_input(0, 3, 2, "1".to_string()).unwrap();
        model.set_user_input(0, 5, 7, "1".to_string()).unwrap();
        let bytes = build_workbook_stream(&model, None, None, None);

        // Walk to the DIMENSIONS record in the sheet substream.
        let mut pos = 0usize;
        let mut seen_globals_eof = false;
        let mut found = false;
        while pos + 4 <= bytes.len() {
            let rt = u16::from_le_bytes([bytes[pos], bytes[pos + 1]]);
            let rl = u16::from_le_bytes([bytes[pos + 2], bytes[pos + 3]]) as usize;
            if rt == R_EOF && !seen_globals_eof {
                seen_globals_eof = true;
            } else if rt == R_DIMENSIONS && seen_globals_eof {
                let body = &bytes[pos + 4..pos + 4 + rl];
                let rw_mic = u32::from_le_bytes([body[0], body[1], body[2], body[3]]);
                let rw_mac = u32::from_le_bytes([body[4], body[5], body[6], body[7]]);
                let col_mic = u16::from_le_bytes([body[8], body[9]]);
                let col_mac = u16::from_le_bytes([body[10], body[11]]);
                assert_eq!(rw_mic, 2);
                assert_eq!(rw_mac, 5);
                assert_eq!(col_mic, 1);
                assert_eq!(col_mac, 7);
                found = true;
                break;
            }
            pos += 4 + rl;
        }
        assert!(found, "DIMENSIONS record not found in sheet substream");
    }

    #[test]
    fn formula_record_has_cached_value_and_rgce() {
        let cached = FormulaCachedValue::Number(123.5);
        let rgce = placeholder_rgce(&cached, None);
        let body = build_formula(0, 0, 15, &cached, &rgce);
        // FORMULA body layout: rw(2) col(2) ixfe(2) val(8) grbit(2) chn(4) cce(2) rgce(N)
        assert_eq!(body.len(), 22 + rgce.len());
        let val_bytes: [u8; 8] = body[6..14].try_into().unwrap();
        assert_eq!(f64::from_le_bytes(val_bytes), 123.5);
        let cce = u16::from_le_bytes([body[20], body[21]]);
        assert_eq!(cce as usize, rgce.len());
    }

    #[test]
    fn boolerr_record_distinguishes_bool_from_error() {
        let b = build_boolerr(1, 1, 15, 1, false);
        assert_eq!(b[7], 0); // is_error flag = 0 (boolean)
        assert_eq!(b[6], 1); // value = TRUE

        let e = build_boolerr(1, 2, 15, 0x07, true);
        assert_eq!(e[7], 1); // is_error flag = 1
        assert_eq!(e[6], 0x07); // #DIV/0!
    }

    #[test]
    fn sst_cst_total_counts_labelsst_refs() {
        let mut model = ironcalc::base::Model::new_empty("t", "en", "UTC", "en").unwrap();
        // Three string cells, two distinct values.
        model.set_user_input(0, 1, 1, "foo".to_string()).unwrap();
        model.set_user_input(0, 2, 1, "foo".to_string()).unwrap();
        model.set_user_input(0, 3, 1, "bar".to_string()).unwrap();
        let bytes = build_workbook_stream(&model, None, None, None);

        // Find the SST record and read cstTotal / cstUnique.
        let mut pos = 0usize;
        while pos + 4 <= bytes.len() {
            let rt = u16::from_le_bytes([bytes[pos], bytes[pos + 1]]);
            let rl = u16::from_le_bytes([bytes[pos + 2], bytes[pos + 3]]) as usize;
            if rt == R_SST {
                let body = &bytes[pos + 4..pos + 4 + rl];
                let cst_total = u32::from_le_bytes([body[0], body[1], body[2], body[3]]);
                let cst_unique = u32::from_le_bytes([body[4], body[5], body[6], body[7]]);
                assert_eq!(cst_total, 3, "3 LABELSST refs across the sheet");
                assert_eq!(cst_unique, 2, "2 distinct strings in the SST");
                return;
            }
            pos += 4 + rl;
        }
        panic!("SST record not found");
    }

    #[test]
    fn formula_with_r1c1_relative_refs_roundtrips() {
        // R1C1-style averaging formula at sheet[1] r25c3:
        //   (R[-10]C[0] + R[-12]C[0] + R[-15]C[0] + R[-17]C[0] + R[-19]C[0]) / 5
        // Refs are R1C1-relative (negative row offsets); the parsed
        // Node stores `row` as the offset directly, and our encoder
        // adds anchor.row to recover the absolute target.
        use crate::xls_load::load_xls;
        let mut model = ironcalc::base::Model::new_empty("rt", "en", "UTC", "en").unwrap();
        // Anchor the formula at C25; refs go back -10..-19 rows in
        // column C. Plant the source values at C6, C8, C10, C13, C15.
        model.set_user_input(0, 6, 3, "10".into()).unwrap();
        model.set_user_input(0, 8, 3, "20".into()).unwrap();
        model.set_user_input(0, 10, 3, "30".into()).unwrap();
        model.set_user_input(0, 13, 3, "40".into()).unwrap();
        model.set_user_input(0, 15, 3, "50".into()).unwrap();
        // Set the formula via A1 notation — IronCalc parses both forms
        // identically (parsed_formulas stores the same Node tree).
        model
            .set_user_input(0, 25, 3, "=(C15+C13+C10+C8+C6)/5".into())
            .unwrap();
        model.evaluate();
        let v_orig = model.get_formatted_cell_value(0, 25, 3).unwrap();
        assert_eq!(v_orig, "30", "expected (10+20+30+40+50)/5 = 30");

        let path = std::env::temp_dir().join("fastsheet_xls_save_r1c1.xls");
        save_xls(&model, &path).expect("save_xls");
        let (mut reloaded, _, _, _) = load_xls(&path.to_string_lossy()).expect("load_xls");
        reloaded.evaluate();
        let v_rt = reloaded.get_formatted_cell_value(0, 25, 3).unwrap();
        assert_eq!(v_rt, "30", "round-trip preserves the divide");
        std::fs::remove_file(&path).ok();
    }

    #[test]
    fn formula_with_defined_name_roundtrips() {
        // =VLOOKUP("a", LookupTable, 2, FALSE) where LookupTable
        // is a defined name pointing at A1:B3. Tests NAME (Lbl)
        // record emit + PtgName encoding + lookup-by-ilbl path.
        use crate::xls_load::load_xls;
        let mut model = ironcalc::base::Model::new_empty("rt", "en", "UTC", "en").unwrap();
        // Plant a 3-row lookup table at A1:B3.
        model.set_user_input(0, 1, 1, "a".into()).unwrap();
        model.set_user_input(0, 1, 2, "10".into()).unwrap();
        model.set_user_input(0, 2, 1, "b".into()).unwrap();
        model.set_user_input(0, 2, 2, "20".into()).unwrap();
        model.set_user_input(0, 3, 1, "c".into()).unwrap();
        model.set_user_input(0, 3, 2, "30".into()).unwrap();
        model.new_defined_name("LookupTable", None, "Sheet1!$A$1:$B$3").unwrap();
        model.set_user_input(0, 5, 1, r#"=VLOOKUP("b", LookupTable, 2, FALSE)"#.into()).unwrap();
        model.evaluate();
        assert_eq!(model.get_formatted_cell_value(0, 5, 1).unwrap(), "20");

        let path = std::env::temp_dir().join("fastsheet_xls_save_dn.xls");
        save_xls(&model, &path).expect("save_xls");
        let (mut reloaded, _, _, _) = load_xls(&path.to_string_lossy()).expect("load_xls");
        reloaded.evaluate();
        assert_eq!(
            reloaded.get_formatted_cell_value(0, 5, 1).unwrap(),
            "20",
            "defined-name-backed VLOOKUP should round-trip"
        );
        std::fs::remove_file(&path).ok();
    }

    #[test]
    fn string_with_embedded_quote_roundtrips() {
        // User reported: Main!C123 (a CellFormulaString resolving via
        // INDEX/INDIRECT lookup) shows 6"" headboard after round-trip
        // when the source has 6" headboard. Doubling is the symptom
        // of a quote-escape bug somewhere in our save+reload.
        // Pin both the literal-string and the lookup-via-formula
        // pathways.
        use crate::xls_load::load_xls;
        let mut model = ironcalc::base::Model::new_empty("rt", "en", "UTC", "en").unwrap();
        // Plant a string with literal " in a cell.
        model.set_user_input(0, 1, 1, r#"6" headboard"#.into()).unwrap();
        // A formula-string cell that copies it.
        model.set_user_input(0, 2, 1, "=A1".into()).unwrap();
        // A formula that builds the value via concat with a literal-
        // quote string in the formula source.
        model.set_user_input(0, 3, 1, r#"="6"" headboard""#.into()).unwrap();
        model.evaluate();
        assert_eq!(model.get_formatted_cell_value(0, 1, 1).unwrap(), r#"6" headboard"#);
        assert_eq!(model.get_formatted_cell_value(0, 2, 1).unwrap(), r#"6" headboard"#);
        assert_eq!(model.get_formatted_cell_value(0, 3, 1).unwrap(), r#"6" headboard"#);

        let path = std::env::temp_dir().join("fastsheet_xls_save_quote.xls");
        save_xls(&model, &path).expect("save_xls");
        let (mut reloaded, _, _, _) = load_xls(&path.to_string_lossy()).expect("load_xls");
        reloaded.evaluate();
        assert_eq!(
            reloaded.get_formatted_cell_value(0, 1, 1).unwrap(),
            r#"6" headboard"#,
            "SharedString literal with embedded quote"
        );
        assert_eq!(
            reloaded.get_formatted_cell_value(0, 2, 1).unwrap(),
            r#"6" headboard"#,
            "=A1 cell mirrors A1"
        );
        assert_eq!(
            reloaded.get_formatted_cell_value(0, 3, 1).unwrap(),
            r#"6" headboard"#,
            "literal-quote in formula source"
        );
        std::fs::remove_file(&path).ok();
    }

    #[test]
    fn formula_with_iferror_roundtrips() {
        // Regression: user reported IFERROR turning into VARPA on
        // round-trip. Cause was iftab table mismatch — calamine
        // reads iftab=481 as IFERROR but the writer emitted iftab=365
        // (which calamine reads as VARPA). This test pins the
        // alignment.
        use crate::xls_load::load_xls;
        let mut model = ironcalc::base::Model::new_empty("rt", "en", "UTC", "en").unwrap();
        model.set_user_input(0, 1, 1, "=IFERROR(1/0, 99)".into()).unwrap();
        model.evaluate();
        assert_eq!(model.get_formatted_cell_value(0, 1, 1).unwrap(), "99");

        let path = std::env::temp_dir().join("fastsheet_xls_save_iferror.xls");
        save_xls(&model, &path).expect("save_xls");
        let (mut reloaded, _, _, _) = load_xls(&path.to_string_lossy()).expect("load_xls");
        reloaded.evaluate();
        assert_eq!(
            reloaded.get_formatted_cell_value(0, 1, 1).unwrap(),
            "99",
            "IFERROR(1/0, 99) should still return 99 after round-trip"
        );
        std::fs::remove_file(&path).ok();
    }

    #[test]
    fn formula_with_3d_ref_roundtrips() {
        // Cross-sheet ref: =Sheet2!A1 + Sheet2!B1.
        // Tests EXTERNSHEET / SUPBOOK emit + PtgRef3d encoding via
        // the xti table.
        use crate::xls_load::load_xls;
        let mut model = ironcalc::base::Model::new_empty("rt", "en", "UTC", "en").unwrap();
        model.add_sheet("Sheet2").unwrap();
        // Plant data on Sheet2.
        model.set_user_input(1, 1, 1, "10".into()).unwrap();
        model.set_user_input(1, 1, 2, "20".into()).unwrap();
        // Formula on Sheet1 referencing Sheet2.
        model.set_user_input(0, 1, 1, "=Sheet2!A1+Sheet2!B1".into()).unwrap();
        model.evaluate();
        assert_eq!(model.get_formatted_cell_value(0, 1, 1).unwrap(), "30");

        let path = std::env::temp_dir().join("fastsheet_xls_save_3d.xls");
        save_xls(&model, &path).expect("save_xls");
        let (mut reloaded, _, _, _) = load_xls(&path.to_string_lossy()).expect("load_xls");
        reloaded.evaluate();
        assert_eq!(
            reloaded.get_formatted_cell_value(0, 1, 1).unwrap(),
            "30",
            "3D refs should round-trip via PtgRef3d + EXTERNSHEET"
        );
        std::fs::remove_file(&path).ok();
    }

    #[test]
    fn formula_with_divide_roundtrips() {
        // =SUM(A1:A5)/5 — average of 5 numbers, divided down.
        // Tests: PtgArea, PtgFuncVar(SUM), PtgInt(5), PtgDiv, PtgRef
        // pathway with relative refs that need anchor resolution.
        use crate::xls_load::load_xls;
        let mut model = ironcalc::base::Model::new_empty("rt", "en", "UTC", "en").unwrap();
        for i in 1..=5 {
            model.set_user_input(0, i, 1, format!("{}", i * 2)).unwrap();
        }
        model.set_user_input(0, 3, 2, "=SUM(A1:A5)/5".to_string()).unwrap();
        model.evaluate();
        let v_orig = model.get_formatted_cell_value(0, 3, 2).unwrap();
        assert_eq!(v_orig, "6", "sum 2+4+6+8+10 = 30; / 5 = 6");

        let path = std::env::temp_dir().join("fastsheet_xls_save_divide.xls");
        save_xls(&model, &path).expect("save_xls");
        let (mut reloaded, _, _, _) = load_xls(&path.to_string_lossy()).expect("load_xls");
        reloaded.evaluate();
        let v_rt = reloaded.get_formatted_cell_value(0, 3, 2).unwrap();
        assert_eq!(v_rt, "6", "after round-trip the divide should still apply");
        std::fs::remove_file(&path).ok();
    }

    #[test]
    fn end_to_end_value_roundtrip_via_load_xls() {
        // Build a model with a number, a string, and a formula. Save to
        // an .xls. Reload via fastsheet's existing xls loader. Assert
        // that the values come back. (Formula text won't survive until
        // phase 4; the cached value should.)
        use crate::xls_load::load_xls;
        let mut model = ironcalc::base::Model::new_empty("rt", "en", "UTC", "en").unwrap();
        model.set_user_input(0, 1, 1, "42".to_string()).unwrap();
        model.set_user_input(0, 1, 2, "hello".to_string()).unwrap();
        model.set_user_input(0, 2, 1, "=A1*2".to_string()).unwrap();
        model.evaluate();

        let dir = std::env::temp_dir();
        let path = dir.join("fastsheet_xls_save_e2e.xls");
        save_xls(&model, &path).expect("save_xls");

        // Reload through our own xls reader.
        let path_str = path.to_string_lossy().into_owned();
        let (mut reloaded, _hidden, _preserved, _prgce) = load_xls(&path_str).expect("load_xls");
        reloaded.evaluate();

        // Pull cells back. IronCalc's get_formatted_cell_value gives us
        // the rendered string per cell.
        let a1 = reloaded.get_formatted_cell_value(0, 1, 1).unwrap_or_default();
        let b1 = reloaded.get_formatted_cell_value(0, 1, 2).unwrap_or_default();
        let a2 = reloaded.get_formatted_cell_value(0, 2, 1).unwrap_or_default();
        assert_eq!(a1, "42", "A1 number value should survive round-trip");
        assert_eq!(b1, "hello", "B1 shared-string value should survive");
        assert_eq!(a2, "84", "A2 formula's cached value should survive");

        std::fs::remove_file(&path).ok();
    }

    #[test]
    fn saved_xls_is_readable_by_our_biff_scanner() {
        // The strongest validation we have without firing up Excel:
        // round-trip the saved file through xls_biff::scan_xls_shape.
        // It should detect our 1 sheet and not panic on any record.
        use crate::xls_biff::scan_xls_shape;
        let mut model = ironcalc::base::Model::new_empty("rt", "en", "UTC", "en").unwrap();
        model.set_user_input(0, 1, 1, "hello".to_string()).unwrap();
        let dir = std::env::temp_dir();
        let path = dir.join("fastsheet_xls_save_scanner_rt.xls");
        save_xls(&model, &path).expect("save_xls");
        let bytes = std::fs::read(&path).expect("read saved xls");
        let shape = scan_xls_shape(&bytes);
        // The reader doesn't expose sheet count directly; the SST and
        // BOUNDSHEET8 paths it walks shouldn't panic. As a soft signal,
        // a non-default palette or any populated map indicates it
        // parsed at least past the globals substream — but with our
        // stub-only file all maps are empty, so the test simply
        // succeeds if the call returned without panicking.
        let _ = shape;
        std::fs::remove_file(&path).ok();
    }

    #[test]
    fn sst_record_emits_with_correct_header() {
        let mut w = BiffWriter::new();
        let strings = vec!["foo".to_string(), "bar".to_string()];
        emit_sst(&mut w, &strings, 5);
        let bytes = w.into_bytes();
        let recs = read_records(&bytes);
        // First record should be SST.
        assert_eq!(recs[0].0, R_SST);
        // SST body: cstTotal=5, cstUnique=2, then 2 string entries.
        let body_start = 4; // record header is 4 bytes
        let cst_total = u32::from_le_bytes([
            bytes[body_start], bytes[body_start + 1],
            bytes[body_start + 2], bytes[body_start + 3],
        ]);
        let cst_unique = u32::from_le_bytes([
            bytes[body_start + 4], bytes[body_start + 5],
            bytes[body_start + 6], bytes[body_start + 7],
        ]);
        assert_eq!(cst_total, 5);
        assert_eq!(cst_unique, 2);
        // EXTSST follows.
        assert_eq!(recs[1].0, R_EXTSST);
    }

    #[test]
    fn xl_unicode_string_compressed_for_ascii() {
        let mut body: Vec<u8> = Vec::new();
        body.put_xl_unicode_string("Sheet1");
        // u16 cch=6 LE + u8 flag=0 + 6 ASCII bytes
        assert_eq!(body, vec![6, 0, 0, b'S', b'h', b'e', b'e', b't', b'1']);
    }

    #[test]
    fn xl_unicode_string_uncompressed_for_unicode() {
        let mut body: Vec<u8> = Vec::new();
        body.put_xl_unicode_string("Σ"); // U+03A3
        // u16 cch=1 + u8 flag=1 + u16 char
        assert_eq!(body, vec![1, 0, 1, 0xA3, 0x03]);
    }

    /// Walk the workbook stream and return every record body whose
    /// type matches `rec_type`. Used by the frozen-pane / hidden-col
    /// tests below to assert on emitted bytes.
    fn collect_bodies(bytes: &[u8], rec_type: u16) -> Vec<&[u8]> {
        let mut out = Vec::new();
        let mut pos = 0usize;
        while pos + 4 <= bytes.len() {
            let rt = u16::from_le_bytes([bytes[pos], bytes[pos + 1]]);
            let rl = u16::from_le_bytes([bytes[pos + 2], bytes[pos + 3]]) as usize;
            if rt == rec_type {
                out.push(&bytes[pos + 4..pos + 4 + rl]);
            }
            pos += 4 + rl;
        }
        out
    }

    #[test]
    fn window2_default_has_no_frozen_bit() {
        let body = build_window2(0, 0, true);
        let grbit = u16::from_le_bytes([body[0], body[1]]);
        assert_eq!(grbit & 0x0008, 0, "fFrozen must be clear when no panes");
        assert_eq!(grbit & 0x0100, 0, "fFrozenNoSplit must be clear when no panes");
    }

    #[test]
    fn window2_sets_frozen_bits_when_rows_frozen() {
        let body = build_window2(2, 0, true);
        let grbit = u16::from_le_bytes([body[0], body[1]]);
        assert_ne!(grbit & 0x0008, 0, "fFrozen must be set");
        assert_ne!(grbit & 0x0100, 0, "fFrozenNoSplit must be set");
    }

    #[test]
    fn window2_selects_only_the_active_sheet() {
        let active = build_window2(0, 0, true);
        let inactive = build_window2(0, 0, false);
        let g_active = u16::from_le_bytes([active[0], active[1]]);
        let g_inactive = u16::from_le_bytes([inactive[0], inactive[1]]);
        assert_eq!(g_active & 0x0600, 0x0600, "active sheet: fSelected+fPaged");
        assert_eq!(g_inactive & 0x0600, 0, "inactive sheet must not be selected (Group mode)");
    }

    #[test]
    fn preserved_rgce_self_containment_gate() {
        // 1+2 → PtgInt PtgInt PtgAdd: self-contained.
        let simple = [0x1Eu8, 1, 0, 0x1E, 2, 0, 0x03];
        assert!(rgce_is_self_contained(&simple));
        // SUM(A1) via PtgRef + PtgFuncVar: self-contained.
        let func = [0x44u8, 0, 0, 0, 0x40, 0x22, 1, 4, 0];
        assert!(rgce_is_self_contained(&func));
        // PtgRef3d references the source EXTERNSHEET — not replayable.
        let ref3d = [0x3Au8, 5, 0, 0, 0, 0, 0x40];
        assert!(!rgce_is_self_contained(&ref3d));
        // PtgName references the source NAME table — not replayable.
        let name = [0x23u8, 0x0A, 0x06, 0, 0];
        assert!(!rgce_is_self_contained(&name));
        // Truncated PtgStr must not read past the end.
        let truncated = [0x17u8, 10, 0, b'a'];
        assert!(!rgce_is_self_contained(&truncated));
    }

    #[test]
    fn pane_encodes_frozen_split_position() {
        let body = build_pane(3, 5);
        // [MS-XLS] 2.4.213: PANE is exactly 10 bytes — pnnAcc is u16.
        // A 9-byte record makes Apache POI reject the whole file.
        assert_eq!(body.len(), 10, "PANE record must be 10 bytes");
        let x = u16::from_le_bytes([body[0], body[1]]);
        let y = u16::from_le_bytes([body[2], body[3]]);
        let rw_top = u16::from_le_bytes([body[4], body[5]]);
        let col_left = u16::from_le_bytes([body[6], body[7]]);
        let pnn_acc = u16::from_le_bytes([body[8], body[9]]);
        assert_eq!(x, 5, "x = frozen_columns");
        assert_eq!(y, 3, "y = frozen_rows");
        assert_eq!(rw_top, 3);
        assert_eq!(col_left, 5);
        // Both panes frozen: data area is lower-right, pnnAcc=0.
        assert_eq!(pnn_acc, 0);
    }

    #[test]
    fn pane_active_pane_is_lower_left_when_only_rows_frozen() {
        let body = build_pane(1, 0);
        assert_eq!(body.len(), 10, "PANE record must be 10 bytes");
        let pnn_acc = u16::from_le_bytes([body[8], body[9]]);
        assert_eq!(pnn_acc, 2, "rows-only frozen → lower-left active");
    }

    #[test]
    fn frozen_pane_round_trip_emits_window2_and_pane_in_correct_order() {
        let mut model = ironcalc::base::Model::new_empty("t", "en", "UTC", "en").unwrap();
        model.workbook.worksheets[0].frozen_rows = 1;
        model.workbook.worksheets[0].frozen_columns = 2;
        let bytes = build_workbook_stream(&model, None, None, None);

        // Locate WINDOW2 + PANE positions (sheet substream only — find
        // them after the globals EOF).
        let mut pos = 0usize;
        let mut globals_eof_passed = false;
        let mut window2_pos = None;
        let mut pane_pos = None;
        while pos + 4 <= bytes.len() {
            let rt = u16::from_le_bytes([bytes[pos], bytes[pos + 1]]);
            let rl = u16::from_le_bytes([bytes[pos + 2], bytes[pos + 3]]) as usize;
            if rt == R_EOF && !globals_eof_passed {
                globals_eof_passed = true;
            } else if globals_eof_passed && rt == R_WINDOW2 && window2_pos.is_none() {
                window2_pos = Some(pos);
            } else if globals_eof_passed && rt == R_PANE && pane_pos.is_none() {
                pane_pos = Some(pos);
            }
            pos += 4 + rl;
        }
        let window2_pos = window2_pos.expect("WINDOW2 missing in sheet substream");
        let pane_pos = pane_pos.expect("PANE missing — should be emitted when frozen");
        assert!(pane_pos > window2_pos, "PANE must follow WINDOW2");

        // WINDOW2 grbit must have fFrozen set; PANE x/y must match the model.
        let w2_body = &bytes[window2_pos + 4..window2_pos + 4 + 18];
        let grbit = u16::from_le_bytes([w2_body[0], w2_body[1]]);
        assert_ne!(grbit & 0x0008, 0);
        let pane_body = &bytes[pane_pos + 4..pane_pos + 4 + 9];
        let x = u16::from_le_bytes([pane_body[0], pane_body[1]]);
        let y = u16::from_le_bytes([pane_body[2], pane_body[3]]);
        assert_eq!(x, 2, "x = frozen_columns");
        assert_eq!(y, 1, "y = frozen_rows");
    }

    #[test]
    fn frozen_pane_omitted_when_no_freeze() {
        let model = ironcalc::base::Model::new_empty("t", "en", "UTC", "en").unwrap();
        let bytes = build_workbook_stream(&model, None, None, None);
        let panes = collect_bodies(&bytes, R_PANE);
        assert!(panes.is_empty(), "PANE must not be emitted without frozen panes");
    }

    #[test]
    fn colinfo_hidden_bit_set_for_hidden_col() {
        let mut model = ironcalc::base::Model::new_empty("t", "en", "UTC", "en").unwrap();
        // Explicit Col entry covering col 5 with custom width — without
        // hidden_cols the COLINFO grbit must be 0; with col 5 in
        // hidden_cols, the same range must come back with grbit bit 0.
        let _ = model.set_column_width(0, 5, 12.0 * 12.0); // 12 chars after factor
        let mut hidden = HashMap::new();
        let mut set = HashSet::new();
        set.insert(5);
        hidden.insert(0u32, set);
        let bytes = build_workbook_stream(&model, Some(&hidden), None, None);
        let bodies = collect_bodies(&bytes, R_COLINFO);
        let mut found_hidden = false;
        for body in bodies {
            let col_first = u16::from_le_bytes([body[0], body[1]]);
            let col_last = u16::from_le_bytes([body[2], body[3]]);
            let grbit = u16::from_le_bytes([body[8], body[9]]);
            if col_first <= 4 && col_last >= 4 && (grbit & 0x0001) != 0 {
                found_hidden = true;
            }
        }
        assert!(found_hidden, "COLINFO covering col 5 must have hidden bit set");
    }

    #[test]
    fn colinfo_emits_for_hidden_col_with_no_explicit_width() {
        let model = ironcalc::base::Model::new_empty("t", "en", "UTC", "en").unwrap();
        let mut hidden = HashMap::new();
        let mut set = HashSet::new();
        set.insert(3);
        hidden.insert(0u32, set);
        let bytes = build_workbook_stream(&model, Some(&hidden), None, None);
        let bodies = collect_bodies(&bytes, R_COLINFO);
        // At least one COLINFO with the hidden bit covering col 3.
        let mut covered = false;
        for body in bodies {
            let col_first = u16::from_le_bytes([body[0], body[1]]);
            let col_last = u16::from_le_bytes([body[2], body[3]]);
            let grbit = u16::from_le_bytes([body[8], body[9]]);
            if col_first <= 2 && col_last >= 2 && (grbit & 0x0001) != 0 {
                covered = true;
            }
        }
        assert!(covered, "COLINFO must be emitted for hidden cols not in worksheet.cols");
    }

    #[test]
    fn colinfo_coalesces_consecutive_hidden_cols_with_same_attrs() {
        let model = ironcalc::base::Model::new_empty("t", "en", "UTC", "en").unwrap();
        let mut hidden = HashMap::new();
        let mut set = HashSet::new();
        set.insert(2);
        set.insert(3);
        set.insert(4);
        hidden.insert(0u32, set);
        let bytes = build_workbook_stream(&model, Some(&hidden), None, None);
        let bodies = collect_bodies(&bytes, R_COLINFO);
        let mut single_range = None;
        for body in bodies {
            let col_first = u16::from_le_bytes([body[0], body[1]]);
            let col_last = u16::from_le_bytes([body[2], body[3]]);
            let grbit = u16::from_le_bytes([body[8], body[9]]);
            if (grbit & 0x0001) != 0 {
                single_range = Some((col_first, col_last));
            }
        }
        let (first, last) = single_range.expect("expected one hidden range");
        assert_eq!(first, 1, "0-based first col");
        assert_eq!(last, 3, "0-based last col");
    }

    #[test]
    fn colinfo_splits_when_hidden_breaks_existing_range() {
        let mut model = ironcalc::base::Model::new_empty("t", "en", "UTC", "en").unwrap();
        // Set width on cols 1..=5 so they form a single Col entry,
        // then hide just col 3 — the writer must split into three
        // COLINFO records: (1..=2, visible), (3, hidden), (4..=5, visible).
        for col in 1..=5 {
            let _ = model.set_column_width(0, col, 10.0 * 12.0);
        }
        let mut hidden = HashMap::new();
        let mut set = HashSet::new();
        set.insert(3);
        hidden.insert(0u32, set);
        let bytes = build_workbook_stream(&model, Some(&hidden), None, None);
        let bodies = collect_bodies(&bytes, R_COLINFO);
        let ranges: Vec<(u16, u16, u16)> = bodies
            .iter()
            .map(|b| {
                let f = u16::from_le_bytes([b[0], b[1]]);
                let l = u16::from_le_bytes([b[2], b[3]]);
                let g = u16::from_le_bytes([b[8], b[9]]);
                (f, l, g)
            })
            .collect();
        // Expect the col-3 range to come out as a single-col hidden
        // record, with sibling visible ranges on either side.
        let hidden_range = ranges.iter().find(|(_, _, g)| g & 0x0001 != 0);
        let (f, l, _) = hidden_range.expect("missing hidden split").clone();
        assert_eq!(f, 2, "hidden col split first (0-based)");
        assert_eq!(l, 2, "hidden col split last (0-based)");
        // At least one visible range covering the cols around it.
        assert!(
            ranges
                .iter()
                .any(|(f, l, g)| g & 0x0001 == 0 && *f <= 1 && *l >= 1),
            "expected visible range covering cols 1..=2"
        );
        assert!(
            ranges
                .iter()
                .any(|(f, l, g)| g & 0x0001 == 0 && *f <= 3 && *l >= 3),
            "expected visible range covering cols 4..=5"
        );
    }

    #[test]
    fn write_record_continues_oversized_payload() {
        let mut w = BiffWriter::new();
        let big = vec![0xABu8; MAX_RECORD_BODY + 100];
        w.write_record(0x9999, &big);
        let bytes = w.into_bytes();
        let recs = read_records(&bytes);
        assert_eq!(recs.len(), 2);
        assert_eq!(recs[0].0, 0x9999);
        assert_eq!(recs[0].1, MAX_RECORD_BODY);
        assert_eq!(recs[1].0, R_CONTINUE);
        assert_eq!(recs[1].1, 100);
    }

    /// Extract the rgce of the FORMULA record at (rw0, col0) — 0-based
    /// — anywhere in the stream. Cell coordinates are unique enough
    /// for these single-sheet-of-interest tests.
    fn formula_rgce_at(buf: &[u8], rw0: u16, col0: u16) -> Option<Vec<u8>> {
        let mut pos = 0;
        while pos + 4 <= buf.len() {
            let rt = u16::from_le_bytes([buf[pos], buf[pos + 1]]);
            let rl = u16::from_le_bytes([buf[pos + 2], buf[pos + 3]]) as usize;
            let body = &buf[pos + 4..pos + 4 + rl];
            if rt == R_FORMULA && body.len() >= 22 {
                let rw = u16::from_le_bytes([body[0], body[1]]);
                let col = u16::from_le_bytes([body[2], body[3]]);
                if rw == rw0 && col == col0 {
                    let cce = u16::from_le_bytes([body[20], body[21]]) as usize;
                    return Some(body[22..22 + cce].to_vec());
                }
            }
            pos += 4 + rl;
        }
        None
    }

    /// Operand-class pins. Excel's recalc engine turns ref-class
    /// PtgName / Ptg*3d in value contexts into #VALUE! on every
    /// dependent (2026-08-14, neutral.xls: proven by flipping the
    /// single class byte 0x3A→0x5A on one cell and watching it
    /// recover). These assertions pin the class byte the encoder
    /// emits per context.
    #[test]
    fn operand_classes_match_excel_conventions() {
        let mut m = ironcalc::base::Model::new_empty("t", "en", "UTC", "en").unwrap();
        m.add_sheet("Data").unwrap();
        m.set_user_input(1, 1, 1, "5".into()).unwrap();
        m.new_defined_name("DataCol", None, "Data!$A$1:$A$3").unwrap();
        // A1: cross-sheet ref at root — must be VALUE class (0x5A).
        m.set_user_input(0, 1, 1, "=Data!A1".into()).unwrap();
        // A2: name at root — VALUE class PtgName (0x43).
        m.set_user_input(0, 2, 1, "=DataCol".into()).unwrap();
        // A3: SUM's arg is declared ref class — Area3d stays 0x3B.
        m.set_user_input(0, 3, 1, "=SUM(Data!A1:A3)".into()).unwrap();
        // A4: same-sheet ref as operator operand — VALUE class (0x44).
        m.set_user_input(0, 4, 1, "=A1+1".into()).unwrap();
        m.evaluate();
        let bytes = build_xls_bytes(&m, None);

        let a1 = formula_rgce_at(&bytes, 0, 0).expect("A1 rgce");
        assert_eq!(a1[0], 0x5A, "root cross-sheet ref must be value-class PtgRef3d");
        let a2 = formula_rgce_at(&bytes, 1, 0).expect("A2 rgce");
        assert_eq!(a2[0], 0x43, "root defined name must be value-class PtgName");
        let a3 = formula_rgce_at(&bytes, 2, 0).expect("A3 rgce");
        assert_eq!(a3[0], 0x3B, "SUM's range arg keeps ref-class PtgArea3d");
        assert_eq!(a3[a3.len() - 4], 0x42, "PtgFuncVar carries value class in cell context");
        let a4 = formula_rgce_at(&bytes, 3, 0).expect("A4 rgce");
        assert_eq!(a4[0], 0x44, "operator operand ref must be value class");
    }

    /// Defined-name definitions are reference contexts: their rgce
    /// must keep base/ref-class ptgs (Excel writes temp1 as 0x3B).
    #[test]
    fn defined_name_rgce_stays_reference_class() {
        let mut m = ironcalc::base::Model::new_empty("t", "en", "UTC", "en").unwrap();
        m.add_sheet("Data").unwrap();
        m.new_defined_name("DataCol", None, "Data!$A$1:$A$3").unwrap();
        m.evaluate();
        let bytes = build_xls_bytes(&m, None);
        let mut pos = 0;
        let mut found = false;
        while pos + 4 <= bytes.len() {
            let rt = u16::from_le_bytes([bytes[pos], bytes[pos + 1]]);
            let rl = u16::from_le_bytes([bytes[pos + 2], bytes[pos + 3]]) as usize;
            if rt == R_LBL {
                let body = &bytes[pos + 4..pos + 4 + rl];
                let cch = body[3] as usize;
                let cce = u16::from_le_bytes([body[4], body[5]]) as usize;
                let hi = body[14] & 1 == 1;
                let name_len = if hi { cch * 2 } else { cch };
                let name = &body[15..15 + cch.min(body.len() - 15)];
                if !hi && name == b"DataCol" {
                    let rgce = &body[15 + name_len..15 + name_len + cce];
                    assert_eq!(rgce[0], 0x3B, "name definition must be ref-class PtgArea3d");
                    found = true;
                }
            }
            pos += 4 + rl;
        }
        assert!(found, "DataCol Lbl record not found");
    }
}
