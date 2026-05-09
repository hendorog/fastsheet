//! Minimal BIFF8 record scanner for .xls files.
//!
//! calamine (our xls value/formula loader) parses BIFF records
//! internally but doesn't expose sheet layout metadata — column widths,
//! row heights, hidden rows/cols, etc. Rather than patch calamine or
//! swap in another reader, we re-open the file ourselves via `cfb`,
//! walk the "Workbook" (or "Book") stream's BIFF records, and extract
//! only the shape data we need to match xlsx behaviour.
//!
//! BIFF8 record layout: each record is 2 bytes of type + 2 bytes of
//! length (little-endian) + N bytes of data. Records larger than 8224
//! bytes use one or more trailing CONTINUE records (0x003C); we don't
//! care about any large record here so we ignore continuations.
//!
//! References for record layouts:
//! * [MS-XLS]: Excel (.xls) Binary File Format
//!   https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/
//! * OpenOffice.org's Excel File Format documentation (much more
//!   readable, covers BIFF5/8 in detail).

use std::collections::{HashMap, HashSet};
use std::io::Read;

/// Result of scanning an .xls file for shape metadata. Keys are 0-based
/// sheet indices; inner indices are 1-based row/col (matching what the
/// rest of the codebase uses when talking to IronCalc).
#[derive(Default, Debug)]
pub struct XlsShape {
    /// col_index → width in pixels (display units, already scaled).
    pub col_widths: HashMap<u32, HashMap<i32, f64>>,
    /// Hidden columns per sheet.
    pub hidden_cols: HashMap<u32, HashSet<i32>>,
    /// row_index → height in points.
    pub row_heights: HashMap<u32, HashMap<i32, f64>>,
    /// Hidden rows per sheet.
    pub hidden_rows: HashMap<u32, HashSet<i32>>,
    /// Frozen pane counts per sheet (rows, cols).
    pub frozen_panes: HashMap<u32, (i32, i32)>,
    /// XF records, indexed by XF position in file. Each entry carries
    /// the format (number-format) index into the `formats` table.
    pub xfs: Vec<XfEntry>,
    /// Custom / built-in number-format strings keyed by their format index.
    /// Built-ins (1..=49) are pre-seeded; FORMAT records in the file
    /// overlay custom ones at index >= 164.
    pub formats: HashMap<u16, String>,
    /// cell → XF index (ixfe). Populated from value-carrying records.
    pub cell_xfs: HashMap<(u32, i32, i32), u16>,
    /// FONT records in file order. XF.ifnt indexes into this.
    pub fonts: Vec<FontEntry>,
    /// Palette color indices → "#RRGGBB" strings. Seeded with the
    /// default 56-entry palette; PALETTE records in the file overlay.
    pub palette: HashMap<u16, String>,
    /// Array-formula spill ranges per sheet: each entry is
    /// (sheet, r1, r2, c1, c2) in 1-based indices. The anchor cell
    /// is (r1, c1); spill cells are everything else in the rectangle.
    pub array_ranges: Vec<(u32, i32, i32, i32, i32)>,
    /// Raw rgce bytes of SHARED formulas, keyed by the anchor cell
    /// referenced by each cell's PtgExp token. The range stored in
    /// the record is the cells that share the formula; each cell in
    /// the range has its own FORMULA record containing PtgExp
    /// pointing to the anchor (the first cell in the range).
    pub shared_formulas: HashMap<(u32, i32, i32), Vec<u8>>,
    /// Raw rgce bytes of each ARRAY formula, keyed by anchor cell.
    /// Calamine's `parse_formula` ignores PtgExp tokens (the anchor
    /// cell's FORMULA record points at an ARRAY record with this
    /// rgce), so without this we never see the `MY*`-style UDF call
    /// that each template uses for array-formula spill. Decoded
    /// separately in `xls_load.rs`.
    pub array_formulas: HashMap<(u32, i32, i32), Vec<u8>>,
    /// Raw rgce bytes per FORMULA cell. Kept for ALL formula cells
    /// so xls_load can run our own ptg decoder and override calamine
    /// whenever calamine has a bug (comparison op swap, PtgExp
    /// shared-formula skip, PtgRef3d column quadrupling, etc.).
    pub ptgexp_cells: HashMap<(u32, i32, i32), Vec<u8>>,
    /// Raw rgce bytes for every FORMULA cell — used only by the
    /// .xls writer to replay an unparsable cell's bytes verbatim
    /// (see xls_load's `preserved_rgce` capture and xls_save's
    /// CellFormulaError emit). Does NOT feed the existing PtgExp
    /// decoder path above; that one stays scoped to PtgExp cells
    /// to avoid behavioural drift on normal formulas (calamine's
    /// rendering remains authoritative for non-PtgExp cells).
    pub formula_rgce: HashMap<(u32, i32, i32), Vec<u8>>,
    /// For each FORMULA cell, the ordered list of comparison
    /// operators (LT, LE, EQ, GE, GT, NE) actually present in the
    /// rgce. Used to correct calamine's GE/GT swap bug (opcodes
    /// 0x0C and 0x0D are swapped in its output).
    pub formula_cmp_ops: HashMap<(u32, i32, i32), Vec<&'static str>>,
    /// For each FORMULA cell, the ordered list of refs (2D + 3D)
    /// extracted from its rgce bytes. Each entry is `(is_3d, text)`
    /// where the text is the A1 portion only — e.g. `$C13` or
    /// `J50:J55` or `$A$1:$B$2`. These replace the corresponding
    /// ref segments in calamine's formula output because calamine's
    /// PtgRef/PtgArea (both 2D and 3D) decoders corrupt the column
    /// index when any flag bit is set: PtgRef3d shifts instead of
    /// masking; PtgRef / PtgArea / PtgArea3d read raw u16 without
    /// masking `& 0x3FFF`. Result on a relative 2D ref like column
    /// J (0x0009) with both flags: output comes out as "$USV$"
    /// (raw 0xC009 pushed through an unmasked column formatter),
    /// mangling thousands of cells.
    pub formula_refs: HashMap<(u32, i32, i32), Vec<(bool, String)>>,
    /// For each FORMULA cell, the ordered list of PtgStr string
    /// literal CONTENTS (raw chars, no surrounding quotes). Used to
    /// correct calamine's string emitter (xls.rs ~line 1265): it
    /// wraps the raw bytes with a leading and trailing `"` but never
    /// doubles an embedded `"` character. A string containing a `"`
    /// like `"`  ` Hexes` therefore comes out as `"" Hexes"` —
    /// Excel's formula parser reads that as "empty string" followed
    /// by a syntax error. By storing the original unescaped bytes we
    /// can re-encode with proper doubling when patching the text.
    pub formula_strings: HashMap<(u32, i32, i32), Vec<String>>,
    /// For each FORMULA cell, the ordered list of PtgNameX
    /// `nameindex` values (1-based into `extern_names`). Calamine
    /// emits a literal `[PtgNameX]` placeholder for these; we use
    /// the indices here to substitute the real function name when
    /// post-processing the formula text. Common case: Analysis
    /// ToolPak functions like MROUND on .xls files come through as
    /// `User([PtgNameX], a, b)` — we rewrite to `MROUND(a, b)`.
    pub formula_name_xs: HashMap<(u32, i32, i32), Vec<u16>>,
    /// Sheet names as they appear in the BIFF BOUNDSHEET8 stream,
    /// indexed by their position in that stream (0-based). This is
    /// the ordering the xti table (EXTERNSHEET) references. It may
    /// or may not match the order `model.workbook.worksheets` ends
    /// up in, so we keep both.
    pub biff_sheet_names: Vec<String>,
    /// Parsed EXTERNSHEET table — maps ixti (index into xti table)
    /// to the sheet index in `biff_sheet_names`. Calamine's
    /// `parse_formula` for PtgArea3d (xls.rs:1187) bypasses this
    /// table and uses ixti directly as a sheet index, which is
    /// wrong; we override with the correct sheet via our own
    /// resolution when patching formulas.
    pub xti_table: Vec<XtiEntry>,
    /// External name records — built-in / add-in function names like
    /// `MROUND`, `CEILING`, `FLOOR`, etc. that older xls files
    /// reference via PtgNameX. Indexed in registration order
    /// (PtgNameX's nameindex is 1-based into this list). Calamine
    /// emits a literal `[PtgNameX]` placeholder for these (xls.rs
    /// line ~1444), so when we see `User([PtgNameX], …)` in
    /// calamine's text we resolve via this list to produce the
    /// proper `MROUND(...)` call.
    pub extern_names: Vec<String>,
    /// For each defined name, the ordered list of 3D refs decoded
    /// correctly from the BIFF NAME (`Lbl`, 0x0218) record's rgce
    /// bytes. Replaces the corresponding refs in calamine's
    /// `defined_names()` string output. Calamine's
    /// `parse_defined_names` for PtgArea3d reads colFirst / colLast
    /// from rgce as raw u16 (no `& 0x3FFF` mask) so flag bits, and
    /// sometimes the actual column index bits that should go in the
    /// top nibble, get mis-interpreted — e.g. `$B$2:$AZ$56` (colLast
    /// = 51 = 0x33) decodes to `$B$2:$Z$56` when bit 0x20 is set in
    /// the high byte. Keyed by name in lowercase.
    pub defined_name_refs: HashMap<String, Vec<String>>,
    /// Excel's cached last-evaluated value for each FORMULA cell —
    /// either a number, a boolean, an error code, or a string (strings
    /// arrive separately in a STRING record, so we only capture the
    /// indicator here). Useful as an oracle: if Excel cached `#N/A`,
    /// the error isn't a fastsheet bug; it's what Excel saw last save.
    pub formula_cache: HashMap<(u32, i32, i32), FormulaCache>,
}

/// Decoded Excel cached formula value, per [MS-XLS] 2.5.133 FormulaValue.
/// The 8-byte `val` field in the FORMULA record is interpreted by its
/// last two bytes (`0xFFFF`) and the first byte (type).
#[derive(Debug, Clone)]
pub enum FormulaCache {
    Number(f64),
    Boolean(bool),
    Error(u8),         // 0x07=DIV/0, 0x0F=VALUE, 0x17=REF, 0x1D=NAME, 0x24=NUM, 0x2A=NA
    StringPending,     // separate STRING record follows
    Blank,
}

impl FormulaCache {
    pub fn display(&self) -> String {
        match self {
            FormulaCache::Number(n) => format!("{n}"),
            FormulaCache::Boolean(b) => if *b { "TRUE".into() } else { "FALSE".into() },
            FormulaCache::Error(e) => match *e {
                0x00 => "#NULL!".into(),
                0x07 => "#DIV/0!".into(),
                0x0F => "#VALUE!".into(),
                0x17 => "#REF!".into(),
                0x1D => "#NAME?".into(),
                0x24 => "#NUM!".into(),
                0x2A => "#N/A".into(),
                _ => format!("#ERR_{:02x}", e),
            },
            FormulaCache::StringPending => "(string)".into(),
            FormulaCache::Blank => "".into(),
        }
    }
}

#[derive(Debug, Clone, Default)]
pub struct FontEntry {
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub strike: bool,
    pub size_pt: i32,
    pub color_idx: u16,
    pub name: String,
}

#[derive(Debug, Clone, Copy, Default)]
pub struct XtiEntry {
    pub isup_book: u16,
    pub itab_first: i16,
    pub itab_last: i16,
}

#[derive(Debug, Clone, Copy, Default)]
pub struct XfEntry {
    pub font_idx: u16,
    pub fmt_idx: u16,
    /// Parent XF index (top 12 bits of cell_options). The cell XF
    /// inherits attributes from its parent style XF when the
    /// corresponding fAtr* flag is 0.
    pub ixf_parent: u16,
    /// fStyle bit of cell_options — this XF describes a style itself
    /// rather than a specific cell's formatting.
    pub is_style: bool,
    /// fAtr* flags: each bit tells us whether the corresponding
    /// attribute on this XF should be used (bit=1) or inherited from
    /// the parent (bit=0).
    pub atr_num: bool,
    pub atr_fnt: bool,
    pub atr_alc: bool,
    pub atr_bdr: bool,
    pub atr_pat: bool,
    pub atr_prot: bool,
    /// Fill pattern (0 = none, 1 = solid, others are hatched patterns).
    pub fill_pattern: u16,
    /// Palette index for fill foreground color (used when pattern = solid).
    pub fill_fg: u16,
    /// Palette index for fill background color.
    pub fill_bg: u16,
    /// Horizontal alignment: 0=general, 1=left, 2=center, 3=right,
    /// 4=fill, 5=justify, 6=center-across-selection, 7=distributed.
    pub h_align: u8,
    /// Vertical alignment: 0=top, 1=center, 2=bottom, 3=justify, 4=distributed.
    pub v_align: u8,
    pub wrap: bool,
    /// Border styles (BIFF codes: 0=none, 1=thin, 2=medium, etc.).
    /// We only use "present or not" so any non-zero = border exists.
    pub border_left: u8,
    pub border_right: u8,
    pub border_top: u8,
    pub border_bottom: u8,
}

// BIFF record types we care about.
const R_BOF: u16 = 0x0809;
const R_EOF: u16 = 0x000A;
const R_COLINFO: u16 = 0x007D;
const R_ROW: u16 = 0x0208;
const R_PANE: u16 = 0x0041;
const R_WINDOW2: u16 = 0x023E;
const R_FORMAT: u16 = 0x041E;
const R_XF: u16 = 0x00E0;
// Cell value records carry the cell's ixfe (XF index).
const R_BLANK: u16 = 0x0201;
const R_MULBLANK: u16 = 0x00BE;
const R_NUMBER: u16 = 0x0203;
const R_LABELSST: u16 = 0x00FD;
const R_RK: u16 = 0x027E;
const R_MULRK: u16 = 0x00BD;
const R_FORMULA: u16 = 0x0006;
const R_FONT: u16 = 0x0031;
const R_PALETTE: u16 = 0x0092;
const R_ARRAY: u16 = 0x0221;
// ShrFmla — shared formula record. Multiple cells reference the same
// formula via PtgExp (0x01) tokens; the actual rgce lives here.
const R_SHRFMLA: u16 = 0x04BC;
// BIFF8 Lbl (defined name) record. MS-XLS calls it 0x0218 but POI and
// some older references list it as 0x0018. Handle both.
const R_NAME: u16 = 0x0218;
const R_NAME_ALT: u16 = 0x0018;
// BoundSheet8 (0x0085) — sheet registration in workbook stream order.
const R_BOUNDSHEET8: u16 = 0x0085;
// ExternSheet (0x0017) — xti table for resolving ixti in 3D refs.
const R_EXTERNSHEET: u16 = 0x0017;
// ExternName — name in an external supporting book or add-in. For
// Analysis ToolPak functions like MROUND/CEILING/etc. the .xls
// stores them as PtgNameX referencing an EXTERNNAME entry whose
// body is the function name string. MS-XLS calls it 0x0223 for
// BIFF8, but real-world BIFF8 files often emit 0x0023 (older
// ExternName opcode) — handle both.
const R_EXTERNNAME: u16 = 0x0223;
const R_EXTERNNAME_ALT: u16 = 0x0023;
const R_CONTINUE: u16 = 0x003C;

// BOF substream types.
const DT_SHEET: u16 = 0x0010;

/// Built-in BIFF number format strings (ECMA-376 / MS-XLS appendix).
/// Indexes 5..8, 14..17, 22 are locale-specific in Excel; we use
/// sensible defaults that round-trip through IronCalc's formatter.
fn builtin_format(idx: u16) -> Option<&'static str> {
    Some(match idx {
        0 => "General",
        1 => "0",
        2 => "0.00",
        3 => "#,##0",
        4 => "#,##0.00",
        5 => "\"$\"#,##0_);(\"$\"#,##0)",
        6 => "\"$\"#,##0_);[Red](\"$\"#,##0)",
        7 => "\"$\"#,##0.00_);(\"$\"#,##0.00)",
        8 => "\"$\"#,##0.00_);[Red](\"$\"#,##0.00)",
        9 => "0%",
        10 => "0.00%",
        11 => "0.00E+00",
        12 => "# ?/?",
        13 => "# ??/??",
        14 => "m/d/yyyy",
        15 => "d-mmm-yy",
        16 => "d-mmm",
        17 => "mmm-yy",
        18 => "h:mm AM/PM",
        19 => "h:mm:ss AM/PM",
        20 => "h:mm",
        21 => "h:mm:ss",
        22 => "m/d/yyyy h:mm",
        37 => "#,##0_);(#,##0)",
        38 => "#,##0_);[Red](#,##0)",
        39 => "#,##0.00_);(#,##0.00)",
        40 => "#,##0.00_);[Red](#,##0.00)",
        45 => "mm:ss",
        46 => "[h]:mm:ss",
        47 => "mm:ss.0",
        48 => "##0.0E+0",
        49 => "@",
        _ => return None,
    })
}

/// Parse the given .xls file and return its layout metadata. On any
/// error the result is a partial (possibly empty) XlsShape — this is
/// best-effort; the caller falls back to calamine-only behaviour.
pub fn scan_xls_shape(bytes: &[u8]) -> XlsShape {
    // Take an in-memory copy so cfb's many small seeks land in RAM
    // rather than re-reading the source over a slow filesystem (e.g.
    // the WSL `\\wsl.localhost` share via 9P).
    let cursor = std::io::Cursor::new(bytes.to_vec());
    let Ok(mut cfb) = cfb::CompoundFile::open(cursor) else {
        return XlsShape::default();
    };
    // The workbook BIFF stream lives in one of these entries depending
    // on the Excel version that wrote the file. Try both.
    let stream_name = if cfb.exists("/Workbook") {
        "/Workbook"
    } else if cfb.exists("/Book") {
        "/Book"
    } else {
        return XlsShape::default();
    };
    let Ok(mut stream) = cfb.open_stream(stream_name) else {
        return XlsShape::default();
    };
    let mut buf = Vec::new();
    if stream.read_to_end(&mut buf).is_err() {
        return XlsShape::default();
    }

    let mut shape = XlsShape::default();
    // Pre-seed built-in number formats so cells referencing them resolve
    // without requiring a FORMAT record in the file.
    for idx in 0u16..=49 {
        if let Some(s) = builtin_format(idx) {
            shape.formats.insert(idx, s.to_string());
        }
    }
    seed_default_palette(&mut shape.palette);
    let mut pos = 0usize;
    // Sheet substream index. -1 means we're in the global substream or
    // between substreams; each sheet substream increments this.
    let mut sheet_idx: i32 = -1;
    // Fused WINDOW2 settings per active sheet — PANE only applies when
    // WINDOW2's fFrozen bit is set.
    let mut pending_frozen: HashMap<i32, bool> = HashMap::new();

    while pos + 4 <= buf.len() {
        let rec_type = u16::from_le_bytes([buf[pos], buf[pos + 1]]);
        let rec_len = u16::from_le_bytes([buf[pos + 2], buf[pos + 3]]) as usize;
        let data_start = pos + 4;
        let data_end = data_start.saturating_add(rec_len);
        if data_end > buf.len() {
            break;
        }
        let data = &buf[data_start..data_end];
        match rec_type {
            R_BOF => {
                if data.len() >= 4 {
                    let dt = u16::from_le_bytes([data[2], data[3]]);
                    if dt == DT_SHEET {
                        sheet_idx += 1;
                    }
                }
            }
            R_EOF => {
                // Leaving the current substream. We don't reset sheet_idx
                // here; the next BOF(sheet) increments it.
            }
            R_COLINFO if sheet_idx >= 0 && data.len() >= 12 => {
                // colFirst(2), colLast(2), colWidth(2), ixfe(2), grbit(2), reserved(2)
                let col_first = u16::from_le_bytes([data[0], data[1]]) as i32;
                let col_last = u16::from_le_bytes([data[2], data[3]]) as i32;
                let col_width_1_256ths = u16::from_le_bytes([data[4], data[5]]) as f64;
                let grbit = u16::from_le_bytes([data[8], data[9]]);
                let hidden = (grbit & 0x0001) != 0;
                // Units: 1/256 of the default font's '0' character width.
                // Excel's default is 7 px per char (Calibri 11), so:
                //   px = (width_1_256ths / 256) * 7
                let px = col_width_1_256ths * 7.0 / 256.0;
                let sheet = sheet_idx as u32;
                for col in col_first..=col_last {
                    let col_1based = col + 1;
                    if hidden {
                        shape.hidden_cols.entry(sheet).or_default().insert(col_1based);
                    }
                    if col_width_1_256ths > 0.0 {
                        shape
                            .col_widths
                            .entry(sheet)
                            .or_default()
                            .insert(col_1based, px);
                    }
                }
            }
            R_ROW if sheet_idx >= 0 && data.len() >= 16 => {
                // rw(2), colMic(2), colMac(2), miyRw(2), reserved1(2),
                // reserved2(2), grbit(2), ixfe(2)
                let rw = u16::from_le_bytes([data[0], data[1]]) as i32;
                let miy_rw = u16::from_le_bytes([data[6], data[7]]) as f64;
                let miy_custom = (miy_rw as u16 & 0x8000) != 0; // top bit = custom
                let height_twips = miy_rw as u16 & 0x7FFF; // bottom 15 bits
                let grbit = u16::from_le_bytes([data[12], data[13]]);
                // grbit bits:
                //   0x00000020 = fDyZero (row hidden)
                //   0x00000040 = fUnsynced (explicit height)
                let hidden = (grbit & 0x0020) != 0;
                let height_pts = (height_twips as f64) / 20.0;
                let sheet = sheet_idx as u32;
                let row_1based = rw + 1;
                if hidden {
                    shape.hidden_rows.entry(sheet).or_default().insert(row_1based);
                }
                if miy_custom && height_twips > 0 {
                    shape
                        .row_heights
                        .entry(sheet)
                        .or_default()
                        .insert(row_1based, height_pts);
                }
            }
            R_WINDOW2 if sheet_idx >= 0 && data.len() >= 2 => {
                // grbit (2) - bit 3 (0x0008) = fFrozen
                let grbit = u16::from_le_bytes([data[0], data[1]]);
                let is_frozen = (grbit & 0x0008) != 0;
                pending_frozen.insert(sheet_idx, is_frozen);
            }
            R_PANE if sheet_idx >= 0 && data.len() >= 8 => {
                // Only respect PANE when the sheet's WINDOW2 had fFrozen set
                // — otherwise PANE is a split (draggable) pane, not a
                // frozen-titles feature.
                if pending_frozen.get(&sheet_idx).copied().unwrap_or(false) {
                    // x(2) = num cols in left pane, y(2) = num rows in top pane
                    let frozen_cols = u16::from_le_bytes([data[0], data[1]]) as i32;
                    let frozen_rows = u16::from_le_bytes([data[2], data[3]]) as i32;
                    shape
                        .frozen_panes
                        .insert(sheet_idx as u32, (frozen_rows, frozen_cols));
                }
            }
            // Global substream: FORMAT and XF records arrive before any
            // sheet substream (sheet_idx still -1). Parse regardless.
            R_FORMAT if data.len() >= 3 => {
                // ifmt(2), cch(2), grbit(1), rgch (variable)
                let ifmt = u16::from_le_bytes([data[0], data[1]]);
                let cch = u16::from_le_bytes([data[2], data[3]]) as usize;
                if data.len() >= 5 + cch {
                    let grbit = data[4];
                    let high_byte = (grbit & 0x01) != 0;
                    let name_bytes = &data[5..];
                    if let Some(s) = parse_biff_string(name_bytes, cch, high_byte) {
                        shape.formats.insert(ifmt, s);
                    }
                }
            }
            R_XF if data.len() >= 20 => {
                // ifnt(2), ifmt(2), fFlags(2), align(1), align2(1),
                // indent(2), borderStyles(2), borderColors+diag(4), fill(4)
                let font_idx = u16::from_le_bytes([data[0], data[1]]);
                let fmt_idx = u16::from_le_bytes([data[2], data[3]]);
                let cell_options = u16::from_le_bytes([data[4], data[5]]);
                let is_style = (cell_options & 0x0004) != 0;
                let ixf_parent = (cell_options >> 4) & 0x0FFF;
                let atr = data[9];
                let atr_num = (atr & 0x01) != 0;
                let atr_fnt = (atr & 0x02) != 0;
                let atr_alc = (atr & 0x04) != 0;
                let atr_bdr = (atr & 0x08) != 0;
                let atr_pat = (atr & 0x10) != 0;
                let atr_prot = (atr & 0x20) != 0;
                // alignment byte 6: bits 0-2 h-align, bit 3 wrap,
                // bits 4-6 v-align, bit 7 justLast
                let align = data[6];
                let h_align = align & 0x07;
                let wrap = (align & 0x08) != 0;
                let v_align = (align >> 4) & 0x07;
                // bytes 10-11 (u16): borderStyles — nibble per side
                let border_styles = u16::from_le_bytes([data[10], data[11]]);
                let border_left = (border_styles & 0x000F) as u8;
                let border_right = ((border_styles >> 4) & 0x000F) as u8;
                let border_top = ((border_styles >> 8) & 0x000F) as u8;
                let border_bottom = ((border_styles >> 12) & 0x000F) as u8;
                // Fill encoding (cross-referenced against Apache POI's
                // HSSF ExtendedFormatRecord):
                //   bytes 14..18 (u32):
                //     bits 0-6   icvTop
                //     bits 7-13  icvBottom
                //     bits 14-20 icvDiag
                //     bits 21-24 dgDiag (diagonal border style)
                //     bits 25    fHasXFExt
                //     bits 26-31 fls (fill pattern)   ← mask 0xfc000000
                //   bytes 18..20 (u16):
                //     bits 0-6  icvFore
                //     bits 7-13 icvBack
                // Fill pattern sits in the top byte: (data[17] >> 2) & 0x3F.
                let fill_pattern = ((data[17] as u16) >> 2) & 0x3F;
                let fill_colors = u16::from_le_bytes([data[18], data[19]]);
                let fill_fg = fill_colors & 0x007F;
                let fill_bg = (fill_colors >> 7) & 0x007F;
                shape.xfs.push(XfEntry {
                    font_idx,
                    fmt_idx,
                    ixf_parent,
                    is_style,
                    atr_num,
                    atr_fnt,
                    atr_alc,
                    atr_bdr,
                    atr_pat,
                    atr_prot,
                    fill_pattern,
                    fill_fg,
                    fill_bg,
                    h_align,
                    v_align,
                    wrap,
                    border_left,
                    border_right,
                    border_top,
                    border_bottom,
                });
            }
            R_FONT if data.len() >= 16 => {
                // dyHeight(2), grbit(2), icv(2), bls(2), sss(2), uls(1),
                // bFamily(1), bCharSet(1), reserved(1), cch(1), grbit2(1),
                // rgch (variable BIFF string)
                let dy_height = u16::from_le_bytes([data[0], data[1]]);
                let grbit = u16::from_le_bytes([data[2], data[3]]);
                let icv = u16::from_le_bytes([data[4], data[5]]);
                let bls = u16::from_le_bytes([data[6], data[7]]);
                let uls = data[10];
                let cch = data[14] as usize;
                let high_byte = (data[15] & 0x01) != 0;
                let name = if data.len() > 16 {
                    parse_biff_string(&data[16..], cch, high_byte).unwrap_or_default()
                } else {
                    String::new()
                };
                shape.fonts.push(FontEntry {
                    bold: bls >= 700,
                    italic: (grbit & 0x0002) != 0,
                    underline: uls != 0,
                    strike: (grbit & 0x0008) != 0,
                    size_pt: (dy_height as i32) / 20,
                    color_idx: icv,
                    name,
                });
            }
            R_BOUNDSHEET8 if data.len() >= 8 => {
                // lbPlyPos(4), hsState(1), dt(1), stName (variable short str)
                let cch = data[6] as usize;
                let high_byte = (data[7] & 0x01) != 0;
                let bytes_needed = 8 + cch * if high_byte { 2 } else { 1 };
                if data.len() >= bytes_needed {
                    let name = parse_biff_string(&data[8..], cch, high_byte)
                        .unwrap_or_default();
                    shape.biff_sheet_names.push(name);
                }
            }
            R_EXTERNSHEET if data.len() >= 2 => {
                // cxti(2), rgixti[cxti] (6 bytes each):
                //   iSupBook(2), itabFirst(2), itabLast(2)
                // Large xti tables span CONTINUE (0x003C) records —
                // concatenate their bodies before parsing.
                let cxti = u16::from_le_bytes([data[0], data[1]]) as usize;
                let mut body: Vec<u8> = data[2..].to_vec();
                let mut scan_pos = data_end;
                while scan_pos + 4 <= buf.len() {
                    let nxt_type = u16::from_le_bytes([buf[scan_pos], buf[scan_pos + 1]]);
                    let nxt_len = u16::from_le_bytes([buf[scan_pos + 2], buf[scan_pos + 3]])
                        as usize;
                    if nxt_type != R_CONTINUE { break; }
                    let nxt_start = scan_pos + 4;
                    let nxt_end = nxt_start + nxt_len;
                    if nxt_end > buf.len() { break; }
                    body.extend_from_slice(&buf[nxt_start..nxt_end]);
                    scan_pos = nxt_end;
                }
                let max_entries = body.len() / 6;
                let n = cxti.min(max_entries);
                for i in 0..n {
                    let off = i * 6;
                    let isup_book = u16::from_le_bytes([body[off], body[off + 1]]);
                    let itab_first = i16::from_le_bytes([body[off + 2], body[off + 3]]);
                    let itab_last = i16::from_le_bytes([body[off + 4], body[off + 5]]);
                    shape.xti_table.push(XtiEntry {
                        isup_book,
                        itab_first,
                        itab_last,
                    });
                }
                // Advance past all merged CONTINUE records so the main
                // loop doesn't re-process them.
                pos = scan_pos;
                continue;
            }
            R_EXTERNNAME | R_EXTERNNAME_ALT if data.len() >= 7 => {
                // ExternName per [MS-XLS] 2.4.69:
                //   options(2)  — bit 0 fBuiltin, bit 1 fWantAdvise, etc.
                //   one(2)      — must be 0x0001 for built-in/add-in
                //                 functions in BIFF8
                //   itab(2)     — sheet table index (0 for fn name)
                //   cch(1)      — name length
                //   grbit(1)    — high-bit flag (bit 0)
                //   rgch(cch * (1 or 2))
                // For Analysis ToolPak / add-in function references
                // (MROUND, CEILING.PRECISE, EOMONTH, etc.), the rgch
                // is the function name as plain text.
                let cch = data[6] as usize;
                if data.len() >= 8 {
                    let high = (data[7] & 0x01) != 0;
                    let needed = 8 + cch * if high { 2 } else { 1 };
                    let name = if data.len() >= needed {
                        parse_biff_string(&data[8..], cch, high).unwrap_or_default()
                    } else {
                        String::new()
                    };
                    shape.extern_names.push(name);
                }
            }
            R_NAME | R_NAME_ALT if data.len() >= 14 => {
                // Lbl record per [MS-XLS] 2.4.180:
                //   grbit(2), chKey(1), cch(1), cce(2), ixals(2),
                //   itab(2), cchCustMenu(1), cchDescription(1),
                //   cchHelpTopic(1), cchStatusText(1), rgch (cch chars
                //   prefixed by 1-byte high-byte flag), rgce (cce
                //   bytes formula), then optional extra strings.
                let grbit = u16::from_le_bytes([data[0], data[1]]);
                let cch = data[3] as usize;
                let cce = u16::from_le_bytes([data[4], data[5]]) as usize;
                let is_builtin = (grbit & 0x0020) != 0;
                let name_flag_byte = 14usize;
                if name_flag_byte >= data.len() { pos = data_end; continue; }
                let high_byte_flag = (data[name_flag_byte] & 0x01) != 0;
                let name_bytes_start = name_flag_byte + 1;
                let name_byte_len = if high_byte_flag { cch * 2 } else { cch };
                if name_bytes_start + name_byte_len + cce > data.len() {
                    pos = data_end;
                    continue;
                }
                let name = parse_biff_string(
                    &data[name_bytes_start..],
                    cch,
                    high_byte_flag,
                )
                .unwrap_or_default();
                let rgce_start = name_bytes_start + name_byte_len;
                let rgce = &data[rgce_start..rgce_start + cce];
                if !is_builtin && !name.is_empty() && cce > 0 {
                    let refs = extract_refs_with_xti(
                        rgce,
                        &shape.xti_table,
                        &shape.biff_sheet_names,
                    );
                    // Defined names historically only got 3D refs
                    // (simple cross-sheet ranges). Keep that shape:
                    // emit full sheet-qualified strings. The patch
                    // pass uses these as a 3D list.
                    let three_d_only: Vec<String> = refs
                        .into_iter()
                        .filter_map(|(is_3d, text)| if is_3d { Some(text) } else { None })
                        .collect();
                    if !three_d_only.is_empty() {
                        shape
                            .defined_name_refs
                            .insert(name.to_lowercase(), three_d_only);
                    }
                }
            }
            R_SHRFMLA if sheet_idx >= 0 && data.len() >= 10 => {
                // rwFirst(2), rwLast(2), colFirst(1), colLast(1),
                // unused(1), cFmla(1), reserved(2), cce(2), rgce(cce).
                // The record's cells all share the same rgce; each
                // FORMULA record in the range contains PtgExp
                // (rwFirst, colFirst) pointing here. We key the
                // shared-formula map by that anchor.
                let rw_first = u16::from_le_bytes([data[0], data[1]]) as i32 + 1;
                let col_first = data[4] as i32 + 1;
                let cce_off = 8;
                if data.len() >= cce_off + 2 {
                    let cce = u16::from_le_bytes([
                        data[cce_off], data[cce_off + 1],
                    ]) as usize;
                    let rgce_off = cce_off + 2;
                    if data.len() >= rgce_off + cce && cce > 0 {
                        let rgce = data[rgce_off..rgce_off + cce].to_vec();
                        shape.shared_formulas.insert(
                            (sheet_idx as u32, rw_first, col_first),
                            rgce,
                        );
                    }
                }
            }
            R_ARRAY if sheet_idx >= 0 && data.len() >= 14 => {
                // rwFirst(2), rwLast(2), colFirst(1), colLast(1),
                // grbit(2), chn(4), cce(2), rgce(cce).
                // calamine decodes the per-cell FORMULA records in the
                // spill range (which contain PtgExp pointing here) but
                // does NOT decode the rgce stored here. That means the
                // actual MY* UDF call only shows up if we decode the
                // ARRAY rgce ourselves.
                let rw_first = u16::from_le_bytes([data[0], data[1]]) as i32 + 1;
                let rw_last = u16::from_le_bytes([data[2], data[3]]) as i32 + 1;
                let col_first = data[4] as i32 + 1;
                let col_last = data[5] as i32 + 1;
                shape.array_ranges.push((
                    sheet_idx as u32,
                    rw_first,
                    rw_last,
                    col_first,
                    col_last,
                ));
                let cce = u16::from_le_bytes([data[12], data[13]]) as usize;
                if data.len() >= 14 + cce && cce > 0 {
                    let rgce = data[14..14 + cce].to_vec();
                    shape.array_formulas.insert(
                        (sheet_idx as u32, rw_first, col_first),
                        rgce,
                    );
                }
            }
            R_PALETTE if data.len() >= 2 => {
                // ccv(2), rgColor[ccv] — each is a LongRGB (4 bytes):
                // Red, Green, Blue, reserved (per [MS-XLS] 2.5.165).
                let ccv = u16::from_le_bytes([data[0], data[1]]) as usize;
                if data.len() >= 2 + ccv * 4 {
                    // Palette records overlay indices 8..8+ccv (standard
                    // BIFF convention — 0..7 are reserved built-ins).
                    for i in 0..ccv {
                        let off = 2 + i * 4;
                        let r = data[off];
                        let g = data[off + 1];
                        let b = data[off + 2];
                        let hex = format!("#{:02X}{:02X}{:02X}", r, g, b);
                        shape.palette.insert(8 + i as u16, hex);
                    }
                }
            }
            R_BLANK if sheet_idx >= 0 && data.len() >= 6 => {
                // rw(2), col(2), ixfe(2)
                let rw = u16::from_le_bytes([data[0], data[1]]) as i32 + 1;
                let col = u16::from_le_bytes([data[2], data[3]]) as i32 + 1;
                let ixfe = u16::from_le_bytes([data[4], data[5]]);
                shape.cell_xfs.insert((sheet_idx as u32, rw, col), ixfe);
            }
            R_NUMBER if sheet_idx >= 0 && data.len() >= 14 => {
                // rw(2), col(2), ixfe(2), num(8)
                let rw = u16::from_le_bytes([data[0], data[1]]) as i32 + 1;
                let col = u16::from_le_bytes([data[2], data[3]]) as i32 + 1;
                let ixfe = u16::from_le_bytes([data[4], data[5]]);
                shape.cell_xfs.insert((sheet_idx as u32, rw, col), ixfe);
            }
            R_LABELSST if sheet_idx >= 0 && data.len() >= 10 => {
                // rw(2), col(2), ixfe(2), isst(4)
                let rw = u16::from_le_bytes([data[0], data[1]]) as i32 + 1;
                let col = u16::from_le_bytes([data[2], data[3]]) as i32 + 1;
                let ixfe = u16::from_le_bytes([data[4], data[5]]);
                shape.cell_xfs.insert((sheet_idx as u32, rw, col), ixfe);
            }
            R_RK if sheet_idx >= 0 && data.len() >= 10 => {
                // rw(2), col(2), ixfe(2), rk(4)
                let rw = u16::from_le_bytes([data[0], data[1]]) as i32 + 1;
                let col = u16::from_le_bytes([data[2], data[3]]) as i32 + 1;
                let ixfe = u16::from_le_bytes([data[4], data[5]]);
                shape.cell_xfs.insert((sheet_idx as u32, rw, col), ixfe);
            }
            R_MULRK if sheet_idx >= 0 && data.len() >= 6 => {
                // rw(2), colFirst(2), rkrec[N] (each 6 bytes: ixfe(2) + rk(4)), colLast(2)
                let rw = u16::from_le_bytes([data[0], data[1]]) as i32 + 1;
                let col_first = u16::from_le_bytes([data[2], data[3]]) as i32;
                // Last 2 bytes are colLast; remainder between 4 and len-2 is rkrec array.
                let rkrec_bytes = &data[4..data.len() - 2];
                let n = rkrec_bytes.len() / 6;
                for i in 0..n {
                    let off = i * 6;
                    let ixfe = u16::from_le_bytes([rkrec_bytes[off], rkrec_bytes[off + 1]]);
                    let col = col_first + i as i32 + 1;
                    shape.cell_xfs.insert((sheet_idx as u32, rw, col), ixfe);
                }
            }
            R_MULBLANK if sheet_idx >= 0 && data.len() >= 6 => {
                // rw(2), colFirst(2), ixfe[N] (each 2 bytes), colLast(2)
                let rw = u16::from_le_bytes([data[0], data[1]]) as i32 + 1;
                let col_first = u16::from_le_bytes([data[2], data[3]]) as i32;
                let ixfe_bytes = &data[4..data.len() - 2];
                let n = ixfe_bytes.len() / 2;
                for i in 0..n {
                    let off = i * 2;
                    let ixfe = u16::from_le_bytes([ixfe_bytes[off], ixfe_bytes[off + 1]]);
                    let col = col_first + i as i32 + 1;
                    shape.cell_xfs.insert((sheet_idx as u32, rw, col), ixfe);
                }
            }
            R_FORMULA if sheet_idx >= 0 && data.len() >= 22 => {
                // rw(2), col(2), ixfe(2), cached(8), grbit(2), chn(4), cce(2), rgce(variable)
                let rw = u16::from_le_bytes([data[0], data[1]]) as i32 + 1;
                let col = u16::from_le_bytes([data[2], data[3]]) as i32 + 1;
                let ixfe = u16::from_le_bytes([data[4], data[5]]);
                shape.cell_xfs.insert((sheet_idx as u32, rw, col), ixfe);
                // Decode the 8-byte cached value. Per [MS-XLS] 2.5.133:
                //   if bytes[6..8] == 0xFFFF, the type is in bytes[0]
                //     (0=string, 1=bool, 2=error, 3=blank)
                //   else the 8 bytes are a little-endian f64
                let val_bytes = &data[6..14];
                let fin = u16::from_le_bytes([val_bytes[6], val_bytes[7]]);
                if std::env::var("FASTSHEET_XLS_DEBUG_CACHE").ok().as_deref()
                    == Some(&format!("{}:{}:{}", sheet_idx, rw, col))
                {
                    eprintln!(
                        "CACHE s{} r{} c{} val8={:02x?} fin={:04x} byte0={:02x} byte2={:02x}",
                        sheet_idx, rw, col,
                        val_bytes, fin, val_bytes[0], val_bytes[2],
                    );
                }
                let cache = if fin == 0xFFFF {
                    match val_bytes[0] {
                        0 => FormulaCache::StringPending,
                        1 => FormulaCache::Boolean(val_bytes[2] != 0),
                        2 => FormulaCache::Error(val_bytes[2]),
                        3 => FormulaCache::Blank,
                        _ => FormulaCache::Blank,
                    }
                } else {
                    let mut arr = [0u8; 8];
                    arr.copy_from_slice(val_bytes);
                    FormulaCache::Number(f64::from_le_bytes(arr))
                };
                shape.formula_cache.insert((sheet_idx as u32, rw, col), cache);
                let cce = u16::from_le_bytes([data[20], data[21]]) as usize;
                if data.len() >= 22 + cce && cce > 0 {
                    let rgce = &data[22..22 + cce];
                    if std::env::var("FASTSHEET_XLS_DEBUG_RGCE_CELL").ok().as_deref()
                        == Some(&format!("{}:{}:{}", sheet_idx, rw, col))
                    {
                        eprintln!("RGCE s{} r{} c{}: {:02x?}", sheet_idx, rw, col, rgce);
                    }
                    // Capture rgce for cells that start with PtgExp —
                    // calamine silently skips those.
                    if !rgce.is_empty() && rgce[0] == 0x01 {
                        shape.ptgexp_cells.insert(
                            (sheet_idx as u32, rw, col),
                            rgce.to_vec(),
                        );
                    }
                    // Always cache the raw rgce regardless of the
                    // first ptg, so the .xls writer can replay it
                    // verbatim for cells whose formulas IronCalc
                    // couldn't parse. Used by BUG-01 round-trip
                    // preservation; the existing PtgExp decoder still
                    // reads from `ptgexp_cells` above so its behaviour
                    // doesn't broaden.
                    if !rgce.is_empty() {
                        shape.formula_rgce.insert(
                            (sheet_idx as u32, rw, col),
                            rgce.to_vec(),
                        );
                    }
                    // Scan rgce for comparison ops. Calamine's
                    // decoder has 0x0C (GE) and 0x0D (GT) swapped;
                    // we capture the CORRECT sequence here so
                    // xls_load can patch the formula text.
                    let ops = scan_comparison_ops(rgce);
                    if !ops.is_empty() {
                        shape
                            .formula_cmp_ops
                            .insert((sheet_idx as u32, rw, col), ops);
                    }
                    let refs = extract_refs_with_xti(
                        rgce,
                        &shape.xti_table,
                        &shape.biff_sheet_names,
                    );
                    if !refs.is_empty() {
                        shape
                            .formula_refs
                            .insert((sheet_idx as u32, rw, col), refs);
                    }
                    // Extract PtgStr literals. Calamine emits these
                    // without doubling internal `"` characters (see
                    // xls.rs line ~1265), which silently corrupts any
                    // formula containing a literal quote char. By
                    // capturing the raw contents here we can re-emit
                    // the correct doubled form when patching.
                    let strings = extract_ptg_strings(rgce);
                    if !strings.is_empty() {
                        shape
                            .formula_strings
                            .insert((sheet_idx as u32, rw, col), strings);
                    }
                    let name_xs = extract_ptg_name_x_indices(rgce);
                    if !name_xs.is_empty() {
                        shape
                            .formula_name_xs
                            .insert((sheet_idx as u32, rw, col), name_xs);
                    }
                }
            }
            _ => {}
        }
        pos = data_end;
    }
    shape
}

/// Walk an rgce and return, in order, the raw content of every
/// PtgStr (0x17) token. The content is what Excel stored — internal
/// `"` characters are present unescaped. Callers produce the
/// properly-escaped formula syntax by doubling `"` chars.
pub fn extract_ptg_strings(rgce: &[u8]) -> Vec<String> {
    let mut out = Vec::new();
    let mut i = 0usize;
    while i < rgce.len() {
        let ptg = rgce[i];
        i += 1;
        let advance: usize = match ptg {
            // Binary / unary / paren / missing-arg ptgs — 0 bytes.
            0x03..=0x11 | 0x12..=0x16 => 0,
            0x01 | 0x02 => 4,
            0x1C | 0x1D => 1,
            0x1E => 2,
            0x1F => 8,
            0x17 => {
                let Some(&cch) = rgce.get(i) else { break };
                let Some(&grbit) = rgce.get(i + 1) else { break };
                let high = (grbit & 0x01) != 0;
                let bytes_per_ch = if high { 2 } else { 1 };
                let payload_len = cch as usize * bytes_per_ch;
                let start = i + 2;
                let end = start + payload_len;
                if end > rgce.len() { break; }
                let s = if high {
                    let mut u16s = Vec::with_capacity(cch as usize);
                    for k in 0..cch as usize {
                        let lo = rgce[start + 2 * k];
                        let hi = rgce[start + 2 * k + 1];
                        u16s.push(u16::from_le_bytes([lo, hi]));
                    }
                    String::from_utf16_lossy(&u16s)
                } else {
                    // BIFF8 compressed strings are Windows-1252.
                    rgce[start..end]
                        .iter()
                        .map(|b| *b as char)
                        .collect::<String>()
                };
                out.push(s);
                2 + payload_len
            }
            0x18 => 5,
            0x19 => {
                let Some(&sub) = rgce.get(i) else { break };
                match sub {
                    0x01 | 0x02 | 0x08 | 0x20 | 0x21 | 0x10 | 0x40 | 0x41 => 3,
                    0x04 => {
                        let Some(lo) = rgce.get(i + 1) else { break };
                        let Some(hi) = rgce.get(i + 2) else { break };
                        let n = u16::from_le_bytes([*lo, *hi]) as usize;
                        3 + 2 * (n + 1)
                    }
                    _ => break,
                }
            }
            0x21 | 0x41 | 0x61 => 2,
            0x22 | 0x42 | 0x62 => 3,
            0x20 | 0x40 | 0x60 => 7,
            0x23 | 0x43 | 0x63 => 4,
            0x24 | 0x44 | 0x64 => 4,
            0x25 | 0x45 | 0x65 => 8,
            0x26 | 0x46 | 0x66 => 6,
            0x27 | 0x47 | 0x67 => 6,
            0x28 | 0x48 | 0x68 => 6,
            0x29 | 0x49 | 0x69 => 2,
            0x2A | 0x4A | 0x6A => 4,
            0x2B | 0x4B | 0x6B => 8,
            0x2C | 0x4C | 0x6C => 4,
            0x2D | 0x4D | 0x6D => 8,
            0x39 | 0x59 | 0x79 => 6,
            0x3A | 0x5A | 0x7A => 6,
            0x3B | 0x5B | 0x7B => 10,
            0x3C | 0x5C | 0x7C => 6,
            0x3D | 0x5D | 0x7D => 10,
            _ => break,
        };
        if i + advance > rgce.len() { break; }
        i += advance;
    }
    out
}

/// Walk an rgce and return, in order, the `nameindex` field of every
/// PtgNameX (0x39 / 0x59 / 0x79) token. PtgNameX layout per
/// [MS-XLS] 2.5.198.92 is:
///   ptg(1), ixti(2), nameindex(2), unused(2)  — 7 bytes total.
/// `nameindex` is 1-based into the workbook's `extern_names` list.
pub fn extract_ptg_name_x_indices(rgce: &[u8]) -> Vec<u16> {
    let mut out = Vec::new();
    let mut i = 0usize;
    while i < rgce.len() {
        let ptg = rgce[i];
        i += 1;
        let advance: usize = match ptg {
            0x03..=0x11 | 0x12..=0x16 => 0,
            0x01 | 0x02 => 4,
            0x1C | 0x1D => 1,
            0x1E => 2,
            0x1F => 8,
            0x17 => {
                let Some(&cch) = rgce.get(i) else { break };
                let Some(&grbit) = rgce.get(i + 1) else { break };
                let high = (grbit & 0x01) != 0;
                2 + cch as usize * if high { 2 } else { 1 }
            }
            0x18 => 5,
            0x19 => {
                let Some(&sub) = rgce.get(i) else { break };
                match sub {
                    0x01 | 0x02 | 0x08 | 0x20 | 0x21 | 0x10 | 0x40 | 0x41 => 3,
                    0x04 => {
                        let Some(lo) = rgce.get(i + 1) else { break };
                        let Some(hi) = rgce.get(i + 2) else { break };
                        let n = u16::from_le_bytes([*lo, *hi]) as usize;
                        3 + 2 * (n + 1)
                    }
                    _ => break,
                }
            }
            0x21 | 0x41 | 0x61 => 2,
            0x22 | 0x42 | 0x62 => 3,
            0x20 | 0x40 | 0x60 => 7,
            0x23 | 0x43 | 0x63 => 4,
            0x24 | 0x44 | 0x64 => 4,
            0x25 | 0x45 | 0x65 => 8,
            0x26 | 0x46 | 0x66 => 6,
            0x27 | 0x47 | 0x67 => 6,
            0x28 | 0x48 | 0x68 => 6,
            0x29 | 0x49 | 0x69 => 2,
            0x2A | 0x4A | 0x6A => 4,
            0x2B | 0x4B | 0x6B => 8,
            0x2C | 0x4C | 0x6C => 4,
            0x2D | 0x4D | 0x6D => 8,
            0x39 | 0x59 | 0x79 => {
                // ixti(2), nameindex(2), unused(2)
                if i + 6 > rgce.len() { break; }
                let nameindex = u16::from_le_bytes([rgce[i + 2], rgce[i + 3]]);
                out.push(nameindex);
                6
            }
            0x3A | 0x5A | 0x7A => 6,
            0x3B | 0x5B | 0x7B => 10,
            0x3C | 0x5C | 0x7C => 6,
            0x3D | 0x5D | 0x7D => 10,
            _ => break,
        };
        if i + advance > rgce.len() { break; }
        i += advance;
    }
    out
}

/// Seed the default BIFF8 palette (64 slots).
/// Indices 0..7 are reserved built-ins; 8..63 are the user-editable
/// palette and 64/65 are system "window" colors.
fn seed_default_palette(palette: &mut HashMap<u16, String>) {
    // Built-ins that show up in XF colour fields as shorthand. Excel's
    // actual resolution for 0..7 is system-theme dependent; these are
    // the documented defaults.
    let defaults: &[(u16, &str)] = &[
        (0, "#000000"),  // Black
        (1, "#FFFFFF"),  // White
        (2, "#FF0000"),  // Red
        (3, "#00FF00"),  // Green
        (4, "#0000FF"),  // Blue
        (5, "#FFFF00"),  // Yellow
        (6, "#FF00FF"),  // Magenta
        (7, "#00FFFF"),  // Cyan
        // User palette 8..63 — defaults from MS-XLS reference
        (8, "#000000"),
        (9, "#FFFFFF"),
        (10, "#FF0000"),
        (11, "#00FF00"),
        (12, "#0000FF"),
        (13, "#FFFF00"),
        (14, "#FF00FF"),
        (15, "#00FFFF"),
        (16, "#800000"),
        (17, "#008000"),
        (18, "#000080"),
        (19, "#808000"),
        (20, "#800080"),
        (21, "#008080"),
        (22, "#C0C0C0"),
        (23, "#808080"),
        (24, "#9999FF"),
        (25, "#993366"),
        (26, "#FFFFCC"),
        (27, "#CCFFFF"),
        (28, "#660066"),
        (29, "#FF8080"),
        (30, "#0066CC"),
        (31, "#CCCCFF"),
        (32, "#000080"),
        (33, "#FF00FF"),
        (34, "#FFFF00"),
        (35, "#00FFFF"),
        (36, "#800080"),
        (37, "#800000"),
        (38, "#008080"),
        (39, "#0000FF"),
        (40, "#00CCFF"),
        (41, "#CCFFFF"),
        (42, "#CCFFCC"),
        (43, "#FFFF99"),
        (44, "#99CCFF"),
        (45, "#FF99CC"),
        (46, "#CC99FF"),
        (47, "#FFCC99"),
        (48, "#3366FF"),
        (49, "#33CCCC"),
        (50, "#99CC00"),
        (51, "#FFCC00"),
        (52, "#FF9900"),
        (53, "#FF6600"),
        (54, "#666699"),
        (55, "#969696"),
        (56, "#003366"),
        (57, "#339966"),
        (58, "#003300"),
        (59, "#333300"),
        (60, "#993300"),
        (61, "#993366"),
        (62, "#333399"),
        (63, "#333333"),
        // System "window" colors — 64 = sheet grid / auto, 65 = bg
        (64, "#000000"),
        (65, "#FFFFFF"),
    ];
    for &(i, hex) in defaults {
        palette.insert(i, hex.to_string());
    }
}

/// Walk a FORMULA record's rgce byte stream and extract every 3D
/// (cross-sheet) reference's ref-portion — the `$C13` or `$A$1:$B$2`
/// text that follows `SheetName!`. Used to patch calamine's output,
/// whose PtgRef3d / PtgArea3d decoders corrupt the column index
/// (quadruples it via a bad `colu << 2` instead of masking with
/// `& 0x3FFF`, and mis-detects the absolute/relative flags).
///
/// We don't bother decoding the sheet name — calamine gets that right
/// via its xti table. We only re-emit the ref portion and substitute
/// it by ordinal position when rewriting the formula text.
/// Walk a FORMULA record's rgce and extract every reference (2D or
/// 3D) in order. For each ref we emit `(is_3d, ref_text)` — the text
/// is the A1 portion only (no sheet prefix). Calamine renders sheet
/// names correctly via its xti table; we just fix the col/row bits
/// and absolute/relative flags, which calamine's decoders corrupt
/// across every ptg variant they handle.
///
/// Covers the common ptgs we actually see in the gop templates:
/// PtgRef (0x24), PtgArea (0x25), PtgRef3d (0x3A), PtgArea3d (0x3B)
/// plus each of their value-class (0x40|) and array-class (0x60|)
/// variants. Other ptg types are skipped by byte count so we stay
/// aligned with the cursor.
pub fn extract_refs(rgce: &[u8]) -> Vec<(bool, String)> {
    extract_refs_with_xti(rgce, &[], &[])
}

/// Like `extract_refs` but resolves 3D-ref sheet names via the xti
/// table. When `xtis` and `sheet_names` are both provided, the
/// returned 3D refs have sheet names embedded ("`SheetName!A1`"),
/// overriding calamine's potentially-wrong emission. When either is
/// empty we emit just the A1 portion for 3D refs — caller keeps
/// whatever calamine gave for the sheet.
pub fn extract_refs_with_xti(
    rgce: &[u8],
    xtis: &[XtiEntry],
    sheet_names: &[String],
) -> Vec<(bool, String)> {
    let resolve_sheet = |ixti: u16| -> Option<String> {
        let xti = xtis.get(ixti as usize)?;
        if xti.itab_first < 0 {
            return None;
        }
        let sheet = sheet_names.get(xti.itab_first as usize)?;
        Some(sheet.clone())
    };
    let mut out = Vec::new();
    let mut i = 0usize;
    while i < rgce.len() {
        let ptg = rgce[i];
        i += 1;
        let advance: usize = match ptg {
            // PtgRef (2D single cell) — col with flags in high bits.
            0x24 | 0x44 | 0x64 => {
                if i + 4 > rgce.len() { break; }
                let rwu = u16::from_le_bytes([rgce[i], rgce[i + 1]]);
                let colu = u16::from_le_bytes([rgce[i + 2], rgce[i + 3]]);
                out.push((false, format_ref(rwu, colu)));
                4
            }
            // PtgArea (2D range).
            0x25 | 0x45 | 0x65 => {
                if i + 8 > rgce.len() { break; }
                let rw1 = u16::from_le_bytes([rgce[i], rgce[i + 1]]);
                let rw2 = u16::from_le_bytes([rgce[i + 2], rgce[i + 3]]);
                let col1 = u16::from_le_bytes([rgce[i + 4], rgce[i + 5]]);
                let col2 = u16::from_le_bytes([rgce[i + 6], rgce[i + 7]]);
                let s = format!("{}:{}", format_ref(rw1, col1), format_ref(rw2, col2));
                out.push((false, s));
                8
            }
            // PtgRef3d — resolve sheet via xti table.
            0x3A | 0x5A | 0x7A => {
                if i + 6 > rgce.len() { break; }
                let ixti = u16::from_le_bytes([rgce[i], rgce[i + 1]]);
                let rwu = u16::from_le_bytes([rgce[i + 2], rgce[i + 3]]);
                let colu = u16::from_le_bytes([rgce[i + 4], rgce[i + 5]]);
                let a1 = format_ref(rwu, colu);
                let text = match resolve_sheet(ixti) {
                    Some(name) => format!("{}!{}", quoted_sheet(&name), a1),
                    None => a1,
                };
                out.push((true, text));
                6
            }
            // PtgArea3d — resolve sheet via xti table.
            0x3B | 0x5B | 0x7B => {
                if i + 10 > rgce.len() { break; }
                let ixti = u16::from_le_bytes([rgce[i], rgce[i + 1]]);
                let rw1 = u16::from_le_bytes([rgce[i + 2], rgce[i + 3]]);
                let rw2 = u16::from_le_bytes([rgce[i + 4], rgce[i + 5]]);
                let col1 = u16::from_le_bytes([rgce[i + 6], rgce[i + 7]]);
                let col2 = u16::from_le_bytes([rgce[i + 8], rgce[i + 9]]);
                let range = format!(
                    "{}:{}",
                    format_ref(rw1, col1),
                    format_ref(rw2, col2)
                );
                let text = match resolve_sheet(ixti) {
                    Some(name) => format!("{}!{}", quoted_sheet(&name), range),
                    None => range,
                };
                out.push((true, text));
                10
            }
            _ => ptg_payload_size(ptg, rgce, i),
        };
        if i + advance > rgce.len() { break; }
        i += advance;
    }
    out
}

/// Decode an ARRAY-record rgce into Excel-readable formula text.
/// Handles the specific ptg mix gop templates use for their MY* UDF
/// calls: PtgName (function + data args), PtgAttrSpace, PtgRef
/// (anchor), PtgFuncVar (User, iftab=0xFF). Returns `None` for ptgs
/// we can't decode — caller falls back to leaving the array
/// unreplicated.
///
/// `defined_names` should be the workbook's defined names in BIFF
/// order (same as calamine's `defined_names()` output). iname
/// references from PtgName are 1-based.
pub fn decode_array_formula(
    rgce: &[u8],
    defined_names: &[String],
    xtis: &[XtiEntry],
    sheet_names: &[String],
    extern_names: &[String],
) -> Option<String> {
    decode_formula_inner(rgce, defined_names, xtis, sheet_names, extern_names, None)
}

/// Decode a FORMULA-record rgce to Excel-readable text, with
/// shared-formula resolution. When the rgce starts with PtgExp we
/// look up the shared formula's rgce keyed by (anchor_row,
/// anchor_col) and decode that instead. Used by `xls_load.rs` to
/// produce text when calamine refuses to decode PtgExp cells.
pub fn decode_full_formula(
    rgce: &[u8],
    defined_names: &[String],
    xtis: &[XtiEntry],
    sheet_names: &[String],
    extern_names: &[String],
    shared_formulas: &HashMap<(u32, i32, i32), Vec<u8>>,
    sheet_idx: u32,
    cell_row: i32,
    cell_col: i32,
) -> Option<String> {
    // PtgExp: opcode(1) + rwAnchor(2) + colAnchor(2). Redirect to
    // the shared-formula rgce, keyed by the anchor coords.
    if rgce.len() == 5 && rgce[0] == 0x01 {
        let rw = u16::from_le_bytes([rgce[1], rgce[2]]) as i32 + 1;
        let col = u16::from_le_bytes([rgce[3], rgce[4]]) as i32 + 1;
        if let Some(shared) = shared_formulas.get(&(sheet_idx, rw, col)) {
            return decode_formula_inner(
                shared,
                defined_names,
                xtis,
                sheet_names,
                extern_names,
                Some((cell_row, cell_col)),
            );
        }
        return None;
    }
    decode_formula_inner(
        rgce,
        defined_names,
        xtis,
        sheet_names,
        extern_names,
        Some((cell_row, cell_col)),
    )
}

fn decode_formula_inner(
    rgce: &[u8],
    defined_names: &[String],
    xtis: &[XtiEntry],
    sheet_names: &[String],
    extern_names: &[String],
    cell_anchor: Option<(i32, i32)>,
) -> Option<String> {
    let mut stack: Vec<String> = Vec::new();
    let mut i = 0usize;
    while i < rgce.len() {
        let ptg = rgce[i];
        i += 1;
        match ptg {
            // PtgInt
            0x1E => {
                if i + 2 > rgce.len() { return None; }
                let n = u16::from_le_bytes([rgce[i], rgce[i + 1]]);
                stack.push(n.to_string());
                i += 2;
            }
            // PtgNum
            0x1F => {
                if i + 8 > rgce.len() { return None; }
                let mut b = [0u8; 8];
                b.copy_from_slice(&rgce[i..i + 8]);
                stack.push(f64::from_le_bytes(b).to_string());
                i += 8;
            }
            // PtgBool
            0x1D => {
                if i + 1 > rgce.len() { return None; }
                stack.push(if rgce[i] != 0 { "TRUE".into() } else { "FALSE".into() });
                i += 1;
            }
            // PtgStr
            0x17 => {
                if i + 2 > rgce.len() { return None; }
                let cch = rgce[i] as usize;
                let grbit = rgce[i + 1];
                let high = (grbit & 0x01) != 0;
                let body_start = i + 2;
                let body_len = cch * if high { 2 } else { 1 };
                if body_start + body_len > rgce.len() { return None; }
                let s = parse_biff_string(&rgce[body_start..], cch, high).unwrap_or_default();
                let escaped = s.replace('"', "\"\"");
                stack.push(format!("\"{escaped}\""));
                i = body_start + body_len;
            }
            // PtgName
            0x23 | 0x43 | 0x63 => {
                if i + 4 > rgce.len() { return None; }
                let iname = u32::from_le_bytes([
                    rgce[i], rgce[i + 1], rgce[i + 2], rgce[i + 3],
                ]) as usize;
                if iname == 0 { return None; }
                let name = defined_names.get(iname - 1)?;
                stack.push(name.clone());
                i += 4;
            }
            // PtgRef (2D, any class)
            0x24 | 0x44 | 0x64 => {
                if i + 4 > rgce.len() { return None; }
                let rwu = u16::from_le_bytes([rgce[i], rgce[i + 1]]);
                let colu = u16::from_le_bytes([rgce[i + 2], rgce[i + 3]]);
                stack.push(format_ref(rwu, colu));
                i += 4;
            }
            // PtgRefN — same layout as PtgRef but row/col are OFFSETS
            // from the cell containing the formula (used in shared
            // formulas). Requires cell_anchor context. Col-offset is
            // 14-bit signed (two's complement in the low 14 bits),
            // NOT zero-extended — negative column offsets like -4
            // are stored as 0x3FFC and must be sign-extended from
            // bit 13.
            0x2C | 0x4C | 0x6C => {
                if i + 4 > rgce.len() { return None; }
                let rw_off = u16::from_le_bytes([rgce[i], rgce[i + 1]]) as i16 as i32;
                let col_word = u16::from_le_bytes([rgce[i + 2], rgce[i + 3]]);
                let col_rel = (col_word & 0x4000) != 0;
                let row_rel = (col_word & 0x8000) != 0;
                // BIFF8 col fits in 8 bits (max IV = 255). For
                // relative references the low 8 bits are a SIGNED
                // offset; for absolute it's unsigned 0..255.
                let col_raw = (col_word & 0x00FF) as u8;
                let col_off = if col_rel {
                    col_raw as i8 as i32
                } else {
                    col_raw as i32
                };
                let (anchor_r, anchor_c) = cell_anchor?;
                let r = if row_rel { anchor_r + rw_off } else { rw_off + 1 };
                let c = if col_rel { anchor_c + col_off } else { col_off + 1 };
                if r <= 0 || c <= 0 { return None; }
                let mut s = String::new();
                if !col_rel { s.push('$'); }
                s.push_str(&col_index_to_letter(c as u32 - 1));
                if !row_rel { s.push('$'); }
                s.push_str(&r.to_string());
                stack.push(s);
                i += 4;
            }
            // PtgAreaN — 2D range with both corners relative to cell.
            0x2D | 0x4D | 0x6D => {
                if i + 8 > rgce.len() { return None; }
                let rw1 = u16::from_le_bytes([rgce[i], rgce[i + 1]]) as i16 as i32;
                let rw2 = u16::from_le_bytes([rgce[i + 2], rgce[i + 3]]) as i16 as i32;
                let cw1 = u16::from_le_bytes([rgce[i + 4], rgce[i + 5]]);
                let cw2 = u16::from_le_bytes([rgce[i + 6], rgce[i + 7]]);
                let (ar, ac) = cell_anchor?;
                let make = |rw: i32, cw: u16| -> Option<String> {
                    let col_rel = (cw & 0x4000) != 0;
                    let row_rel = (cw & 0x8000) != 0;
                    let col_raw = (cw & 0x00FF) as u8;
                    let col_off = if col_rel {
                        col_raw as i8 as i32
                    } else {
                        col_raw as i32
                    };
                    let r = if row_rel { ar + rw } else { rw + 1 };
                    let c = if col_rel { ac + col_off } else { col_off + 1 };
                    if r <= 0 || c <= 0 { return None; }
                    let mut s = String::new();
                    if !col_rel { s.push('$'); }
                    s.push_str(&col_index_to_letter(c as u32 - 1));
                    if !row_rel { s.push('$'); }
                    s.push_str(&r.to_string());
                    Some(s)
                };
                let a = make(rw1, cw1)?;
                let b = make(rw2, cw2)?;
                stack.push(format!("{a}:{b}"));
                i += 8;
            }
            // PtgArea (2D, any class)
            0x25 | 0x45 | 0x65 => {
                if i + 8 > rgce.len() { return None; }
                let rw1 = u16::from_le_bytes([rgce[i], rgce[i + 1]]);
                let rw2 = u16::from_le_bytes([rgce[i + 2], rgce[i + 3]]);
                let col1 = u16::from_le_bytes([rgce[i + 4], rgce[i + 5]]);
                let col2 = u16::from_le_bytes([rgce[i + 6], rgce[i + 7]]);
                stack.push(format!(
                    "{}:{}",
                    format_ref(rw1, col1),
                    format_ref(rw2, col2)
                ));
                i += 8;
            }
            // PtgNameX (any class) — references an EXTERNNAME entry.
            // Common case in xls files: Analysis ToolPak functions
            // (MROUND, CONVERT, etc.) come through here. Layout:
            //   ptg(1), ixti(2), nameindex(2), unused(2)  — 7 bytes.
            // We push `_xlfn.<name>` so unwrap_user_xlfn handles it
            // uniformly with the regular `_xlfn.IFERROR` cases the
            // post-2007 xlsx workbook also emits.
            0x39 | 0x59 | 0x79 => {
                if i + 6 > rgce.len() { return None; }
                let nameindex = u16::from_le_bytes([rgce[i + 2], rgce[i + 3]]);
                if nameindex == 0 { return None; }
                let name = extern_names
                    .get((nameindex as usize).saturating_sub(1))
                    .filter(|s| !s.is_empty())
                    .cloned()
                    .unwrap_or_else(|| "_unresolved_externname_".to_string());
                stack.push(format!("_xlfn.{name}"));
                i += 6;
            }
            // PtgRef3d
            0x3A | 0x5A | 0x7A => {
                if i + 6 > rgce.len() { return None; }
                let ixti = u16::from_le_bytes([rgce[i], rgce[i + 1]]);
                let rwu = u16::from_le_bytes([rgce[i + 2], rgce[i + 3]]);
                let colu = u16::from_le_bytes([rgce[i + 4], rgce[i + 5]]);
                let a1 = format_ref(rwu, colu);
                let sheet = xtis
                    .get(ixti as usize)
                    .and_then(|xti| {
                        if xti.itab_first >= 0 {
                            sheet_names.get(xti.itab_first as usize).cloned()
                        } else { None }
                    })
                    .unwrap_or_else(|| "#REF".into());
                stack.push(format!("{}!{}", quoted_sheet(&sheet), a1));
                i += 6;
            }
            // PtgArea3d
            0x3B | 0x5B | 0x7B => {
                if i + 10 > rgce.len() { return None; }
                let ixti = u16::from_le_bytes([rgce[i], rgce[i + 1]]);
                let rw1 = u16::from_le_bytes([rgce[i + 2], rgce[i + 3]]);
                let rw2 = u16::from_le_bytes([rgce[i + 4], rgce[i + 5]]);
                let col1 = u16::from_le_bytes([rgce[i + 6], rgce[i + 7]]);
                let col2 = u16::from_le_bytes([rgce[i + 8], rgce[i + 9]]);
                let range = format!(
                    "{}:{}",
                    format_ref(rw1, col1),
                    format_ref(rw2, col2)
                );
                let sheet = xtis
                    .get(ixti as usize)
                    .and_then(|xti| {
                        if xti.itab_first >= 0 {
                            sheet_names.get(xti.itab_first as usize).cloned()
                        } else { None }
                    })
                    .unwrap_or_else(|| "#REF".into());
                stack.push(format!("{}!{}", quoted_sheet(&sheet), range));
                i += 10;
            }
            // Binary ops (0x03..=0x11)
            0x03..=0x11 => {
                if stack.len() < 2 { return None; }
                let b = stack.pop().unwrap();
                let a = stack.pop().unwrap();
                let op = match ptg {
                    0x03 => "+", 0x04 => "-", 0x05 => "*", 0x06 => "/",
                    0x07 => "^", 0x08 => "&",
                    0x09 => "<",  0x0A => "<=", 0x0B => "=",
                    0x0C => ">=", 0x0D => ">",  0x0E => "<>",
                    _ => return None,
                };
                stack.push(format!("{a}{op}{b}"));
            }
            // Unary minus / plus / percent / parens
            0x12 => {
                let a = stack.pop()?;
                stack.push(format!("+{a}"));
            }
            0x13 => {
                let a = stack.pop()?;
                stack.push(format!("-{a}"));
            }
            0x14 => {
                let a = stack.pop()?;
                stack.push(format!("{a}%"));
            }
            0x15 => {
                let a = stack.pop()?;
                stack.push(format!("({a})"));
            }
            0x16 => stack.push(String::new()), // PtgMissArg
            // PtgAttr — most variants are just hints we can skip,
            // but sub=0x10 is PtgAttrSum: a one-operand SUM()
            // shortcut Excel emits for simple `=SUM(range)` formulas
            // (stored as PtgArea + PtgAttrSum instead of
            // PtgArea + PtgFuncVar SUM). Without wrapping, decoded
            // shared formulas come out as bare `=AR18:AR29` ranges
            // which IronCalc correctly rejects as "Implicit
            // Intersection not implemented".
            0x19 => {
                if i >= rgce.len() { return None; }
                let sub = rgce[i];
                if sub == 0x10 {
                    let operand = stack.pop()?;
                    stack.push(format!("SUM({operand})"));
                }
                let skip = match sub {
                    0x01 | 0x02 | 0x08 | 0x20 | 0x21 | 0x10 | 0x40 | 0x41 => 3,
                    0x04 => {
                        if i + 3 > rgce.len() { return None; }
                        let n = u16::from_le_bytes([rgce[i + 1], rgce[i + 2]]) as usize;
                        3 + 2 * (n + 1)
                    }
                    _ => return None,
                };
                i += skip;
            }
            // PtgFunc (fixed arg count from FTAB).
            0x21 | 0x41 | 0x61 => {
                if i + 2 > rgce.len() { return None; }
                let iftab = u16::from_le_bytes([rgce[i], rgce[i + 1]]);
                i += 2;
                let (fname, argc) = ftab_lookup(iftab)?;
                if stack.len() < argc { return None; }
                let args: Vec<String> = stack.split_off(stack.len() - argc);
                stack.push(format!("{fname}({})", args.join(",")));
            }
            // PtgFuncVar (variable arg count)
            0x22 | 0x42 | 0x62 => {
                if i + 3 > rgce.len() { return None; }
                let argc = rgce[i] as usize;
                let iftab = u16::from_le_bytes([rgce[i + 1], rgce[i + 2]]);
                i += 3;
                if stack.len() < argc { return None; }
                let args: Vec<String> = stack.split_off(stack.len() - argc);
                let (fname, real_args) = if iftab == 0xFF {
                    // User-defined: first arg IS the function name.
                    if args.is_empty() { return None; }
                    let mut a = args.into_iter();
                    let fname = a.next().unwrap();
                    (fname, a.collect::<Vec<_>>())
                } else {
                    let (f, _) = ftab_lookup(iftab)?;
                    (f.to_string(), args)
                };
                stack.push(format!("{fname}({})", real_args.join(",")));
            }
            _ => return None,
        }
    }
    stack.pop()
}

/// BIFF FTAB lookup — mirrors calamine 0.26's FTAB + FTAB_ARGC
/// (`calamine-0.26.1/src/utils.rs:75`). Calamine is authoritative
/// for [MS-XLS] FTAB indices, and since calamine does our PtgFuncVar
/// rendering for non-shared formulas, mismatches between this table
/// and calamine's create round-trip drift in shared formulas. Index
/// is the iftab from PtgFunc/PtgFuncVar.
///
/// The argc value follows calamine's FTAB_ARGC convention: `255` =
/// variable args (any count); `254` = variable but at least one
/// fixed-arity sentinel; otherwise the exact arg count. Only the
/// PtgFunc (0x21) caller uses argc; PtgFuncVar embeds its own count
/// in the rgce stream.
fn ftab_lookup(iftab: u16) -> Option<(&'static str, usize)> {
    // Empty entries ("", "User", and gaps in calamine's table) return
    // None so callers bail rather than emit garbage.
    let i = iftab as usize;
    if i >= FTAB_NAMES.len() { return None; }
    let name = FTAB_NAMES[i];
    if name.is_empty() || name == "User" { return None; }
    Some((name, FTAB_ARGC_TABLE[i] as usize))
}

const FTAB_NAMES: [&str; 485] = [
    "COUNT",
    "IF",
    "ISNA",
    "ISERROR",
    "SUM",
    "AVERAGE",
    "MIN",
    "MAX",
    "ROW",
    "COLUMN",
    "NA",
    "NPV",
    "STDEV",
    "DOLLAR",
    "FIXED",
    "SIN",
    "COS",
    "TAN",
    "ATAN",
    "PI",
    "SQRT",
    "EXP",
    "LN",
    "LOG10",
    "ABS",
    "INT",
    "SIGN",
    "ROUND",
    "LOOKUP",
    "INDEX",
    "REPT",
    "MID",
    "LEN",
    "VALUE",
    "TRUE",
    "FALSE",
    "AND",
    "OR",
    "NOT",
    "MOD",
    "DCOUNT",
    "DSUM",
    "DAVERAGE",
    "DMIN",
    "DMAX",
    "DSTDEV",
    "VAR",
    "DVAR",
    "TEXT",
    "LINEST",
    "TREND",
    "LOGEST",
    "GROWTH",
    "GOTO",
    "HALT",
    "RETURN",
    "PV",
    "FV",
    "NPER",
    "PMT",
    "RATE",
    "MIRR",
    "IRR",
    "RAND",
    "MATCH",
    "DATE",
    "TIME",
    "DAY",
    "MONTH",
    "YEAR",
    "WEEKDAY",
    "HOUR",
    "MINUTE",
    "SECOND",
    "NOW",
    "AREAS",
    "ROWS",
    "COLUMNS",
    "OFFSET",
    "ABSREF",
    "RELREF",
    "ARGUMENT",
    "SEARCH",
    "TRANSPOSE",
    "ERROR",
    "STEP",
    "TYPE",
    "ECHO",
    "SET.NAME",
    "CALLER",
    "DEREF",
    "WINDOWS",
    "SERIES",
    "DOCUMENTS",
    "ACTIVE.CELL",
    "SELECTION",
    "RESULT",
    "ATAN2",
    "ASIN",
    "ACOS",
    "CHOOSE",
    "HLOOKUP",
    "VLOOKUP",
    "LINKS",
    "INPUT",
    "ISREF",
    "GET.FORMULA",
    "GET.NAME",
    "SET.VALUE",
    "LOG",
    "EXEC",
    "CHAR",
    "LOWER",
    "UPPER",
    "PROPER",
    "LEFT",
    "RIGHT",
    "EXACT",
    "TRIM",
    "REPLACE",
    "SUBSTITUTE",
    "CODE",
    "NAMES",
    "DIRECTORY",
    "FIND",
    "CELL",
    "ISERR",
    "ISTEXT",
    "ISNUMBER",
    "ISBLANK",
    "T",
    "N",
    "FOPEN",
    "FCLOSE",
    "FSIZE",
    "FREADLN",
    "FREAD",
    "FWRITELN",
    "FWRITE",
    "FPOS",
    "DATEVALUE",
    "TIMEVALUE",
    "SLN",
    "SYD",
    "DDB",
    "GET.DEF",
    "REFTEXT",
    "TEXTREF",
    "INDIRECT",
    "REGISTER",
    "CALL",
    "ADD.BAR",
    "ADD.MENU",
    "ADD.COMMAND",
    "ENABLE.COMMAND",
    "CHECK.COMMAND",
    "RENAME.COMMAND",
    "SHOW.BAR",
    "DELETE.MENU",
    "DELETE.COMMAND",
    "GET.CHART.ITEM",
    "DIALOG.BOX",
    "CLEAN",
    "MDETERM",
    "MINVERSE",
    "MMULT",
    "FILES",
    "IPMT",
    "PPMT",
    "COUNTA",
    "CANCEL.KEY",
    "FOR",
    "WHILE",
    "BREAK",
    "NEXT",
    "INITIATE",
    "REQUEST",
    "POKE",
    "EXECUTE",
    "TERMINATE",
    "RESTART",
    "HELP",
    "GET.BAR",
    "PRODUCT",
    "FACT",
    "GET.CELL",
    "GET.WORKSPACE",
    "GET.WINDOW",
    "GET.DOCUMENT",
    "DPRODUCT",
    "ISNONTEXT",
    "GET.NOTE",
    "NOTE",
    "STDEVP",
    "VARP",
    "DSTDEVP",
    "DVARP",
    "TRUNC",
    "ISLOGICAL",
    "DCOUNTA",
    "DELETE.BAR",
    "UNREGISTER",
    "",
    "",
    "USDOLLAR",
    "FINDB",
    "SEARCHB",
    "REPLACEB",
    "LEFTB",
    "RIGHTB",
    "MIDB",
    "LENB",
    "ROUNDUP",
    "ROUNDDOWN",
    "ASC",
    "DBCS",
    "RANK",
    "",
    "",
    "ADDRESS",
    "DAYS360",
    "TODAY",
    "VDB",
    "ELSE",
    "ELSE.IF",
    "END.IF",
    "FOR.CELL",
    "MEDIAN",
    "SUMPRODUCT",
    "SINH",
    "COSH",
    "TANH",
    "ASINH",
    "ACOSH",
    "ATANH",
    "DGET",
    "CREATE.OBJECT",
    "VOLATILE",
    "LAST.ERROR",
    "CUSTOM.UNDO",
    "CUSTOM.REPEAT",
    "FORMULA.CONVERT",
    "GET.LINK.INFO",
    "TEXT.BOX",
    "INFO",
    "GROUP",
    "GET.OBJECT",
    "DB",
    "PAUSE",
    "",
    "",
    "RESUME",
    "FREQUENCY",
    "ADD.TOOLBAR",
    "DELETE.TOOLBAR",
    "User",
    "RESET.TOOLBAR",
    "EVALUATE",
    "GET.TOOLBAR",
    "GET.TOOL",
    "SPELLING.CHECK",
    "ERROR.TYPE",
    "APP.TITLE",
    "WINDOW.TITLE",
    "SAVE.TOOLBAR",
    "ENABLE.TOOL",
    "PRESS.TOOL",
    "REGISTER.ID",
    "GET.WORKBOOK",
    "AVEDEV",
    "BETADIST",
    "GAMMALN",
    "BETAINV",
    "BINOMDIST",
    "CHIDIST",
    "CHIINV",
    "COMBIN",
    "CONFIDENCE",
    "CRITBINOM",
    "EVEN",
    "EXPONDIST",
    "FDIST",
    "FINV",
    "FISHER",
    "FISHERINV",
    "FLOOR",
    "GAMMADIST",
    "GAMMAINV",
    "CEILING",
    "HYPGEOMDIST",
    "LOGNORMDIST",
    "LOGINV",
    "NEGBINOMDIST",
    "NORMDIST",
    "NORMSDIST",
    "NORMINV",
    "NORMSINV",
    "STANDARDIZE",
    "ODD",
    "PERMUT",
    "POISSON",
    "TDIST",
    "WEIBULL",
    "SUMXMY2",
    "SUMX2MY2",
    "SUMX2PY2",
    "CHITEST",
    "CORREL",
    "COVAR",
    "FORECAST",
    "FTEST",
    "INTERCEPT",
    "PEARSON",
    "RSQ",
    "STEYX",
    "SLOPE",
    "TTEST",
    "PROB",
    "DEVSQ",
    "GEOMEAN",
    "HARMEAN",
    "SUMSQ",
    "KURT",
    "SKEW",
    "ZTEST",
    "LARGE",
    "SMALL",
    "QUARTILE",
    "PERCENTILE",
    "PERCENTRANK",
    "MODE",
    "TRIMMEAN",
    "TINV",
    "",
    "MOVIE.COMMAND",
    "GET.MOVIE",
    "CONCATENATE",
    "POWER",
    "PIVOT.ADD.DATA",
    "GET.PIVOT.TABLE",
    "GET.PIVOT.FIELD",
    "GET.PIVOT.ITEM",
    "RADIANS",
    "DEGREES",
    "SUBTOTAL",
    "SUMIF",
    "COUNTIF",
    "COUNTBLANK",
    "SCENARIO.GET",
    "OPTIONS.LISTS.GET",
    "ISPMT",
    "DATEDIF",
    "DATESTRING",
    "NUMBERSTRING",
    "ROMAN",
    "OPEN.DIALOG",
    "SAVE.DIALOG",
    "VIEW.GET",
    "GETPIVOTDATA",
    "HYPERLINK",
    "PHONETIC",
    "AVERAGEA",
    "MAXA",
    "MINA",
    "STDEVPA",
    "VARPA",
    "STDEVA",
    "VARA",
    "BAHTTEXT",
    "THAIDAYOFWEEK",
    "THAIDIGIT",
    "THAIMONTHOFYEAR",
    "THAINUMSOUND",
    "THAINUMSTRING",
    "THAISTRINGLENGTH",
    "ISTHAIDIGIT",
    "ROUNDBAHTDOWN",
    "ROUNDBAHTUP",
    "THAIYEAR",
    "RTD",
    "CUBEVALUE",
    "CUBEMEMBER",
    "CUBEMEMBERPROPERTY",
    "CUBERANKEDMEMBER",
    "HEX2BIN",
    "HEX2DEC",
    "HEX2OCT",
    "DEC2BIN",
    "DEC2HEX",
    "DEC2OCT",
    "OCT2BIN",
    "OCT2HEX",
    "OCT2DEC",
    "BIN2DEC",
    "BIN2OCT",
    "BIN2HEX",
    "IMSUB",
    "IMDIV",
    "IMPOWER",
    "IMABS",
    "IMSQRT",
    "IMLN",
    "IMLOG2",
    "IMLOG10",
    "IMSIN",
    "IMCOS",
    "IMEXP",
    "IMARGUMENT",
    "IMCONJUGATE",
    "IMAGINARY",
    "IMREAL",
    "COMPLEX",
    "IMSUM",
    "IMPRODUCT",
    "SERIESSUM",
    "FACTDOUBLE",
    "SQRTPI",
    "QUOTIENT",
    "DELTA",
    "GESTEP",
    "ISEVEN",
    "ISODD",
    "MROUND",
    "ERF",
    "ERFC",
    "BESSELJ",
    "BESSELK",
    "BESSELY",
    "BESSELI",
    "XIRR",
    "XNPV",
    "PRICEMAT",
    "YIELDMAT",
    "INTRATE",
    "RECEIVED",
    "DISC",
    "PRICEDISC",
    "YIELDDISC",
    "TBILLEQ",
    "TBILLPRICE",
    "TBILLYIELD",
    "PRICE",
    "YIELD",
    "DOLLARDE",
    "DOLLARFR",
    "NOMINAL",
    "EFFECT",
    "CUMPRINC",
    "CUMIPMT",
    "EDATE",
    "EOMONTH",
    "YEARFRAC",
    "COUPDAYBS",
    "COUPDAYS",
    "COUPDAYSNC",
    "COUPNCD",
    "COUPNUM",
    "COUPPCD",
    "DURATION",
    "MDURATION",
    "ODDLPRICE",
    "ODDLYIELD",
    "ODDFPRICE",
    "ODDFYIELD",
    "RANDBETWEEN",
    "WEEKNUM",
    "AMORDEGRC",
    "AMORLINC",
    "CONVERT",
    "ACCRINT",
    "ACCRINTM",
    "WORKDAY",
    "NETWORKDAYS",
    "GCD",
    "MULTINOMIAL",
    "LCM",
    "FVSCHEDULE",
    "CUBEKPIMEMBER",
    "CUBESET",
    "CUBESETCOUNT",
    "IFERROR",
    "COUNTIFS",
    "SUMIFS",
    "AVERAGEIF",
    "AVERAGEIFS",
];

const FTAB_ARGC_TABLE: [u8; 485] = [
    255,
    3,
    1,
    1,
    255,
    255,
    255,
    255,
    1,
    1,
    0,
    254,
    255,
    2,
    3,
    1,
    1,
    1,
    1,
    0,
    1,
    1,
    1,
    1,
    1,
    1,
    1,
    2,
    3,
    4,
    2,
    3,
    1,
    1,
    0,
    0,
    255,
    255,
    1,
    2,
    3,
    3,
    3,
    3,
    3,
    3,
    255,
    3,
    2,
    4,
    4,
    4,
    4,
    1,
    1,
    1,
    5,
    5,
    5,
    5,
    6,
    3,
    2,
    0,
    3,
    3,
    3,
    1,
    1,
    1,
    2,
    1,
    1,
    1,
    0,
    1,
    1,
    1,
    5,
    2,
    2,
    3,
    3,
    1,
    2,
    0,
    1,
    1,
    2,
    0,
    1,
    2,
    2,
    2,
    0,
    0,
    1,
    2,
    1,
    1,
    255,
    4,
    4,
    2,
    7,
    1,
    1,
    2,
    2,
    2,
    4,
    1,
    1,
    1,
    1,
    2,
    2,
    2,
    1,
    4,
    4,
    1,
    3,
    1,
    3,
    2,
    1,
    1,
    1,
    1,
    1,
    1,
    2,
    1,
    1,
    1,
    2,
    2,
    2,
    2,
    1,
    1,
    3,
    4,
    5,
    3,
    2,
    2,
    2,
    255,
    255,
    1,
    4,
    5,
    5,
    5,
    5,
    1,
    3,
    4,
    3,
    1,
    1,
    1,
    1,
    1,
    2,
    6,
    6,
    255,
    2,
    4,
    1,
    0,
    0,
    2,
    2,
    3,
    2,
    1,
    1,
    1,
    4,
    255,
    1,
    2,
    1,
    2,
    2,
    3,
    1,
    3,
    4,
    255,
    255,
    3,
    3,
    2,
    1,
    3,
    1,
    1,
    0,
    0,
    2,
    3,
    3,
    4,
    2,
    2,
    3,
    3,
    2,
    2,
    1,
    1,
    3,
    0,
    0,
    5,
    3,
    0,
    7,
    0,
    1,
    0,
    3,
    255,
    255,
    1,
    1,
    1,
    1,
    1,
    1,
    3,
    11,
    1,
    0,
    2,
    3,
    5,
    4,
    4,
    1,
    0,
    5,
    5,
    1,
    0,
    0,
    1,
    2,
    2,
    1,
    255,
    1,
    1,
    2,
    3,
    3,
    1,
    1,
    1,
    2,
    3,
    3,
    3,
    2,
    255,
    5,
    1,
    5,
    4,
    2,
    2,
    2,
    3,
    3,
    1,
    3,
    3,
    3,
    1,
    1,
    2,
    4,
    3,
    2,
    4,
    3,
    3,
    3,
    4,
    1,
    3,
    1,
    3,
    1,
    2,
    3,
    3,
    4,
    2,
    2,
    2,
    2,
    2,
    2,
    3,
    2,
    2,
    2,
    2,
    2,
    2,
    4,
    4,
    255,
    255,
    255,
    255,
    255,
    255,
    3,
    2,
    2,
    2,
    2,
    3,
    255,
    2,
    2,
    4,
    4,
    3,
    255,
    2,
    9,
    2,
    3,
    4,
    1,
    1,
    255,
    3,
    2,
    1,
    2,
    1,
    4,
    3,
    1,
    2,
    2,
    4,
    5,
    2,
    128,
    2,
    1,
    255,
    255,
    255,
    255,
    255,
    255,
    255,
    1,
    1,
    1,
    1,
    1,
    1,
    1,
    1,
    1,
    1,
    1,
    255,
    255,
    3,
    3,
    4,
    2,
    1,
    2,
    2,
    2,
    2,
    2,
    2,
    1,
    1,
    2,
    2,
    2,
    2,
    2,
    1,
    1,
    1,
    1,
    1,
    1,
    1,
    1,
    1,
    1,
    1,
    1,
    3,
    255,
    255,
    4,
    1,
    1,
    2,
    2,
    2,
    1,
    1,
    2,
    2,
    1,
    2,
    2,
    2,
    2,
    3,
    3,
    6,
    6,
    5,
    5,
    5,
    5,
    5,
    3,
    3,
    3,
    7,
    7,
    2,
    2,
    2,
    2,
    6,
    6,
    2,
    2,
    3,
    4,
    4,
    4,
    4,
    4,
    4,
    6,
    6,
    8,
    8,
    8,
    8,
    2,
    2,
    7,
    7,
    8,
    8,
    5,
    3,
    3,
    255,
    255,
    255,
    2,
    4,
    5,
    1,
    2,
    128,
    129,
    3,
    129,
];

/// Wrap a sheet name in single quotes if Excel grammar requires it
/// (spaces, hyphens, leading digit, any non-ident char). Calamine's
/// quoting convention matches this.
fn quoted_sheet(name: &str) -> String {
    let needs = name.is_empty()
        || name.chars().next().map_or(false, |c| c.is_ascii_digit())
        || name
            .chars()
            .any(|c| !(c.is_ascii_alphanumeric() || c == '_'));
    if needs {
        format!("'{}'", name.replace('\'', "''"))
    } else {
        name.to_string()
    }
}

fn ptg_payload_size(ptg: u8, rgce: &[u8], i: usize) -> usize {
    match ptg {
        0x03..=0x11 | 0x12 | 0x13 | 0x14 | 0x15 | 0x16 => 0,
        0x01 | 0x02 => 4,
        0x1C | 0x1D => 1,
        0x1E => 2,
        0x1F => 8,
        0x17 => {
            let Some(&cch) = rgce.get(i) else { return 0; };
            let Some(&grbit) = rgce.get(i + 1) else { return 0; };
            let high = (grbit & 0x01) != 0;
            2 + cch as usize * if high { 2 } else { 1 }
        }
        0x18 => 5,
        0x19 => {
            let Some(&sub) = rgce.get(i) else { return 0; };
            match sub {
                0x01 | 0x02 | 0x08 | 0x20 | 0x21 | 0x10 | 0x40 | 0x41 => 3,
                0x04 => {
                    let Some(lo) = rgce.get(i + 1) else { return 0; };
                    let Some(hi) = rgce.get(i + 2) else { return 0; };
                    let n = u16::from_le_bytes([*lo, *hi]) as usize;
                    3 + 2 * (n + 1)
                }
                _ => 0,
            }
        }
        0x21 | 0x41 | 0x61 => 2,
        0x22 | 0x42 | 0x62 => 3,
        0x20 | 0x40 | 0x60 => 7,
        0x23 | 0x43 | 0x63 => 4,
        0x26 | 0x46 | 0x66 => 6,
        0x27 | 0x47 | 0x67 => 6,
        0x28 | 0x48 | 0x68 => 6,
        0x29 | 0x49 | 0x69 => 2,
        0x2A | 0x4A | 0x6A => 4,
        0x2B | 0x4B | 0x6B => 8,
        0x2C | 0x4C | 0x6C => 4,
        0x2D | 0x4D | 0x6D => 8,
        0x39 | 0x59 | 0x79 => 6,
        0x3C | 0x5C | 0x7C => 6,
        0x3D | 0x5D | 0x7D => 10,
        _ => 0,
    }
}


/// Format a PtgRef-style (rw, col) pair into an A1 reference portion
/// (without the sheet prefix). Handles the fRwRel (bit 15 of col) and
/// fColRel (bit 14 of col) flags per [MS-XLS] 2.5.198.
///
/// In PtgRef3d / PtgArea3d the col word carries the absolute/relative
/// flags in its top two bits; calamine's buggy decoder ignores this
/// and shifts left by 2 instead of masking, which turns col C (2) into
/// col I (8) on every xls file with cross-sheet refs.
fn format_ref(rwu: u16, colu: u16) -> String {
    // Use u32 so rowu=0xFFFF (defined-name "whole column" idiom)
    // doesn't wrap to 0.
    let row: u32 = rwu as u32 + 1;
    let col_abs = (colu & 0x4000) == 0;
    let row_abs = (colu & 0x8000) == 0;
    let col_idx = (colu & 0x3FFF) as u32;
    let col_letter = col_index_to_letter(col_idx);
    let mut s = String::new();
    if col_abs { s.push('$'); }
    s.push_str(&col_letter);
    if row_abs { s.push('$'); }
    s.push_str(&row.to_string());
    s
}

/// Walk a formula rgce and return every comparison operator in the
/// order they appear. Used to patch calamine's output when its
/// decoder has `0x0C` (PtgGE) and `0x0D` (PtgGT) transposed —
/// meaning every `>` in a formula shows up as `>=` and vice versa.
pub fn scan_comparison_ops(rgce: &[u8]) -> Vec<&'static str> {
    let mut out = Vec::new();
    let mut i = 0usize;
    while i < rgce.len() {
        let ptg = rgce[i];
        i += 1;
        let advance: usize = match ptg {
            0x09 => { out.push("<"); 0 }
            0x0A => { out.push("<="); 0 }
            0x0B => { out.push("="); 0 }
            0x0C => { out.push(">="); 0 }
            0x0D => { out.push(">"); 0 }
            0x0E => { out.push("<>"); 0 }
            // Other binary ops, unary, parens.
            0x03..=0x08 | 0x0F..=0x11 | 0x12..=0x16 => 0,
            0x01 | 0x02 => 4,
            0x1C | 0x1D => 1,
            0x1E => 2,
            0x1F => 8,
            0x17 => {
                let Some(&cch) = rgce.get(i) else { break };
                let Some(&grbit) = rgce.get(i + 1) else { break };
                let high = (grbit & 0x01) != 0;
                2 + cch as usize * if high { 2 } else { 1 }
            }
            0x18 => 5,
            0x19 => {
                let Some(&sub) = rgce.get(i) else { break };
                match sub {
                    0x01 | 0x02 | 0x08 | 0x20 | 0x21 | 0x10 | 0x40 | 0x41 => 3,
                    0x04 => {
                        let Some(lo) = rgce.get(i + 1) else { break };
                        let Some(hi) = rgce.get(i + 2) else { break };
                        let n = u16::from_le_bytes([*lo, *hi]) as usize;
                        3 + 2 * (n + 1)
                    }
                    _ => break,
                }
            }
            0x21 | 0x41 | 0x61 => 2,
            0x22 | 0x42 | 0x62 => 3,
            0x20 | 0x40 | 0x60 => 7,
            0x23 | 0x43 | 0x63 => 4,
            0x24 | 0x44 | 0x64 => 4,
            0x25 | 0x45 | 0x65 => 8,
            0x26 | 0x46 | 0x66 => 6,
            0x27 | 0x47 | 0x67 => 6,
            0x28 | 0x48 | 0x68 => 6,
            0x29 | 0x49 | 0x69 => 2,
            0x2A | 0x4A | 0x6A => 4,
            0x2B | 0x4B | 0x6B => 8,
            0x2C | 0x4C | 0x6C => 4,
            0x2D | 0x4D | 0x6D => 8,
            0x39 | 0x59 | 0x79 => 6,
            0x3A | 0x5A | 0x7A => 6,
            0x3B | 0x5B | 0x7B => 10,
            0x3C | 0x5C | 0x7C => 6,
            0x3D | 0x5D | 0x7D => 10,
            _ => break,
        };
        if i + advance > rgce.len() { break; }
        i += advance;
    }
    out
}

fn col_index_to_letter(mut col: u32) -> String {
    // col here is 0-based (A = 0). Produce A, B, ..., Z, AA, AB, ...
    let mut out = Vec::new();
    loop {
        out.push((b'A' + (col % 26) as u8) as char);
        if col < 26 { break; }
        col = col / 26 - 1;
    }
    out.iter().rev().collect()
}

/// Parse a BIFF8 character string where `cch` characters follow either
/// as 1-byte each (when `high_byte` is false) or 2-byte UTF-16LE each
/// (when true). Returns None if the buffer isn't long enough.
fn parse_biff_string(buf: &[u8], cch: usize, high_byte: bool) -> Option<String> {
    if high_byte {
        if buf.len() < cch * 2 {
            return None;
        }
        let mut units: Vec<u16> = Vec::with_capacity(cch);
        for i in 0..cch {
            units.push(u16::from_le_bytes([buf[i * 2], buf[i * 2 + 1]]));
        }
        String::from_utf16(&units).ok()
    } else {
        if buf.len() < cch {
            return None;
        }
        // BIFF single-byte strings are actually windows-1252-ish; treat
        // as latin-1 to preserve any high bytes present (good enough
        // for the ASCII-dominated content we expect).
        Some(buf[..cch].iter().map(|&b| b as char).collect())
    }
}
