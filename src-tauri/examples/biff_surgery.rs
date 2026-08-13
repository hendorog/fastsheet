//! biff_surgery — dev diagnostic: copy an xls, applying record-level
//! edits to one worksheet substream, fixing BOUNDSHEET8 lbPlyPos.
//! Used to bisect Office File Validation failures by mutating an
//! Excel-authored file toward fastsheet's output one difference at a
//! time.
//!
//! Usage:
//!   biff_surgery <in.xls> <out.xls> <sheet_index_0based> <drop_opcodes_hex_csv>
//!   biff_surgery <in.xls> <out.xls> <sheet_index_0based> transplant <donor.xls>
//!
//! transplant replaces the target substream in <in.xls> with the same-
//! index substream from <donor.xls>.

use std::io::{Read, Write};

fn substream_spans(wb: &[u8]) -> Vec<(usize, usize)> {
    let mut spans = Vec::new();
    let mut pos = 0usize;
    let mut depth = 0u32;
    let mut start = 0usize;
    while pos + 4 <= wb.len() {
        let sid = u16::from_le_bytes([wb[pos], wb[pos + 1]]);
        let ln = u16::from_le_bytes([wb[pos + 2], wb[pos + 3]]) as usize;
        if sid == 0x0809 {
            if depth == 0 {
                start = pos;
            }
            depth += 1;
        }
        pos += 4 + ln;
        if sid == 0x000A {
            depth = depth.saturating_sub(1);
            if depth == 0 {
                spans.push((start, pos));
            }
        }
    }
    spans
}

fn main() {
    let mut args = std::env::args().skip(1);
    let inp = args.next().expect("in.xls");
    let outp = args.next().expect("out.xls");
    let sheet_idx: usize = args.next().expect("sheet idx").parse().unwrap();
    let mode = args.next().expect("opcodes or 'transplant'");

    let bytes = std::fs::read(&inp).expect("read");

    if mode == "transplant" {
        let donor = args.next().expect("donor.xls");
        let donor_bytes = std::fs::read(&donor).expect("read donor");
        let mut src_cfb = cfb::CompoundFile::open(std::io::Cursor::new(bytes.clone())).expect("cfb");
        let mut wb = Vec::new();
        src_cfb.open_stream("/Workbook").unwrap().read_to_end(&mut wb).unwrap();
        let mut dn_cfb = cfb::CompoundFile::open(std::io::Cursor::new(donor_bytes)).expect("donor cfb");
        let mut dwb = Vec::new();
        dn_cfb.open_stream("/Workbook").unwrap().read_to_end(&mut dwb).unwrap();
        let spans = substream_spans(&wb);
        let dspans = substream_spans(&dwb);
        let (a, b) = spans[sheet_idx + 1]; // +1: globals
        let (da, db) = dspans[sheet_idx + 1];
        let mut out = Vec::with_capacity(wb.len());
        out.extend_from_slice(&wb[..a]);
        out.extend_from_slice(&dwb[da..db]);
        out.extend_from_slice(&wb[b..]);
        let delta = (db - da) as i64 - (b - a) as i64;
        // Fix lbPlyPos of BOUNDSHEET8s pointing past the replaced span.
        let mut pos = 0usize;
        while pos + 4 <= out.len() {
            let sid = u16::from_le_bytes([out[pos], out[pos + 1]]);
            let ln = u16::from_le_bytes([out[pos + 2], out[pos + 3]]) as usize;
            if sid == 0x0085 {
                let bp = pos + 4;
                let old = u32::from_le_bytes([out[bp], out[bp + 1], out[bp + 2], out[bp + 3]]) as i64;
                if old > a as i64 {
                    let newv = (old + delta) as u32;
                    out[bp..bp + 4].copy_from_slice(&newv.to_le_bytes());
                }
            }
            pos += 4 + ln;
            if sid == 0x0809 && pos > a && pos < a + (db - da) {
                // inside transplanted region — skip scanning? records
                // parse the same way; nothing to do.
            }
        }
        let entries: Vec<(std::path::PathBuf, bool)> = src_cfb
            .walk()
            .map(|e| (e.path().to_path_buf(), e.is_storage()))
            .collect();
        let out_file = std::fs::File::create(&outp).expect("create out");
        let mut dst = cfb::CompoundFile::create_with_version(cfb::Version::V3, out_file).unwrap();
        for (path, is_storage) in entries {
            if path.as_os_str().is_empty() || path == std::path::Path::new("/") {
                continue;
            }
            if is_storage {
                let _ = dst.create_storage_all(&path);
            } else {
                let mut data = Vec::new();
                src_cfb.open_stream(&path).unwrap().read_to_end(&mut data).unwrap();
                if path == std::path::Path::new("/Workbook") {
                    data = out.clone();
                }
                if let Some(parent) = path.parent() {
                    if !parent.as_os_str().is_empty() {
                        let _ = dst.create_storage_all(parent);
                    }
                }
                let mut s = dst.create_stream(&path).unwrap();
                s.write_all(&data).unwrap();
            }
        }
        dst.flush().unwrap();
        println!("transplanted substream {} (delta {} bytes) -> {}", sheet_idx, delta, outp);
        return;
    }

    let drop: Vec<u16> = mode
        .split(',')
        .map(|s| u16::from_str_radix(s.trim_start_matches("0x"), 16).unwrap())
        .collect();
    let cursor = std::io::Cursor::new(bytes.clone());
    let mut src_cfb = cfb::CompoundFile::open(cursor).expect("cfb open");
    let mut wb = Vec::new();
    src_cfb
        .open_stream("/Workbook")
        .expect("wb stream")
        .read_to_end(&mut wb)
        .expect("read wb");

    // Walk records; track top-level substream index via depth. Records
    // in target substream with dropped opcodes are removed. Collect
    // (old_offset -> removed_bytes_before) mapping for lbPlyPos fixes.
    let mut out: Vec<u8> = Vec::with_capacity(wb.len());
    let mut removed_before: Vec<(usize, usize)> = Vec::new(); // (old offset, cumulative removed)
    let mut boundsheet_positions: Vec<usize> = Vec::new(); // offsets of lbPlyPos fields in OUT
    let mut pos = 0usize;
    let mut depth = 0u32;
    let mut substream = -1i32;
    let mut removed = 0usize;
    while pos + 4 <= wb.len() {
        let sid = u16::from_le_bytes([wb[pos], wb[pos + 1]]);
        let ln = u16::from_le_bytes([wb[pos + 2], wb[pos + 3]]) as usize;
        if sid == 0x0809 {
            if depth == 0 {
                substream += 1;
            }
            depth += 1;
        }
        let in_target = depth > 0 && substream == sheet_idx as i32 + 1; // +1: globals is substream 0
        let dropped = in_target && drop.contains(&sid);
        if sid == 0x0085 && depth > 0 && substream == 0 {
            // BOUNDSHEET8 in globals — lbPlyPos at body offset 0
            boundsheet_positions.push(out.len() + 4);
        }
        if dropped {
            removed += 4 + ln;
        } else {
            out.extend_from_slice(&wb[pos..pos + 4 + ln]);
        }
        removed_before.push((pos, removed));
        if sid == 0x000A {
            depth = depth.saturating_sub(1);
        }
        pos += 4 + ln;
    }
    println!("removed {} bytes", removed);

    // Fix lbPlyPos: for each BOUNDSHEET8, read old offset, compute
    // removed-bytes before that offset, subtract.
    for &bp in &boundsheet_positions {
        let old = u32::from_le_bytes([out[bp], out[bp + 1], out[bp + 2], out[bp + 3]]) as usize;
        let rem = removed_before
            .iter()
            .take_while(|&&(o, _)| o < old)
            .last()
            .map(|&(_, r)| r)
            .unwrap_or(0);
        let newv = (old - rem) as u32;
        out[bp..bp + 4].copy_from_slice(&newv.to_le_bytes());
    }

    // Write new CFB: copy every stream, replacing /Workbook.
    let entries: Vec<(std::path::PathBuf, bool)> = src_cfb
        .walk()
        .map(|e| (e.path().to_path_buf(), e.is_storage()))
        .collect();
    let out_file = std::fs::File::create(&outp).expect("create out");
    let mut dst = cfb::CompoundFile::create_with_version(cfb::Version::V3, out_file).expect("cfb create");
    for (path, is_storage) in entries {
        if path.as_os_str().is_empty() || path == std::path::Path::new("/") {
            continue;
        }
        if is_storage {
            let _ = dst.create_storage_all(&path);
        } else {
            let mut data = Vec::new();
            src_cfb.open_stream(&path).unwrap().read_to_end(&mut data).unwrap();
            if path == std::path::Path::new("/Workbook") {
                data = out.clone();
            }
            if let Some(parent) = path.parent() {
                if !parent.as_os_str().is_empty() {
                    let _ = dst.create_storage_all(parent);
                }
            }
            let mut s = dst.create_stream(&path).unwrap();
            s.write_all(&data).unwrap();
        }
    }
    dst.flush().unwrap();
    println!("wrote {}", outp);
}
