# fastsheet

A Lotus 1-2-3-inspired desktop spreadsheet with native Excel file
compatibility (`.xlsx` and `.xls`) and a keyboard-first UI. Built on
Tauri 2 + SvelteKit + Rust + [IronCalc](https://github.com/ironcalc/IronCalc).

**Status:** early development, but xlsx and xls round-trip on real-world
templates.

## Why

Modern spreadsheet UIs are built around a mouse. fastsheet is built
around the keyboard:

- **Lotus-style menu tree** — `/` opens the menu, single-letter keys
  navigate. No clicks required for any operation.
- **Address-bar goto** — `F5` and a sheet-qualified address jumps
  anywhere in the workbook instantly.
- **Modal hints in the menu bar** — when an operation takes over the
  keyboard (e.g. range pick for copy/move), the keys you can press are
  shown where you'd expect them.
- **Arrow-key responsiveness as a first-class goal** — virtualized
  rendering, rAF-coalesced movement, overlay-based selection. Stays
  smooth on 100k-cell workbooks.

## Stack

- **UI**: [SvelteKit](https://svelte.dev) (SPA mode via
  `adapter-static`) + TypeScript
- **Shell**: [Tauri 2](https://tauri.app) (cross-platform, native
  webview, small binary)
- **Calc engine**: [IronCalc](https://github.com/ironcalc/IronCalc)
  0.7.x, vendored under `vendor/` with a small set of additions for
  pseudo-spill UDFs and a couple of evaluation-consistency fixes
- **xls reader / writer**: hand-rolled BIFF8 over
  [calamine](https://github.com/tafia/calamine) + [cfb](https://github.com/mdsteele/rust-cfb) —
  see `src-tauri/src/xls_load.rs` and `xls_save.rs`
- **xlsx**: read via IronCalc; saves go through an in-place
  byte-patcher that preserves charts / pivots / drawings / VBA when
  the original is available

## Build + run

This project cross-builds to a Windows `.exe` from WSL, because WSLg's
WebKit rendering has been unreliable on the dev machine. The Windows
binary runs perfectly under WSL via `cmd.exe /c start`.

```bash
# Rebuild (Rust + frontend + embed)
npx tauri build --runner cargo-xwin --target x86_64-pc-windows-msvc --no-bundle

# Launch
cmd.exe /c start "" "$(wslpath -w src-tauri/target/x86_64-pc-windows-msvc/release/fastsheet.exe)"

# Kill a stuck window
cmd.exe /c "taskkill /F /IM fastsheet.exe"
```

For native Linux / macOS / Windows host development, the standard
Tauri commands work too:

```bash
npm install
npm run tauri dev
```

Quick checks without a full build:

```bash
cd src-tauri && cargo check --release      # Rust
npm run check                               # Svelte / TS
```

## Headless probes

Two CLI tools exist for validating behaviour against real spreadsheet
files:

```bash
# Load a file, dump styles / col widths / hidden ranges.
cd src-tauri && cargo run --release --bin probe -- ../fixtures/<file>.xlsx

# Round-trip an .xls file through save → reload, report cell-by-cell diff.
cargo run --release --bin save_probe -- <file>.xls
```

`save_probe` accepts diagnostic env vars:

- `FASTSHEET_PROBE_FULL_DIFF=1` — group remaining diffs by category
- `FASTSHEET_PROBE_COMPARE=sheet:row:col[,...]` — dump cell variants
  side-by-side
- `FASTSHEET_PROBE_CELL=sheet:row:col` — print the parsed Node tree
  for a specific cell

## Tests

```bash
cd src-tauri && cargo test
```

Includes 21 unit tests for the BIFF8 writer plus 10 integration tests
exercising the .xls round-trip path. Real-file round-trip tests can
be opted into via env vars:

```bash
FASTSHEET_RT_FILE=/path/to/file.xls FASTSHEET_RT_THRESHOLD=5 \
  cargo test --test xls_roundtrip rt_external_fixture
```

## Repo layout

```
src/                 SvelteKit frontend
  routes/+page.svelte  Top-level page (keyboard dispatch, layout)
  lib/                 Reusable components + utilities
src-tauri/           Rust backend + Tauri shell
  src/
    workbook.rs        Open / new / save / backup commands
    cells.rs           Cell read / write / layout
    xls_load.rs        .xls loader (calamine + custom BIFF scanner)
    xls_save.rs        .xls writer (from-scratch BIFF8 emit)
    xlsx_load.rs       .xlsx loader (IronCalc + preprocessing)
    xlsx_save.rs       .xlsx in-place byte patcher
    ...
  tests/               Integration tests
  bin/                 Headless probes
vendor/              Vendored IronCalc with fastsheet additions
fixtures/            Generated test workbooks
```

## License

MIT.

IronCalc is vendored under `vendor/` (crates `ironcalc` and
`ironcalc_base`, version 0.7.x) and is licensed under
`MIT OR Apache-2.0`. See https://github.com/ironcalc/IronCalc.
