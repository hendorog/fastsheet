# fastsheet

A Lotus 1-2-3-inspired desktop spreadsheet with native Excel file
compatibility (`.xlsx` and `.xls`) and a keyboard-first UI. Built on
Tauri 2 + SvelteKit + Rust + [IronCalc](https://github.com/ironcalc/IronCalc).

**Status:** early development, but xlsx and xls round-trip on real-world
templates including macros and array-formula UDFs.

## Install (Windows)

Grab the latest release from
[Releases](https://github.com/hendorog/fastsheet/releases):

- **`fastsheet_x.y.z_x64-setup.exe`** — NSIS installer (recommended;
  installs to your user profile, no admin required, adds a Start
  Menu entry).
- **`fastsheet_x.y.z_x64_en-US.msi`** — same thing as an MSI for
  group-policy-friendly deployments.

On first launch Windows SmartScreen may show a "Windows protected
your PC" prompt — fastsheet builds aren't yet code-signed. Click
**More info → Run anyway**.

**Requirements:** Windows 10 1809+ or Windows 11. The WebView2
runtime ships with both, so no extra install needed.

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
- **Formula-trace popup** (`/T T`) — walks a cell's full dependency
  chain across sheets and named ranges so you can see *why* a formula
  evaluates to what it does. F2 edit-mode highlights every reference
  the formula points at, including named ranges.

## Stack

- **UI**: [SvelteKit](https://svelte.dev) (SPA mode via
  `adapter-static`) + TypeScript
- **Shell**: [Tauri 2](https://tauri.app) (native WebView2 on Windows,
  small binary, fast cold-start)
- **Calc engine**: [IronCalc](https://github.com/ironcalc/IronCalc)
  0.7.x, vendored under `vendor/` with additions for pseudo-spill
  UDFs and a couple of evaluation-consistency fixes
- **xls reader / writer**: hand-rolled BIFF8 over
  [calamine](https://github.com/tafia/calamine) + [cfb](https://github.com/mdsteele/rust-cfb)
  — see `src-tauri/src/xls_load.rs` and `xls_save.rs`. Macros / VBA
  storages are captured on load and replayed on save so Excel users
  keep working files.
- **xlsx**: read via IronCalc; saves go through an in-place
  byte-patcher that preserves charts / pivots / drawings / VBA when
  the original is available

## Build from source (Windows)

```bash
git clone https://github.com/hendorog/fastsheet
cd fastsheet
npm install
npm run tauri build
```

Output: `src-tauri/target/release/bundle/nsis/*.exe` and
`src-tauri/target/release/bundle/msi/*.msi`.

For dev, `npm run tauri dev` launches a hot-reload build pointed at
`http://localhost:1420`.

## Headless probes

Three CLI tools live under `src-tauri/src/bin/` for validating
behaviour against real spreadsheet files without launching the GUI:

```bash
cd src-tauri

# Load a file, dump styles / col widths / hidden ranges.
cargo run --release --bin probe -- ../fixtures/<file>.xlsx

# Round-trip an .xls file through save → reload, report cell-by-cell diff.
cargo run --release --bin save_probe -- <file>.xls

# Walk the dependency tree of a specific cell — the same trace the
# /T T menu produces, but printed to stdout.
cargo run --release --bin trace_probe -- <file>.xls 0:4:7
```

`save_probe` accepts diagnostic env vars:

- `FASTSHEET_PROBE_FULL_DIFF=1` — group remaining diffs by category
- `FASTSHEET_PROBE_COMPARE=sheet:row:col[,...]` — dump cell variants
  side-by-side
- `FASTSHEET_PROBE_CELL=sheet:row:col` — print the parsed Node tree
  + formatted value for a specific cell

## Tests

```bash
cd src-tauri && cargo test
```

21 unit tests for the BIFF8 writer plus integration tests exercising
the .xls round-trip path (including VBA / macro preservation). Real-
file round-trip tests opt-in via env vars:

```bash
FASTSHEET_RT_FILE=/path/to/file.xls FASTSHEET_RT_THRESHOLD=5 \
  cargo test --test xls_roundtrip rt_external_fixture
```

## Repo layout

```
src/                 SvelteKit frontend
  routes/+page.svelte  Top-level page (keyboard dispatch, layout)
  lib/                 Components + utilities
    FormulaTrace.svelte  /T T popup
    Grid.svelte          virtualized cell grid
    Navigator.svelte     /F R file picker
    menu.ts              Lotus-style menu tree
src-tauri/           Rust backend + Tauri shell
  src/
    workbook.rs        Open / new / save / backup commands
    cells.rs           Cell read / write / layout / trace commands
    trace.rs           Formula dependency walker
    xls_load.rs        .xls loader (calamine + custom BIFF scanner)
    xls_save.rs        .xls writer (from-scratch BIFF8 emit)
    xls_preserve.rs    VBA / macro storage capture+replay
    xlsx_load.rs       .xlsx loader (IronCalc + preprocessing)
    xlsx_save.rs       .xlsx in-place byte patcher
    ...
  tests/               Integration tests
  bin/                 Headless probes
vendor/              Vendored IronCalc with fastsheet additions
.github/workflows/   CI: release.yml builds Windows artefacts on tag push
fixtures/            Generated test workbooks
scripts/             Maintainer scripts (Linux/WSL dev workflow)
```

## Development (WSL → Windows)

The project's primary author develops in WSL and cross-builds to
Windows via `cargo-xwin`. If you're on Linux/macOS and want to do the
same:

```bash
# Build + install to %USERPROFILE%\Tools\fastsheet\ + drop a Start
# Menu shortcut. Idempotent — re-run after every change.
scripts/install-windows.sh
```

This runs:
```bash
npx tauri build --runner cargo-xwin --target x86_64-pc-windows-msvc --no-bundle
```
under the hood. See `scripts/install-windows.sh` for the install
details (path resolution, MOTW stripping, shortcut creation).

`FASTSHEET_PROFILE_LOAD=1` enables phase timing for load / evaluate /
boot — written to `fastsheet_profile.log` next to the running .exe.

## License

MIT.

IronCalc is vendored under `vendor/` (crates `ironcalc` and
`ironcalc_base`, version 0.7.x) and is licensed under
`MIT OR Apache-2.0`. See https://github.com/ironcalc/IronCalc.
