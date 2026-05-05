<script lang="ts">
  import { invoke } from "@tauri-apps/api/core";
  import { onMount, tick } from "svelte";
  // Best-effort boot-trace mark. Active only when FASTSHEET_PROFILE_LOAD
  // is set on the .exe — otherwise this Tauri command is a no-op.
  invoke("profile_mark", { label: "script_loaded" }).catch(() => {});
  import Grid from "$lib/Grid.svelte";
  import Navigator from "$lib/Navigator.svelte";
  import SheetTabs from "$lib/SheetTabs.svelte";
  import FormulaTrace from "$lib/FormulaTrace.svelte";
  import CompareDiff from "$lib/CompareDiff.svelte";
  import {
    buildMenu,
    saveMenuItems,
    currentLevel,
    breadcrumb,
    type FormatKind,
  } from "$lib/menu";
  import type {
    CellView,
    LayoutData,
    MenuItem,
    WorkbookInfo,
    SaveResult,
    BackupResult,
    TraceNode,
    NamedRangeInfo,
    CompareResult,
    CompareDiff as CompareDiffType,
    CompareSheetMissing,
  } from "$lib/types";
  import {
    addr,
    autoFitColumnPx,
    autoFitRowPx,
    colWidthPx,
    key,
    rowHeightPx,
  } from "$lib/utils";

  // Viewport sized to the active sheet's dimension (clamped backend-side
  // to [100..1048576] × [60..16384]). Refreshed on sheet switch + initial
  // open via get_sheet_dim. The cursor isn't bound by viewportRows /
  // viewportCols — navigating past the dimension grows the viewport
  // (see growViewportToInclude). Cells are loaded lazily per visible row
  // band — see ensureRowsLoaded / handleBandChange below.
  let viewportRows = $state(100);
  let viewportCols = $state(60);
  // Excel's absolute hard limits. Cursor position is clamped to these,
  // not to viewportRows / viewportCols.
  const ABS_MAX_ROWS = 1_048_576;
  const ABS_MAX_COLS = 16_384;
  // Buffer rows/cols added when growing the viewport so the user can
  // keep typing past the new edge without immediately re-growing.
  const VIEWPORT_GROW_ROWS = 100;
  const VIEWPORT_GROW_COLS = 20;
  let frozenRows = $state(0);
  let frozenCols = $state(0);
  let mergedRanges = $state<string[]>([]);

  let workbook = $state<WorkbookInfo | null>(null);
  let activeSheet = $state(0);
  let cells = $state<Map<string, CellView>>(new Map());
  let colWidths = $state<Map<number, number>>(new Map());
  let rowHeights = $state<Map<number, number>>(new Map());
  let selRow = $state(1);
  let selCol = $state(1);
  // Opposite corner of the active selection rectangle. When equal to
  // (selRow, selCol) the selection is a single cell. Shift+arrow / click+
  // drag moves this end without disturbing the anchor.
  let rangeEndRow = $state(1);
  let rangeEndCol = $state(1);
  let editing = $state(false);
  let editValue = $state("");
  let editorEl: HTMLInputElement | null = $state(null);
  let gridWrapEl: HTMLDivElement | null = $state(null);
  let dirtyEdits = $state(0);
  let statusMsg = $state("Ready. Press / for menu.");
  let currentPath = $state("");

  // /-menu state
  let menuOpen = $state(false);
  let menuPath = $state<number[]>([]);
  let menuHighlight = $state(0);
  // Optional ad-hoc menu level (e.g. the Save Replace/SaveAs/Backup/Cancel
  // picker). When set, replaces what currentLevel would walk to. Title is
  // shown in the menu prompt instead of breadcrumb.
  let dynamicLevel = $state<MenuItem[] | null>(null);
  let dynamicTitle = $state<string>("");

  // file navigator state
  let navOpen = $state(false);
  let navMode = $state<"open" | "save">("open");

  /// Formula trace popup — set to a TraceNode to show the popup,
  /// null to close. Driven by the /T menu items.
  let traceRoot = $state<TraceNode | null>(null);
  /// Layout flags for the popup, two-way bound. `docked` flips it
  /// from centered modal to a right-side panel; `hidden` collapses
  /// it to a tiny bar so the user can keep interacting with the
  /// grid without dropping trace state.
  let traceDocked = $state(false);
  let traceHidden = $state(false);
  /// Active workbook comparison. Null = compare mode off. The dock
  /// shows the diff list; trace popups in this mode pick up
  /// compare_value on each node so deps render `left | right`.
  let compareResult = $state<CompareResult | null>(null);
  let compareHidden = $state(false);
  /// When true the navigator is being used to pick the right-side
  /// file for a compare; on activate, it routes to compareOpen instead
  /// of opening the workbook.
  let navCompareTarget = $state(false);
  /// Snapshot of (sheet, cursor) at the moment the trace popup
  /// opened. Restored on Esc close so the user lands back where they
  /// started. Cleared if they Enter-jump to a previewed cell instead.
  let traceOriginCursor = $state<{
    sheet: number;
    selRow: number;
    selCol: number;
    rangeEndRow: number;
    rangeEndCol: number;
  } | null>(null);
  /// Active reference highlights overlaid on the grid. Used by the
  /// trace popup (single-element preview) and F2 edit-mode (one
  /// per ref in the in-progress formula).
  let refHighlights = $state<
    {
      sheet: number;
      r1: number;
      c1: number;
      r2: number;
      c2: number;
      color: string;
      label?: string;
    }[]
  >([]);
  /// Cache of defined names → resolved range bounds. Refreshed on
  /// workbook open and after any name create / delete. Lets F2 edit
  /// mode highlight named-range references in real time without a
  /// per-keystroke backend round-trip.
  type ParsedNameRange = {
    sheet: number;
    r1: number;
    c1: number;
    r2: number;
    c2: number;
  };
  let nameCache = $state<Map<string, ParsedNameRange>>(new Map());
  /// Tells Grid to scroll a non-cursor cell into view. Set to a new
  /// object identity each time we want the scroll-on-target effect to
  /// re-fire.
  let scrollTarget = $state<{ row: number; col: number } | null>(null);

  /// Inline menu prompt — when set, the menu description bar becomes a
  /// single-line input. Used by menu actions that need a value (column
  /// width, row height, etc.). Esc cancels, Enter submits.
  let menuPrompt = $state<{
    label: string;
    value: string;
    onSubmit: (v: string) => void | Promise<void>;
    onCancel: (() => void | Promise<void>) | null;
    /// Optional autocomplete pool. When non-empty, the prompt renders a
    /// dropdown of matches. Tab cycles through them; Up/Down highlight;
    /// Enter submits with the highlighted suggestion (or raw input).
    candidates: string[] | null;
  } | null>(null);
  let menuPromptHighlight = $state(-1);

  /// Modal-mode instruction bar. Set to a string while a non-text modal
  /// is active (axis pick, /Copy or /Move destination cursor) so the
  /// keys-and-actions prompt renders in the menu-bar slot — same place
  /// as the inline menu prompt and the breadcrumb. Status bar is for
  /// non-modal feedback only; modal instructions belong here so the
  /// user always knows where to look. Cleared on commit/cancel.
  let menuMessage = $state<string | null>(null);

  /// Right-click context menu position (viewport pixels). Closed via Esc
  /// or any window click after the next render.
  let contextMenu = $state<{ x: number; y: number } | null>(null);
  let tabContextMenu = $state<{ x: number; y: number; sheet: number } | null>(null);
  function openContextMenu(x: number, y: number) {
    contextMenu = { x, y };
    tabContextMenu = null;
  }
  function closeContextMenu() {
    contextMenu = null;
  }
  function openTabContextMenu(x: number, y: number, sheet: number) {
    tabContextMenu = { x, y, sheet };
    contextMenu = null;
  }
  function closeTabContextMenu() {
    tabContextMenu = null;
  }
  $effect(() => {
    if (!contextMenu) return;
    const onClick = () => contextMenu = null;
    const t = setTimeout(() => window.addEventListener("click", onClick, { once: true }), 0);
    return () => {
      clearTimeout(t);
      window.removeEventListener("click", onClick);
    };
  });
  $effect(() => {
    if (!tabContextMenu) return;
    const onClick = () => tabContextMenu = null;
    const t = setTimeout(() => window.addEventListener("click", onClick, { once: true }), 0);
    return () => {
      clearTimeout(t);
      window.removeEventListener("click", onClick);
    };
  });
  let menuPromptEl: HTMLInputElement | null = $state(null);

  function openMenuPrompt(
    label: string,
    defaultValue: string,
    onSubmit: (v: string) => void | Promise<void>,
    onCancel?: () => void | Promise<void>,
    candidates?: string[],
  ) {
    menuPrompt = {
      label,
      value: defaultValue,
      onSubmit,
      onCancel: onCancel ?? null,
      candidates: candidates ?? null,
    };
    menuPromptHighlight = -1;
    tick().then(() => {
      menuPromptEl?.focus();
      menuPromptEl?.select();
    });
  }

  /// Filtered candidates list for the active prompt. Case-insensitive
  /// substring match keeps it permissive ("sum" matches "Summary").
  let promptMatches = $derived.by<string[]>(() => {
    if (!menuPrompt?.candidates) return [];
    const q = menuPrompt.value.toLowerCase();
    if (!q) return menuPrompt.candidates.slice(0, 12);
    return menuPrompt.candidates
      .filter((c) => c.toLowerCase().includes(q))
      .slice(0, 12);
  });

  async function submitMenuPrompt() {
    if (!menuPrompt) return;
    // If a candidate is highlighted, prefer it over the raw input.
    let value = menuPrompt.value;
    if (menuPromptHighlight >= 0 && promptMatches[menuPromptHighlight]) {
      value = promptMatches[menuPromptHighlight];
    }
    const p = menuPrompt;
    menuPrompt = null;
    await p.onSubmit(value);
  }

  /// Tab / Shift+Tab and Up/Down cycle through the suggestion list,
  /// like bash menu-complete. The input value stays as what the user
  /// typed so the match set doesn't collapse on the first Tab; the
  /// highlighted candidate is what `submitMenuPrompt` actually returns.
  function moveMenuPromptHighlight(delta: number) {
    if (!menuPrompt?.candidates) return;
    const n = promptMatches.length;
    if (n === 0) return;
    if (menuPromptHighlight < 0) {
      menuPromptHighlight = delta > 0 ? 0 : n - 1;
    } else {
      menuPromptHighlight = ((menuPromptHighlight + delta) % n + n) % n;
    }
  }

  async function cancelMenuPrompt() {
    if (!menuPrompt) return;
    const cb = menuPrompt.onCancel;
    menuPrompt = null;
    if (cb) {
      await cb();
    } else {
      statusMsg = "Cancelled";
      focusGrid();
    }
  }

  /// Grow the viewport to include (row, col) if either is past the
  /// current bound. Excel lets you navigate to any cell up to the
  /// absolute limits regardless of `dimension`; we mirror that by
  /// extending viewportRows / viewportCols on demand. Each grow adds
  /// a buffer so the user can keep moving without re-growing on every
  /// keystroke. Capped at Excel's hard limits.
  function growViewportToInclude(row: number, col: number) {
    if (row > viewportRows) {
      viewportRows = Math.min(ABS_MAX_ROWS, row + VIEWPORT_GROW_ROWS);
    }
    if (col > viewportCols) {
      viewportCols = Math.min(ABS_MAX_COLS, col + VIEWPORT_GROW_COLS);
    }
  }

  /// Ask the backend for the active sheet's used range and resize the
  /// viewport to match (the backend clamps to a sane min/max). Must be
  /// called after any sheet switch — otherwise the previous sheet's
  /// dimension would size the next one's request.
  async function resizeViewportToSheet() {
    if (!workbook) return;
    try {
      const [r, c] = await invoke<[number, number]>("get_sheet_dim", {
        sheet: activeSheet,
      });
      viewportRows = r;
      viewportCols = c;
    } catch {
      // Defaults already applied; non-fatal.
    }
  }

  // Paged cell fetch state. Rows are loaded lazily as the visible band
  // slides — Grid emits onBandChange and we top up missing ranges. The
  // loaded set is what's currently in the `cells` Map; clearing it
  // forces a re-fetch (used after edits / recalc / sheet switch).
  let loadedRows = new Set<number>();
  let bandStart = 1;
  let bandEnd = 0;
  let bandLoadRaf: number | null = null;
  // Currently-loaded column range. fetchRowBand fetches only this slice
  // of cols so navigating far right doesn't blow up per-fetch cost.
  // When the visible col band moves outside this range we invalidate
  // every loaded row and refetch with the new col window.
  let loadedColStart = 1;
  let loadedColEnd = 0;
  let colBandStart = 1;
  let colBandEnd = 0;
  let colBandLoadRaf: number | null = null;
  // Ticks each time loaded data is invalidated. fetchRowBand calls
  // started before the bump are stale and discard their results — without
  // this an in-flight fetch from the previous sheet (or pre-edit state)
  // would clobber the just-loaded fresh data when its IPC returns.
  let loadGen = 0;

  /// Fetch sheet-wide layout: column widths, frozen panes, merged ranges.
  /// Row heights are NOT pulled here — they come per-band via fetchRowBand.
  async function loadSheetLayout() {
    if (!workbook) return;
    const layout = await invoke<LayoutData>("get_layout", {
      sheet: activeSheet,
      startRow: 1,
      endRow: 1,
      startCol: 1,
      endCol: viewportCols,
    });
    const cw = new Map<number, number>();
    for (const [c, w] of layout.col_widths) cw.set(c, colWidthPx(w));
    colWidths = cw;
    frozenRows = layout.frozen_rows;
    frozenCols = layout.frozen_cols;
    mergedRanges = layout.merged_ranges;
  }

  /// Fetch one contiguous row range — but only across the currently-
  /// loaded column slice (loadedColStart..loadedColEnd). Cells outside
  /// that slice are left out of the Map; they render as empty, which is
  /// usually correct. Horizontal scroll triggers a separate path
  /// (handleColBandChange) that refreshes the loaded col window when
  /// the visible band moves outside it.
  async function fetchRowBand(r1: number, r2: number) {
    if (!workbook || r1 > r2 || loadedColStart > loadedColEnd) return;
    const sheetAtFetch = activeSheet;
    const genAtFetch = loadGen;
    const c1 = loadedColStart;
    const c2 = loadedColEnd;
    const [list, layout] = await Promise.all([
      invoke<CellView[]>("get_cells", {
        sheet: sheetAtFetch,
        startRow: r1,
        endRow: r2,
        startCol: c1,
        endCol: c2,
      }),
      invoke<LayoutData>("get_layout", {
        sheet: sheetAtFetch,
        startRow: r1,
        endRow: r2,
        startCol: c1,
        endCol: c2,
      }),
    ]);
    if (sheetAtFetch !== activeSheet || genAtFetch !== loadGen) return;
    // Drop any previously-cached cells in (r1..r2, c1..c2) first so
    // deletions / formula re-evaluations propagate (an empty-cell-elided
    // cell that BECAME empty would otherwise persist).
    const newCells = new Map(cells);
    for (let r = r1; r <= r2; r++) {
      for (let c = c1; c <= c2; c++) newCells.delete(key(r, c));
    }
    for (const c of list) newCells.set(key(c.row, c.col), c);
    cells = newCells;
    const newRH = new Map(rowHeights);
    for (const [r, h] of layout.row_heights) newRH.set(r, rowHeightPx(h));
    rowHeights = newRH;
    for (let r = r1; r <= r2; r++) loadedRows.add(r);
  }

  /// Walk [r1..r2], find each contiguous gap of un-loaded rows, fetch it.
  /// Awaits sequentially — the gaps are short enough (one band at most)
  /// that parallelism isn't worth the IPC overhead.
  async function ensureRowsLoaded(r1: number, r2: number) {
    if (!workbook || r1 > r2) return;
    let r = r1;
    while (r <= r2) {
      while (r <= r2 && loadedRows.has(r)) r++;
      if (r > r2) break;
      const start = r;
      while (r <= r2 && !loadedRows.has(r)) r++;
      await fetchRowBand(start, r - 1);
    }
  }

  /// Invalidate just the rows in [r1..r2] and re-fetch them. Used by
  /// edit paths that touched a known row range — much cheaper than a
  /// full refreshViewport because we don't drop the rest of the loaded
  /// band. Spilled-formula dependents in OTHER rows aren't refreshed
  /// (set_cell doesn't recalc — that's F9), so this is correct: only
  /// the directly-edited cells changed.
  async function refreshRows(r1: number, r2: number) {
    if (!workbook) return;
    const lo = Math.max(1, Math.min(r1, r2));
    const hi = Math.min(viewportRows, Math.max(r1, r2));
    if (lo > hi) return;
    for (let r = lo; r <= hi; r++) loadedRows.delete(r);
    await ensureRowsLoaded(lo, hi);
  }

  /// Invalidate cached cell data and pull the currently-visible band.
  /// Used after any operation that may have changed cell values: edits,
  /// recalc, sheet switch, sort, fill, insert/delete rows, etc. Callers
  /// that know the affected range can use refreshRows for a narrower
  /// invalidation, but the safe default is a full reload.
  ///
  /// `clear`:
  ///   - `false` (default) — leave the existing cells map visible while
  ///     the fresh data fetches in the background. Recalc and edits
  ///     stay on-screen during the IPC; the new data overwrites in
  ///     place once it arrives. This is what stops the "screen flash"
  ///     on F9 / save / fill / etc.
  ///   - `true` — wipe the map up front. Use this only when the cells
  ///     showing now would be wrong to keep visible: sheet switch,
  ///     opening a different workbook, etc.
  async function refreshViewport(opts: { clear?: boolean } = {}) {
    if (!workbook) return;
    const clear = opts.clear ?? false;
    loadGen++;
    if (clear) {
      cells = new Map();
      rowHeights = new Map();
    }
    loadedRows = new Set();
    await loadSheetLayout();
    loadedColStart = Math.max(1, colBandStart);
    loadedColEnd = Math.min(viewportCols, colBandEnd > 0 ? colBandEnd : 60);
    const r1 = Math.max(1, bandStart);
    const r2 = Math.min(viewportRows, bandEnd > 0 ? bandEnd : 60);
    await ensureRowsLoaded(r1, r2);
  }

  /// Grid's visible-band edges shifted (scroll, resize, sheet switch).
  /// rAF-coalesce so a fast scroll doesn't fire one fetch per frame.
  function handleBandChange(start: number, end: number) {
    bandStart = start;
    bandEnd = end;
    if (bandLoadRaf !== null) return;
    bandLoadRaf = requestAnimationFrame(async () => {
      bandLoadRaf = null;
      if (!workbook) return;
      const r1 = Math.max(1, bandStart);
      const r2 = Math.min(viewportRows, bandEnd);
      await ensureRowsLoaded(r1, r2);
    });
  }

  /// Visible col band shifted. If it's already inside the loaded col
  /// window, nothing to do. Otherwise: WIDEN the loaded col window to
  /// include the new band and fetch only the missing slice — keeping
  /// already-loaded cells visible while the new ones load.
  ///
  /// The previous version cleared `cells = new Map()` synchronously
  /// before awaiting the IPC, which painted the grid blank for the
  /// fetch's full duration (often tens of ms on a heavy sheet — the
  /// "Main! goes blank when scrolling" the user sees, only fixed by a
  /// recalc which triggers a full refetch). Cells outside the visible
  /// band don't render anyway, so leaving them in the Map is free
  /// memory-wise (~277 cols * 1000 rows worst case is fine) and lets
  /// the grid keep showing the cells it has while we top up.
  function handleColBandChange(start: number, end: number) {
    colBandStart = start;
    colBandEnd = end;
    if (colBandLoadRaf !== null) return;
    colBandLoadRaf = requestAnimationFrame(async () => {
      colBandLoadRaf = null;
      if (!workbook) return;
      const want_start = Math.max(1, colBandStart);
      const want_end = Math.min(viewportCols, colBandEnd);
      if (want_start > want_end) return;
      if (loadedColStart <= want_start && loadedColEnd >= want_end) return;
      // Compute which slice(s) of cols still need fetching. Most
      // common case is "scrolled right by one viewport" → one slice
      // on the right; can be either side or both.
      const sheetAtFetch = activeSheet;
      const r1 = Math.max(1, bandStart);
      const r2 = Math.min(viewportRows, bandEnd);
      const slices: Array<[number, number]> = [];
      if (loadedColEnd < loadedColStart) {
        // No window loaded yet (initial open) — fetch everything.
        slices.push([want_start, want_end]);
      } else {
        if (want_start < loadedColStart) {
          slices.push([want_start, Math.min(loadedColStart - 1, want_end)]);
        }
        if (want_end > loadedColEnd) {
          slices.push([Math.max(loadedColEnd + 1, want_start), want_end]);
        }
      }
      const newColStart = Math.min(loadedColStart, want_start);
      const newColEnd = Math.max(loadedColEnd, want_end);
      const genAtFetch = loadGen;
      // Fetch each missing col slice across the currently-visible row
      // band only. Other rows scroll into view via handleBandChange,
      // which now sees an expanded loaded col window and fetches the
      // right span without dropping anything.
      for (const [c1, c2] of slices) {
        const [list, layout] = await Promise.all([
          invoke<CellView[]>("get_cells", {
            sheet: sheetAtFetch,
            startRow: r1,
            endRow: r2,
            startCol: c1,
            endCol: c2,
          }),
          invoke<LayoutData>("get_layout", {
            sheet: sheetAtFetch,
            startRow: r1,
            endRow: r2,
            startCol: c1,
            endCol: c2,
          }),
        ]);
        if (sheetAtFetch !== activeSheet || genAtFetch !== loadGen) return;
        const newCells = new Map(cells);
        for (const c of list) newCells.set(key(c.row, c.col), c);
        cells = newCells;
        const newRH = new Map(rowHeights);
        for (const [r, h] of layout.row_heights) newRH.set(r, rowHeightPx(h));
        rowHeights = newRH;
      }
      // Now the new col span is loaded for the visible row band; the
      // loaded col window expands to cover it, and the loaded-rows
      // set is invalidated for rows we haven't yet refetched at the
      // new wider span (so they get topped up next handleBandChange).
      loadedColStart = newColStart;
      loadedColEnd = newColEnd;
      // Mark only the rows we just fetched as covered for this wider
      // col window. Rows OUTSIDE [r1, r2] still hold stale narrower
      // data, but they're off-screen — when they scroll into view
      // ensureRowsLoaded will refetch them at the new span.
      loadedRows = new Set();
      for (let r = r1; r <= r2; r++) loadedRows.add(r);
    });
  }

  async function newWorkbook() {
    workbook = await invoke<WorkbookInfo>("new_workbook");
    activeSheet = 0;
    selRow = 1;
    selCol = 1;
    rangeEndRow = 1;
    rangeEndCol = 1;
    currentPath = "";
    history = [];
    historyIdx = 0;
    sheetCursors = new Map();
    await resizeViewportToSheet();
    await refreshViewport({ clear: true });
    await refreshNameCache();
    statusMsg = "New workbook";
  }

  async function openWorkbookFromPath(path: string) {
    const p = path.trim();
    if (!p) {
      statusMsg = "Enter a path then press Enter";
      return;
    }
    try {
      workbook = await invoke<WorkbookInfo>("open_workbook", { path: p });
      activeSheet = 0;
      selRow = 1;
      selCol = 1;
      rangeEndRow = 1;
      rangeEndCol = 1;
      currentPath = p;
      history = [];
      historyIdx = 0;
      sheetCursors = new Map();
      await resizeViewportToSheet();
      await refreshViewport({ clear: true });
      await refreshNameCache();
      statusMsg = `Opened ${p} (sheets: ${workbook.sheet_names.join(", ")})`;
      focusGrid();
    } catch (e) {
      statusMsg = `Open failed: ${e}`;
    }
  }

  function focusGrid() {
    tick().then(() => gridWrapEl?.focus());
  }

  function describeSave(r: SaveResult): string {
    if (r.mode === "preserved") {
      return `Saved ${r.path} · ${r.cells_patched} cell${r.cells_patched === 1 ? "" : "s"} patched in place (charts/pivots/drawings preserved)`;
    }
    if (r.mode === "xls") {
      return `Saved ${r.path} · BIFF8 .xls (formulas currently saved as cached values; charts/macros not preserved)`;
    }
    return `Saved ${r.path} · ⚠ written via IronCalc — features it doesn't understand (charts, pivots, drawings) were lost`;
  }

  async function saveWorkbookToPath(path: string) {
    if (!workbook) {
      statusMsg = "No workbook to save";
      return;
    }
    // The Rust save_workbook command dispatches on extension: .xls →
    // BIFF8 writer (xls_save.rs), .xlsx → preservation patcher or
    // IronCalc save_to_xlsx fallback.
    try {
      const r = await invoke<SaveResult>("save_workbook", { path });
      currentPath = path;
      statusMsg = describeSave(r);
    } catch (e) {
      statusMsg = `Save failed: ${e}`;
    }
  }

  /// Save-As path picked from the navigator: if it already exists, pop a
  /// Replace/Cancel confirm before clobbering. New filename → save direct.
  async function saveAsWithConfirm(path: string) {
    let exists = false;
    try {
      exists = await invoke<boolean>("file_exists", { path });
    } catch {
      exists = false;
    }
    if (!exists) {
      await saveWorkbookToPath(path);
      return;
    }
    dynamicTitle = `Replace ${path}`;
    dynamicLevel = [
      {
        letter: "R",
        label: "Replace",
        description: `Overwrite existing ${path}`,
        action: () => saveWorkbookToPath(path),
      },
      {
        letter: "C",
        label: "Cancel",
        description: "Don't save",
        action: () => { statusMsg = "Save cancelled"; },
      },
    ];
    menuOpen = true;
    menuPath = [];
    menuHighlight = 0;
  }

  async function backupAndSave(path: string) {
    try {
      const r = await invoke<BackupResult>("backup_and_save", { path });
      currentPath = path;
      statusMsg = `${describeSave(r.save)} · backup: ${r.backup_path}`;
    } catch (e) {
      statusMsg = `Backup failed: ${e}`;
    }
  }

  function openRetrieveNavigator() {
    navMode = "open";
    navOpen = true;
  }

  function openSaveAsNavigator() {
    navMode = "save";
    navOpen = true;
  }

  /// Lotus /F S flow. With no current path → Save As navigator.
  /// With current path that exists → Replace/SaveAs/Backup/Cancel picker.
  /// With current path that doesn't exist → straight save (new file).
  async function fileSaveFlow() {
    if (!currentPath) {
      openSaveAsNavigator();
      return;
    }
    let exists = false;
    try {
      exists = await invoke<boolean>("file_exists", { path: currentPath });
    } catch {
      exists = false;
    }
    if (!exists) {
      await saveWorkbookToPath(currentPath);
      return;
    }
    dynamicTitle = `Save ${currentPath}`;
    dynamicLevel = saveMenuItems({
      replace: () => saveWorkbookToPath(currentPath),
      saveAs: openSaveAsNavigator,
      backup: () => backupAndSave(currentPath),
      cancel: () => { statusMsg = "Save cancelled"; },
    });
    menuOpen = true;
    menuPath = [];
    menuHighlight = 0;
  }

  async function eraseCurrentCell() {
    const r1 = Math.min(selRow, rangeEndRow);
    const r2 = Math.max(selRow, rangeEndRow);
    const c1 = Math.min(selCol, rangeEndCol);
    const c2 = Math.max(selCol, rangeEndCol);
    const sheet = activeSheet;
    const edits: EditOp[] = [];
    for (let r = r1; r <= r2; r++) {
      for (let c = c1; c <= c2; c++) {
        const prev = cells.get(key(r, c))?.input ?? "";
        if (prev !== "") edits.push({ row: r, col: c, prev, next: "" });
      }
    }
    try {
      for (const op of edits) {
        await invoke("set_cell", { sheet, row: op.row, col: op.col, value: "" });
      }
      await refreshRows(r1, r2);
      dirtyEdits += edits.length;
      const span =
        r1 === r2 && c1 === c2 ? addr(r1, c1) : `${addr(r1, c1)}:${addr(r2, c2)}`;
      pushHistory({ description: `Erase ${span}`, sheet, edits });
      statusMsg = `Erased ${span}`;
    } catch (e) {
      statusMsg = `Erase failed: ${e}`;
    }
  }

  async function quitApp() {
    try {
      const w = await import("@tauri-apps/api/window");
      await w.getCurrentWindow().close();
    } catch (e) {
      statusMsg = `Quit failed: ${e}`;
    }
  }

  async function commitEdit() {
    if (!editing) return;
    const prev = cells.get(key(selRow, selCol))?.input ?? "";
    const next = editValue;
    const sheet = activeSheet;
    const r = selRow;
    const c = selCol;
    try {
      await invoke<string>("set_cell", { sheet, row: r, col: c, value: next });
      await refreshRows(r, r);
      dirtyEdits += 1;
      if (prev !== next) {
        pushHistory({
          description: `Set ${addr(r, c)}`,
          sheet,
          edits: [{ row: r, col: c, prev, next }],
        });
      }
      statusMsg = `Set ${addr(r, c)} = ${next} · F9 to recalc`;
    } catch (e) {
      statusMsg = `Set failed: ${e}`;
    }
    editing = false;
    editValue = "";
    refHighlights = [];
    focusGrid();
  }

  async function jumpEdge(dr: number, dc: number) {
    if (!workbook) return;
    try {
      const [r, c] = await invoke<[number, number]>("jump_edge", {
        sheet: activeSheet,
        row: selRow,
        col: selCol,
        dr,
        dc,
      });
      selRow = r;
      selCol = c;
      rangeEndRow = r;
      rangeEndCol = c;
    } catch (e) {
      statusMsg = `jump failed: ${e}`;
    }
  }

  async function recalcWorkbook() {
    if (!workbook) return;
    statusMsg = "Recalculating…";
    try {
      const ms = await invoke<number>("recalc");
      await refreshViewport();
      dirtyEdits = 0;
      statusMsg = `Recalc done in ${ms} ms`;
    } catch (e) {
      statusMsg = `Recalc failed: ${e}`;
    }
  }

  /// /File Compare Open — pick a file via the navigator, then load+diff.
  function compareOpen() {
    navMode = "open";
    navCompareTarget = true;
    navOpen = true;
    statusMsg = "Pick a file to compare against";
  }

  async function compareOpenWith(p: string) {
    try {
      const r = await invoke<CompareResult>("compare_open", { path: p });
      compareResult = r;
      compareHidden = false;
      const cap = r.diffs_capped ? ` (showing ${r.diffs.length})` : "";
      statusMsg = `Compare: ${r.total_diffs} diff${r.total_diffs === 1 ? "" : "s"}${cap}`;
    } catch (e) {
      statusMsg = `Compare failed: ${e}`;
      compareResult = null;
    }
  }

  /// /File Compare Exit — close the comparison and clear the dock.
  async function compareExit() {
    if (!compareResult) {
      statusMsg = "No active comparison";
      return;
    }
    try {
      await invoke("compare_close");
    } catch (e) {
      // Backend should never fail to close; not worth blocking the UI.
      console.warn("compare_close failed:", e);
    }
    compareResult = null;
    compareHidden = false;
    statusMsg = "Compare closed";
    focusGrid();
  }

  /// Jump cursor to a diff cell. Switches sheet if the diff is on a
  /// different one, then moves selection.
  async function compareJumpTo(d: CompareDiffType) {
    if (d.sheet_idx === null) return;
    if (d.sheet_idx !== activeSheet) {
      await switchSheet(d.sheet_idx);
    }
    selRow = d.row;
    selCol = d.col;
    rangeEndRow = d.row;
    rangeEndCol = d.col;
    growViewportToInclude(d.row, d.col);
  }

  /// Preview = highlight without moving cursor. Just like trace
  /// preview: switch the active sheet so the user sees context, but
  /// leave the actual selection alone (it's restored when compare
  /// exits via the trace popup's origin-snapshot pattern is not
  /// applicable here — the user explicitly entered compare mode).
  async function comparePreview(d: CompareDiffType) {
    if (d.sheet_idx === null) return;
    if (d.sheet_idx !== activeSheet) {
      await switchSheet(d.sheet_idx);
    }
    growViewportToInclude(d.row, d.col);
    // Fire a one-shot scroll target so the grid centers on the diff
    // cell without changing the cursor.
    scrollTarget = { row: d.row, col: d.col };
  }

  /// Highlight overlays for every visible diff on the active sheet.
  /// Color tints by kind so the user can spot value vs formula vs
  /// missing at a glance. Capped at the diff list size — not
  /// lazy-rendered yet because the grid handles ~5000 overlays fine
  /// (see existing refHighlights pipeline).
  let compareHighlights = $derived.by(() => {
    if (!compareResult) return [];
    const out: typeof refHighlights = [];
    for (const d of compareResult.diffs) {
      if (d.sheet_idx === null) continue;
      let color: string;
      switch (d.kind) {
        case "value":
          color = "#f88";
          break;
        case "formula":
          color = "#6cf";
          break;
        case "missing-left":
        case "missing-right":
          color = "#c8a060";
          break;
      }
      out.push({
        sheet: d.sheet_idx,
        r1: d.row,
        c1: d.col,
        r2: d.row,
        c2: d.col,
        color,
      });
    }
    return out;
  });

  /// /Trace Trace — open the dependency-tree popup for the current cell.
  async function traceFormula() {
    if (!workbook) return;
    try {
      const root = await invoke<TraceNode>("trace_formula", {
        sheet: activeSheet,
        row: selRow,
        col: selCol,
      });
      // Snapshot where the user was so Esc-close can put them back.
      traceOriginCursor = {
        sheet: activeSheet,
        selRow,
        selCol,
        rangeEndRow,
        rangeEndCol,
      };
      traceRoot = root;
    } catch (e) {
      statusMsg = `Trace failed: ${e}`;
    }
  }

  /// Called by FormulaTrace as the user moves through the list. We
  /// switch sheets if needed and scroll the grid to the highlighted
  /// item's cell — without changing the active selection cursor on
  /// either sheet. Highlights show what the popup is pointing at;
  /// for `name` kind nodes the highlight covers the full resolved
  /// range and carries the name as a tag.
  async function tracePreview(node: TraceNode) {
    if (node.kind === "name") {
      // Defined name — value field carries the resolved formula text
      // (e.g. "Discount!$B$24:$W$35"). Parse it and highlight the
      // full range with the name as a label.
      const range = parseNameFormula(node.value);
      if (!range) {
        refHighlights = [];
        scrollTarget = null;
        return;
      }
      if (range.sheet !== activeSheet) await switchSheet(range.sheet);
      refHighlights = [
        {
          sheet: range.sheet,
          r1: range.r1,
          c1: range.c1,
          r2: range.r2,
          c2: range.c2,
          color: "#ff8800",
          label: node.address,
        },
      ];
      scrollTarget = { row: range.r1, col: range.c1 };
      return;
    }
    if (node.sheet === null || node.row === null || node.col === null) {
      refHighlights = [];
      scrollTarget = null;
      return;
    }
    const targetSheet = node.sheet;
    if (targetSheet !== activeSheet) {
      await switchSheet(targetSheet);
    }
    refHighlights = [
      {
        sheet: targetSheet,
        r1: node.row,
        c1: node.col,
        r2: node.row,
        c2: node.col,
        color: "#ff8800",
      },
    ];
    scrollTarget = { row: node.row, col: node.col };
  }

  async function closeTrace(restoreCursor: boolean) {
    traceRoot = null;
    traceDocked = false;
    traceHidden = false;
    refHighlights = [];
    scrollTarget = null;
    if (restoreCursor && traceOriginCursor) {
      const o = traceOriginCursor;
      if (o.sheet !== activeSheet) await switchSheet(o.sheet);
      selRow = o.selRow;
      selCol = o.selCol;
      rangeEndRow = o.rangeEndRow;
      rangeEndCol = o.rangeEndCol;
    }
    traceOriginCursor = null;
    focusGrid();
  }

  /// /Trace Goto — list the top-level dependencies of the current cell's
  /// formula and prompt the user to jump to one. Reuses the menu prompt
  /// candidates pattern so typing filters the list.
  async function traceGoto() {
    if (!workbook) return;
    let root: TraceNode;
    try {
      root = await invoke<TraceNode>("trace_formula", {
        sheet: activeSheet,
        row: selRow,
        col: selCol,
      });
    } catch (e) {
      statusMsg = `Trace failed: ${e}`;
      return;
    }
    if (root.deps.length === 0) {
      statusMsg = `${root.address} has no dependencies (kind: ${root.kind})`;
      return;
    }
    // Build candidate strings — "Discount!E20 = Hybrid Structured Luff" —
    // and a parallel index of jump targets so submission can resolve
    // back to the chosen dep.
    const candidates: string[] = [];
    const targets: TraceNode[] = [];
    for (const d of root.deps) {
      const value = d.value || "(empty)";
      candidates.push(`${d.address}  =  ${value}`);
      targets.push(d);
    }
    openMenuPrompt(
      `Jump to dependency of ${root.address}:`,
      "",
      async (v) => {
        // Match by exact candidate (the user picked from the list)
        // or by leading-substring on address.
        const idx = candidates.findIndex((c) => c === v) >= 0
          ? candidates.findIndex((c) => c === v)
          : targets.findIndex((t) => t.address.toLowerCase().startsWith(v.trim().toLowerCase()));
        if (idx < 0) {
          statusMsg = `No matching dependency: ${v}`;
          focusGrid();
          return;
        }
        const t = targets[idx];
        if (t.kind === "name") {
          // Defined name → jumpToAddress will resolve it via list_names.
          const ok = await jumpToAddress(t.address);
          if (!ok) statusMsg = `Could not resolve ${t.address}`;
        } else if (t.sheet !== null && t.row !== null && t.col !== null) {
          if (t.sheet !== activeSheet) await switchSheet(t.sheet);
          selRow = t.row;
          selCol = t.col;
          rangeEndRow = t.row;
          rangeEndCol = t.col;
          growViewportToInclude(t.row, t.col);
        }
        focusGrid();
      },
      undefined,
      candidates,
    );
  }

  /// /Trace Names — browse all defined names with their resolved
  /// locations and jump to the chosen one.
  async function traceNames() {
    if (!workbook) return;
    let names: NamedRangeInfo[];
    try {
      names = await invoke<NamedRangeInfo[]>("list_named_ranges");
    } catch (e) {
      statusMsg = `List names failed: ${e}`;
      return;
    }
    if (names.length === 0) {
      statusMsg = "No defined names in this workbook";
      return;
    }
    const candidates = names.map(
      (n) => `${n.name}  →  ${n.formula}  ${n.scope}`,
    );
    openMenuPrompt(
      `Jump to named range (${names.length} total):`,
      "",
      async (v) => {
        const exact = candidates.indexOf(v);
        const idx = exact >= 0
          ? exact
          : names.findIndex((n) => n.name.toLowerCase() === v.trim().toLowerCase());
        if (idx < 0) {
          statusMsg = `No matching name: ${v}`;
          focusGrid();
          return;
        }
        const n = names[idx];
        if (n.jump_sheet === null || n.jump_row === null || n.jump_col === null) {
          statusMsg = `${n.name}: cannot jump (formula = ${n.formula})`;
          focusGrid();
          return;
        }
        if (n.jump_sheet !== activeSheet) await switchSheet(n.jump_sheet);
        selRow = n.jump_row;
        selCol = n.jump_col;
        rangeEndRow = n.jump_row;
        rangeEndCol = n.jump_col;
        growViewportToInclude(n.jump_row, n.jump_col);
        focusGrid();
      },
      undefined,
      candidates,
    );
  }

  function startEdit(seed?: string) {
    if (!workbook) return;
    if (seed !== undefined) {
      editValue = seed;
    } else {
      const c = cells.get(key(selRow, selCol));
      editValue = c?.input ?? "";
    }
    editing = true;
    updateEditHighlights();
    // Explicit focus after the input is in the DOM — autofocus is flaky
    // in Svelte 5 when the element conditionally renders.
    tick().then(() => {
      editorEl?.focus();
      if (seed === undefined) editorEl?.select();
    });
  }

  function cancelEdit() {
    editing = false;
    editValue = "";
    refHighlights = [];
  }

  /// Parse `editValue` as an in-progress formula and update
  /// `refHighlights` so the user can see what cells / ranges the
  /// formula points at while they type. Limited to A1-style refs
  /// (with optional sheet prefix) — defined names are not resolved
  /// here yet (would need a per-keystroke list_names lookup).
  /// Strings inside `"..."` are stripped first so cells named like
  /// "M3" inside a string literal don't trigger highlights.
  ///
  /// Color cycles through a small palette so multiple distinct refs
  /// in the same formula get different colors — easier to read at a
  /// glance.
  function updateEditHighlights() {
    if (!editing || !workbook || !editValue.startsWith("=")) {
      refHighlights = [];
      return;
    }
    const palette = ["#0a84ff", "#34c759", "#ff9500", "#bf5af2", "#ff375f", "#5ac8fa"];
    const stripped = editValue.replace(/"[^"]*"/g, '""');
    const out: typeof refHighlights = [];
    let colorIdx = 0;

    // Pass 1: A1-style cells and ranges, with optional sheet prefix.
    // The sheet prefix can be quoted ('Sheet One') or bare (Sheet1).
    // We track consumed character spans so the name pass below can
    // skip over identifiers already accounted for here.
    const consumed: Array<[number, number]> = [];
    const cellRe = /(?:'([^']+)'!|([A-Za-z_][A-Za-z_0-9.]*)!)?(\$?[A-Z]{1,2}\$?\d+)(?::(\$?[A-Z]{1,2}\$?\d+))?/g;
    let m: RegExpExecArray | null;
    while ((m = cellRe.exec(stripped)) !== null) {
      consumed.push([m.index, m.index + m[0].length]);
      const sheetName = m[1] ?? m[2] ?? null;
      let sheetIdx = activeSheet;
      if (sheetName) {
        const idx = workbook.sheet_names.indexOf(sheetName);
        if (idx < 0) continue;
        sheetIdx = idx;
      }
      const startCell = parseA1Frontend(m[3]);
      if (!startCell) continue;
      const endCell = m[4] ? parseA1Frontend(m[4]) : startCell;
      if (!endCell) continue;
      out.push({
        sheet: sheetIdx,
        r1: Math.min(startCell.row, endCell.row),
        c1: Math.min(startCell.col, endCell.col),
        r2: Math.max(startCell.row, endCell.row),
        c2: Math.max(startCell.col, endCell.col),
        color: palette[colorIdx % palette.length],
      });
      colorIdx++;
    }

    // Pass 2: bare identifiers that match a defined name in the
    // cache. Skip identifiers that are followed by `(` (function
    // calls), preceded by `!` (sheet qualifier — already handled),
    // or fall inside a span consumed by a cell match above.
    const nameRe = /[A-Za-z_][A-Za-z_0-9.]*/g;
    while ((m = nameRe.exec(stripped)) !== null) {
      const start = m.index;
      const end = start + m[0].length;
      // Skip if any consumed span overlaps this identifier.
      if (consumed.some(([a, b]) => start < b && end > a)) continue;
      // Skip function calls.
      if (stripped[end] === "(") continue;
      // Skip sheet qualifier head (e.g. "Sheet1!" — should have been
      // handled in Pass 1, but be defensive).
      if (stripped[end] === "!") continue;
      const range = nameCache.get(m[0].toLowerCase());
      if (!range) continue;
      out.push({
        sheet: range.sheet,
        r1: range.r1,
        c1: range.c1,
        r2: range.r2,
        c2: range.c2,
        color: palette[colorIdx % palette.length],
        label: m[0],
      });
      colorIdx++;
    }

    refHighlights = out;
  }

  // Keep edit-mode highlights in sync with the input value.
  $effect(() => {
    editValue;
    if (editing) updateEditHighlights();
  });

  /// Excel F4: cycle the cell reference under the caret through
  /// A1 → $A$1 → A$1 → $A1 → A1. Token detection is liberal — anything
  /// shaped like [letters][digits] with optional $ prefixes counts.
  function f4Toggle() {
    if (!editorEl) return;
    const text = editValue;
    const caret = editorEl.selectionStart ?? text.length;
    const isTokenChar = (c: string) => /[A-Za-z0-9$]/.test(c);
    let start = caret;
    while (start > 0 && isTokenChar(text[start - 1])) start--;
    let end = caret;
    while (end < text.length && isTokenChar(text[end])) end++;
    const token = text.slice(start, end);
    const m = token.match(/^(\$?)([A-Za-z]+)(\$?)(\d+)$/);
    if (!m) return;
    const [, dollarCol, letters, dollarRow, digits] = m;
    const colLocked = dollarCol === "$";
    const rowLocked = dollarRow === "$";
    let newColLocked: boolean;
    let newRowLocked: boolean;
    if (!colLocked && !rowLocked) {
      newColLocked = true;
      newRowLocked = true;
    } else if (colLocked && rowLocked) {
      newColLocked = false;
      newRowLocked = true;
    } else if (!colLocked && rowLocked) {
      newColLocked = true;
      newRowLocked = false;
    } else {
      newColLocked = false;
      newRowLocked = false;
    }
    const newToken =
      (newColLocked ? "$" : "") + letters + (newRowLocked ? "$" : "") + digits;
    editValue = text.slice(0, start) + newToken + text.slice(end);
    const newCaret = start + newToken.length;
    tick().then(() => editorEl?.setSelectionRange(newCaret, newCaret));
  }

  /// Lotus `.` (period) — cycle which corner of the active selection
  /// rectangle is the anchor. After a press, Shift+arrow extends from
  /// the new anchor; the fill handle (which renders at the rangeEnd
  /// corner) moves to the opposite corner. Cycles clockwise:
  /// TL→TR→BR→BL→TL. With a single-cell selection, this is a no-op.
  function cycleAnchor() {
    if (selRow === rangeEndRow && selCol === rangeEndCol) return;
    const r1 = Math.min(selRow, rangeEndRow);
    const r2 = Math.max(selRow, rangeEndRow);
    const c1 = Math.min(selCol, rangeEndCol);
    const c2 = Math.max(selCol, rangeEndCol);
    // Identify current anchor corner.
    const atTop = selRow === r1;
    const atLeft = selCol === c1;
    let next: { r: number; c: number; er: number; ec: number };
    if (atTop && atLeft) {
      // TL → TR
      next = { r: r1, c: c2, er: r2, ec: c1 };
    } else if (atTop && !atLeft) {
      // TR → BR
      next = { r: r2, c: c2, er: r1, ec: c1 };
    } else if (!atTop && !atLeft) {
      // BR → BL
      next = { r: r2, c: c1, er: r1, ec: c2 };
    } else {
      // BL → TL
      next = { r: r1, c: c1, er: r2, ec: c2 };
    }
    selRow = next.r;
    selCol = next.c;
    rangeEndRow = next.er;
    rangeEndCol = next.ec;
  }

  // rAF-coalesced arrow moves. Key repeat at 30-60Hz can outrun our
  // per-press render cycle (Svelte reactive update + scroll-into-view
  // + band shift + flushSync). If that happens, events stack up and
  // scrolling goes jerky because the browser can't drain the queue.
  // Collect pending delta into a scratch pair and apply once per
  // animation frame — the browser paces us to display refresh rate,
  // events never queue past one frame's worth of input.
  let pendingMoveDr = 0;
  let pendingMoveDc = 0;
  let pendingMoveExtend = false;
  let moveRafId: number | null = null;

  function moveSel(dr: number, dc: number, extend = false) {
    pendingMoveDr += dr;
    pendingMoveDc += dc;
    pendingMoveExtend = extend;
    if (moveRafId !== null) return;
    moveRafId = requestAnimationFrame(() => {
      moveRafId = null;
      const r = pendingMoveDr;
      const c = pendingMoveDc;
      const ext = pendingMoveExtend;
      pendingMoveDr = 0;
      pendingMoveDc = 0;
      if (r === 0 && c === 0) return;
      if (ext) {
        rangeEndRow = Math.max(1, Math.min(ABS_MAX_ROWS, rangeEndRow + r));
        rangeEndCol = Math.max(1, Math.min(ABS_MAX_COLS, rangeEndCol + c));
        growViewportToInclude(rangeEndRow, rangeEndCol);
      } else {
        selRow = Math.max(1, Math.min(ABS_MAX_ROWS, selRow + r));
        selCol = Math.max(1, Math.min(ABS_MAX_COLS, selCol + c));
        rangeEndRow = selRow;
        rangeEndCol = selCol;
        growViewportToInclude(selRow, selCol);
      }
    });
  }

  /// Client-side undo/redo. Each completed user edit pushes one
  /// `UndoEntry` recording the prior + new value of every affected cell;
  /// undo replays prev values, redo replays next. We capture before the
  /// mutation runs so prev is always accurate even after refreshViewport
  /// has overwritten our local cells map.
  type EditOp = { row: number; col: number; prev: string; next: string };
  type StyleEdit = {
    r1: number; c1: number; r2: number; c2: number;
    prev_indices: number[];
    next_indices: number[];
  };
  type UndoEntry =
    | { kind: "value"; description: string; sheet: number; edits: EditOp[] }
    | { kind: "style"; description: string; sheet: number; edit: StyleEdit };
  let history = $state<UndoEntry[]>([]);
  let historyIdx = $state(0);

  function pushHistory(entry: { description: string; sheet: number; edits: EditOp[] } | UndoEntry) {
    let normalized: UndoEntry;
    if ("kind" in entry) {
      normalized = entry;
    } else {
      if (entry.edits.length === 0) return;
      normalized = { kind: "value", ...entry };
    }
    if (historyIdx < history.length) history = history.slice(0, historyIdx);
    history = [...history, normalized];
    historyIdx = history.length;
  }

  async function applyEdits(sheet: number, ops: { row: number; col: number; value: string }[]) {
    let failed = 0;
    let lastErr: unknown = null;
    for (const op of ops) {
      try {
        await invoke("set_cell", { sheet, row: op.row, col: op.col, value: op.value });
      } catch (e) {
        failed++;
        lastErr = e;
      }
    }
    if (ops.length > 0) {
      const r1 = Math.min(...ops.map((o) => o.row));
      const r2 = Math.max(...ops.map((o) => o.row));
      await refreshRows(r1, r2);
    }
    if (failed > 0) {
      // Surface the failure so undo/redo doesn't silently lie. Most
      // common cause: a formula now references a deleted defined name
      // / sheet, or the cell is in a sheet that no longer exists.
      statusMsg = `Edit failed for ${failed}/${ops.length} cell${ops.length === 1 ? "" : "s"}: ${lastErr}`;
    }
  }

  async function applyStyleIndices(sheet: number, edit: StyleEdit, restore: "prev" | "next") {
    const indices = restore === "prev" ? edit.prev_indices : edit.next_indices;
    try {
      await invoke("apply_style_indices", {
        sheet,
        r1: edit.r1, c1: edit.c1, r2: edit.r2, c2: edit.c2,
        indices,
      });
      await refreshRows(edit.r1, edit.r2);
    } catch (e) {
      statusMsg = `Style restore failed: ${e}`;
    }
  }

  async function undo() {
    if (historyIdx <= 0) { statusMsg = "Nothing to undo"; return; }
    const entry = history[historyIdx - 1];
    if (entry.sheet !== activeSheet) await switchSheet(entry.sheet);
    if (entry.kind === "value") {
      await applyEdits(
        entry.sheet,
        entry.edits.map((e) => ({ row: e.row, col: e.col, value: e.prev })),
      );
      dirtyEdits += entry.edits.length;
    } else {
      await applyStyleIndices(entry.sheet, entry.edit, "prev");
    }
    historyIdx -= 1;
    statusMsg = `Undid: ${entry.description}`;
  }

  async function redo() {
    if (historyIdx >= history.length) { statusMsg = "Nothing to redo"; return; }
    const entry = history[historyIdx];
    if (entry.sheet !== activeSheet) await switchSheet(entry.sheet);
    if (entry.kind === "value") {
      await applyEdits(
        entry.sheet,
        entry.edits.map((e) => ({ row: e.row, col: e.col, value: e.next })),
      );
      dirtyEdits += entry.edits.length;
    } else {
      await applyStyleIndices(entry.sheet, entry.edit, "next");
    }
    historyIdx += 1;
    statusMsg = `Redid: ${entry.description}`;
  }

  /// Copy / cut the active selection to the OS clipboard as TSV (Excel-
  /// and Google-Sheets-compatible). Cut additionally clears the cells.
  /// Values-only — formula text is not preserved (to round-trip formulas
  /// internally we'd need relative-ref adjustment, which IronCalc has but
  /// we haven't wired yet).
  async function copySelection(cut: boolean) {
    const r1 = Math.min(selRow, rangeEndRow);
    const r2 = Math.max(selRow, rangeEndRow);
    const c1 = Math.min(selCol, rangeEndCol);
    const c2 = Math.max(selCol, rangeEndCol);
    const lines: string[] = [];
    for (let r = r1; r <= r2; r++) {
      const cols: string[] = [];
      for (let c = c1; c <= c2; c++) {
        const cell = cells.get(`${r}:${c}`);
        cols.push(cell?.text ?? "");
      }
      lines.push(cols.join("\t"));
    }
    const tsv = lines.join("\n");
    try {
      await navigator.clipboard.writeText(tsv);
    } catch (e) {
      statusMsg = `${cut ? "Cut" : "Copy"} failed: ${e}`;
      return;
    }
    if (cut) {
      const sheet = activeSheet;
      const edits: EditOp[] = [];
      for (let r = r1; r <= r2; r++) {
        for (let c = c1; c <= c2; c++) {
          const prev = cells.get(key(r, c))?.input ?? "";
          if (prev !== "") edits.push({ row: r, col: c, prev, next: "" });
        }
      }
      for (const op of edits) {
        try {
          await invoke("set_cell", { sheet, row: op.row, col: op.col, value: "" });
        } catch {}
      }
      await refreshRows(r1, r2);
      dirtyEdits += edits.length;
      const span =
        r1 === r2 && c1 === c2 ? addr(r1, c1) : `${addr(r1, c1)}:${addr(r2, c2)}`;
      pushHistory({ description: `Cut ${span}`, sheet, edits });
    }
    const w = c2 - c1 + 1;
    const h = r2 - r1 + 1;
    statusMsg = `${cut ? "Cut" : "Copied"} ${h}×${w} range to clipboard`;
  }

  /// Paste TSV from the OS clipboard at the cursor. Each \t becomes a
  /// column boundary, each \n a row boundary. Values land verbatim — for
  /// formula text starting with `=` IronCalc parses it back into a
  /// formula on set_user_input.
  async function pasteFromClipboard() {
    let text: string;
    try {
      text = await navigator.clipboard.readText();
    } catch (e) {
      statusMsg = `Paste failed: ${e}`;
      return;
    }
    if (!text) {
      statusMsg = "Clipboard empty";
      return;
    }
    const rows = text.replace(/\r\n/g, "\n").split("\n");
    while (rows.length > 0 && rows[rows.length - 1] === "") rows.pop();
    const sheet = activeSheet;
    const edits: EditOp[] = [];
    for (let i = 0; i < rows.length; i++) {
      const cols = rows[i].split("\t");
      for (let j = 0; j < cols.length; j++) {
        const r = selRow + i;
        const c = selCol + j;
        const prev = cells.get(key(r, c))?.input ?? "";
        const next = cols[j];
        if (prev !== next) edits.push({ row: r, col: c, prev, next });
      }
    }
    for (const op of edits) {
      try {
        await invoke("set_cell", { sheet, row: op.row, col: op.col, value: op.next });
      } catch {}
    }
    if (edits.length > 0) {
      const editR1 = Math.min(...edits.map((e) => e.row));
      const editR2 = Math.max(...edits.map((e) => e.row));
      await refreshRows(editR1, editR2);
    }
    dirtyEdits += edits.length;
    const w = rows[0]?.split("\t").length ?? 0;
    pushHistory({
      description: `Paste ${rows.length}×${w} at ${addr(selRow, selCol)}`,
      sheet,
      edits,
    });
    statusMsg = `Pasted ${rows.length}×${w} range · F9 to recalc`;
  }

  async function refreshWorkbookInfo() {
    try {
      workbook = await invoke<WorkbookInfo>("workbook_info");
    } catch (e) {
      statusMsg = `Sheet info refresh failed: ${e}`;
    }
  }

  async function addSheet() {
    try {
      const [name, idx] = await invoke<[string, number]>("add_sheet");
      await refreshWorkbookInfo();
      await switchSheet(idx);
      statusMsg = `Added sheet "${name}"`;
    } catch (e) {
      statusMsg = `Add sheet failed: ${e}`;
    }
  }

  function renameSheetPrompt(sheet: number) {
    const current = workbook?.sheet_names[sheet] ?? "";
    openMenuPrompt(`Rename sheet "${current}" to:`, current, async (v) => {
      const name = v.trim();
      if (!name || name === current) {
        focusGrid();
        return;
      }
      try {
        await invoke("rename_sheet", { sheet, newName: name });
        await refreshWorkbookInfo();
        statusMsg = `Renamed "${current}" → "${name}"`;
      } catch (e) {
        statusMsg = `Rename failed: ${e}`;
      }
      focusGrid();
    });
  }

  async function deleteSheetConfirm(sheet: number) {
    const current = workbook?.sheet_names[sheet] ?? "";
    if (!workbook || workbook.sheet_names.length <= 1) {
      statusMsg = "Can't delete the only sheet";
      return;
    }
    dynamicTitle = `Delete sheet "${current}"?`;
    dynamicLevel = [
      {
        letter: "Y",
        label: "Yes",
        description: `Delete "${current}" and all its data`,
        action: async () => {
          try {
            await invoke("delete_sheet", { sheet });
            await refreshWorkbookInfo();
            // Sheet indices above the deleted one shift down by 1, so
            // the cursor map keyed by index is now misaligned. Cheapest
            // correct fix is to drop it and let positions repopulate as
            // the user revisits sheets.
            sheetCursors = new Map();
            // If we deleted the active sheet, switch to a neighbor.
            if (sheet <= activeSheet && activeSheet > 0) {
              await switchSheet(activeSheet - 1);
            } else {
              await refreshViewport();
            }
            statusMsg = `Deleted sheet "${current}"`;
          } catch (e) {
            statusMsg = `Delete sheet failed: ${e}`;
          }
        },
      },
      {
        letter: "N",
        label: "No",
        description: "Keep the sheet",
        action: () => { statusMsg = "Delete cancelled"; },
      },
    ];
    menuOpen = true;
    menuPath = [];
    menuHighlight = 0;
  }

  /// Open the F5 goto prompt with sheet + defined-name autocomplete.
  /// Pulled into a helper so the top-of-onKey reload-defang block can
  /// trigger it without duplicating the candidate-fetching logic.
  async function openF5GotoPrompt() {
    const candidates: string[] = [];
    if (workbook) {
      for (const n of workbook.sheet_names) candidates.push(`${n}!`);
    }
    try {
      const names = await invoke<[string, string][]>("list_names");
      for (const [n] of names) candidates.push(n);
    } catch {}
    openMenuPrompt(
      "Goto cell / sheet / name (Tab to complete):",
      addr(selRow, selCol),
      async (v) => {
        const ok = await jumpToAddress(v);
        if (!ok) statusMsg = `Invalid address: ${v}`;
        else statusMsg = `Jumped to ${v.trim()}`;
        focusGrid();
      },
      undefined,
      candidates,
    );
  }

  /// Goto address — accepts plain "B22", sheet-qualified "Sheet2!CK99"
  /// (single-quoted sheet names allowed), bare "Sheet2!" (jumps to A1 of
  /// that sheet), or a defined-name identifier (jumps to the top-left of
  /// the range that name resolves to). Returns false if nothing matches.
  async function jumpToAddress(spec: string): Promise<boolean> {
    const trimmed = spec.trim();
    if (!trimmed) return false;
    let sheetIdx = activeSheet;
    let cellPart = trimmed;
    const bang = trimmed.indexOf("!");
    if (bang >= 0) {
      let sheetName = trimmed.slice(0, bang);
      if (sheetName.startsWith("'") && sheetName.endsWith("'")) {
        sheetName = sheetName.slice(1, -1);
      }
      const idx = workbook?.sheet_names.indexOf(sheetName) ?? -1;
      if (idx < 0) return false;
      sheetIdx = idx;
      cellPart = trimmed.slice(bang + 1);
      // Bare "Sheet!" → jump to A1 of that sheet.
      if (!cellPart) {
        if (sheetIdx !== activeSheet) await switchSheet(sheetIdx);
        selRow = 1; selCol = 1;
        rangeEndRow = 1; rangeEndCol = 1;
        return true;
      }
    }
    const m = cellPart.match(/^\$?([A-Za-z]+)\$?(\d+)$/);
    let row: number, col: number;
    if (m) {
      const colLetters = m[1].toUpperCase();
      row = parseInt(m[2], 10);
      col = 0;
      for (const ch of colLetters) col = col * 26 + (ch.charCodeAt(0) - 64);
    } else if (bang < 0) {
      // Try as a defined name.
      const resolved = await resolveDefinedName(trimmed);
      if (!resolved) return false;
      sheetIdx = resolved.sheet;
      row = resolved.row;
      col = resolved.col;
    } else {
      return false;
    }
    if (col < 1 || row < 1) return false;
    if (sheetIdx !== activeSheet) await switchSheet(sheetIdx);
    selRow = Math.min(row, ABS_MAX_ROWS);
    selCol = Math.min(col, ABS_MAX_COLS);
    rangeEndRow = selRow;
    rangeEndCol = selCol;
    growViewportToInclude(selRow, selCol);
    return true;
  }

  /// Look up a defined name and parse the top-left of the range it
  /// points at. Formulas look like "='Sheet Name'!$A$1:$B$10" or
  /// "=Sheet!$A$1". Returns null if not found / unparseable.
  async function resolveDefinedName(name: string): Promise<{ sheet: number; row: number; col: number } | null> {
    let names: [string, string][];
    try {
      names = await invoke<[string, string][]>("list_names");
    } catch {
      return null;
    }
    const lower = name.toLowerCase();
    const hit = names.find(([n]) => n.toLowerCase() === lower);
    if (!hit) return null;
    let formula = hit[1];
    if (formula.startsWith("=")) formula = formula.slice(1);
    const bang = formula.indexOf("!");
    if (bang < 0) return null;
    let sheetName = formula.slice(0, bang);
    if (sheetName.startsWith("'") && sheetName.endsWith("'")) {
      sheetName = sheetName.slice(1, -1).replace(/''/g, "'");
    }
    const sheetIdx = workbook?.sheet_names.indexOf(sheetName) ?? -1;
    if (sheetIdx < 0) return null;
    let cellRef = formula.slice(bang + 1);
    const colon = cellRef.indexOf(":");
    if (colon >= 0) cellRef = cellRef.slice(0, colon);
    const parsed = parseA1Frontend(cellRef);
    if (!parsed) return null;
    return { sheet: sheetIdx, row: parsed.row, col: parsed.col };
  }

  /// Ctrl+Home → A1 on the active sheet.
  function goHome() {
    selRow = 1;
    selCol = 1;
    rangeEndRow = 1;
    rangeEndCol = 1;
  }

  /// Ctrl+End → bottom-right of the sheet's used range. Falls back to
  /// the current viewport bounds if the dimension can't be read.
  async function goEnd() {
    if (!workbook) return;
    try {
      const [r, c] = await invoke<[number, number]>("get_used_range", { sheet: activeSheet });
      selRow = Math.max(1, Math.min(viewportRows, r));
      selCol = Math.max(1, Math.min(viewportCols, c));
    } catch {
      selRow = viewportRows;
      selCol = viewportCols;
    }
    rangeEndRow = selRow;
    rangeEndCol = selCol;
  }

  /// Page Up/Down — move the cursor by one viewport's worth of rows.
  /// Use the gridWrap's clientHeight divided by the average rendered row
  /// height (default 19px) as an estimate. With shift held, extend
  /// rangeEnd instead of moving the anchor.
  function pageScroll(direction: 1 | -1, extend: boolean) {
    const px = gridWrapEl?.clientHeight ?? 600;
    const rowH = 19;
    const step = Math.max(1, Math.floor(px / rowH) - 1);
    moveSel(direction * step, 0, extend);
  }

  /// Per-sheet cursor + scroll memory. Switching sheets stashes the
  /// outgoing sheet's position here and restores the incoming sheet's
  /// (defaulting to A1 / scroll-zero on first visit). Cleared on
  /// workbook open / new — entries from a previous workbook are stale.
  let sheetCursors = new Map<
    number,
    { selRow: number; selCol: number; rangeEndRow: number; rangeEndCol: number; scrollTop: number; scrollLeft: number }
  >();

  /// Ctrl+PgUp/PgDn sheet navigation. Clamps at the ends (no wrap).
  /// Restores the cursor + scroll position from the last visit to the
  /// target sheet — Excel behaves the same; jumping back to a sheet
  /// where you were deep in the data shouldn't dump you back at A1.
  async function switchSheet(target: number) {
    if (!workbook) return;
    const n = workbook.sheet_names.length;
    if (n === 0) return;
    const next = Math.max(0, Math.min(n - 1, target));
    if (next === activeSheet) return;
    // Stash outgoing position. Read scroll from the live element since
    // bandStart/End may be one frame stale relative to user scrolling.
    sheetCursors.set(activeSheet, {
      selRow,
      selCol,
      rangeEndRow,
      rangeEndCol,
      scrollTop: gridWrapEl?.scrollTop ?? 0,
      scrollLeft: gridWrapEl?.scrollLeft ?? 0,
    });
    activeSheet = next;
    const restore = sheetCursors.get(next);
    selRow = restore?.selRow ?? 1;
    selCol = restore?.selCol ?? 1;
    rangeEndRow = restore?.rangeEndRow ?? selRow;
    rangeEndCol = restore?.rangeEndCol ?? selCol;
    // Scroll restore happens after refreshViewport — we need viewportRows
    // (and the colgroup widths) settled first so the scroll position is
    // meaningful. bandStart/End are zeroed to force ensureRowsLoaded to
    // fall back to its 1..60 default; the post-scroll handleBandChange
    // emit will top up whatever the restored band actually needs.
    bandStart = 1;
    bandEnd = 0;
    if (gridWrapEl) {
      gridWrapEl.scrollTop = 0;
      gridWrapEl.scrollLeft = 0;
    }
    await resizeViewportToSheet();
    await refreshViewport({ clear: true });
    // Apply the restored scroll AFTER the layout settles, then wait a
    // tick so Grid recomputes its band off the new scrollTop and emits
    // onBandChange — that pulls in the rows the user expects to see.
    if (restore && gridWrapEl) {
      await tick();
      gridWrapEl.scrollTop = restore.scrollTop;
      gridWrapEl.scrollLeft = restore.scrollLeft;
    }
    statusMsg = `Sheet: ${workbook.sheet_names[next]}`;
  }

  // ---- menu ----

  async function hideCurrentColumn() {
    try {
      await invoke("set_column_hidden", { sheet: activeSheet, col: selCol, hidden: true });
      await refreshViewport();
      statusMsg = `Hid column ${addr(1, selCol).replace("1", "")}`;
    } catch (e) {
      statusMsg = `Hide failed: ${e}`;
    }
  }

  async function hideCurrentRow() {
    try {
      await invoke("set_row_hidden", { sheet: activeSheet, row: selRow, hidden: true });
      await refreshViewport();
      statusMsg = `Hid row ${selRow}`;
    } catch (e) {
      statusMsg = `Hide failed: ${e}`;
    }
  }

  async function showAllColumns() {
    try {
      const n = await invoke<number>("show_all_cols", { sheet: activeSheet });
      await refreshViewport();
      statusMsg = n > 0 ? `Displayed ${n} hidden column${n === 1 ? "" : "s"}` : "No hidden columns";
    } catch (e) {
      statusMsg = `Display failed: ${e}`;
    }
  }

  async function showAllRows() {
    try {
      const n = await invoke<number>("show_all_rows", { sheet: activeSheet });
      await refreshViewport();
      statusMsg = n > 0 ? `Displayed ${n} hidden row${n === 1 ? "" : "s"}` : "No hidden rows";
    } catch (e) {
      statusMsg = `Display failed: ${e}`;
    }
  }

  function setColumnWidthPrompt() { startAxisPick("col", "set-size"); }
  function autoColumnWidth() { startAxisPick("col", "auto"); }

  /// Find / Replace. v1 scans the loaded `cells` map (which equals the
  /// active sheet for sheets within the viewport cap), case-insensitive
  /// substring on cell.text and cell.input. F3 / Shift+F3 cycle.
  type Match = { row: number; col: number };
  let findResults = $state<Match[]>([]);
  let findIdx = $state(0);
  let findNeedle = $state("");

  function scanForMatches(needle: string): Match[] {
    const lower = needle.toLowerCase();
    const out: Match[] = [];
    for (const cv of cells.values()) {
      const t = cv.text?.toLowerCase() ?? "";
      const i = cv.input?.toLowerCase() ?? "";
      if (t.includes(lower) || i.includes(lower)) {
        out.push({ row: cv.row, col: cv.col });
      }
    }
    out.sort((a, b) => a.row - b.row || a.col - b.col);
    return out;
  }

  function jumpToMatch(idx: number) {
    if (findResults.length === 0) return;
    const i = ((idx % findResults.length) + findResults.length) % findResults.length;
    findIdx = i;
    const m = findResults[i];
    selRow = m.row;
    selCol = m.col;
    rangeEndRow = m.row;
    rangeEndCol = m.col;
    statusMsg = `Find "${findNeedle}" — ${i + 1} of ${findResults.length} at ${addr(m.row, m.col)} · F3 next, Shift+F3 prev`;
  }

  function doFind(needle: string) {
    if (!needle) { focusGrid(); return; }
    findNeedle = needle;
    findResults = scanForMatches(needle);
    if (findResults.length === 0) {
      statusMsg = `Find "${needle}" — no matches`;
      focusGrid();
      return;
    }
    jumpToMatch(0);
    focusGrid();
  }

  async function doFindReplace(needle: string, replacement: string) {
    if (!needle) { focusGrid(); return; }
    findNeedle = needle;
    const matches = scanForMatches(needle);
    if (matches.length === 0) {
      statusMsg = `Replace "${needle}" — no matches`;
      focusGrid();
      return;
    }
    const lower = needle.toLowerCase();
    const sheet = activeSheet;
    const edits: EditOp[] = [];
    for (const m of matches) {
      const cv = cells.get(`${m.row}:${m.col}`);
      if (!cv) continue;
      const prev = cv.input;
      // Case-preserving substring replace via case-insensitive scan.
      const lowerInput = prev.toLowerCase();
      let next = "";
      let cursor = 0;
      while (cursor < prev.length) {
        const i = lowerInput.indexOf(lower, cursor);
        if (i < 0) {
          next += prev.slice(cursor);
          break;
        }
        next += prev.slice(cursor, i) + replacement;
        cursor = i + needle.length;
      }
      if (next !== prev) edits.push({ row: m.row, col: m.col, prev, next });
    }
    for (const op of edits) {
      try {
        await invoke("set_cell", { sheet, row: op.row, col: op.col, value: op.next });
      } catch {}
    }
    await refreshViewport();
    dirtyEdits += edits.length;
    if (edits.length > 0) {
      pushHistory({
        description: `Replace "${needle}" → "${replacement}"`,
        sheet,
        edits,
      });
    }
    statusMsg = `Replaced ${edits.length} occurrence${edits.length === 1 ? "" : "s"}`;
    focusGrid();
  }

  function openFindReplace() {
    openMenuPrompt(
      "Find:",
      findNeedle,
      (find) => {
        if (!find) { focusGrid(); return; }
        // Chain into a Replace prompt; Esc on the second falls back to find-only.
        openMenuPrompt(
          `Replace "${find}" with (Esc = find only):`,
          "",
          (replace) => doFindReplace(find, replace),
          () => doFind(find),
        );
      },
    );
  }

  /// Apply a number format to the active selection. Variants that take
  /// a decimal count first prompt for it; the others write directly.
  async function applyNumberFormat(format: string) {
    const r1 = Math.min(selRow, rangeEndRow);
    const r2 = Math.max(selRow, rangeEndRow);
    const c1 = Math.min(selCol, rangeEndCol);
    const c2 = Math.max(selCol, rangeEndCol);
    try {
      const n = await invoke<number>("set_range_number_format", {
        sheet: activeSheet,
        r1, c1, r2, c2,
        format,
      });
      await refreshRows(r1, r2);
      const span = r1 === r2 && c1 === c2 ? addr(r1, c1) : `${addr(r1, c1)}:${addr(r2, c2)}`;
      statusMsg = `Formatted ${span} (${n} cell${n === 1 ? "" : "s"}) as ${format}`;
    } catch (e) {
      statusMsg = `Format failed: ${e}`;
    }
  }

  function buildFixedFormat(decimals: number): string {
    if (decimals <= 0) return "0";
    return "0." + "0".repeat(Math.min(15, decimals));
  }
  function buildCommaFormat(decimals: number): string {
    if (decimals <= 0) return "#,##0";
    return "#,##0." + "0".repeat(Math.min(15, decimals));
  }
  function buildCurrencyFormat(decimals: number): string {
    if (decimals <= 0) return "$#,##0";
    return "$#,##0." + "0".repeat(Math.min(15, decimals));
  }
  function buildPercentFormat(decimals: number): string {
    if (decimals <= 0) return "0%";
    return "0." + "0".repeat(Math.min(15, decimals)) + "%";
  }

  /// Drag-fill: source pattern wraps to fill the extended range, skipping
  /// cells inside the source itself (they keep their values).
  async function fillFromHandle(
    src: { r1: number; c1: number; r2: number; c2: number },
    dest: { r1: number; c1: number; r2: number; c2: number },
  ) {
    const srcH = src.r2 - src.r1 + 1;
    const srcW = src.c2 - src.c1 + 1;
    const sheet = activeSheet;
    const edits: EditOp[] = [];
    for (let r = dest.r1; r <= dest.r2; r++) {
      for (let c = dest.c1; c <= dest.c2; c++) {
        if (r >= src.r1 && r <= src.r2 && c >= src.c1 && c <= src.c2) continue;
        const srcR = src.r1 + (((r - src.r1) % srcH) + srcH) % srcH;
        const srcC = src.c1 + (((c - src.c1) % srcW) + srcW) % srcW;
        const srcVal = cells.get(`${srcR}:${srcC}`)?.input ?? "";
        const prev = cells.get(`${r}:${c}`)?.input ?? "";
        if (prev !== srcVal) edits.push({ row: r, col: c, prev, next: srcVal });
      }
    }
    for (const op of edits) {
      try {
        await invoke("set_cell", { sheet, row: op.row, col: op.col, value: op.next });
      } catch {}
    }
    await refreshRows(dest.r1, dest.r2);
    dirtyEdits += edits.length;
    pushHistory({
      description: `Fill from ${addr(src.r1, src.c1)}:${addr(src.r2, src.c2)}`,
      sheet,
      edits,
    });
    statusMsg = `Filled ${edits.length} cell${edits.length === 1 ? "" : "s"}`;
  }

  function nameCreate() {
    const r1 = Math.min(selRow, rangeEndRow);
    const r2 = Math.max(selRow, rangeEndRow);
    const c1 = Math.min(selCol, rangeEndCol);
    const c2 = Math.max(selCol, rangeEndCol);
    openMenuPrompt("Defined-name to create:", "", async (name) => {
      const t = name.trim();
      if (!t) { focusGrid(); return; }
      try {
        await invoke("define_name", { name: t, sheet: activeSheet, r1, c1, r2, c2 });
        statusMsg = `Named "${t}" → ${addr(r1, c1)}:${addr(r2, c2)}`;
        await refreshNameCache();
      } catch (e) {
        statusMsg = `Define name failed: ${e}`;
      }
      focusGrid();
    });
  }

  function nameDelete() {
    openMenuPrompt("Defined-name to delete:", "", async (name) => {
      const t = name.trim();
      if (!t) { focusGrid(); return; }
      try {
        await invoke("delete_name", { name: t });
        statusMsg = `Deleted name "${t}"`;
        await refreshNameCache();
      } catch (e) {
        statusMsg = `Delete name failed: ${e}`;
      }
      focusGrid();
    });
  }

  /// Parse a defined-name's resolved formula text — typically
  /// "Discount!$B$24:$W$35" or "'Sheet 1'!$A$1" — into concrete range
  /// bounds. Returns null when the formula isn't a simple sheet-
  /// qualified A1 / A1:B2 reference (e.g. names defined as
  /// expressions like OFFSET(...) — those get skipped).
  function parseNameFormula(formula: string): ParsedNameRange | null {
    if (!workbook) return null;
    const stripped = formula.startsWith("=") ? formula.slice(1) : formula;
    // Pull off the sheet name first (quoted or bare).
    let sheet = "";
    let rest = "";
    if (stripped.startsWith("'")) {
      const close = stripped.indexOf("'!");
      if (close < 0) return null;
      sheet = stripped.slice(1, close);
      rest = stripped.slice(close + 2);
    } else {
      const bang = stripped.indexOf("!");
      if (bang < 0) return null;
      sheet = stripped.slice(0, bang);
      rest = stripped.slice(bang + 1);
    }
    const sheetIdx = workbook.sheet_names.indexOf(sheet);
    if (sheetIdx < 0) return null;
    const cells = rest.split(":");
    const start = parseA1Frontend(cells[0]);
    if (!start) return null;
    const end = cells.length === 2 ? parseA1Frontend(cells[1]) : start;
    if (!end) return null;
    return {
      sheet: sheetIdx,
      r1: Math.min(start.row, end.row),
      c1: Math.min(start.col, end.col),
      r2: Math.max(start.row, end.row),
      c2: Math.max(start.col, end.col),
    };
  }

  /// Pull the current defined-name list from the backend and rebuild
  /// the cache used by F2 edit-mode highlighting.
  async function refreshNameCache() {
    try {
      const names = await invoke<NamedRangeInfo[]>("list_named_ranges");
      const next = new Map<string, ParsedNameRange>();
      for (const n of names) {
        const range = parseNameFormula(n.formula);
        if (range) next.set(n.name.toLowerCase(), range);
      }
      nameCache = next;
    } catch {
      nameCache = new Map();
    }
  }

  /// Lotus /R/N/L — drop the names list into the worksheet as a 2-column
  /// table starting at the chosen address. Useful for sorting / copying /
  /// inspecting all defined names. The status bar isn't suitable for
  /// more than a couple of names.
  async function nameList() {
    let names: [string, string][];
    try {
      names = await invoke<[string, string][]>("list_names");
    } catch (e) {
      statusMsg = `List names failed: ${e}`;
      return;
    }
    if (names.length === 0) {
      statusMsg = "No defined names to list";
      return;
    }
    openMenuPrompt(
      `Insert ${names.length} name${names.length === 1 ? "" : "s"} starting at:`,
      addr(selRow, selCol),
      async (v) => {
        const target = parseA1Frontend(v);
        if (!target) {
          statusMsg = `Invalid destination: ${v}`;
          focusGrid();
          return;
        }
        const sheet = activeSheet;
        const edits: EditOp[] = [];
        for (let i = 0; i < names.length; i++) {
          const r = target.row + i;
          for (const [c, val] of [
            [target.col, names[i][0]],
            [target.col + 1, names[i][1]],
          ] as [number, string][]) {
            const prev = cells.get(`${r}:${c}`)?.input ?? "";
            if (prev !== val) edits.push({ row: r, col: c, prev, next: val });
          }
        }
        for (const op of edits) {
          try {
            await invoke("set_cell", { sheet, row: op.row, col: op.col, value: op.next });
          } catch {}
        }
        await refreshViewport();
        dirtyEdits += edits.length;
        pushHistory({
          description: `Insert ${names.length} name${names.length === 1 ? "" : "s"} list`,
          sheet,
          edits,
        });
        statusMsg = `Inserted ${names.length} name${names.length === 1 ? "" : "s"} at ${addr(target.row, target.col)}`;
        focusGrid();
      },
    );
  }

  /// /Data/Fill — three-step prompt chain (start → step → done).
  /// Fills the selection in row-major order with start + i*step.
  function dataFill() {
    const r1 = Math.min(selRow, rangeEndRow);
    const r2 = Math.max(selRow, rangeEndRow);
    const c1 = Math.min(selCol, rangeEndCol);
    const c2 = Math.max(selCol, rangeEndCol);
    openMenuPrompt("Fill start:", "1", (startV) => {
      const start = Number(startV);
      if (!Number.isFinite(start)) {
        statusMsg = `Invalid start: ${startV}`;
        focusGrid();
        return;
      }
      openMenuPrompt("Fill step:", "1", async (stepV) => {
        const step = Number(stepV);
        if (!Number.isFinite(step)) {
          statusMsg = `Invalid step: ${stepV}`;
          focusGrid();
          return;
        }
        const sheet = activeSheet;
        const edits: EditOp[] = [];
        let i = 0;
        for (let r = r1; r <= r2; r++) {
          for (let c = c1; c <= c2; c++) {
            const next = String(start + i * step);
            const prev = cells.get(`${r}:${c}`)?.input ?? "";
            if (prev !== next) edits.push({ row: r, col: c, prev, next });
            i++;
          }
        }
        for (const op of edits) {
          try {
            await invoke("set_cell", { sheet, row: op.row, col: op.col, value: op.next });
          } catch {}
        }
        await refreshViewport();
        dirtyEdits += edits.length;
        const span = r1 === r2 && c1 === c2 ? addr(r1, c1) : `${addr(r1, c1)}:${addr(r2, c2)}`;
        pushHistory({
          description: `Fill ${span} (${start}, +${step})`,
          sheet,
          edits,
        });
        statusMsg = `Filled ${edits.length} cell${edits.length === 1 ? "" : "s"} in ${span}`;
        focusGrid();
      });
    });
  }

  /// /Data/Sort — sort the selection's rows by a column. Numeric vs
  /// lexical comparator picked per cell-pair (numbers compare numerically
  /// when both parse as f64; otherwise case-insensitive string compare).
  function dataSort() {
    const r1 = Math.min(selRow, rangeEndRow);
    const r2 = Math.max(selRow, rangeEndRow);
    const c1 = Math.min(selCol, rangeEndCol);
    const c2 = Math.max(selCol, rangeEndCol);
    if (r2 - r1 < 1) {
      statusMsg = "Select at least 2 rows to sort";
      return;
    }
    const defaultKey = addr(1, c1).replace("1", "");
    openMenuPrompt(`Sort by column (${defaultKey}-${addr(1, c2).replace("1", "")}):`, defaultKey, (keyV) => {
      const t = keyV.trim().toUpperCase();
      let keyCol = 0;
      for (const ch of t) {
        if (!/[A-Z]/.test(ch)) { keyCol = 0; break; }
        keyCol = keyCol * 26 + (ch.charCodeAt(0) - 64);
      }
      if (keyCol < c1 || keyCol > c2) {
        statusMsg = `Column ${keyV} not in selection`;
        focusGrid();
        return;
      }
      openMenuPrompt("Order: A=ascending, D=descending", "A", async (orderV) => {
        const desc = orderV.trim().toUpperCase().startsWith("D");
        // Snapshot rows.
        const snap: Array<{ origRow: number; vals: string[] }> = [];
        for (let r = r1; r <= r2; r++) {
          const vals: string[] = [];
          for (let c = c1; c <= c2; c++) {
            vals.push(cells.get(`${r}:${c}`)?.input ?? "");
          }
          snap.push({ origRow: r, vals });
        }
        const keyIdx = keyCol - c1;
        snap.sort((a, b) => {
          const av = a.vals[keyIdx] ?? "";
          const bv = b.vals[keyIdx] ?? "";
          const an = Number(av);
          const bn = Number(bv);
          let cmp = 0;
          if (Number.isFinite(an) && Number.isFinite(bn) && av.trim() !== "" && bv.trim() !== "") {
            cmp = an - bn;
          } else {
            cmp = av.toLowerCase().localeCompare(bv.toLowerCase());
          }
          return desc ? -cmp : cmp;
        });
        // Write back.
        const sheet = activeSheet;
        const edits: EditOp[] = [];
        for (let i = 0; i < snap.length; i++) {
          const targetR = r1 + i;
          for (let c = 0; c < snap[i].vals.length; c++) {
            const targetC = c1 + c;
            const prev = cells.get(`${targetR}:${targetC}`)?.input ?? "";
            const next = snap[i].vals[c];
            if (prev !== next) edits.push({ row: targetR, col: targetC, prev, next });
          }
        }
        for (const op of edits) {
          try {
            await invoke("set_cell", { sheet, row: op.row, col: op.col, value: op.next });
          } catch {}
        }
        await refreshViewport();
        dirtyEdits += edits.length;
        const span = `${addr(r1, c1)}:${addr(r2, c2)}`;
        pushHistory({ description: `Sort ${span} by ${t} ${desc ? "desc" : "asc"}`, sheet, edits });
        statusMsg = `Sorted ${span} by ${t} ${desc ? "↓" : "↑"}`;
        focusGrid();
      });
    });
  }

  /// /Range/Value — replace each formula in the selection with its
  /// evaluated value. One undo entry covers the whole range.
  async function rangeValue() {
    const r1 = Math.min(selRow, rangeEndRow);
    const r2 = Math.max(selRow, rangeEndRow);
    const c1 = Math.min(selCol, rangeEndCol);
    const c2 = Math.max(selCol, rangeEndCol);
    const sheet = activeSheet;
    const edits: EditOp[] = [];
    for (let r = r1; r <= r2; r++) {
      for (let c = c1; c <= c2; c++) {
        const cv = cells.get(`${r}:${c}`);
        if (!cv?.is_formula) continue;
        const prev = cv.input;
        const next = cv.text;
        if (prev !== next) edits.push({ row: r, col: c, prev, next });
      }
    }
    for (const op of edits) {
      try {
        await invoke("set_cell", { sheet, row: op.row, col: op.col, value: op.next });
      } catch {}
    }
    await refreshRows(r1, r2);
    dirtyEdits += edits.length;
    const span = r1 === r2 && c1 === c2 ? addr(r1, c1) : `${addr(r1, c1)}:${addr(r2, c2)}`;
    if (edits.length > 0) {
      pushHistory({ description: `Convert ${span} to values`, sheet, edits });
    }
    statusMsg = `Converted ${edits.length} formula${edits.length === 1 ? "" : "s"} in ${span} to values`;
  }

  /// /Range/Trans — read the selection as a matrix, write back transposed
  /// (rows ↔ cols) starting at the same anchor. Pushes one undo entry.
  async function rangeTrans() {
    const r1 = Math.min(selRow, rangeEndRow);
    const r2 = Math.max(selRow, rangeEndRow);
    const c1 = Math.min(selCol, rangeEndCol);
    const c2 = Math.max(selCol, rangeEndCol);
    const h = r2 - r1 + 1;
    const w = c2 - c1 + 1;
    const sheet = activeSheet;
    // Snapshot the source values via cell.input so formulas survive the
    // transpose (Excel's Paste-Special-Transpose preserves them).
    const src: string[][] = [];
    for (let r = 0; r < h; r++) {
      const row: string[] = [];
      for (let c = 0; c < w; c++) {
        row.push(cells.get(`${r1 + r}:${c1 + c}`)?.input ?? "");
      }
      src.push(row);
    }
    const newH = w;
    const newW = h;
    const edits: EditOp[] = [];
    for (let r = 0; r < newH; r++) {
      for (let c = 0; c < newW; c++) {
        const targetR = r1 + r;
        const targetC = c1 + c;
        const prev = cells.get(`${targetR}:${targetC}`)?.input ?? "";
        // Transposed source[c][r]
        const next = src[c]?.[r] ?? "";
        if (prev !== next) edits.push({ row: targetR, col: targetC, prev, next });
      }
    }
    // Also clear cells in the original rectangle that are outside the
    // new (transposed) rectangle.
    for (let r = r1; r <= r2; r++) {
      for (let c = c1; c <= c2; c++) {
        const within = r < r1 + newH && c < c1 + newW;
        if (!within) {
          const prev = cells.get(`${r}:${c}`)?.input ?? "";
          if (prev !== "") edits.push({ row: r, col: c, prev, next: "" });
        }
      }
    }
    for (const op of edits) {
      try {
        await invoke("set_cell", { sheet, row: op.row, col: op.col, value: op.next });
      } catch {}
    }
    // Transpose can write outside the original rectangle when newH > h
    // or newW > w, so refresh the union of the source and target boxes.
    await refreshRows(r1, Math.max(r2, r1 + newH - 1));
    dirtyEdits += edits.length;
    const span = r1 === r2 && c1 === c2 ? addr(r1, c1) : `${addr(r1, c1)}:${addr(r2, c2)}`;
    pushHistory({ description: `Transpose ${span}`, sheet, edits });
    rangeEndRow = r1 + newH - 1;
    rangeEndCol = c1 + newW - 1;
    statusMsg = `Transposed ${span} → ${addr(r1, c1)}:${addr(rangeEndRow, rangeEndCol)}`;
  }

  /// /Copy and /Move enter "destination cursor" mode — Lotus convention.
  /// The arrow keys move a destination anchor highlighted as a dashed
  /// outline at source-rectangle dimensions; Enter commits; Esc cancels.
  let pendingMove = $state<{
    kind: "copy" | "move";
    source: { r1: number; c1: number; r2: number; c2: number };
    anchor: { row: number; col: number };
  } | null>(null);

  /// Visible rectangle of the pending destination (passed to Grid as a
  /// "ghost" overlay).
  let ghostRange = $derived.by<{ r1: number; c1: number; r2: number; c2: number } | null>(() => {
    if (!pendingMove) return null;
    const h = pendingMove.source.r2 - pendingMove.source.r1;
    const w = pendingMove.source.c2 - pendingMove.source.c1;
    return {
      r1: pendingMove.anchor.row,
      c1: pendingMove.anchor.col,
      r2: pendingMove.anchor.row + h,
      c2: pendingMove.anchor.col + w,
    };
  });

  function startCopyMove(kind: "copy" | "move") {
    const r1 = Math.min(selRow, rangeEndRow);
    const r2 = Math.max(selRow, rangeEndRow);
    const c1 = Math.min(selCol, rangeEndCol);
    const c2 = Math.max(selCol, rangeEndCol);
    pendingMove = {
      kind,
      source: { r1, c1, r2, c2 },
      anchor: { row: r1, col: c1 },
    };
    const span = r1 === r2 && c1 === c2 ? addr(r1, c1) : `${addr(r1, c1)}:${addr(r2, c2)}`;
    menuMessage = `${kind === "move" ? "Move" : "Copy"} ${span} — arrow keys to position destination, Enter to confirm, Esc to cancel`;
  }

  function copyRange() { startCopyMove("copy"); }
  function moveRange() { startCopyMove("move"); }

  function pendingMoveStep(dr: number, dc: number) {
    if (!pendingMove) return;
    pendingMove.anchor.row = Math.max(1, Math.min(ABS_MAX_ROWS, pendingMove.anchor.row + dr));
    pendingMove.anchor.col = Math.max(1, Math.min(ABS_MAX_COLS, pendingMove.anchor.col + dc));
    growViewportToInclude(pendingMove.anchor.row, pendingMove.anchor.col);
  }

  async function commitPendingMove() {
    if (!pendingMove) return;
    const { kind, source, anchor } = pendingMove;
    pendingMove = null;
    menuMessage = null;
    const sheet = activeSheet;
    const h = source.r2 - source.r1 + 1;
    const w = source.c2 - source.c1 + 1;
    const edits: EditOp[] = [];
    for (let r = 0; r < h; r++) {
      for (let c = 0; c < w; c++) {
        const srcVal = cells.get(`${source.r1 + r}:${source.c1 + c}`)?.input ?? "";
        const tgtR = anchor.row + r;
        const tgtC = anchor.col + c;
        const prev = cells.get(`${tgtR}:${tgtC}`)?.input ?? "";
        if (prev !== srcVal) edits.push({ row: tgtR, col: tgtC, prev, next: srcVal });
      }
    }
    if (kind === "move") {
      for (let r = source.r1; r <= source.r2; r++) {
        for (let c = source.c1; c <= source.c2; c++) {
          const overlap =
            r >= anchor.row && r < anchor.row + h &&
            c >= anchor.col && c < anchor.col + w;
          if (overlap) continue;
          const prev = cells.get(`${r}:${c}`)?.input ?? "";
          if (prev !== "") edits.push({ row: r, col: c, prev, next: "" });
        }
      }
    }
    for (const op of edits) {
      try {
        await invoke("set_cell", { sheet, row: op.row, col: op.col, value: op.next });
      } catch {}
    }
    // Both source and destination row ranges may have changed (move
    // clears source, both touch dest). Span the union.
    await refreshRows(
      Math.min(source.r1, anchor.row),
      Math.max(source.r2, anchor.row + h - 1),
    );
    dirtyEdits += edits.length;
    const span = source.r1 === source.r2 && source.c1 === source.c2
      ? addr(source.r1, source.c1)
      : `${addr(source.r1, source.c1)}:${addr(source.r2, source.c2)}`;
    pushHistory({
      description: `${kind === "move" ? "Move" : "Copy"} ${span} → ${addr(anchor.row, anchor.col)}`,
      sheet,
      edits,
    });
    statusMsg = `${kind === "move" ? "Moved" : "Copied"} ${span} → ${addr(anchor.row, anchor.col)}`;
    // Move the cursor to the destination so the user can keep working there.
    selRow = anchor.row;
    selCol = anchor.col;
    rangeEndRow = anchor.row + h - 1;
    rangeEndCol = anchor.col + w - 1;
  }

  function cancelPendingMove() {
    if (!pendingMove) return;
    pendingMove = null;
    menuMessage = null;
    statusMsg = "Cancelled";
  }

  function parseA1Frontend(s: string): { row: number; col: number } | null {
    const m = s.replace(/\$/g, "").trim().match(/^([A-Za-z]+)(\d+)$/);
    if (!m) return null;
    let col = 0;
    for (const ch of m[1].toUpperCase()) col = col * 26 + (ch.charCodeAt(0) - 64);
    const row = parseInt(m[2], 10);
    if (col < 1 || row < 1) return null;
    return { row, col };
  }

  /// Apply a generic style op (B/I/U/align/colour) to the active selection.
  /// Captures prev / next style indices so the change is undoable.
  async function applyStyleOp(op: object, label: string) {
    const r1 = Math.min(selRow, rangeEndRow);
    const r2 = Math.max(selRow, rangeEndRow);
    const c1 = Math.min(selCol, rangeEndCol);
    const c2 = Math.max(selCol, rangeEndCol);
    try {
      const result = await invoke<{ count: number; prev_indices: number[]; next_indices: number[] }>(
        "set_range_style",
        { sheet: activeSheet, r1, c1, r2, c2, op },
      );
      await refreshRows(r1, r2);
      const span = r1 === r2 && c1 === c2 ? addr(r1, c1) : `${addr(r1, c1)}:${addr(r2, c2)}`;
      // Only push undo if any cell's style index actually changed.
      const changed = result.prev_indices.some((p, i) => p !== result.next_indices[i]);
      if (changed) {
        pushHistory({
          kind: "style",
          description: `${label} ${span}`,
          sheet: activeSheet,
          edit: {
            r1, c1, r2, c2,
            prev_indices: result.prev_indices,
            next_indices: result.next_indices,
          },
        });
      }
      statusMsg = `${label} ${span} (${result.count} cell${result.count === 1 ? "" : "s"})`;
    } catch (e) {
      statusMsg = `Style failed: ${e}`;
    }
  }

  function alignRange(h: "left" | "center" | "right" | "general") {
    const map = {
      left: { kind: "align_left" },
      center: { kind: "align_center" },
      right: { kind: "align_right" },
      general: { kind: "align_general" },
    } as const;
    applyStyleOp(map[h], `Aligned ${h}`);
  }

  function promptColor(label: string, onColor: (color: string) => void, onClear: () => void) {
    openMenuPrompt(label, "#FFD966", (v) => {
      const t = v.trim();
      if (t === "") {
        onClear();
      } else if (/^#[0-9A-Fa-f]{6}$/.test(t)) {
        onColor(t.toUpperCase());
      } else {
        statusMsg = `Invalid colour: ${v} (use #RRGGBB or leave empty)`;
        focusGrid();
      }
    });
  }

  function setFillColor() {
    promptColor(
      "Fill colour (#RRGGBB, empty to clear):",
      (color) => applyStyleOp({ kind: "set_fill_color", color }, `Filled`),
      () => applyStyleOp({ kind: "clear_fill_color" }, `Cleared fill`),
    );
  }

  function setBorder(sides: "all" | "outline" | "top" | "bottom" | "left" | "right" | "none") {
    applyStyleOp({ kind: "set_border", sides }, `Border (${sides})`);
  }

  function setTextColor() {
    promptColor(
      "Text colour (#RRGGBB, empty to clear):",
      (color) => applyStyleOp({ kind: "set_text_color", color }, `Set text colour`),
      () => applyStyleOp({ kind: "clear_text_color" }, `Cleared text colour`),
    );
  }

  function formatRange(kind: FormatKind) {
    switch (kind) {
      case "general":
        applyNumberFormat("General");
        return;
      case "date":
        applyNumberFormat("yyyy-mm-dd");
        return;
      case "time":
        applyNumberFormat("h:mm:ss");
        return;
    }
    // Decimal-prompted variants
    const builders: Record<string, (n: number) => string> = {
      fixed: buildFixedFormat,
      currency: buildCurrencyFormat,
      comma: buildCommaFormat,
      percent: buildPercentFormat,
    };
    const builder = builders[kind];
    openMenuPrompt(`Decimals (0..15) for ${kind}:`, "2", async (v) => {
      const n = Number(v);
      if (!Number.isFinite(n) || n < 0 || n > 15) {
        statusMsg = `Invalid decimals: ${v}`;
        focusGrid();
        return;
      }
      await applyNumberFormat(builder(Math.floor(n)));
      focusGrid();
    });
  }

  /// Insert / delete operations. Count is the selection size on the
  /// matching axis: e.g. selecting rows 5..9 inserts 5 rows at row 5.
  /// Single-cell selection means count=1.
  async function insertRowsAtSel() {
    const r1 = Math.min(selRow, rangeEndRow);
    const r2 = Math.max(selRow, rangeEndRow);
    const count = r2 - r1 + 1;
    try {
      await invoke("insert_rows", { sheet: activeSheet, row: r1, count });
      await refreshViewport();
      statusMsg = `Inserted ${count} row${count === 1 ? "" : "s"} at ${r1}`;
    } catch (e) {
      statusMsg = `Insert failed: ${e}`;
    }
  }

  async function deleteRowsAtSel() {
    const r1 = Math.min(selRow, rangeEndRow);
    const r2 = Math.max(selRow, rangeEndRow);
    const count = r2 - r1 + 1;
    try {
      await invoke("delete_rows", { sheet: activeSheet, row: r1, count });
      await refreshViewport();
      statusMsg = `Deleted ${count} row${count === 1 ? "" : "s"} from ${r1}`;
    } catch (e) {
      statusMsg = `Delete failed: ${e}`;
    }
  }

  async function insertColumnsAtSel() {
    const c1 = Math.min(selCol, rangeEndCol);
    const c2 = Math.max(selCol, rangeEndCol);
    const count = c2 - c1 + 1;
    try {
      await invoke("insert_columns", { sheet: activeSheet, col: c1, count });
      await refreshViewport();
      statusMsg = `Inserted ${count} column${count === 1 ? "" : "s"} at ${addr(1, c1).replace("1", "")}`;
    } catch (e) {
      statusMsg = `Insert failed: ${e}`;
    }
  }

  async function deleteColumnsAtSel() {
    const c1 = Math.min(selCol, rangeEndCol);
    const c2 = Math.max(selCol, rangeEndCol);
    const count = c2 - c1 + 1;
    try {
      await invoke("delete_columns", { sheet: activeSheet, col: c1, count });
      await refreshViewport();
      statusMsg = `Deleted ${count} column${count === 1 ? "" : "s"} from ${addr(1, c1).replace("1", "")}`;
    } catch (e) {
      statusMsg = `Delete failed: ${e}`;
    }
  }

  async function setTitles(kind: "both" | "horizontal" | "vertical" | "clear") {
    let rows = 0;
    let cols = 0;
    if (kind === "both" || kind === "horizontal") rows = selRow - 1;
    if (kind === "both" || kind === "vertical") cols = selCol - 1;
    try {
      await invoke("set_frozen_panes", { sheet: activeSheet, rows, cols });
      await refreshViewport();
      statusMsg =
        kind === "clear"
          ? "Titles cleared"
          : `Frozen ${rows} row${rows === 1 ? "" : "s"} / ${cols} col${cols === 1 ? "" : "s"}`;
    } catch (e) {
      statusMsg = `Titles failed: ${e}`;
    }
  }

  function setRowHeightPrompt() { startAxisPick("row", "set-size"); }
  function autoRowHeight() { startAxisPick("row", "auto"); }

  // ---- axis-pick (multi-row / multi-column resize) ----

  /// /Worksheet/{Row,Column}/Set-* and /Auto enter "axis pick" mode —
  /// the user extends a row or column range with arrow keys before the
  /// resize fires. Defaults to the current row/col only; Up/Down (rows)
  /// or Left/Right (cols) extend by 1, Ctrl+arrow extends by 10.
  ///
  /// We mutate selRow/selCol/rangeEndRow/rangeEndCol so the existing
  /// range-tint overlay shows what's being picked, and stash the
  /// originals to restore on Esc.
  let pendingAxisPick = $state<{
    axis: "row" | "col";
    op: "set-size" | "auto";
    start: number;
    end: number;
    origSelRow: number;
    origSelCol: number;
    origRangeEndRow: number;
    origRangeEndCol: number;
  } | null>(null);

  const AXIS_PICK_JUMP = 10;

  function startAxisPick(axis: "row" | "col", op: "set-size" | "auto") {
    const idx = axis === "row" ? selRow : selCol;
    pendingAxisPick = {
      axis,
      op,
      start: idx,
      end: idx,
      origSelRow: selRow,
      origSelCol: selCol,
      origRangeEndRow: rangeEndRow,
      origRangeEndCol: rangeEndCol,
    };
    paintAxisPickSelection();
    describeAxisPick();
    focusGrid();
  }

  function paintAxisPickSelection() {
    if (!pendingAxisPick) return;
    const { axis, start, end } = pendingAxisPick;
    // Active cell follows the moving end (`end` becomes selRow/selCol)
    // so the blue active-cell outline leads the pick. Grid's existing
    // selRow/selCol effect calls scrollSelIntoView, so the viewport
    // tracks the cursor as it walks off-screen — no extra scroll wiring
    // needed. The opposite axis stays anchored at the original cell so
    // we don't yank the horizontal scroll while picking rows (or vice
    // versa).
    if (axis === "row") {
      selRow = end;
      rangeEndRow = start;
      selCol = pendingAxisPick.origSelCol;
      rangeEndCol = pendingAxisPick.origSelCol;
    } else {
      selCol = end;
      rangeEndCol = start;
      selRow = pendingAxisPick.origSelRow;
      rangeEndRow = pendingAxisPick.origSelRow;
    }
  }

  function describeAxisPick() {
    if (!pendingAxisPick) return;
    const { axis, op, start, end } = pendingAxisPick;
    const lo = Math.min(start, end);
    const hi = Math.max(start, end);
    const verb = op === "auto" ? "Auto-fit" : "Resize";
    const label = axis === "row" ? "row" : "column";
    const span = axis === "row"
      ? (lo === hi ? `${lo}` : `${lo}–${hi}`)
      : (lo === hi
          ? addr(1, lo).replace(/\d+$/, "")
          : `${addr(1, lo).replace(/\d+$/, "")}–${addr(1, hi).replace(/\d+$/, "")}`);
    const arrows = axis === "row" ? "↑/↓" : "←/→";
    menuMessage = `${verb} ${label}${lo === hi ? "" : "s"} ${span} — ${arrows} extend, Ctrl+${arrows} jump, Enter confirm, Esc cancel`;
  }

  function axisPickStep(delta: number, jump: boolean) {
    if (!pendingAxisPick) return;
    const step = (jump ? AXIS_PICK_JUMP : 1) * delta;
    const max = pendingAxisPick.axis === "row" ? ABS_MAX_ROWS : ABS_MAX_COLS;
    pendingAxisPick.end = Math.max(1, Math.min(max, pendingAxisPick.end + step));
    if (pendingAxisPick.axis === "row") {
      growViewportToInclude(pendingAxisPick.end, 1);
    } else {
      growViewportToInclude(1, pendingAxisPick.end);
    }
    paintAxisPickSelection();
    describeAxisPick();
  }

  function cancelAxisPick() {
    if (!pendingAxisPick) return;
    selRow = pendingAxisPick.origSelRow;
    selCol = pendingAxisPick.origSelCol;
    rangeEndRow = pendingAxisPick.origRangeEndRow;
    rangeEndCol = pendingAxisPick.origRangeEndCol;
    pendingAxisPick = null;
    menuMessage = null;
    statusMsg = "Cancelled";
    focusGrid();
  }

  async function commitAxisPick() {
    if (!pendingAxisPick) return;
    const { axis, op, start, end } = pendingAxisPick;
    const lo = Math.min(start, end);
    const hi = Math.max(start, end);
    pendingAxisPick = null;
    menuMessage = null;
    if (op === "auto") {
      await applyAxisResize(axis, lo, hi, "auto");
      focusGrid();
      return;
    }
    // set-size: prompt for the value. "auto"/"a"/"0" all trigger auto-fit.
    const labelSpan = axis === "row"
      ? (lo === hi ? `row ${lo}` : `rows ${lo}–${hi}`)
      : (lo === hi
          ? `col ${addr(1, lo).replace(/\d+$/, "")}`
          : `cols ${addr(1, lo).replace(/\d+$/, "")}–${addr(1, hi).replace(/\d+$/, "")}`);
    const current = axis === "row"
      ? (rowHeights.get(lo) ?? 19)
      : (colWidths.get(lo) ?? 73);
    const what = axis === "row" ? "Height" : "Width";
    openMenuPrompt(
      `${what} for ${labelSpan} (px, or "auto"):`,
      String(current),
      async (v) => {
        const trimmed = v.trim().toLowerCase();
        if (trimmed === "auto" || trimmed === "a" || trimmed === "0") {
          await applyAxisResize(axis, lo, hi, "auto");
          focusGrid();
          return;
        }
        const px = Number(trimmed);
        if (!Number.isFinite(px) || px <= 0) {
          statusMsg = `Invalid ${axis === "row" ? "height" : "width"}: ${v}`;
          focusGrid();
          return;
        }
        await applyAxisResize(axis, lo, hi, px);
        focusGrid();
      },
    );
  }

  /// Apply a uniform size (or per-index auto-fit) across [lo..hi] on the
  /// chosen axis. For auto-fit we ensure the range's cells are loaded
  /// first, since paged fetch may have only the visible band — without
  /// the cell data autoFit would fall back to the per-row default.
  async function applyAxisResize(
    axis: "row" | "col",
    lo: number,
    hi: number,
    size: number | "auto",
  ) {
    const sheet = activeSheet;
    if (axis === "row" && size === "auto") {
      await ensureRowsLoaded(lo, hi);
    }
    let count = 0;
    try {
      for (let i = lo; i <= hi; i++) {
        const px = size === "auto"
          ? (axis === "row" ? autoFitRowPx(cells, i) : autoFitColumnPx(cells, i))
          : size;
        if (axis === "row") {
          await invoke("set_row_height", { sheet, row: i, px });
        } else {
          await invoke("set_column_width", { sheet, col: i, px });
        }
        count++;
      }
    } catch (e) {
      statusMsg = `Resize failed at ${axis} ${lo + count}: ${e}`;
      return;
    }
    await refreshViewport();
    const label = axis === "row" ? "row" : "column";
    const span = lo === hi ? `${lo}` : `${lo}–${hi}`;
    statusMsg = size === "auto"
      ? `Auto-fit ${count} ${label}${count === 1 ? "" : "s"} (${span})`
      : `Set ${count} ${label}${count === 1 ? "" : "s"} ${axis === "row" ? "height" : "width"} = ${Math.round(size)}px (${span})`;
  }

  const menu = buildMenu({
    newWorkbook,
    eraseCurrentCell,
    openRetrieveNavigator,
    fileSaveFlow,
    quitApp,
    setStatus: (m) => { statusMsg = m; },
    hideColumn: hideCurrentColumn,
    showAllColumns,
    setColumnWidth: setColumnWidthPrompt,
    autoColumnWidth,
    hideRow: hideCurrentRow,
    showAllRows,
    setRowHeight: setRowHeightPrompt,
    autoRowHeight,
    setTitles,
    insertRows: insertRowsAtSel,
    deleteRows: deleteRowsAtSel,
    insertColumns: insertColumnsAtSel,
    deleteColumns: deleteColumnsAtSel,
    formatRange,
    searchRange: openFindReplace,
    alignRange,
    setFillColor,
    setTextColor,
    rangeValue,
    rangeTrans,
    copyRange,
    moveRange,
    dataFill,
    dataSort,
    nameCreate,
    nameDelete,
    nameList,
    setBorder,
    sheetNew: addSheet,
    sheetDelete: () => deleteSheetConfirm(activeSheet),
    sheetRename: () => renameSheetPrompt(activeSheet),
    traceFormula,
    traceGoto,
    traceNames,
    compareOpen,
    compareExit,
  });

  let levelItems = $derived(
    menuOpen ? currentLevel(menu, menuPath, dynamicLevel) : [],
  );
  let levelHighlight = $derived(
    Math.min(menuHighlight, Math.max(0, levelItems.length - 1)),
  );
  let levelDescription = $derived(
    levelItems[levelHighlight]?.description ?? "",
  );
  let breadcrumbText = $derived(
    breadcrumb(menu, menuPath, dynamicLevel, dynamicTitle),
  );

  function openMenu() {
    menuOpen = true;
    menuPath = [];
    menuHighlight = 0;
  }
  function closeMenu() {
    menuOpen = false;
    menuPath = [];
    menuHighlight = 0;
    dynamicLevel = null;
    dynamicTitle = "";
    focusGrid();
  }

  async function selectMenuItem(idx: number) {
    const items = currentLevel(menu, menuPath, dynamicLevel);
    const item = items[idx];
    if (!item) return;
    if (item.children) {
      menuPath = [...menuPath, idx];
      menuHighlight = 0;
    } else if (item.action) {
      const action = item.action;
      closeMenu();
      await action();
    } else {
      closeMenu();
    }
  }

  function popMenu() {
    // Ad-hoc dynamic level escapes straight back to the grid, since it has
    // no parent in the static MENU tree.
    if (dynamicLevel) {
      closeMenu();
      return;
    }
    if (menuPath.length === 0) {
      closeMenu();
    } else {
      menuPath = menuPath.slice(0, -1);
      menuHighlight = 0;
    }
  }

  async function handleMenuKey(e: KeyboardEvent) {
    e.preventDefault();
    const items = currentLevel(menu, menuPath, dynamicLevel);
    switch (e.key) {
      case "Escape":
        popMenu();
        return;
      case "ArrowLeft":
        menuHighlight = Math.max(0, menuHighlight - 1);
        return;
      case "ArrowRight":
        menuHighlight = Math.min(items.length - 1, menuHighlight + 1);
        return;
      case "Home":
        menuHighlight = 0;
        return;
      case "End":
        menuHighlight = items.length - 1;
        return;
      case "Enter":
        await selectMenuItem(menuHighlight);
        return;
    }
    if (e.key.length === 1) {
      const ch = e.key.toUpperCase();
      const idx = items.findIndex((it) => it.letter === ch);
      if (idx >= 0) {
        await selectMenuItem(idx);
      }
    }
  }

  function onKey(e: KeyboardEvent) {
    // Always defang webview reload shortcuts — they'd reset the workbook
    // process state. Must run BEFORE any modal early-return below; F5
    // also reaches here when the menuPrompt input is focused and the
    // window listener sees a bubble.
    if (e.key === "F5" || (e.ctrlKey && (e.key === "r" || e.key === "R"))) {
      e.preventDefault();
      // Only open the goto prompt when nothing else owns the keyboard.
      if (
        e.key === "F5" &&
        !menuPrompt && !navOpen && !menuOpen && !pendingMove && !pendingAxisPick && !editing
      ) {
        openF5GotoPrompt();
      }
      return;
    }
    // Inline menu prompt (Set-Width / Set-Height etc.) owns all keys while
    // visible — it has its own input.
    if (menuPrompt) {
      return;
    }
    // Trace popup owns all keys while visible (modal or docked) —
    // its own listener is on the capture phase, but we also gate
    // here so selection-moving keys don't bleed through. While
    // hidden the popup is collapsed to a tiny status bar and gives
    // up the keyboard, so the grid runs as normal.
    if (traceRoot && !traceHidden) {
      return;
    }
    // Compare panel is also keyboard-modal when visible. When hidden
    // (its tiny bar) it gives up keys to the grid the same way the
    // trace popup does — H brings it back.
    if (compareResult && !compareHidden) {
      return;
    }
    // Pending /Copy or /Move — arrow keys steer the destination, Enter
    // commits, Esc cancels. All other keys are swallowed so they can't
    // disturb the source selection.
    if (pendingMove) {
      e.preventDefault();
      switch (e.key) {
        case "ArrowUp": pendingMoveStep(-1, 0); return;
        case "ArrowDown": pendingMoveStep(1, 0); return;
        case "ArrowLeft": pendingMoveStep(0, -1); return;
        case "ArrowRight": pendingMoveStep(0, 1); return;
        case "Enter": commitPendingMove(); return;
        case "Escape": cancelPendingMove(); return;
      }
      return;
    }
    // Axis pick (Set-Width/Set-Height/Auto). Arrows extend the row or
    // column range; Ctrl+arrow jumps; Enter confirms; Esc cancels. Wrong-
    // axis arrows are ignored so a misaimed key doesn't lose the pick.
    if (pendingAxisPick) {
      e.preventDefault();
      const isRow = pendingAxisPick.axis === "row";
      switch (e.key) {
        case "ArrowUp":   if (isRow)  axisPickStep(-1, e.ctrlKey); return;
        case "ArrowDown": if (isRow)  axisPickStep( 1, e.ctrlKey); return;
        case "ArrowLeft": if (!isRow) axisPickStep(-1, e.ctrlKey); return;
        case "ArrowRight":if (!isRow) axisPickStep( 1, e.ctrlKey); return;
        case "Enter":     commitAxisPick(); return;
        case "Escape":    cancelAxisPick(); return;
      }
      return;
    }
    // Menu is modal: it owns all keys while open.
    if (menuOpen) {
      handleMenuKey(e);
      return;
    }
    // Navigator owns all keys while open (it has its own filter input).
    if (navOpen) {
      return;
    }
    // Cell editing — handle Enter/Esc here. Must come BEFORE the generic
    // INPUT-focused guard below, otherwise the editor input swallows Enter
    // and our commit logic never runs.
    if (editing) {
      if (e.key === "Enter") {
        e.preventDefault();
        commitEdit().then(() => moveSel(1, 0));
      } else if (e.key === "Tab") {
        e.preventDefault();
        commitEdit().then(() => moveSel(0, e.shiftKey ? -1 : 1));
      } else if (e.key === "Escape") {
        e.preventDefault();
        cancelEdit();
      } else if (e.key === "F4") {
        e.preventDefault();
        f4Toggle();
      }
      return;
    }
    // Tab/Shift+Tab outside the editor: move horizontally without
    // editing. Excel/Lotus convention.
    if (e.key === "Tab") {
      e.preventDefault();
      moveSel(0, e.shiftKey ? -1 : 1);
      return;
    }
    // The path input (and any other ad-hoc input) handles its own keys —
    // we just let them through, except Escape, which blurs back to the grid.
    if ((document.activeElement as HTMLElement | null)?.tagName === "INPUT") {
      if (e.key === "Escape") {
        (document.activeElement as HTMLElement).blur();
        e.preventDefault();
      }
      return;
    }
    if (e.key === "/") {
      e.preventDefault();
      openMenu();
      return;
    }
    // Lotus `.` — cycle the anchor among the corners of the active
    // selection. Lets the user extend the range in any direction with
    // shift+arrow without first collapsing it.
    if (e.key === ".") {
      e.preventDefault();
      cycleAnchor();
      return;
    }
    // Clipboard shortcuts. Ctrl+C/X copy/cut the selection; Ctrl+V pastes
    // at the cursor. Excel/Google-Sheets-compatible TSV via the OS
    // clipboard. Doesn't preempt single-cell editing (the editing branch
    // returned earlier).
    if (e.ctrlKey && !e.altKey && !e.metaKey) {
      const k = e.key.toLowerCase();
      if (!e.shiftKey) {
        switch (k) {
          case "c":
            e.preventDefault();
            copySelection(false);
            return;
          case "x":
            e.preventDefault();
            copySelection(true);
            return;
          case "v":
            e.preventDefault();
            pasteFromClipboard();
            return;
          case "z":
            e.preventDefault();
            undo();
            return;
          case "y":
            e.preventDefault();
            redo();
            return;
          case "f":
            e.preventDefault();
            openFindReplace();
            return;
          case "b":
            e.preventDefault();
            applyStyleOp({ kind: "toggle_bold" }, "Toggled bold");
            return;
          case "i":
            e.preventDefault();
            applyStyleOp({ kind: "toggle_italic" }, "Toggled italic");
            return;
          case "u":
            e.preventDefault();
            applyStyleOp({ kind: "toggle_underline" }, "Toggled underline");
            return;
        }
      } else {
        // Ctrl+Shift+Z is the alternative redo binding (matches the
        // common cross-app convention).
        if (k === "z") {
          e.preventDefault();
          redo();
          return;
        }
      }
    }
    switch (e.key) {
      case "ArrowUp":
        e.preventDefault();
        if (e.ctrlKey) jumpEdge(-1, 0);
        else moveSel(-1, 0, e.shiftKey);
        break;
      case "ArrowDown":
        e.preventDefault();
        if (e.ctrlKey) jumpEdge(1, 0);
        else moveSel(1, 0, e.shiftKey);
        break;
      case "ArrowLeft":
        e.preventDefault();
        if (e.ctrlKey) jumpEdge(0, -1);
        else moveSel(0, -1, e.shiftKey);
        break;
      case "ArrowRight":
        e.preventDefault();
        if (e.ctrlKey) jumpEdge(0, 1);
        else moveSel(0, 1, e.shiftKey);
        break;
      case "Enter":
      case "F2":
        e.preventDefault();
        startEdit();
        break;
      case "F3":
        e.preventDefault();
        if (findResults.length === 0) {
          statusMsg = "No active find. Use Ctrl+F to start one.";
        } else {
          jumpToMatch(findIdx + (e.shiftKey ? -1 : 1));
        }
        break;
      // F5 is now handled at the top of onKey (the reload-defang block);
      // case kept here just to avoid falling into the printable-char
      // catch-all below.
      case "F5":
        break;
      case "F9":
        e.preventDefault();
        recalcWorkbook();
        break;
      case "PageUp":
        e.preventDefault();
        if (e.ctrlKey) switchSheet(activeSheet - 1);
        else pageScroll(-1, e.shiftKey);
        break;
      case "PageDown":
        e.preventDefault();
        if (e.ctrlKey) switchSheet(activeSheet + 1);
        else pageScroll(1, e.shiftKey);
        break;
      case "Home":
        if (e.ctrlKey) {
          e.preventDefault();
          goHome();
        }
        break;
      case "End":
        if (e.ctrlKey) {
          e.preventDefault();
          goEnd();
        }
        break;
      case "Delete":
      case "Backspace":
        e.preventDefault();
        eraseCurrentCell();
        break;
      default:
        if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
          e.preventDefault();
          startEdit(e.key);
        }
    }
  }

  function blockContextMenu(e: MouseEvent) {
    // Block the browser's native right-click menu app-wide so the only
    // context menu in fastsheet is ours. Our cell handler intercepts
    // before bubbling here.
    e.preventDefault();
  }

  onMount(() => {
    window.addEventListener("keydown", onKey);
    window.addEventListener("contextmenu", blockContextMenu);
    invoke("profile_mark", { label: "onMount" }).catch(() => {});
    // Seed something usable on launch — either the file passed via
    // "Open with" / shell association or, failing that, a blank
    // untitled workbook. Without one of these, the grid renders a
    // "ghost" blank: looks like an empty spreadsheet but `workbook`
    // is null on the frontend and no Model exists on the backend,
    // so save fails and the selection overlay misaligns (colWidths
    // and rowHeights are empty maps).
    invoke<string | null>("take_startup_path")
      .then(async (p) => {
        if (p) {
          await openWorkbookFromPath(p);
          // If openWorkbookFromPath failed (file missing, parse
          // error, ...) workbook is still null — drop into a blank
          // so the launch state stays usable.
          if (!workbook) await newWorkbook();
        } else {
          await newWorkbook();
        }
      })
      .catch((e) => {
        console.warn("initial workbook seed failed:", e);
        newWorkbook().catch(() => {});
      });
    // After first paint — the moment the user can actually see and
    // interact with the grid. This is the metric they care about.
    tick().then(() => {
      requestAnimationFrame(() => {
        invoke("profile_mark", { label: "first_paint" }).catch(() => {});
      });
    });
    return () => {
      window.removeEventListener("keydown", onKey);
      window.removeEventListener("contextmenu", blockContextMenu);
    };
  });

  let selCell = $derived(cells.get(key(selRow, selCol)));

  /// Reflect open file + dirty state in the window title.
  $effect(() => {
    const sep = currentPath.includes("\\") ? "\\" : "/";
    const idx = currentPath.lastIndexOf(sep);
    const base = currentPath ? currentPath.slice(idx + 1) : "untitled";
    const dirty = dirtyEdits > 0 ? "● " : "";
    if (typeof document !== "undefined") {
      document.title = `${dirty}${base} — fastsheet`;
    }
  });

  /// Excel-style status-bar summary: when the selection covers more than
  /// one cell, count non-empty cells and show sum + average across the
  /// numeric subset. Cheap — iterates the local cells map, no backend.
  let selectionSummary = $derived.by(() => {
    if (selRow === rangeEndRow && selCol === rangeEndCol) return "";
    const r1 = Math.min(selRow, rangeEndRow);
    const r2 = Math.max(selRow, rangeEndRow);
    const c1 = Math.min(selCol, rangeEndCol);
    const c2 = Math.max(selCol, rangeEndCol);
    let count = 0;
    let numCount = 0;
    let sum = 0;
    for (let r = r1; r <= r2; r++) {
      for (let c = c1; c <= c2; c++) {
        const cell = cells.get(`${r}:${c}`);
        const t = cell?.text;
        if (!t) continue;
        count++;
        const n = Number(t);
        if (Number.isFinite(n) && t.trim() !== "") {
          sum += n;
          numCount++;
        }
      }
    }
    if (count === 0) return `Count: 0`;
    if (numCount === 0) return `Count: ${count}`;
    const avg = sum / numCount;
    const fmt = (x: number) => Number.isInteger(x) ? `${x}` : x.toFixed(4).replace(/\.?0+$/, "");
    return `Count: ${count} · Sum: ${fmt(sum)} · Avg: ${fmt(avg)}`;
  });
</script>

<div class="app">
  {#if menuOpen}
    <div class="menubar" role="menu">
      <span class="menu-prompt">{breadcrumbText}&gt;</span>
      {#each levelItems as item, i}
        <span class="menu-item" class:hi={i === levelHighlight}>
          <span class="letter">{item.letter}</span>{item.label.slice(1)}
        </span>
      {/each}
    </div>
    <div class="menu-desc">{levelDescription}</div>
  {/if}
  {#if menuMessage && !menuOpen && !menuPrompt}
    <div class="menu-prompt-bar">
      <span class="menu-prompt-label">{menuMessage}</span>
    </div>
  {/if}
  {#if menuPrompt}
    <div class="menu-prompt-bar">
      <span class="menu-prompt-label">{menuPrompt.label}</span>
      <input
        class="menu-prompt-input"
        bind:this={menuPromptEl}
        bind:value={menuPrompt.value}
        oninput={() => menuPromptHighlight = -1}
        onkeydown={(e) => {
          if (e.key === "Enter") { e.preventDefault(); submitMenuPrompt(); }
          else if (e.key === "Escape") { e.preventDefault(); cancelMenuPrompt(); }
          else if (menuPrompt?.candidates && (e.key === "Tab" || e.key === "ArrowDown")) {
            e.preventDefault();
            moveMenuPromptHighlight(e.shiftKey ? -1 : 1);
          } else if (menuPrompt?.candidates && e.key === "ArrowUp") {
            e.preventDefault();
            moveMenuPromptHighlight(-1);
          }
        }}
      />
      {#if menuPrompt.candidates && promptMatches.length > 0}
        <div class="menu-prompt-suggestions">
          {#each promptMatches as cand, i}
            <span
              class="menu-prompt-cand"
              class:hi={i === menuPromptHighlight}
            >{cand}</span>
          {/each}
        </div>
      {/if}
    </div>
  {/if}

  <header class="topbar">
    <span class="brand">fastsheet</span>
    <span class="addr">{addr(selRow, selCol)}</span>
    <span class="formula-bar">
      {#if editing}
        <input
          class="editor"
          bind:this={editorEl}
          bind:value={editValue}
          onblur={commitEdit}
        />
      {:else}
        {selCell?.input ?? ""}
      {/if}
    </span>
  </header>

  {#if navOpen}
    <Navigator
      mode={navMode}
      {currentPath}
      onOpenFile={async (p) => {
        navOpen = false;
        if (navCompareTarget) {
          navCompareTarget = false;
          await compareOpenWith(p);
        } else {
          await openWorkbookFromPath(p);
        }
      }}
      onSaveFile={async (p) => {
        navOpen = false;
        await saveAsWithConfirm(p);
      }}
      onClose={() => {
        navOpen = false;
        focusGrid();
      }}
      onStatus={(m) => (statusMsg = m)}
    />
  {/if}

  {#if traceRoot}
    <FormulaTrace
      root={traceRoot}
      onClose={() => closeTrace(true)}
      onPreview={tracePreview}
      bind:docked={traceDocked}
      bind:hidden={traceHidden}
      onJump={async (node) => {
        // Resolve coords by kind — name nodes carry their range in
        // node.value; cell / range nodes have explicit (sheet,row,col).
        let target: { sheet: number; row: number; col: number } | null = null;
        if (node.kind === "name") {
          const range = parseNameFormula(node.value);
          if (range) target = { sheet: range.sheet, row: range.r1, col: range.c1 };
        } else if (node.sheet !== null && node.row !== null && node.col !== null) {
          target = { sheet: node.sheet, row: node.row, col: node.col };
        }
        if (!target) return;
        if (target.sheet !== activeSheet) await switchSheet(target.sheet);
        selRow = target.row;
        selCol = target.col;
        rangeEndRow = target.row;
        rangeEndCol = target.col;
        growViewportToInclude(target.row, target.col);
        await closeTrace(false);
      }}
    />
  {/if}

  {#if compareResult}
    <CompareDiff
      diffs={compareResult.diffs}
      missingSheets={compareResult.missing_sheets}
      rightPath={compareResult.right_path}
      totalDiffs={compareResult.total_diffs}
      capped={compareResult.diffs_capped}
      onClose={compareExit}
      onJump={compareJumpTo}
      onPreview={comparePreview}
      bind:hidden={compareHidden}
    />
  {/if}

  <Grid
    {cells}
    {colWidths}
    {rowHeights}
    rows={viewportRows}
    cols={viewportCols}
    {frozenRows}
    {frozenCols}
    {mergedRanges}
    {ghostRange}
    highlights={[
      ...refHighlights.filter((h) => h.sheet === activeSheet),
      ...compareHighlights.filter((h) => h.sheet === activeSheet),
    ]}
    {scrollTarget}
    bind:selRow
    bind:selCol
    bind:rangeEndRow
    bind:rangeEndCol
    bind:gridWrapEl
    onDblClick={() => startEdit()}
    onResizeRow={async (r, px) => {
      try {
        await invoke("set_row_height", { sheet: activeSheet, row: r, px });
        await refreshViewport();
      } catch (e) {
        statusMsg = `Resize row failed: ${e}`;
      }
    }}
    onContextMenu={(x, y) => openContextMenu(x, y)}
    onFill={fillFromHandle}
    onResizeCol={async (c, px) => {
      try {
        await invoke("set_column_width", { sheet: activeSheet, col: c, px });
        await refreshViewport();
      } catch (e) {
        statusMsg = `Resize col failed: ${e}`;
      }
    }}
    onBandChange={handleBandChange}
    onColBandChange={handleColBandChange}
  />

  {#if workbook}
    <SheetTabs
      sheetNames={workbook.sheet_names}
      activeIndex={activeSheet}
      onSelect={(i) => switchSheet(i)}
      onAddSheet={addSheet}
      onTabContextMenu={openTabContextMenu}
    />
  {/if}

  {#if tabContextMenu}
    <div
      class="ctx-menu"
      style={`left:${tabContextMenu.x}px; top:${tabContextMenu.y}px;`}
    >
      <button type="button" onclick={() => { const s = tabContextMenu!.sheet; closeTabContextMenu(); renameSheetPrompt(s); }}>Rename…</button>
      <button type="button" onclick={() => { const s = tabContextMenu!.sheet; closeTabContextMenu(); deleteSheetConfirm(s); }}>Delete</button>
      <hr />
      <button type="button" onclick={() => { closeTabContextMenu(); addSheet(); }}>Add new sheet</button>
    </div>
  {/if}

  {#if contextMenu}
    <div
      class="ctx-menu"
      style={`left:${contextMenu.x}px; top:${contextMenu.y}px;`}
    >
      <button type="button" onclick={() => { closeContextMenu(); copySelection(true); }}>Cut</button>
      <button type="button" onclick={() => { closeContextMenu(); copySelection(false); }}>Copy</button>
      <button type="button" onclick={() => { closeContextMenu(); pasteFromClipboard(); }}>Paste</button>
      <hr />
      <button type="button" onclick={() => { closeContextMenu(); insertRowsAtSel(); }}>Insert row above</button>
      <button type="button" onclick={() => { closeContextMenu(); insertColumnsAtSel(); }}>Insert column left</button>
      <button type="button" onclick={() => { closeContextMenu(); deleteRowsAtSel(); }}>Delete row</button>
      <button type="button" onclick={() => { closeContextMenu(); deleteColumnsAtSel(); }}>Delete column</button>
      <hr />
      <button type="button" onclick={() => { closeContextMenu(); hideCurrentRow(); }}>Hide row</button>
      <button type="button" onclick={() => { closeContextMenu(); hideCurrentColumn(); }}>Hide column</button>
      <button type="button" onclick={() => { closeContextMenu(); eraseCurrentCell(); }}>Erase</button>
    </div>
  {/if}

  <footer class="status">
    <span class="status-cell">
      {workbook?.sheet_names[activeSheet] ?? ""}!{addr(selRow, selCol)}
    </span>
    {#if currentPath}<span class="path-tag">{currentPath}</span>{/if}
    {#if dirtyEdits > 0}
      <span class="dirty-tag"
        >● {dirtyEdits} edit{dirtyEdits > 1 ? "s" : ""} pending recalc (F9)</span
      >
    {/if}
    {#if selectionSummary}
      <span class="sel-summary">{selectionSummary}</span>
    {/if}
    {statusMsg}
  </footer>
</div>

<style>
  :global(html, body) {
    margin: 0;
    padding: 0;
    height: 100%;
    overflow: hidden;
    /* Light mode by default — matches Excel so file colors render with the
       contrast they were authored for (e.g., black text on a yellow fill). */
    font-family: "Calibri", "Segoe UI", Arial, sans-serif;
    font-size: 11pt;
    background: #ffffff;
    color: #111111;
  }
  .app {
    display: flex;
    flex-direction: column;
    height: 100vh;
  }
  /* Lotus-style menu stays yellow on dark — this is the defining visual
     and contrasts well against the light grid below. */
  .menubar {
    background: #f0c419;
    color: #111;
    padding: 0.25rem 0.6rem;
    font-size: 12px;
    display: flex;
    flex-wrap: wrap;
    align-items: center;
    gap: 1.2rem;
    border-bottom: 1px solid #b88a00;
  }
  .menu-prompt {
    color: #111;
    font-weight: 700;
    margin-right: 0.4rem;
  }
  .menu-item {
    padding: 0.05rem 0.3rem;
    border-radius: 2px;
  }
  .menu-item.hi {
    background: #111;
    color: #f0c419;
  }
  .menu-item .letter {
    text-decoration: underline;
    font-weight: 700;
  }
  .menu-desc {
    background: #2b2b2b;
    color: #ddd;
    padding: 0.2rem 0.6rem;
    font-size: 11px;
    border-bottom: 1px solid #444;
    min-height: 1.1em;
  }
  /* Inline menu prompt bar — appears below the menu while waiting for an
     input value (e.g. /Worksheet/Column/Set-Width). Mirrors the menu desc
     bar's dark style with a writable input on the right. */
  .menu-prompt-bar {
    background: #2b2b2b;
    color: #ddd;
    padding: 0.2rem 0.6rem;
    font-size: 12px;
    border-bottom: 1px solid #444;
    display: flex;
    align-items: center;
    gap: 0.5rem;
  }
  .menu-prompt-label {
    color: #f0c419;
    font-weight: 600;
  }
  .menu-prompt-input {
    flex: 1;
    background: #111;
    color: #fff;
    border: 1px solid #444;
    padding: 0.1rem 0.4rem;
    font: inherit;
  }
  /* Suggestion strip under the prompt input. Tab cycles; the highlighted
     entry is what Enter would submit. */
  .menu-prompt-suggestions {
    flex: 0 0 100%;
    display: flex;
    flex-wrap: wrap;
    gap: 0.4rem;
    padding-top: 0.2rem;
    color: #aaa;
    font-size: 11px;
  }
  .menu-prompt-cand {
    padding: 0.05rem 0.4rem;
    border: 1px solid transparent;
    border-radius: 2px;
  }
  .menu-prompt-cand.hi {
    background: #f0c419;
    color: #111;
    border-color: #b88a00;
    font-weight: 700;
  }
  /* Right-click cell context menu — floating over the grid. */
  .ctx-menu {
    position: fixed;
    z-index: 100;
    background: #fff;
    border: 1px solid #c0c0c0;
    box-shadow: 0 4px 12px rgba(0, 0, 0, 0.18);
    padding: 0.2rem 0;
    min-width: 11rem;
    font-size: 12px;
  }
  .ctx-menu button {
    display: block;
    width: 100%;
    text-align: left;
    background: none;
    border: none;
    padding: 0.25rem 0.8rem;
    color: #222;
    font: inherit;
    cursor: pointer;
  }
  .ctx-menu button:hover {
    background: #1f6feb;
    color: #fff;
  }
  .ctx-menu hr {
    border: none;
    border-top: 1px solid #e0e0e0;
    margin: 0.2rem 0;
  }
  .topbar {
    display: flex;
    align-items: center;
    gap: 0.5rem;
    padding: 0.3rem 0.5rem;
    background: #f3f3f3;
    border-bottom: 1px solid #d0d0d0;
    font-size: 12px;
  }
  .brand {
    font-weight: 700;
    color: #b88a00;
    letter-spacing: 0.05em;
  }
  .addr {
    margin-left: 1rem;
    color: #444;
    font-weight: 600;
    min-width: 4rem;
  }
  .formula-bar {
    flex: 1;
    background: #fff;
    color: #111;
    border: 1px solid #c0c0c0;
    padding: 0.15rem 0.4rem;
    min-height: 1.3rem;
    overflow: hidden;
    white-space: nowrap;
  }
  .editor {
    width: 100%;
    background: transparent;
    color: #111;
    border: none;
    outline: none;
    font: inherit;
  }
  .status {
    background: #f3f3f3;
    border-top: 1px solid #d0d0d0;
    padding: 0.25rem 0.6rem;
    font-size: 11px;
    color: #444;
    display: flex;
    gap: 0.6rem;
    align-items: center;
  }
  .path-tag {
    color: #1f6feb;
    font-weight: 600;
  }
  .dirty-tag {
    color: #d4691e;
    font-weight: 600;
  }
  .sel-summary {
    color: #1f6feb;
    font-weight: 600;
    font-variant-numeric: tabular-nums;
  }
  .status-cell {
    color: #444;
    font-weight: 600;
    font-variant-numeric: tabular-nums;
  }
</style>
