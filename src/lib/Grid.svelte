<script lang="ts">
  import { flushSync, tick, untrack } from "svelte";
  import type { CellView } from "./types";
  import {
    autoFitColumnPx,
    autoFitRowPx,
    cellTdStyle,
    cellContentStyle,
    colLetter,
    key,
    parseA1,
  } from "./utils";

  type Props = {
    cells: Map<string, CellView>;
    colWidths: Map<number, number>;
    rowHeights: Map<number, number>;
    rows: number;
    cols: number;
    selRow: number;
    selCol: number;
    /// Opposite corner of the active selection rectangle. When equal to
    /// (selRow, selCol) the selection is a single cell.
    rangeEndRow: number;
    rangeEndCol: number;
    /// Number of rows pinned to the top (sticky). 0 = no horizontal freeze.
    frozenRows: number;
    /// Number of cols pinned to the left (sticky). 0 = no vertical freeze.
    frozenCols: number;
    showGridLines?: boolean;
    pageBreakRows?: Set<number>;
    pageBreakCols?: Set<number>;
    /// A1-style merged ranges from worksheet.merge_cells (e.g. "A1:B2").
    mergedRanges: string[];
    /// Optional "ghost" overlay rectangle — used for /Copy /Move
    /// destination preview. Cells in this range get a dashed outline
    /// without affecting the active selection.
    ghostRange: { r1: number; c1: number; r2: number; c2: number } | null;
    /// Reference highlights — used by the formula-trace popup and F2
    /// edit-mode to outline cells/ranges the formula refers to. Each
    /// entry is one outlined box, independent of the selection. Cells
    /// here aren't part of any range; they're purely visual hints.
    /// Optional `label` renders a small tag at the top-left — used to
    /// show the name of a defined-range reference.
    highlights?: { r1: number; c1: number; r2: number; c2: number; color: string; label?: string }[];
    /// When set, the grid auto-scrolls to bring (row, col) into view
    /// without changing the selection cursor. Used by trace preview
    /// to follow the highlighted entry through the workbook. The page
    /// is expected to set this once per highlight transition; the
    /// effect re-runs on each new identity (set to null and back to
    /// retrigger).
    scrollTarget?: { row: number; col: number } | null;
    gridWrapEl?: HTMLDivElement | null;
    onDblClick: () => void;
    onResizeRow: (row: number, px: number) => void | Promise<void>;
    onResizeCol: (col: number, px: number) => void | Promise<void>;
    onContextMenu: (x: number, y: number, row: number, col: number) => void;
    onFill: (
      src: { r1: number; c1: number; r2: number; c2: number },
      dest: { r1: number; c1: number; r2: number; c2: number },
    ) => void;
    /// Fired whenever the visible row band changes (scroll, resize, sheet
    /// switch). The page uses this to drive paged cell-range fetching —
    /// rows entering the viewport are loaded on demand. Callback is
    /// expected to be cheap (rAF-coalesced); this $effect runs every time
    /// the band edges shift.
    onBandChange?: (start: number, end: number) => void;
    /// Same as onBandChange, but for the visible column band. The page
    /// uses this to scope per-fetch cell ranges so navigating far right
    /// doesn't make every row-band fetch span thousands of empty cols.
    onColBandChange?: (start: number, end: number) => void;
  };

  let {
    cells,
    colWidths,
    rowHeights,
    rows,
    cols,
    selRow = $bindable(),
    selCol = $bindable(),
    rangeEndRow = $bindable(),
    rangeEndCol = $bindable(),
    frozenRows,
    frozenCols,
    showGridLines = true,
    pageBreakRows = new Set(),
    pageBreakCols = new Set(),
    mergedRanges,
    ghostRange,
    highlights = [],
    scrollTarget = null,
    gridWrapEl = $bindable(null),
    onDblClick,
    onResizeRow,
    onResizeCol,
    onContextMenu,
    onFill,
    onBandChange,
    onColBandChange,
  }: Props = $props();

  // Cumulative geometry. rowOffsets[r] = top-edge px of row r (within the
  // table body, NOT counting the colhdr). colLefts[c] = left-edge px of
  // col c (counting the row-header column at index 0). Both arrays are
  // 1-based for natural row/col indexing.
  //
  // colhdrH is measured after mount — the rendered column-header height
  // depends on the browser's baseline font metrics, so a hardcoded
  // constant was off by a few px and made the selection outline drift
  // down from the cell. Initial value is a best-guess so the first
  // frame isn't too far off; the real value lands on mount.
  let colhdrH = $state(22);
  const ROWHDR_W = 42; // matches the colgroup row-header width
  const DEFAULT_ROW_H = 19;
  const DEFAULT_COL_W = 73;
  /// Row heights MEASURED off the DOM after render. Wrap-text cells
  /// auto-grow their row past the configured rowHeights value (the
  /// browser respects `<tr style="height:N">` only as a min height
  /// when content needs more), so rowOffsets built from rowHeights
  /// alone would under-predict and the cursor would shrink relative
  /// to the actual cell. Populated by the post-render $effect below;
  /// rowOffsets prefers measured values when present.
  let measuredRowHeights = $state<Map<number, number>>(new Map());
  let rowOffsets = $derived.by(() => {
    const out: number[] = new Array(rows + 2);
    out[1] = 0;
    for (let r = 1; r <= rows; r++) {
      const h = measuredRowHeights.get(r) ?? rowHeights.get(r) ?? DEFAULT_ROW_H;
      out[r + 1] = out[r] + h;
    }
    return out;
  });
  let totalRowH = $derived(rowOffsets[rows + 1] ?? 0);
  let colLefts = $derived.by(() => {
    const out: number[] = new Array(cols + 2);
    out[1] = ROWHDR_W;
    for (let c = 1; c <= cols; c++) {
      out[c + 1] = out[c] + (colWidths.get(c) ?? DEFAULT_COL_W);
    }
    return out;
  });

  let frozenRowTops = $derived.by(() => {
    const out: number[] = new Array(frozenRows + 1);
    let acc = colhdrH;
    for (let r = 1; r <= frozenRows; r++) {
      out[r] = acc;
      acc += rowHeights.get(r) ?? DEFAULT_ROW_H;
    }
    return out;
  });
  let frozenColLefts = $derived.by(() => {
    const out: number[] = new Array(frozenCols + 1);
    let acc = ROWHDR_W;
    for (let c = 1; c <= frozenCols; c++) {
      out[c] = acc;
      acc += colWidths.get(c) ?? DEFAULT_COL_W;
    }
    return out;
  });

  // Drag state for column / row resize. Live preview is the new size shown
  // via reactive style overrides; we only call the backend on mouseup.
  type DragState =
    | { kind: "col"; index: number; startCoord: number; original: number; current: number }
    | { kind: "row"; index: number; startCoord: number; original: number; current: number }
    | null;
  let drag = $state<DragState>(null);

  function startColResize(col: number, e: MouseEvent) {
    const w = colWidths.get(col) ?? 73;
    drag = { kind: "col", index: col, startCoord: e.clientX, original: w, current: w };
    e.preventDefault();
    e.stopPropagation();
  }

  function startRowResize(row: number, e: MouseEvent) {
    const h = rowHeights.get(row) ?? 19;
    drag = { kind: "row", index: row, startCoord: e.clientY, original: h, current: h };
    e.preventDefault();
    e.stopPropagation();
  }

  function onWindowMouseMove(e: MouseEvent) {
    if (!drag) return;
    const coord = drag.kind === "col" ? e.clientX : e.clientY;
    const next = Math.max(8, drag.original + (coord - drag.startCoord));
    drag = { ...drag, current: next };
  }

  async function onWindowMouseUp() {
    if (!drag) return;
    const final = drag;
    // Don't clear `drag` yet — the parent's resize handler is async
    // (it round-trips to the backend, then refreshViewport repopulates
    // colWidths / rowHeights from the new layout). Clearing drag too
    // early flips colWidthList back to the OLD value for the few
    // frames between drag=null and the layout refresh, producing a
    // visible "snap back to original then jump to new size" glitch.
    if (final.kind === "col") {
      await onResizeCol(final.index, final.current);
    } else {
      await onResizeRow(final.index, final.current);
    }
    drag = null;
  }

  // Auto-fit on resize-handle double-click. Measurement lives in utils
  // so the menu's Set-Width/Set-Height (with 0 input) can use the same
  // logic without re-implementing the canvas measurement.
  function autoFitColumn(col: number) {
    onResizeCol(col, autoFitColumnPx(cells, col));
  }
  function autoFitRow(row: number) {
    onResizeRow(row, autoFitRowPx(cells, row));
  }

  $effect(() => {
    if (drag) {
      window.addEventListener("mousemove", onWindowMouseMove);
      window.addEventListener("mouseup", onWindowMouseUp);
      return () => {
        window.removeEventListener("mousemove", onWindowMouseMove);
        window.removeEventListener("mouseup", onWindowMouseUp);
      };
    }
  });

  // Pre-computed per-column widths in render order so the colgroup is a
  // simple iteration over a $derived array (more reliable than reading the
  // Map from inside an {@const} which Svelte 5 doesn't always re-evaluate
  // when the Map is replaced).
  let colWidthList = $derived.by(() => {
    const out: number[] = new Array(cols);
    for (let i = 1; i <= cols; i++) {
      out[i - 1] =
        drag?.kind === "col" && drag.index === i
          ? drag.current
          : (colWidths.get(i) ?? 73);
    }
    return out;
  });

  function rowHeightFor(r: number): number | undefined {
    if (drag?.kind === "row" && drag.index === r) return drag.current;
    return rowHeights.get(r);
  }

  // Merged-cell processing: walk worksheet.merge_cells once per change
  // and produce two lookups — anchors get colspan/rowspan, all other
  // cells inside the merge are skipped at render time.
  type MergeAnchor = { colspan: number; rowspan: number };
  let mergeMap = $derived.by(() => {
    const anchors = new Map<string, MergeAnchor>();
    const skip = new Set<string>();
    for (const range of mergedRanges) {
      const parts = range.split(":");
      const a = parseA1(parts[0]);
      const b = parseA1(parts[1] ?? parts[0]);
      if (!a || !b) continue;
      const r1 = Math.min(a.row, b.row);
      const r2 = Math.max(a.row, b.row);
      const c1 = Math.min(a.col, b.col);
      const c2 = Math.max(a.col, b.col);
      anchors.set(`${r1}:${c1}`, { colspan: c2 - c1 + 1, rowspan: r2 - r1 + 1 });
      for (let r = r1; r <= r2; r++) {
        for (let c = c1; c <= c2; c++) {
          if (r === r1 && c === c1) continue;
          skip.add(`${r}:${c}`);
        }
      }
    }
    return { anchors, skip };
  });

  // Inclusive bounds of the current selection rectangle (anchor → opposite).
  let rangeBounds = $derived.by(() => ({
    r1: Math.min(selRow, rangeEndRow),
    r2: Math.max(selRow, rangeEndRow),
    c1: Math.min(selCol, rangeEndCol),
    c2: Math.max(selCol, rangeEndCol),
  }));

  // Which corner of the selection rectangle holds the "free" cell
  // (rangeEnd) — the fill handle renders there. Cycled by Lotus `.` in
  // the parent. With a single-cell selection rangeEnd === anchor and
  // the corner is "br" by default.
  let freeCorner = $derived.by<"tl" | "tr" | "bl" | "br">(() => {
    const top = rangeEndRow === rangeBounds.r1;
    const left = rangeEndCol === rangeBounds.c1;
    if (top && left) return "tl";
    if (top && !left) return "tr";
    if (!top && left) return "bl";
    return "br";
  });
  function inRange(r: number, c: number): boolean {
    return (
      r >= rangeBounds.r1 &&
      r <= rangeBounds.r2 &&
      c >= rangeBounds.c1 &&
      c <= rangeBounds.c2
    );
  }

  function inGhost(r: number, c: number): boolean {
    if (!ghostRange) return false;
    return (
      r >= ghostRange.r1 && r <= ghostRange.r2 &&
      c >= ghostRange.c1 && c <= ghostRange.c2
    );
  }

  // Active drag-to-select. Mousedown on a cell sets the anchor (or the
  // opposite corner if shift held) and arms drag mode; mouseenter on
  // other cells while armed updates the opposite corner; window-level
  // mouseup disarms.
  let dragSelecting = $state(false);

  function onCellMouseDown(r: number, c: number, e: MouseEvent) {
    if (e.button !== 0) return;
    if (e.shiftKey) {
      rangeEndRow = r;
      rangeEndCol = c;
    } else {
      selRow = r;
      selCol = c;
      rangeEndRow = r;
      rangeEndCol = c;
    }
    dragSelecting = true;
    e.preventDefault();
  }

  function onCellMouseEnter(r: number, c: number) {
    if (fillSource) {
      // Extend the selection from the source's anchor in whichever
      // dimension(s) the user is dragging into. We constrain to ≥ source
      // (Excel doesn't allow filling backwards from the handle).
      const src = fillSource;
      rangeEndRow = Math.max(src.r2, r);
      rangeEndCol = Math.max(src.c2, c);
      return;
    }
    if (!dragSelecting) return;
    rangeEndRow = r;
    rangeEndCol = c;
  }

  function endDragSelect() {
    dragSelecting = false;
  }

  $effect(() => {
    if (dragSelecting) {
      window.addEventListener("mouseup", endDragSelect);
      return () => window.removeEventListener("mouseup", endDragSelect);
    }
  });

  // Drag-fill ("fill handle"): user grabs the small square at the
  // bottom-right of the active selection and extends. On release, copy
  // the source pattern into the new (extended) area. v1 is copy-only —
  // arithmetic-progression detection is a future refinement.
  let fillSource = $state<{ r1: number; c1: number; r2: number; c2: number } | null>(null);

  function startFillDrag(e: MouseEvent) {
    e.preventDefault();
    e.stopPropagation();
    fillSource = { ...rangeBounds };
  }

  function endFillDrag() {
    if (!fillSource) return;
    const src = fillSource;
    fillSource = null;
    const dest = { ...rangeBounds };
    const same =
      dest.r1 === src.r1 && dest.c1 === src.c1 &&
      dest.r2 === src.r2 && dest.c2 === src.c2;
    if (same) return;
    onFill(src, dest);
  }

  $effect(() => {
    if (fillSource) {
      window.addEventListener("mouseup", endFillDrag);
      return () => window.removeEventListener("mouseup", endFillDrag);
    }
  });
  // Total table width so table-layout:fixed honours each col's width
  // instead of distributing leftover space evenly.
  let tableWidthPx = $derived(42 + colWidthList.reduce((a, b) => a + b, 0));

  // Track the wrap's scroll position so the visible-band derivation
  // can react to it. cached-scrollTop / cached-scrollLeft are mirrored
  // into $state via a scroll listener attached on first mount.
  let scrollTop = $state(0);
  let scrollLeft = $state(0);
  let viewportH = $state(600);
  let viewportW = $state(800);

  // Event delegation: cells carry data-r/data-c attrs, not per-cell
  // handlers. The table-level listeners below find the target <td> via
  // event.target.closest and dispatch to the existing logic. This cuts
  // ~4 handler attaches per cell created, which adds up to milliseconds
  // per scroll-band shift on content-heavy sheets.
  function cellFromEvent(e: Event): { r: number; c: number; el: HTMLTableCellElement } | null {
    const target = e.target as HTMLElement | null;
    if (!target) return null;
    const td = target.closest("td.cell") as HTMLTableCellElement | null;
    if (!td) return null;
    const r = parseInt(td.dataset.r ?? "0", 10);
    const c = parseInt(td.dataset.c ?? "0", 10);
    if (!r || !c) return null;
    return { r, c, el: td };
  }

  let lastHoveredCell: { r: number; c: number } | null = null;

  function onTableMouseDown(e: MouseEvent) {
    const hit = cellFromEvent(e);
    if (!hit) return;
    onCellMouseDown(hit.r, hit.c, e);
  }
  function onTableMouseOver(e: MouseEvent) {
    const hit = cellFromEvent(e);
    if (!hit) return;
    if (lastHoveredCell?.r === hit.r && lastHoveredCell?.c === hit.c) return;
    lastHoveredCell = { r: hit.r, c: hit.c };
    onCellMouseEnter(hit.r, hit.c);
  }
  function onTableDblClick(e: MouseEvent) {
    const hit = cellFromEvent(e);
    if (!hit) return;
    onDblClick();
  }
  function onTableContextMenu(e: MouseEvent) {
    const hit = cellFromEvent(e);
    if (!hit) return;
    e.preventDefault();
    if (!inRange(hit.r, hit.c)) {
      selRow = hit.r;
      selCol = hit.c;
      rangeEndRow = hit.r;
      rangeEndCol = hit.c;
    }
    onContextMenu(e.clientX, e.clientY, hit.r, hit.c);
  }

  $effect(() => {
    if (!gridWrapEl) return;
    const wrap = gridWrapEl;
    const onScroll = () => {
      scrollTop = wrap.scrollTop;
      scrollLeft = wrap.scrollLeft;
    };
    const onResize = () => {
      viewportH = wrap.clientHeight;
      viewportW = wrap.clientWidth;
      // Measure the actual top of the first tbody row, NOT
      // thead.offsetHeight. With border-collapse: collapse, the
      // border between thead and tbody is shared and offsetHeight
      // can land 1-2px short of the first row's actual offsetTop —
      // visible to the user as a "cursor sits 2px above the cell"
      // long-standing offset. tbody.offsetTop is the position the
      // browser actually places the first row at, including any
      // border-collapse rounding.
      const tbody = wrap.querySelector("tbody") as HTMLElement | null;
      if (tbody) colhdrH = tbody.offsetTop;
    };
    onResize();
    onScroll();
    wrap.addEventListener("scroll", onScroll, { passive: true });
    const ro = new ResizeObserver(onResize);
    ro.observe(wrap);
    return () => {
      wrap.removeEventListener("scroll", onScroll);
      ro.disconnect();
    };
  });

  /// After every render, walk the visible-band tbody rows and
  /// reconcile their actual rendered heights into measuredRowHeights.
  /// Wrap-text cells auto-grow their row past the configured height
  /// (browsers grow rows to fit multi-line content, ignoring
  /// `style="height:N"` as a minimum-only hint), so rowOffsets built
  /// from rowHeights alone under-predicts in those cases. Without
  /// this reconciliation the cursor overlay (top + height computed
  /// from rowOffsets) shrinks relative to the actual cell, and rows
  /// below a wrap-grown row drift.
  ///
  /// Only push back a measured value when it differs from the
  /// configured rowHeights — saves churn on the (vast majority of)
  /// rows that match exactly. Tolerance ±1 swallows browser
  /// sub-pixel rounding.
  $effect(() => {
    if (!gridWrapEl) return;
    // Re-trigger when geometry-affecting state changes. colWidths
    // matters because wrap-text content reflows when its column's
    // width changes — e.g. narrowing a wrap-text column makes the
    // text wrap to more lines and the row grows. viewportW catches
    // window resizes for the same reason.
    rowHeights;
    cells;
    bandStart;
    bandEnd;
    colWidths;
    viewportW;
    requestAnimationFrame(() => {
      const wrap = gridWrapEl;
      if (!wrap) return;
      const trs = wrap.querySelectorAll<HTMLTableRowElement>(
        "tbody tr:not(.virt-spacer)",
      );
      let next: Map<number, number> | null = null;
      for (const tr of trs) {
        const dataR = tr.querySelector<HTMLElement>("[data-r]");
        if (!dataR) continue;
        const r = parseInt(dataR.dataset.r ?? "0", 10);
        if (!r) continue;
        const actual = tr.offsetHeight;
        const configured = rowHeights.get(r) ?? DEFAULT_ROW_H;
        if (actual <= 0) continue;
        if (Math.abs(actual - configured) > 1) {
          if (!next) next = new Map(measuredRowHeights);
          if (next.get(r) !== actual) next.set(r, actual);
        } else if (measuredRowHeights.has(r)) {
          if (!next) next = new Map(measuredRowHeights);
          next.delete(r);
        }
      }
      if (next) measuredRowHeights = next;
    });
  });

  // Visible row band — only these rows + frozen + buffer end up in the
  // DOM. Cuts Svelte's per-arrow {@const} evaluations from
  // (rows × cols) down to roughly (visible × cols). A bigger buffer
  // means fewer band shifts while arrowing, so amortised arrow-cost
  // stays low. Depends ONLY on scrollTop + viewportH — not on selRow.
  // scrollSelIntoView keeps the cursor in the visible area, so the band
  // naturally covers it; skipping the selRow dep means vertical arrows
  // within the band don't touch the band at all.
  const ROW_BUFFER = 30;
  let visibleBand = $derived.by(() => {
    if (rows === 0) return { start: 1, end: 0 };
    const top = scrollTop;
    const bottom = scrollTop + viewportH;
    let start = 1;
    while (start <= rows && rowOffsets[start + 1] <= top) start++;
    let end = start;
    while (end <= rows && rowOffsets[end] < bottom) end++;
    start = Math.max(1, start - ROW_BUFFER);
    end = Math.min(rows, end + ROW_BUFFER);
    return { start, end };
  });

  // First row of the visible-band block (always > frozenRows so the
  // frozen rows aren't rendered twice). Spacers above/below absorb the
  // skipped rows so the scrollbar accurately represents totalRowH.
  let bandStart = $derived(Math.max(visibleBand.start, frozenRows + 1));
  let bandEnd = $derived(visibleBand.end);
  let topSpacerH = $derived(
    bandStart > frozenRows + 1
      ? rowOffsets[bandStart] - rowOffsets[frozenRows + 1]
      : 0,
  );
  let bottomSpacerH = $derived(
    bandEnd < rows ? rowOffsets[rows + 1] - rowOffsets[bandEnd + 1] : 0,
  );

  // Column virtualisation — same approach as rows. Skipped cells get
  // collapsed into colspan-based spacer <td>s on the left and right of
  // the visible band. Frozen cols always render. Merge anchors that
  // straddle the band's left edge force the band to expand leftward
  // so their content stays visible (otherwise the anchor cell vanishes
  // when its starting col scrolls off-screen).
  const COL_BUFFER = 30;
  let visibleColBand = $derived.by(() => {
    if (cols === 0) return { start: 1, end: 0 };
    const left = scrollLeft;
    const right = scrollLeft + viewportW;
    let start = 1;
    while (start <= cols && colLefts[start + 1] <= left) start++;
    let end = start;
    while (end <= cols && colLefts[end] < right) end++;
    start = Math.max(1, start - COL_BUFFER);
    end = Math.min(cols, end + COL_BUFFER);
    return { start, end };
  });
  let colBandStart = $derived(Math.max(visibleColBand.start, frozenCols + 1));
  let colBandEnd = $derived(visibleColBand.end);
  // Pull the band's left edge in to cover any merge anchor whose range
  // crosses the boundary. Without this, merges visually disappear when
  // their anchor scrolls left of the band.
  let effectiveColBandStart = $derived.by(() => {
    let s = colBandStart;
    for (const [key, anchor] of mergeMap.anchors) {
      const colon = key.indexOf(":");
      if (colon < 0) continue;
      const c = parseInt(key.slice(colon + 1), 10);
      if (c >= s) continue;
      if (c + anchor.colspan - 1 >= s) s = Math.min(s, c);
    }
    return Math.max(frozenCols + 1, s);
  });
  let leftSpacerW = $derived(
    effectiveColBandStart > frozenCols + 1
      ? colLefts[effectiveColBandStart] - colLefts[frozenCols + 1]
      : 0,
  );
  let rightSpacerW = $derived(
    colBandEnd < cols ? colLefts[cols + 1] - colLefts[colBandEnd + 1] : 0,
  );
  let leftSpacerSpan = $derived(
    Math.max(0, effectiveColBandStart - frozenCols - 1),
  );
  let rightSpacerSpan = $derived(Math.max(0, cols - colBandEnd));

  // Emit band edges to the page so it can lazily fetch rows entering view.
  // Page is responsible for rAF-coalescing — this fires every reactive
  // cycle that changes the band, which is cheap to ignore but expensive
  // to actually fetch on every emit.
  $effect(() => {
    onBandChange?.(bandStart, bandEnd);
  });
  $effect(() => {
    onColBandChange?.(effectiveColBandStart, colBandEnd);
  });

  // Calculated scroll-into-view — no layout-forcing DOM queries.
  // Reads cumulative geometry (rowOffsets/colLefts) and the wrap's
  // scroll position directly (.scrollTop / .scrollLeft are cached
  // properties, cheap). Deliberately DOESN'T read the scrollTop /
  // scrollLeft $state mirrors — otherwise the scroll event fired by
  // scrollBy would re-run this effect for no reason. When a scroll
  // IS needed we write the $state mirrors synchronously too so the
  // visibleBand derivation sees the new position in the same reactive
  // cycle as the selRow change (the later scroll event becomes a
  // no-op).
  /// Scroll a (row, col) into view, mutating `scrollTop`/`scrollLeft`
  /// and the wrap's scroll position. Caller is responsible for
  /// passing valid 1-based indices; out-of-range indices fall back to
  /// 0 via `?? 0` so a stale call doesn't throw mid-render.
  ///
  /// `flushSync()` before `scrollBy` is load-bearing — without it the
  /// browser scrolls before Svelte's pending reactive updates land,
  /// paints empty rows for one frame, then fills them in. The blank
  /// flash is visible at any scroll speed.
  function scrollCellIntoView(row: number, col: number) {
    if (!gridWrapEl) return;
    const wrap = gridWrapEl;
    const cellTop = (rowOffsets[row] ?? 0) + colhdrH;
    const cellBottom = (rowOffsets[row + 1] ?? 0) + colhdrH;
    const cellLeft = colLefts[col] ?? 0;
    const cellRight = colLefts[col + 1] ?? 0;
    const sTop = wrap.scrollTop;
    const sLeft = wrap.scrollLeft;
    const visTop = sTop + colhdrH;
    const visBottom = sTop + viewportH;
    const visLeft = sLeft + ROWHDR_W;
    const visRight = sLeft + viewportW;
    let dy = 0;
    if (cellTop < visTop) dy = cellTop - visTop;
    else if (cellBottom > visBottom) dy = cellBottom - visBottom;
    let dx = 0;
    if (cellLeft < visLeft) dx = cellLeft - visLeft;
    else if (cellRight > visRight) dx = cellRight - visRight;
    if (dx !== 0 || dy !== 0) {
      scrollTop = sTop + dy;
      scrollLeft = sLeft + dx;
      flushSync();
      wrap.scrollBy({ left: dx, top: dy, behavior: "instant" as ScrollBehavior });
    }
  }

  function scrollSelIntoView() {
    scrollCellIntoView(selRow, selCol);
  }

  // Only fire scrollSelIntoView when selRow/selCol *change* — not when
  // the reactive reads INSIDE scrollSelIntoView (rowOffsets, colhdrH,
  // viewportH, ...) invalidate. Without `untrack` the parent's
  // band-driven rowHeights map mutation during a mouse-wheel scroll
  // re-derives rowOffsets, which re-fires this effect, which scrolls
  // the cursor back into view — visible "snap back" while the user is
  // mid-wheel.
  $effect(() => {
    selRow;
    selCol;
    untrack(() => {
      if (!gridWrapEl) return;
      scrollSelIntoView();
    });
  });


  $effect(() => {
    if (!scrollTarget) return;
    const { row, col } = scrollTarget;
    untrack(() => {
      if (!gridWrapEl) return;
      scrollCellIntoView(row, col);
    });
  });

  // Overlay-layer geometry. Computed once per reactive update and read
  // by the absolute-positioned overlay divs in the markup. Cells stay
  // independent of selRow/selCol so arrow keys don't trigger any
  // per-cell re-evaluation.
  let isMultiCell = $derived(selRow !== rangeEndRow || selCol !== rangeEndCol);
  // Selection cursor overlay. `colhdrH` is the actual offsetTop of
  // the first tbody row (measured below) so the cursor sits flush
  // with the cell's top edge regardless of how border-collapse
  // resolves the thead/tbody border boundary in any given browser.
  // `rowOffsets` cumulative sums use measured row heights when
  // available so wrap cells (which grow the row beyond its
  // configured height) still match the actual cell box.
  let selBox = $derived({
    top: colhdrH + (rowOffsets[selRow] ?? 0),
    left: colLefts[selCol] ?? 0,
    width: (colLefts[selCol + 1] ?? 0) - (colLefts[selCol] ?? 0),
    height: (rowOffsets[selRow + 1] ?? 0) - (rowOffsets[selRow] ?? 0),
  });
  let rangeBox = $derived({
    top: colhdrH + (rowOffsets[rangeBounds.r1] ?? 0),
    left: colLefts[rangeBounds.c1] ?? 0,
    width: (colLefts[rangeBounds.c2 + 1] ?? 0) - (colLefts[rangeBounds.c1] ?? 0),
    height: (rowOffsets[rangeBounds.r2 + 1] ?? 0) - (rowOffsets[rangeBounds.r1] ?? 0),
  });
  let ghostBox = $derived(
    ghostRange
      ? {
          top: colhdrH + (rowOffsets[ghostRange.r1] ?? 0),
          left: colLefts[ghostRange.c1] ?? 0,
          width: (colLefts[ghostRange.c2 + 1] ?? 0) - (colLefts[ghostRange.c1] ?? 0),
          height: (rowOffsets[ghostRange.r2 + 1] ?? 0) - (rowOffsets[ghostRange.r1] ?? 0),
        }
      : null,
  );
  let highlightBoxes = $derived.by(() =>
    (highlights ?? []).map((h) => ({
      top: colhdrH + (rowOffsets[h.r1] ?? 0),
      left: colLefts[h.c1] ?? 0,
      width: (colLefts[h.c2 + 1] ?? 0) - (colLefts[h.c1] ?? 0),
      height: (rowOffsets[h.r2 + 1] ?? 0) - (rowOffsets[h.r1] ?? 0),
      color: h.color,
      label: h.label ?? null,
    })),
  );
  let fillBox = $derived.by(() => {
    const r = freeCorner === "tl" || freeCorner === "tr" ? rangeBounds.r1 : rangeBounds.r2;
    const c = freeCorner === "tl" || freeCorner === "bl" ? rangeBounds.c1 : rangeBounds.c2;
    const cellTop = colhdrH + (rowOffsets[r] ?? 0);
    const cellLeft = colLefts[c] ?? 0;
    const cellW = (colLefts[c + 1] ?? 0) - (colLefts[c] ?? 0);
    const cellH = (rowOffsets[r + 1] ?? 0) - (rowOffsets[r] ?? 0);
    const top = freeCorner === "tl" || freeCorner === "tr" ? cellTop - 3 : cellTop + cellH - 4;
    const left = freeCorner === "tl" || freeCorner === "bl" ? cellLeft - 3 : cellLeft + cellW - 4;
    return { top, left };
  });
</script>

{#snippet cellTpl(rowIdx: number, colIdx: number, isFrozenRow: boolean, rowTop: number | null)}
  {@const cellKey = `${rowIdx}:${colIdx}`}
  {#if !mergeMap.skip.has(cellKey)}
    {@const cell = cells.get(cellKey)}
    {@const isFrozenCol = colIdx <= frozenCols}
    {@const colLeft = isFrozenCol ? frozenColLefts[colIdx] : null}
    {@const merge = mergeMap.anchors.get(cellKey)}
    {@const right = cells.get(`${rowIdx}:${colIdx + 1}`)}
    {@const clipRight = !!right?.text}
    <td
      class="cell"
      class:frozen-row={isFrozenRow}
      class:frozen-col={isFrozenCol}
      class:frozen-corner={isFrozenRow && isFrozenCol}
      class:page-break-row={pageBreakRows.has(rowIdx)}
      class:page-break-col={pageBreakCols.has(colIdx)}
      class:merged={merge != null}
      colspan={merge?.colspan ?? null}
      rowspan={merge?.rowspan ?? null}
      data-r={rowIdx}
      data-c={colIdx}
      style={(isFrozenRow ? `top:${rowTop}px;` : "") +
        (isFrozenCol ? `left:${colLeft}px;` : "") +
        cellTdStyle(cell)}
    >
      {#if cell?.text || cell?.style?.bg}
        <div
          class="cell-content"
          class:wrap={cell.style?.wrap}
          class:clip-right={clipRight}
          style={cellContentStyle(cell)}
        >
          {cell?.text ?? ""}
        </div>
      {/if}
    </td>
  {/if}
{/snippet}

{#snippet rowTpl(rowIdx: number)}
  {@const rh = rowHeightFor(rowIdx)}
  {@const hiddenRow = rh === 0}
  {@const isFrozenRow = rowIdx <= frozenRows}
  {@const rowTop = isFrozenRow ? frozenRowTops[rowIdx] : null}
  <tr
    class:hidden-row={hiddenRow}
    style={rh != null && rh > 0 ? `height:${rh}px` : ""}
  >
    <th
      class="rowhdr"
      class:frozen-row={isFrozenRow}
      class:active-hdr={selRow === rowIdx}
      style={isFrozenRow ? `top:${rowTop}px;` : ""}
    >
      {rowIdx}
      <span
        class="row-resize"
        onmousedown={(e) => startRowResize(rowIdx, e)}
        ondblclick={(e) => { e.preventDefault(); e.stopPropagation(); autoFitRow(rowIdx); }}
      ></span>
    </th>
    <!-- Frozen cols always render; sticky positioning keeps them onscreen. -->
    {#each Array(frozenCols) as _, fi (fi)}
      {@render cellTpl(rowIdx, fi + 1, isFrozenRow, rowTop)}
    {/each}
    <!-- Left spacer absorbs cols skipped by column virtualisation. -->
    {#if leftSpacerSpan > 0}
      <td class="virt-spacer-cell" colspan={leftSpacerSpan} aria-hidden="true"></td>
    {/if}
    <!-- Visible col band — the actual rendered cells. -->
    {#each Array(Math.max(0, colBandEnd - effectiveColBandStart + 1)) as _, bi (effectiveColBandStart + bi)}
      {@render cellTpl(rowIdx, effectiveColBandStart + bi, isFrozenRow, rowTop)}
    {/each}
    <!-- Right spacer for cols off the right edge of the viewport. -->
    {#if rightSpacerSpan > 0}
      <td class="virt-spacer-cell" colspan={rightSpacerSpan} aria-hidden="true"></td>
    {/if}
  </tr>
{/snippet}

<div class:grid-no-lines={!showGridLines} class="grid-wrap" tabindex="0" bind:this={gridWrapEl}>
  <div
    class="grid-inner"
    style={`width:${tableWidthPx}px;`}
    onmousedown={onTableMouseDown}
    onmouseover={onTableMouseOver}
    ondblclick={onTableDblClick}
    oncontextmenu={onTableContextMenu}
  >
  <table class="grid" style={`width:${tableWidthPx}px`}>
    <colgroup>
      <!-- Explicit width on the row-header col so table-layout:fixed
           has widths for every column (otherwise the browser falls
           back to auto-distribution and the <col> widths get ignored). -->
      <col class="rowhdr-col" style="width:42px" />
      {#each colWidthList as w}
        <col class:hidden-col={w === 0} style={`width:${w}px`} />
      {/each}
    </colgroup>
    <thead>
      <tr>
        <th class="rowhdr"></th>
        <!-- Frozen col headers — sticky like their data cells. -->
        {#each Array(frozenCols) as _, fi (fi)}
          {@const i = fi + 1}
          <th
            class="colhdr frozen-col"
            class:active-hdr={selCol === i}
            style={`left:${frozenColLefts[i]}px`}
          >
            {colLetter(i)}
            <span
              class="col-resize"
              onmousedown={(e) => startColResize(i, e)}
              ondblclick={(e) => { e.preventDefault(); e.stopPropagation(); autoFitColumn(i); }}
            ></span>
          </th>
        {/each}
        <!-- Header spacer: matches the data-row spacer so column geometry
             stays aligned across <thead> and <tbody>. -->
        {#if leftSpacerSpan > 0}
          <th class="virt-spacer-cell" colspan={leftSpacerSpan} aria-hidden="true"></th>
        {/if}
        {#each Array(Math.max(0, colBandEnd - effectiveColBandStart + 1)) as _, bi (effectiveColBandStart + bi)}
          {@const i = effectiveColBandStart + bi}
          <th
            class="colhdr"
            class:active-hdr={selCol === i}
          >
            {colLetter(i)}
            <span
              class="col-resize"
              onmousedown={(e) => startColResize(i, e)}
              ondblclick={(e) => { e.preventDefault(); e.stopPropagation(); autoFitColumn(i); }}
            ></span>
          </th>
        {/each}
        {#if rightSpacerSpan > 0}
          <th class="virt-spacer-cell" colspan={rightSpacerSpan} aria-hidden="true"></th>
        {/if}
      </tr>
    </thead>
    <tbody>
      <!-- Frozen rows always render so sticky positioning has anchors. -->
      {#each Array(frozenRows) as _, r}
        {@render rowTpl(r + 1)}
      {/each}
      <!-- Spacer for rows skipped between frozen and visible band. -->
      {#if topSpacerH > 0}
        <tr aria-hidden="true" class="virt-spacer">
          <td colspan={cols + 1} style={`height:${topSpacerH}px`}></td>
        </tr>
      {/if}
      <!-- Visible band — virtualisation's payoff: only ~50 rows in DOM. -->
      {#each Array(Math.max(0, bandEnd - bandStart + 1)) as _, i (bandStart + i)}
        {@render rowTpl(bandStart + i)}
      {/each}
      <!-- Spacer for rows below visible band. Sized so the table's total
           height equals the sum of all row heights → scrollbar stays
           accurate. -->
      {#if bottomSpacerH > 0}
        <tr aria-hidden="true" class="virt-spacer">
          <td colspan={cols + 1} style={`height:${bottomSpacerH}px`}></td>
        </tr>
      {/if}
    </tbody>
  </table>
  <!-- Selection overlay layer. These divs sit above the table cells and
       update independently — no per-cell class:directives means arrow
       keys don't trigger any cell re-evaluation, only style mutations on
       a handful of overlay elements. -->
  {#if isMultiCell}
    <div
      class="sel-range-tint"
      style={`top:${rangeBox.top}px; left:${rangeBox.left}px; width:${rangeBox.width}px; height:${rangeBox.height}px;`}
    ></div>
  {/if}
  {#if ghostBox}
    <div
      class="sel-ghost"
      style={`top:${ghostBox.top}px; left:${ghostBox.left}px; width:${ghostBox.width}px; height:${ghostBox.height}px;`}
    ></div>
  {/if}
  {#each highlightBoxes as h}
    <div
      class="ref-highlight"
      style={`top:${h.top}px; left:${h.left}px; width:${h.width}px; height:${h.height}px; border-color: ${h.color}; box-shadow: 0 0 0 1px ${h.color} inset;`}
    ></div>
    {#if h.label}
      <div
        class="ref-tag"
        style={`top:${Math.max(h.top - 14, colhdrH)}px; left:${h.left}px; background:${h.color};`}
      >{h.label}</div>
    {/if}
  {/each}
  <div
    class="sel-cell-outline"
    style={`top:${selBox.top}px; left:${selBox.left}px; width:${selBox.width}px; height:${selBox.height}px;`}
  ></div>
  <span
    class="fill-handle"
    style={`top:${fillBox.top}px; left:${fillBox.left}px;`}
    onmousedown={startFillDrag}
    title="Drag to fill — press . to move"
  ></span>
  </div>
</div>

<style>
  .grid-wrap {
    flex: 1;
    overflow: auto;
    background: #fff;
    min-height: 0;
    /* Hint the compositor to keep scroll-position on the GPU path. */
    will-change: scroll-position;
  }
  .grid {
    /* Fixed layout so colgroup widths are authoritative — content can
       overflow visually (Excel-style spill into empty neighbors) without
       resizing the column. */
    border-collapse: collapse;
    table-layout: fixed;
  }
  .rowhdr,
  .colhdr {
    background: #f3f3f3;
    color: #555;
    font-weight: 500;
    padding: 0.05rem 0.4rem;
    border: 1px solid #c0c0c0;
    text-align: center;
    user-select: none;
    font-size: 11px;
    position: relative;
  }
  /* Active header highlight — Excel/Lotus convention. The colhdr at the
     active column and the rowhdr at the active row get a darker tint and
     bold weight so the user can locate the cursor on a wide sheet. */
  .colhdr.active-hdr,
  .rowhdr.active-hdr {
    background: #e0e8f5;
    color: #1f6feb;
    font-weight: 700;
  }
  /* Drag handles. Anchored to the right edge of column headers and the
     bottom edge of row headers; ~5px wide hot-zone with the col-/row-
     resize cursor. mousedown bubbles to the parent <th>'s onclick — we
     stopPropagation in startColResize/startRowResize. */
  .col-resize {
    position: absolute;
    top: 0;
    bottom: 0;
    right: -3px;
    width: 6px;
    cursor: col-resize;
    z-index: 4;
  }
  .row-resize {
    position: absolute;
    left: 0;
    right: 0;
    bottom: -3px;
    height: 6px;
    cursor: row-resize;
    z-index: 4;
  }
  /* No min-width on headers — colgroup widths are authoritative under
     table-layout:fixed. min-width here would silently force cols wider. */
  .rowhdr {
    position: sticky;
    left: 0;
    z-index: 3;
  }
  .colhdr {
    position: sticky;
    top: 0;
    z-index: 2;
  }
  .cell {
    border: 1px solid #e0e0e0;
    padding: 0;
    /* No background here — empty cells must be transparent so a long
       string in a left neighbor can spill visually through them. */
    /* No `position: relative`, no `contain`. Both turn each cell into
       a self-contained painting unit that paints in DOM order, which
       breaks the CSS table painting model that normally splits "all
       cell backgrounds first (step 6), all cell content second (step
       7)" across siblings. Once that split is preserved, A1's
       overflowing text in `.cell-content` paints at step 7 on top of
       B1's bg painted at step 6 — which is exactly Excel's spill
       behaviour. The earlier reasons we'd reached for `position:
       relative` (`.cell-content { position: absolute; inset: 0 }`)
       and `contain: layout` (perf) both turned out to break this
       interaction silently. */
  }
  .grid-no-lines .cell {
    border-color: transparent;
  }
  /* Page-break markers. Using `box-shadow inset` instead of a thicker
     border because border-collapse: collapse adds the extra pixel to
     the row's actual height — `rowOffsets` (computed from the
     rowHeights map alone) wouldn't see it, and the selection overlay
     would drift down by 1px per page break in the rows above the
     cursor. Inset shadows render purely on top of the cell paint and
     don't affect layout. Dashed-look approximated via gradient. */
  .cell.page-break-row {
    box-shadow: inset 0 2px 0 #2563eb;
  }
  .cell.page-break-col {
    box-shadow: inset 2px 0 0 #2563eb;
  }
  .cell.page-break-row.page-break-col {
    box-shadow: inset 2px 0 0 #2563eb, inset 0 2px 0 #2563eb;
  }
  /* Spacer rows that absorb the rows skipped by row virtualisation.
     Their height keeps the table's total geometry equal to the sum of
     all row heights so the scrollbar maps 1:1 to row offsets. */
  tr.virt-spacer > td,
  .virt-spacer-cell {
    padding: 0;
    border: none;
    background: transparent;
  }
  /* Frozen rows / cols. Sticky-position the cells (TR can't be sticky)
     so they stay visible while the body scrolls under them. The top/left
     offsets are set inline per-cell from the cumulative-size derived
     arrays. z-index: rowhdr (3) < frozen-col (4) < frozen-row (5) <
     frozen-corner (6) < colhdr (existing 2 — bumped via class) so
     overlapping panes layer correctly. */
  td.frozen-row,
  th.frozen-row {
    position: sticky;
    z-index: 5;
    background: #fff;
  }
  td.frozen-col {
    position: sticky;
    z-index: 4;
    background: #fff;
  }
  td.frozen-corner {
    z-index: 6;
  }
  th.rowhdr.frozen-row {
    z-index: 6;
    background: #f3f3f3;
  }
  /* Hidden rows fully collapse — adjacent rows flush. Visibility is
     surfaced through /Worksheet/Row/Display in the menu. */
  tr.hidden-row {
    display: none;
  }
  /* Hidden columns. visibility:collapse is the spec'd way for table-column
     elements. Surfaced via /Worksheet/Column/Display. */
  col.hidden-col {
    visibility: collapse;
  }
  .cell-content {
    /* Normal flow, NOT absolute. `position: absolute` requires
       `position: relative` on the parent `.cell`, and that forces
       table cells out of the table painting model that puts cell
       backgrounds before cell content — breaking spill into bg-
       styled neighbours. Filling the cell via height: 100% works
       because the parent <tr> sets a fixed pixel height, so the
       <td>'s height resolves and the percentage child can compute. */
    box-sizing: border-box;
    height: 100%;
    width: 100%;
    padding: 1px 4px;
    /* Line-height 1 is required so the cell content fits in the
       configured row height. Browsers default to ~1.2 for sans-serif,
       which makes a Calibri-11 cell's natural min-height ~22px even
       though Excel renders the same text at 20px. With 1.2 line-height
       the row stretches past the height attribute we set on the <tr>,
       and rowOffsets (built from the rowHeights map) drifts by ~2-3px
       per row — accumulating into a visible cursor-vs-grid offset
       further down the sheet. line-height: 1 collapses that gap. */
    line-height: 1;
    white-space: nowrap;
    overflow: visible;
    /* Opaque background masks any overflow leaking in from the left
       neighbor when this cell has its own content. Inline `background:` set
       by cellStyle() overrides for filled cells. */
    background: #fff;
    /* Excel default vertical alignment is bottom. Use flex-column +
       justify-content:flex-end so text sits at the bottom of the cell and
       text-align still drives horizontal alignment. */
    display: flex;
    flex-direction: column;
    justify-content: flex-end;
  }
  /* Wrap-text cells: render wrapped content in natural flow so the
     row auto-fits to fit it (Excel-style row auto-grow). Top-align
     the visible lines so reading order is natural. The actual
     rendered row height is captured by the ResizeObserver below
     and pushed back into rowHeights, so rowOffsets stays in sync
     with the visible geometry and the cursor matches. */
  .cell-content.wrap {
    justify-content: flex-start;
  }
  /* Excel spill rule: text overflow stops at the first neighbour
     cell with its own text. `overflow: visible` is the default so
     spill flows through truly empty cells AND through bg-only
     styled cells (matches gop.xlsx row 1 "Panelled sail Discount/
     Lead time lookup" spilling through grey-tinted cells). When
     the immediate right neighbour has text, we clip on the LEFT
     cell to stop the spill before it paints over the neighbour's
     glyphs. Keyed on `right.text` only (NOT `right.style?.bg`) — a
     bg-only neighbour does not block spill. */
  .cell-content.clip-right {
    overflow: hidden;
  }
  /* Inner positioning context for the overlay layer. The table sits
     here too; both scroll together as the user drags grid-wrap. The
     overlays (sel-cell-outline, sel-range-tint, sel-ghost, fill-handle)
     are absolute-positioned children of grid-inner and update only
     their inline `top/left/width/height` as selRow/selCol change.
     This avoids any per-cell class:directive churn on arrow keys. */
  .grid-inner {
    position: relative;
  }
  .sel-cell-outline {
    position: absolute;
    box-sizing: border-box;
    border: 2px solid #1f6feb;
    pointer-events: none;
    z-index: 10;
  }
  .sel-range-tint {
    position: absolute;
    background: rgba(31, 111, 235, 0.16);
    pointer-events: none;
    z-index: 9;
  }
  .sel-ghost {
    position: absolute;
    background: rgba(240, 196, 25, 0.22);
    outline: 1px dashed #b88a00;
    outline-offset: -1px;
    pointer-events: none;
    z-index: 9;
  }
  /* Reference highlight — outline-only box used by the formula trace
     popup and F2 edit mode to show what cells/ranges a formula points
     at. Border color is set per-instance so multiple refs in one
     formula can be color-coded. */
  .ref-highlight {
    position: absolute;
    border: 2px solid;
    border-radius: 2px;
    pointer-events: none;
    z-index: 8;
  }
  /* Small badge anchored to the top-left of a highlight, used to
     show the name of a defined-range reference. Background color
     matches the highlight border so the visual association is
     immediate. */
  .ref-tag {
    position: absolute;
    height: 14px;
    padding: 0 4px;
    font-size: 10px;
    font-weight: bold;
    color: #fff;
    line-height: 14px;
    white-space: nowrap;
    pointer-events: none;
    z-index: 9;
    border-radius: 2px 2px 0 0;
  }
  /* Excel-style fill handle — moves to whichever corner of the
     selection rectangle is the "free" one (cycled by `.` in the
     parent). Position is computed in the overlay layer above. */
  .fill-handle {
    position: absolute;
    width: 7px;
    height: 7px;
    background: #1f6feb;
    border: 1px solid #fff;
    cursor: crosshair;
    z-index: 11;
  }
</style>
