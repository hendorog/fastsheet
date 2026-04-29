<script lang="ts">
  /// Docked panel listing the cell-level diffs between the active
  /// workbook and a side-loaded comparison file. Mirrors the layout
  /// pattern of FormulaTrace.svelte so the user lands on a familiar
  /// dock on the right side of the screen.
  ///
  /// Keyboard:
  ///   Esc      close compare mode entirely
  ///   ↑ / ↓    move highlight between diffs
  ///   Enter    jump cursor to the highlighted diff
  ///   H        hide / show this dock (compare mode stays active)
  ///   D        toggle dock width (currently fixed; reserved)

  import { onMount, tick } from "svelte";
  import type { CompareDiff, CompareSheetMissing } from "./types";

  type Props = {
    diffs: CompareDiff[];
    missingSheets: CompareSheetMissing[];
    rightPath: string;
    totalDiffs: number;
    capped: boolean;
    onJump: (d: CompareDiff) => void | Promise<void>;
    onPreview?: (d: CompareDiff) => void;
    onClose: () => void;
    hidden?: boolean;
  };
  let {
    diffs,
    missingSheets,
    rightPath,
    totalDiffs,
    capped,
    onJump,
    onPreview,
    onClose,
    hidden = $bindable(false),
  }: Props = $props();

  let highlight = $state(0);

  /// Group diffs by sheet for display, but preserve a flat row index
  /// so keyboard navigation and onPreview wiring stays simple.
  type FlatRow =
    | { kind: "header"; sheet: string }
    | { kind: "diff"; diff: CompareDiff }
    | { kind: "missing"; missing: CompareSheetMissing };

  let rows = $derived.by<FlatRow[]>(() => {
    const out: FlatRow[] = [];
    for (const m of missingSheets) {
      out.push({ kind: "missing", missing: m });
    }
    let lastSheet: string | null = null;
    for (const d of diffs) {
      if (d.sheet !== lastSheet) {
        out.push({ kind: "header", sheet: d.sheet });
        lastSheet = d.sheet;
      }
      out.push({ kind: "diff", diff: d });
    }
    return out;
  });

  /// Keep highlight pointing at a "diff" row — skip headers when
  /// moving with arrow keys so Enter always means something.
  function nextDiffRow(from: number, dir: 1 | -1): number {
    let i = from + dir;
    while (i >= 0 && i < rows.length && rows[i].kind !== "diff") i += dir;
    if (i < 0 || i >= rows.length) return from;
    return i;
  }

  function activate(idx: number) {
    const r = rows[idx];
    if (!r || r.kind !== "diff") return;
    onJump(r.diff);
  }

  function onKey(e: KeyboardEvent) {
    if (e.key === "Escape") {
      e.preventDefault();
      e.stopPropagation();
      onClose();
      return;
    }
    if (e.key === "h" || e.key === "H") {
      e.preventDefault();
      e.stopPropagation();
      hidden = !hidden;
      return;
    }
    if (hidden) return;
    if (e.key === "ArrowDown") {
      e.preventDefault();
      e.stopPropagation();
      highlight = nextDiffRow(highlight, 1);
      scrollHighlightIntoView();
      return;
    }
    if (e.key === "ArrowUp") {
      e.preventDefault();
      e.stopPropagation();
      highlight = nextDiffRow(highlight, -1);
      scrollHighlightIntoView();
      return;
    }
    if (e.key === "Enter") {
      e.preventDefault();
      e.stopPropagation();
      activate(highlight);
      return;
    }
  }

  let listEl: HTMLDivElement | undefined = $state(undefined);
  function scrollHighlightIntoView() {
    tick().then(() => {
      const item = listEl?.querySelector<HTMLElement>(
        `[data-idx="${highlight}"]`,
      );
      item?.scrollIntoView({ block: "nearest", behavior: "instant" });
    });
  }

  let mounted = $state(false);
  $effect(() => {
    highlight;
    if (!mounted) return;
    const r = rows[highlight];
    if (r && r.kind === "diff" && onPreview) onPreview(r.diff);
  });

  onMount(() => {
    // Land highlight on the first diff (skip any leading missing-sheet
    // headers) so Enter immediately jumps to a real cell.
    for (let i = 0; i < rows.length; i++) {
      if (rows[i].kind === "diff") {
        highlight = i;
        break;
      }
    }
    window.addEventListener("keydown", onKey, true);
    mounted = true;
    return () => window.removeEventListener("keydown", onKey, true);
  });

  function basename(p: string): string {
    const i = Math.max(p.lastIndexOf("/"), p.lastIndexOf("\\"));
    return i < 0 ? p : p.slice(i + 1);
  }

  function kindLabel(k: CompareDiff["kind"]): string {
    switch (k) {
      case "value":
        return "≠";
      case "formula":
        return "ƒ";
      case "missing-left":
        return "+R";
      case "missing-right":
        return "+L";
    }
  }
</script>

{#if hidden}
  <div class="collapsed-bar" role="status">
    <span class="collapsed-icon">↹</span>
    <span class="collapsed-text"
      >Compare: {basename(rightPath)} · {totalDiffs} diff{totalDiffs === 1
        ? ""
        : "s"}{capped ? " (capped)" : ""}</span
    >
    <span class="collapsed-hint">H show · Esc exit</span>
  </div>
{:else}
  <div class="popup" role="dialog" aria-label="Compare diffs">
    <div class="header">
      <div class="title-block">
        <span class="title">Compare</span>
        <span class="path">vs {basename(rightPath)}</span>
      </div>
      <div class="meta">
        <span
          >{totalDiffs} diff{totalDiffs === 1 ? "" : "s"}{capped
            ? ` (showing ${diffs.length})`
            : ""}</span
        >
        <span class="hint">↑↓ Enter · H hide · Esc exit</span>
      </div>
    </div>
    {#if rows.length === 0}
      <div class="empty">
        No differences. Both workbooks are equal in formatted values
        and formula text.
      </div>
    {:else}
      <div class="list" bind:this={listEl}>
        {#each rows as row, i}
          {#if row.kind === "header"}
            <div class="sheet-header" data-idx={i}>{row.sheet}</div>
          {:else if row.kind === "missing"}
            <div class="missing-row" data-idx={i}>
              {row.missing.side === "left"
                ? `(only on right) ${row.missing.sheet}`
                : `(only on left)  ${row.missing.sheet}`}
            </div>
          {:else}
            {@const d = row.diff}
            <div
              class="row kind-{d.kind}"
              class:hl={i === highlight}
              data-idx={i}
              role="button"
              tabindex="-1"
              onclick={() => {
                highlight = i;
                activate(i);
              }}
              onmouseenter={() => {
                highlight = i;
              }}
            >
              <span class="badge">{kindLabel(d.kind)}</span>
              <span class="addr">{d.address}</span>
              <span class="left-value">{d.left_value || "(empty)"}</span>
              <span class="arrow">→</span>
              <span class="right-value">{d.right_value || "(empty)"}</span>
              {#if d.kind === "formula"}
                <span class="formula-hint"
                  >{d.left_formula} → {d.right_formula}</span
                >
              {/if}
            </div>
          {/if}
        {/each}
      </div>
    {/if}
  </div>
{/if}

<style>
  .popup {
    position: fixed;
    background: #1a1a1a;
    color: #ddd;
    border: 1px solid #444;
    display: flex;
    flex-direction: column;
    font-family: monospace;
    font-size: 13px;
    box-shadow: 0 8px 32px rgba(0, 0, 0, 0.6);
    z-index: 998;
    /* Docked: pinned to the right edge as a full-height panel. */
    top: 0;
    left: auto;
    right: 0;
    bottom: 0;
    width: min(560px, 50vw);
    max-height: 100vh;
    border-radius: 0;
    border-left: 1px solid #444;
    border-top: none;
    border-right: none;
    border-bottom: none;
  }
  .collapsed-bar {
    position: fixed;
    bottom: 0;
    right: 0;
    background: #2a2a2a;
    border-top: 1px solid #444;
    border-left: 1px solid #444;
    border-radius: 4px 0 0 0;
    padding: 4px 10px;
    font-family: monospace;
    font-size: 12px;
    color: #ddd;
    display: flex;
    align-items: center;
    gap: 8px;
    z-index: 999;
    pointer-events: none;
  }
  .collapsed-icon {
    color: #f99;
    font-weight: bold;
  }
  .collapsed-text {
    color: #f99;
    font-weight: bold;
  }
  .collapsed-hint {
    color: #888;
    font-size: 11px;
  }
  .header {
    padding: 8px 12px;
    border-bottom: 1px solid #333;
    display: flex;
    flex-direction: column;
    gap: 4px;
    background: #2a2a2a;
  }
  .title-block {
    display: flex;
    align-items: baseline;
    gap: 10px;
  }
  .title {
    font-weight: bold;
    color: #f99;
  }
  .path {
    color: #aaa;
  }
  .meta {
    display: flex;
    align-items: center;
    justify-content: space-between;
    color: #aaa;
    font-size: 11px;
  }
  .hint {
    color: #888;
  }
  .empty {
    padding: 24px 16px;
    color: #888;
    font-style: italic;
  }
  .list {
    overflow-y: auto;
    padding: 6px 0;
  }
  .sheet-header {
    padding: 6px 12px 2px 12px;
    color: #6cf;
    font-weight: bold;
    border-top: 1px solid #2a2a2a;
  }
  .missing-row {
    padding: 4px 12px;
    color: #c8a060;
    font-style: italic;
  }
  .row {
    display: flex;
    align-items: center;
    gap: 6px;
    padding: 2px 12px;
    cursor: pointer;
    white-space: nowrap;
    user-select: none;
  }
  .row.hl {
    background: #3a2a2a;
  }
  .badge {
    display: inline-block;
    width: 24px;
    text-align: center;
    color: #888;
    font-size: 11px;
    font-weight: bold;
  }
  .row.kind-value .badge {
    color: #f88;
  }
  .row.kind-formula .badge {
    color: #6cf;
  }
  .row.kind-missing-left .badge,
  .row.kind-missing-right .badge {
    color: #c8a060;
  }
  .addr {
    color: #ddd;
    font-weight: bold;
    width: 64px;
    flex-shrink: 0;
  }
  .left-value {
    color: #e8c068;
    overflow: hidden;
    text-overflow: ellipsis;
    max-width: 180px;
  }
  .arrow {
    color: #666;
  }
  .right-value {
    color: #6c6;
    overflow: hidden;
    text-overflow: ellipsis;
    max-width: 180px;
  }
  .formula-hint {
    color: #888;
    margin-left: 8px;
    overflow: hidden;
    text-overflow: ellipsis;
    max-width: 320px;
    font-style: italic;
  }
</style>
