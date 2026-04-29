<script lang="ts">
  /// Floating popup that renders a formula's dependency tree.
  ///
  /// Keyboard:
  ///   Esc            close
  ///   ↑ / ↓          move highlight between visible cell rows
  ///   Enter          jump to the highlighted cell (if it has a sheet)
  ///   Left / Right   collapse / expand the highlighted node
  ///   * (asterisk)   expand all
  ///   /              collapse all
  ///
  /// Click on any cell-kind node to jump there.

  import { onMount, tick } from "svelte";
  import type { TraceNode } from "./types";

  type Props = {
    root: TraceNode;
    onClose: () => void;
    /// Called when user activates (Enter / click) a row. The page
    /// handler decides what to do based on node.kind — cells and
    /// ranges jump to (sheet,row,col); names parse the resolved
    /// formula text and jump to that range's top-left.
    onJump: (node: TraceNode) => void | Promise<void>;
    /// Called every time the keyboard highlight changes to a different
    /// row. The page uses this to switch sheets / scroll the grid to
    /// the previewed cell + show a reference outline, without changing
    /// the active cursor.
    onPreview?: (node: TraceNode) => void;
    /// Two-way bindable layout flags. `docked` shifts the popup from
    /// centered modal to a right-side panel that doesn't dim the
    /// grid. `hidden` collapses the popup to a tiny status bar so
    /// the user can interact with the grid without losing trace
    /// state. Both default false; toggled by D and H respectively.
    docked?: boolean;
    hidden?: boolean;
  };
  let {
    root,
    onClose,
    onJump,
    onPreview,
    docked = $bindable(false),
    hidden = $bindable(false),
  }: Props = $props();

  /// Flat list of (node, depth, parent-collapse-key) — derived from
  /// the tree + the collapsed set so keyboard navigation is O(N) not
  /// recursive. Rebuilt whenever `collapsed` mutates.
  type FlatRow = { node: TraceNode; depth: number; key: string };
  let collapsed = $state(new Set<string>());
  let highlight = $state(0);

  function nodeKey(node: TraceNode, path: string): string {
    return `${path}/${node.address}|${node.row ?? "_"}|${node.col ?? "_"}`;
  }

  let rows = $derived.by<FlatRow[]>(() => {
    const out: FlatRow[] = [];
    const walk = (n: TraceNode, depth: number, path: string) => {
      const key = nodeKey(n, path);
      out.push({ node: n, depth, key });
      if (collapsed.has(key)) return;
      for (const d of n.deps) walk(d, depth + 1, key);
    };
    walk(root, 0, "");
    return out;
  });

  function toggle(row: FlatRow) {
    if (row.node.deps.length === 0) return;
    const next = new Set(collapsed);
    if (next.has(row.key)) next.delete(row.key);
    else next.add(row.key);
    collapsed = next;
  }

  function expandAll() {
    collapsed = new Set();
  }

  function collapseAll() {
    // Collapse everything except the root.
    const next = new Set<string>();
    const walk = (n: TraceNode, path: string) => {
      const key = nodeKey(n, path);
      if (n !== root && n.deps.length > 0) next.add(key);
      for (const d of n.deps) walk(d, key);
    };
    walk(root, "");
    collapsed = next;
  }

  async function jumpRow(row: FlatRow) {
    const n = row.node;
    // Skip jump for literals + cycles (no useful target). Everything
    // else — cells, ranges, named ranges — is handled by the parent.
    if (n.kind === "literal" && n.sheet === null) return;
    if (n.cycle) return;
    await onJump(n);
    onClose();
  }

  function onKey(e: KeyboardEvent) {
    // Esc always closes, regardless of mode — escape hatch from any
    // state. Stop propagation so the grid's onKey doesn't see it.
    if (e.key === "Escape") {
      e.preventDefault();
      e.stopPropagation();
      onClose();
      return;
    }
    // H toggles hidden — works in all modes. When hidden the popup
    // becomes a tiny status bar and gives up the keyboard so the
    // grid runs normally; H brings it back.
    if (e.key === "h" || e.key === "H") {
      e.preventDefault();
      e.stopPropagation();
      hidden = !hidden;
      return;
    }
    // D toggles docked. If hidden, also un-hide (D = "show me docked").
    if (e.key === "d" || e.key === "D") {
      e.preventDefault();
      e.stopPropagation();
      docked = !docked;
      hidden = false;
      return;
    }
    // While hidden the popup gives up all OTHER keys to the grid.
    // Only Esc / H / D above apply.
    if (hidden) return;
    if (e.key === "ArrowDown") {
      e.preventDefault();
      e.stopPropagation();
      highlight = Math.min(highlight + 1, rows.length - 1);
      scrollHighlightIntoView();
      return;
    }
    if (e.key === "ArrowUp") {
      e.preventDefault();
      e.stopPropagation();
      highlight = Math.max(highlight - 1, 0);
      scrollHighlightIntoView();
      return;
    }
    if (e.key === "ArrowRight") {
      e.preventDefault();
      e.stopPropagation();
      const r = rows[highlight];
      if (r && collapsed.has(r.key)) toggle(r);
      return;
    }
    if (e.key === "ArrowLeft") {
      e.preventDefault();
      e.stopPropagation();
      const r = rows[highlight];
      if (!r) return;
      if (r.node.deps.length > 0 && !collapsed.has(r.key)) {
        toggle(r);
      } else {
        // Jump up to parent (find the row whose depth is one less).
        for (let i = highlight - 1; i >= 0; i--) {
          if (rows[i].depth < r.depth) {
            highlight = i;
            scrollHighlightIntoView();
            break;
          }
        }
      }
      return;
    }
    if (e.key === "Enter") {
      e.preventDefault();
      e.stopPropagation();
      const r = rows[highlight];
      if (r) jumpRow(r);
      return;
    }
    if (e.key === "*") {
      e.preventDefault();
      e.stopPropagation();
      expandAll();
      return;
    }
    if (e.key === "/") {
      e.preventDefault();
      e.stopPropagation();
      collapseAll();
      return;
    }
  }

  let listEl: HTMLDivElement | undefined = $state(undefined);
  function scrollHighlightIntoView() {
    tick().then(() => {
      const item = listEl?.querySelector<HTMLElement>(`[data-idx="${highlight}"]`);
      item?.scrollIntoView({ block: "nearest", behavior: "instant" });
    });
  }

  /// Fire onPreview whenever the keyboard highlight settles on a new
  /// row. Skips the initial mount (no movement happened yet — the user
  /// shouldn't be teleported to wherever they last were on each open).
  let mounted = $state(false);
  $effect(() => {
    highlight;
    if (!mounted) return;
    const row = rows[highlight];
    if (row && onPreview) onPreview(row.node);
  });

  onMount(() => {
    // Capture-phase listener so the popup gets keys before the grid.
    window.addEventListener("keydown", onKey, true);
    mounted = true;
    return () => window.removeEventListener("keydown", onKey, true);
  });

  function iconFor(n: TraceNode): string {
    if (n.cycle) return "↺";
    if (n.truncated) return "…";
    if (n.is_error) return "⚠";
    if (n.kind === "name") return "ƒ";
    if (n.kind === "range") return "▦";
    return "•";
  }
</script>

{#if hidden}
  <!-- Collapsed: tiny status bar at the bottom-right. The popup
       gives up keyboard focus to the grid; only H / D / Esc still
       hit the popup's listener. -->
  <div class="collapsed-bar" role="status">
    <span class="collapsed-icon">↶</span>
    <span class="collapsed-text">Trace: {root.address}</span>
    <span class="collapsed-hint">H show · D dock · Esc close</span>
  </div>
{:else}
  {#if !docked}
    <!-- Modal: dark overlay, centered. Click outside doesn't close
         (matches Esc-only convention used elsewhere in the app). -->
    <div class="overlay" role="presentation"></div>
  {/if}
  <div class="popup" class:docked role="dialog" aria-label="Formula trace">
    <div class="header">
      <span class="title">Formula trace — {root.address}</span>
      <span class="hint">↑↓ Enter ←→ · * / · H hide · D {docked ? "undock" : "dock"} · Esc close</span>
    </div>
    <div class="list" bind:this={listEl}>
      {#each rows as row, i}
        <div
          class="row"
          class:hl={i === highlight}
          class:err={row.node.is_error}
          class:cycle={row.node.cycle}
          data-idx={i}
          role="button"
          tabindex="-1"
          onclick={() => { highlight = i; jumpRow(row); }}
          onmouseenter={() => { highlight = i; }}
        >
          <span class="indent" style={`width: ${row.depth * 14}px;`}></span>
          {#if row.node.deps.length > 0}
            <span class="caret"
              role="button"
              tabindex="-1"
              onclick={(e) => { e.stopPropagation(); toggle(row); }}
            >{collapsed.has(row.key) ? "▶" : "▼"}</span>
          {:else}
            <span class="caret-spacer"></span>
          {/if}
          <span class="icon kind-{row.node.kind}">{iconFor(row.node)}</span>
          <span class="addr">{row.node.address}</span>
          <span class="eq">=</span>
          <span class="value" class:diff={row.node.compare_differs}
            >{row.node.value || "(empty)"}</span
          >
          {#if row.node.compare_value !== null}
            <span class="cmp-sep">|</span>
            <span class="cmp-value" class:diff={row.node.compare_differs}
              >{row.node.compare_value || "(empty)"}</span
            >
          {/if}
          {#if row.node.formula}
            <span class="formula">{row.node.formula}</span>
          {/if}
          {#if row.node.note}
            <span class="note">{row.node.note}</span>
          {/if}
        </div>
      {/each}
    </div>
  </div>
{/if}

<style>
  .overlay {
    position: fixed;
    inset: 0;
    background: rgba(0, 0, 0, 0.55);
    z-index: 999;
  }
  .popup {
    position: fixed;
    background: #1a1a1a;
    color: #ddd;
    border: 1px solid #444;
    border-radius: 4px;
    display: flex;
    flex-direction: column;
    font-family: monospace;
    font-size: 13px;
    box-shadow: 0 8px 32px rgba(0, 0, 0, 0.6);
    z-index: 1000;
    /* Modal default: centered. */
    top: 50%;
    left: 50%;
    transform: translate(-50%, -50%);
    width: min(1100px, 95vw);
    max-height: 85vh;
  }
  /* Docked: pinned to the right edge as a full-height panel. No
     overlay backdrop — user can see the grid behind. */
  .popup.docked {
    top: 0;
    left: auto;
    right: 0;
    bottom: 0;
    transform: none;
    width: min(560px, 50vw);
    max-height: 100vh;
    border-radius: 0;
    border-left: 1px solid #444;
    border-top: none;
    border-right: none;
    border-bottom: none;
  }
  /* Collapsed: tiny status bar at bottom-right. The popup is hidden
     in this state; this bar is the only visible reminder that a
     trace is still active. */
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
    z-index: 1000;
    pointer-events: none;
  }
  .collapsed-icon {
    color: #e8c068;
    font-weight: bold;
  }
  .collapsed-text {
    color: #e8c068;
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
    align-items: center;
    justify-content: space-between;
    gap: 12px;
    background: #2a2a2a;
  }
  .title {
    font-weight: bold;
    color: #e8c068;
  }
  .hint {
    font-size: 11px;
    color: #888;
  }
  .list {
    overflow-y: auto;
    padding: 6px 0;
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
    background: #2d3a4d;
  }
  .row.err {
    color: #f55;
  }
  .row.cycle {
    color: #aaa;
    font-style: italic;
  }
  .indent {
    flex-shrink: 0;
  }
  .caret, .caret-spacer {
    display: inline-block;
    width: 12px;
    text-align: center;
    color: #888;
    font-size: 10px;
  }
  .icon {
    width: 14px;
    text-align: center;
    color: #888;
  }
  .icon.kind-name {
    color: #6cf;
  }
  .icon.kind-range {
    color: #c8a060;
  }
  .icon.kind-cell {
    color: #6c6;
  }
  .icon.kind-literal {
    color: #888;
  }
  .addr {
    color: #ddd;
    font-weight: bold;
  }
  .eq {
    color: #666;
  }
  .value {
    color: #e8c068;
    overflow: hidden;
    text-overflow: ellipsis;
    max-width: 320px;
  }
  .value.diff {
    color: #f88;
  }
  .cmp-sep {
    color: #666;
    margin: 0 4px;
  }
  .cmp-value {
    color: #888;
    overflow: hidden;
    text-overflow: ellipsis;
    max-width: 320px;
  }
  .cmp-value.diff {
    color: #6c6;
    font-weight: bold;
  }
  .formula {
    color: #6cf;
    margin-left: 8px;
    overflow: hidden;
    text-overflow: ellipsis;
    max-width: 480px;
  }
  .note {
    color: #888;
    margin-left: 8px;
    font-style: italic;
  }
</style>
