<script lang="ts">
  import { invoke } from "@tauri-apps/api/core";
  import { onMount, tick } from "svelte";
  import type { DirListing, RecentEntry, RecentDir, NavRow, DirEntry } from "./types";

  type Props = {
    mode: "open" | "save";
    /// What file kind the picker is for. "workbook" lists .xlsx / .xls
    /// and shows the recents-of-workbooks block. "text" lists
    /// .csv / .tsv / .txt and suppresses recents (workbook recents
    /// would be misleading and there's no text-file recents store).
    fileKind?: "workbook" | "text";
    currentPath: string;
    startDir?: string;
    onOpenFile: (path: string) => void | Promise<void>;
    onSaveFile: (path: string) => void | Promise<void>;
    onDirectoryChange?: (dir: string) => void;
    onClose: () => void;
    onStatus: (msg: string) => void;
  };

  let {
    mode,
    fileKind = "workbook",
    currentPath,
    startDir = "",
    onOpenFile,
    onSaveFile,
    onDirectoryChange,
    onClose,
    onStatus,
  }: Props = $props();

  let listing = $state<DirListing | null>(null);
  let filter = $state("");
  let selectedIdx = $state(0);
  let recents = $state<RecentEntry[]>([]);
  let recentDirs = $state<RecentDir[]>([]);
  let filterEl: HTMLInputElement | null = $state(null);
  /// Recents and recent dirs are shown on the initial open so users
  /// can re-pick a frequently-used file or jump to a known location
  /// without browsing. Once they navigate to a different directory
  /// these lists stop being useful — hide them so the entry list
  /// isn't padded with stale rows and tab-style filtering (e.g. `u`
  /// → `ubuntu`) resolves to a single match. Same flag gates both
  /// lists.
  let hasInteracted = $state(false);

  /// Recognise filter strings that mean "navigate to this location" rather
  /// than "filter the current list" — Enter on these jumps directly.
  function isNavToken(s: string): boolean {
    const t = s.trim();
    if (!t) return false;
    if (t === ".." || t === "~" || t === "/") return true;
    if (/^[a-zA-Z]:$/.test(t)) return true;
    if (t.startsWith("\\\\") || t.startsWith("//")) return true;
    if (t.startsWith("/") || t.startsWith("~/") || t.startsWith("~\\")) return true;
    if (/^[a-zA-Z]:[\\/]/.test(t)) return true;
    return false;
  }

  async function navTo(path: string, cwd: string | null) {
    try {
      const previous = listing?.dir ?? null;
      const result = await invoke<DirListing>("list_dir", { path, cwd, kind: fileKind });
      listing = result;
      onDirectoryChange?.(result.dir);
      filter = "";
      selectedIdx = 0;
      // Crossing into a new directory means the user has acted on
      // the listing, so further recents become noise. Typing in the
      // filter alone does NOT count — recents should remain
      // searchable so `77` can match a recent file containing 77.
      if (previous !== null && previous !== result.dir) hasInteracted = true;
      await tick();
      filterEl?.focus();
    } catch (e) {
      onStatus(`cd failed: ${e}`);
    }
  }

  async function refreshRecents() {
    // Recents are workbook-opens. For text pickers (/D Import) those
    // would be misleading entries that the user can't actually pick
    // (the backend filter rejects them), so suppress entirely.
    if (fileKind === "text") {
      recents = [];
      recentDirs = [];
      return;
    }
    try {
      const [files, dirs] = await Promise.all([
        invoke<RecentEntry[]>("query_recents", { query: filter, limit: 8 }),
        invoke<RecentDir[]>("query_recent_dirs", { query: filter, limit: 7 }),
      ]);
      recents = files;
      recentDirs = dirs;
    } catch {
      recents = [];
      recentDirs = [];
    }
  }

  function startDirFromCurrentPath(): string | null {
    if (!currentPath) return null;
    const sep = currentPath.includes("\\") ? "\\" : "/";
    const idx = currentPath.lastIndexOf(sep);
    return idx >= 0 ? currentPath.slice(0, idx) : ".";
  }

  onMount(async () => {
    const preferred = startDir || startDirFromCurrentPath();
    const start = preferred ?? (await invoke<string>("start_dir"));
    await Promise.all([navTo(start, null), refreshRecents()]);
  });

  let rows = $derived.by<NavRow[]>(() => {
    const out: NavRow[] = [];
    if (!hasInteracted) {
      // Recent files first (newest open first, sorted server-side).
      // Then recent directories — a fallback "I want to go to a place
      // I've been recently" list. Both hidden once the user crosses
      // into a new directory.
      for (const r of recents) out.push({ kind: "recent", recent: r });
      for (const d of recentDirs) out.push({ kind: "recent_dir", recent_dir: d });
    }
    if (!listing) return out;
    const q = filter.trim().toLowerCase();
    if (q) {
      for (const e of listing.entries) {
        if (e.name.toLowerCase().includes(q)) {
          out.push({ kind: "entry", entry: e });
        }
      }
    } else {
      if (listing.parent) {
        out.push({
          kind: "entry",
          entry: { name: "..", is_dir: true, modified: null, size: null },
        });
      }
      for (const e of listing.entries) out.push({ kind: "entry", entry: e });
    }
    return out;
  });

  // Filter changes: reset selection and refresh recents to match.
  $effect(() => {
    filter; // dep
    selectedIdx = 0;
    refreshRecents();
  });

  /// Typing `\\` in the filter auto-jumps to the WSL UNC root so the user
  /// gets a distro list to filter against. `\\ubuntu` lands you in that root
  /// with the filter set to `ubuntu`. A fully-typed UNC like
  /// `\\server\share` (which contains a third backslash) does NOT trigger —
  /// hit Enter to navigate the full path, same as before.
  $effect(() => {
    const f = filter;
    if (!listing) return;
    if (
      f.startsWith("\\\\") &&
      !f.slice(2).includes("\\") &&
      !listing.dir.toLowerCase().includes("wsl.localhost")
    ) {
      const rest = f.slice(2);
      (async () => {
        await navTo("\\\\wsl.localhost\\", null);
        if (filter === "" && rest !== "") filter = rest;
      })();
    }
  });

  function joinPath(dir: string, name: string): string {
    const sep = dir.includes("\\") ? "\\" : "/";
    return dir.endsWith(sep) ? dir + name : dir + sep + name;
  }

  async function activateEntry(entry: DirEntry) {
    if (!listing) return;
    if (entry.name === "..") {
      if (listing.parent) await navTo(listing.parent, null);
      return;
    }
    if (entry.is_dir) {
      await navTo(joinPath(listing.dir, entry.name), null);
    } else {
      const full = joinPath(listing.dir, entry.name);
      await onOpenFile(full);
    }
  }

  async function activateRow(row: NavRow) {
    if (row.kind === "recent") {
      await onOpenFile(row.recent.path);
      return;
    }
    if (row.kind === "recent_dir") {
      // Re-base navigation at the chosen directory; navTo flips
      // hasInteracted so both recents lists collapse and the user
      // continues browsing from there. No file is opened yet.
      await navTo(row.recent_dir.dir, null);
      return;
    }
    await activateEntry(row.entry);
  }

  function saveTargetFromFilter(): string | null {
    if (!listing) return null;
    const raw = filter.trim();
    if (!raw) return null;
    // Default to .xlsx, but accept .xls explicitly — fastsheet has a
    // BIFF writer for that path now (formulas round-trip as cached
    // values until the ptg encoder lands; values + styles round-trip
    // fully). Any other extension passes through as typed.
    const lower = raw.toLowerCase();
    const filename = (lower.endsWith(".xlsx") || lower.endsWith(".xls"))
      ? raw
      : `${raw}.xlsx`;
    return joinPath(listing.dir, filename);
  }

  async function submit() {
    if (isNavToken(filter)) {
      await navTo(filter, listing?.dir ?? null);
      return;
    }
    // Save-mode: filter has no matching entries — save to <dir>/<filter>[.xlsx]
    if (mode === "save" && filter.trim() && rows.length === 0) {
      const target = saveTargetFromFilter();
      if (target) await onSaveFile(target);
      return;
    }
    const sel = rows[selectedIdx];
    if (!sel) {
      if (mode === "save" && filter.trim()) {
        const target = saveTargetFromFilter();
        if (target) {
          await onSaveFile(target);
          return;
        }
      }
      onStatus("no match");
      return;
    }
    // In save mode, picking a file uses it as the save target.
    if (mode === "save" && sel.kind === "entry" && !sel.entry.is_dir && listing) {
      await onSaveFile(joinPath(listing.dir, sel.entry.name));
      return;
    }
    await activateRow(sel);
  }

  function move(delta: number) {
    if (rows.length === 0) return;
    selectedIdx = ((selectedIdx + delta) % rows.length + rows.length) % rows.length;
  }

  function onFilterKey(e: KeyboardEvent) {
    if (e.key === "Escape") {
      e.preventDefault();
      onClose();
    } else if (e.key === "ArrowDown") {
      e.preventDefault();
      move(1);
    } else if (e.key === "ArrowUp") {
      e.preventDefault();
      move(-1);
    } else if (e.key === "Tab") {
      e.preventDefault();
      move(e.shiftKey ? -1 : 1);
    } else if (e.key === "Enter") {
      e.preventDefault();
      submit();
    }
  }
</script>

{#if listing}
  <div class="navigator">
    <div class="nav-dir">{listing.dir}</div>
    <input
      class="nav-filter"
      bind:this={filterEl}
      bind:value={filter}
      placeholder="filter (or type ..  c:  ~  /abs/path)"
      onkeydown={onFilterKey}
      autofocus
    />
    <div class="nav-list" role="listbox">
      {#each rows as row, i (row.kind + ":" + (row.kind === "recent" ? row.recent.path : row.kind === "recent_dir" ? row.recent_dir.dir : row.entry.name) + ":" + i)}
        {#if row.kind === "recent"}
          <div
            class="nav-row recent"
            class:sel={i === selectedIdx}
            role="option"
            aria-selected={i === selectedIdx}
            tabindex="-1"
            onclick={() => {
              selectedIdx = i;
              activateRow(row);
            }}
            onkeydown={(e) => {
              if (e.key === "Enter") {
                selectedIdx = i;
                activateRow(row);
              }
            }}
          >
            <span class="nav-tag">recent</span>
            <span class="nav-name">{row.recent.basename}</span>
            <span class="nav-dim">{row.recent.dir}</span>
          </div>
        {:else if row.kind === "recent_dir"}
          <div
            class="nav-row recent-dir"
            class:sel={i === selectedIdx}
            role="option"
            aria-selected={i === selectedIdx}
            tabindex="-1"
            onclick={() => {
              selectedIdx = i;
              activateRow(row);
            }}
            onkeydown={(e) => {
              if (e.key === "Enter") {
                selectedIdx = i;
                activateRow(row);
              }
            }}
          >
            <span class="nav-tag dir-tag">recent dir</span>
            <span class="nav-name">{row.recent_dir.dir}</span>
          </div>
        {:else}
          <div
            class="nav-row"
            class:sel={i === selectedIdx}
            class:dir={row.entry.is_dir}
            role="option"
            aria-selected={i === selectedIdx}
            tabindex="-1"
            onclick={() => {
              selectedIdx = i;
              activateRow(row);
            }}
            onkeydown={(e) => {
              if (e.key === "Enter") {
                selectedIdx = i;
                activateRow(row);
              }
            }}
          >
            <span class="nav-name"
              >{row.entry.name}{row.entry.is_dir && row.entry.name !== ".."
                ? "/"
                : ""}</span
            >
            {#if !row.entry.is_dir && row.entry.size != null}
              <span class="nav-size">{Math.ceil(row.entry.size / 1024)} KB</span>
            {/if}
          </div>
        {/if}
      {/each}
      {#if rows.length === 0}
        <div class="nav-empty">no matches</div>
      {/if}
    </div>
    <div class="nav-hint">↑/↓ Tab to select · Enter open · Esc cancel</div>
  </div>
{/if}

<style>
  .navigator {
    background: #f8f8f8;
    border-bottom: 1px solid #c0c0c0;
    padding: 0.4rem 0.6rem 0.5rem;
    display: flex;
    flex-direction: column;
    gap: 0.3rem;
    max-height: 60vh;
  }
  .nav-dir {
    color: #1f6feb;
    font-size: 12px;
    font-weight: 600;
  }
  .nav-filter {
    background: #fff;
    color: #111;
    border: 1px solid #c0c0c0;
    padding: 0.2rem 0.5rem;
    font: inherit;
    font-size: 12px;
  }
  .nav-list {
    overflow-y: auto;
    background: #fff;
    border: 1px solid #d0d0d0;
    min-height: 4rem;
    max-height: 40vh;
    font-size: 12px;
  }
  .nav-row {
    display: flex;
    justify-content: space-between;
    padding: 0.1rem 0.5rem;
    cursor: pointer;
    color: #222;
  }
  .nav-row.dir .nav-name {
    color: #1f6feb;
    font-weight: 600;
  }
  .nav-row.sel {
    background: #1f6feb;
    color: #fff;
  }
  .nav-row.sel.dir .nav-name {
    color: #fff;
  }
  .nav-size {
    color: #888;
    font-size: 11px;
  }
  .nav-row.sel .nav-size {
    color: #d0e0ff;
  }
  .nav-row.recent,
  .nav-row.recent-dir {
    gap: 0.6rem;
  }
  .nav-row.recent-dir .nav-name {
    color: #1f6feb;
    font-weight: 600;
    flex: 1;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
  }
  .nav-row.recent-dir.sel .nav-name {
    color: #fff;
  }
  .nav-tag {
    color: #b88a00;
    font-size: 10px;
    text-transform: uppercase;
    font-weight: 700;
    min-width: 5rem;
  }
  .nav-tag.dir-tag {
    color: #1f6feb;
  }
  .nav-row.sel .nav-tag {
    color: #fff;
  }
  .nav-dim {
    flex: 1;
    color: #888;
    font-size: 11px;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
  }
  .nav-row.sel .nav-dim {
    color: #d0e0ff;
  }
  .nav-empty {
    color: #888;
    padding: 0.6rem;
    text-align: center;
  }
  .nav-hint {
    color: #888;
    font-size: 11px;
  }
</style>
