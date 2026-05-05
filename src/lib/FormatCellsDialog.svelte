<script lang="ts">
  import { invoke } from "@tauri-apps/api/core";
  import { onMount, tick } from "svelte";

  type Props = {
    sheet: number;
    /// Active cell at the time the modal opened — used to seed the
    /// initial form state. Subsequent applies hit the full selection
    /// rectangle (rangeR1..rangeR2 × rangeC1..rangeC2).
    cellRow: number;
    cellCol: number;
    rangeR1: number;
    rangeC1: number;
    rangeR2: number;
    rangeC2: number;
    onClose: () => void;
    onApplied: () => void | Promise<void>;
    onStatus: (msg: string) => void;
  };

  type FormatInfo = {
    num_fmt: string;
    bold: boolean;
    italic: boolean;
    underline: boolean;
    strike: boolean;
    font_size: number;
    font_name: string;
    font_color?: string;
    fill_color?: string;
    align_h: "general" | "left" | "center" | "right" | "justify";
    align_v: "top" | "middle" | "bottom";
    wrap: boolean;
    border_top: boolean;
    border_bottom: boolean;
    border_left: boolean;
    border_right: boolean;
  };

  let {
    sheet,
    cellRow,
    cellCol,
    rangeR1,
    rangeC1,
    rangeR2,
    rangeC2,
    onClose,
    onApplied,
    onStatus,
  }: Props = $props();

  type Tab = "number" | "font" | "border" | "fill" | "alignment";
  let activeTab = $state<Tab>("number");
  let initial = $state<FormatInfo | null>(null);
  let busy = $state(false);
  let cardEl: HTMLDivElement | null = $state(null);

  // Form state — seeded from `initial` once the read returns.
  let numCategory = $state<
    "general" | "number" | "currency" | "percent" | "scientific" | "date" | "time" | "custom"
  >("general");
  let numDecimals = $state(2);
  let numUseSep = $state(true);
  let numCustom = $state("");

  let bold = $state(false);
  let italic = $state(false);
  let underline = $state(false);
  let strike = $state(false);
  let fontColor = $state("");
  let fillColor = $state("");

  let borderTop = $state(false);
  let borderRight = $state(false);
  let borderBottom = $state(false);
  let borderLeft = $state(false);

  let alignH = $state<"general" | "left" | "center" | "right" | "justify">("general");
  let alignV = $state<"top" | "middle" | "bottom">("bottom");
  let wrap = $state(false);

  onMount(async () => {
    try {
      const info = await invoke<FormatInfo>("get_cell_format", {
        sheet,
        row: cellRow,
        col: cellCol,
      });
      initial = info;
      seedFromInfo(info);
    } catch (e) {
      onStatus(`Failed to read cell format: ${e}`);
    }
    await tick();
    cardEl?.focus();
  });

  function seedFromInfo(info: FormatInfo) {
    // Seed the Number tab by classifying the existing num_fmt string.
    const fmt = info.num_fmt;
    const m = classifyFormat(fmt);
    numCategory = m.category;
    numDecimals = m.decimals;
    numUseSep = m.useSep;
    numCustom = fmt;

    bold = info.bold;
    italic = info.italic;
    underline = info.underline;
    strike = info.strike;
    fontColor = info.font_color ?? "";
    fillColor = info.fill_color ?? "";

    borderTop = info.border_top;
    borderRight = info.border_right;
    borderBottom = info.border_bottom;
    borderLeft = info.border_left;

    alignH = info.align_h;
    alignV = info.align_v;
    wrap = info.wrap;
  }

  function classifyFormat(fmt: string): {
    category: typeof numCategory;
    decimals: number;
    useSep: boolean;
  } {
    const f = fmt.trim();
    if (!f || f === "General") return { category: "general", decimals: 2, useSep: true };
    if (/^h:?mm(:ss)?(\s*AM\/PM)?/i.test(f)) return { category: "time", decimals: 0, useSep: false };
    if (/[dmy]/.test(f) && !/[#0]/.test(f)) return { category: "date", decimals: 0, useSep: false };
    const decMatch = f.match(/\.(0+)/);
    const decimals = decMatch ? decMatch[1].length : 0;
    if (/E\+?00/.test(f)) return { category: "scientific", decimals, useSep: false };
    if (f.includes("%")) return { category: "percent", decimals, useSep: false };
    if (f.startsWith("$") || /^\[\$/.test(f)) {
      return { category: "currency", decimals, useSep: f.includes("#,##0") };
    }
    if (/[#0]/.test(f)) {
      return { category: "number", decimals, useSep: f.includes("#,##0") };
    }
    return { category: "custom", decimals, useSep: false };
  }

  function buildNumFormat(): string {
    const dz = (n: number) => (n <= 0 ? "" : "." + "0".repeat(Math.min(15, n)));
    const base = (sep: boolean) => (sep ? "#,##0" : "0");
    switch (numCategory) {
      case "general":
        return "General";
      case "number":
        return base(numUseSep) + dz(numDecimals);
      case "currency":
        return "$" + base(numUseSep) + dz(numDecimals);
      case "percent":
        return base(false) + dz(numDecimals) + "%";
      case "scientific":
        return base(false) + dz(numDecimals) + "E+00";
      case "date":
        return "yyyy-mm-dd";
      case "time":
        return "h:mm:ss";
      case "custom":
        return numCustom.trim() || "General";
    }
  }

  function previewSample(): string {
    const fmt = buildNumFormat();
    const sample = 1234.5678;
    if (numCategory === "general" || numCategory === "custom") {
      return `Format: ${fmt}`;
    }
    if (numCategory === "date") return "Today as 2026-05-05";
    if (numCategory === "time") return "Noon as 12:00:00";
    // Best-effort numeric preview.
    if (numCategory === "percent") return `${(0.1234).toLocaleString(undefined, { minimumFractionDigits: numDecimals, maximumFractionDigits: numDecimals })}%`;
    if (numCategory === "currency") {
      const s = sample.toLocaleString(undefined, {
        minimumFractionDigits: numDecimals,
        maximumFractionDigits: numDecimals,
        useGrouping: numUseSep,
      });
      return `$${s}`;
    }
    if (numCategory === "scientific") return sample.toExponential(numDecimals).replace("e", "E");
    return sample.toLocaleString(undefined, {
      minimumFractionDigits: numDecimals,
      maximumFractionDigits: numDecimals,
      useGrouping: numUseSep,
    });
  }

  function isHexColor(s: string): boolean {
    return /^#[0-9A-Fa-f]{6}$/.test(s.trim());
  }

  /// Apply the form against the active selection. Each section only
  /// fires the corresponding tauri call when its values DIFFER from
  /// the initial snapshot — so opening the dialog and pressing Apply
  /// without changes is a no-op. This keeps the undo stack clean.
  async function applyAll() {
    if (!initial || busy) return;
    busy = true;
    let dirty = false;
    try {
      // 1. Number format.
      const newFmt = buildNumFormat();
      if (newFmt !== initial.num_fmt) {
        await invoke("set_range_number_format", {
          sheet,
          r1: rangeR1,
          c1: rangeC1,
          r2: rangeR2,
          c2: rangeC2,
          format: newFmt,
        });
        dirty = true;
      }

      // 2. Font attributes. Use semantic setters so the result does
      // not depend on whichever cell happens to be top-left in the
      // selected range.
      const setIfDiffers = async (
        formVal: boolean,
        initVal: boolean,
        opKind: string,
      ) => {
        if (formVal !== initVal) {
          await invoke("set_range_style", {
            sheet,
            r1: rangeR1,
            c1: rangeC1,
            r2: rangeR2,
            c2: rangeC2,
            op: { kind: opKind, enabled: formVal },
          });
          dirty = true;
        }
      };
      await setIfDiffers(bold, initial.bold, "set_bold");
      await setIfDiffers(italic, initial.italic, "set_italic");
      await setIfDiffers(underline, initial.underline, "set_underline");
      await setIfDiffers(strike, initial.strike, "set_strike");

      // 3. Font color.
      if (fontColor.trim() !== (initial.font_color ?? "")) {
        if (fontColor.trim() === "") {
          await invoke("set_range_style", {
            sheet, r1: rangeR1, c1: rangeC1, r2: rangeR2, c2: rangeC2,
            op: { kind: "clear_text_color" },
          });
        } else if (isHexColor(fontColor)) {
          await invoke("set_range_style", {
            sheet, r1: rangeR1, c1: rangeC1, r2: rangeR2, c2: rangeC2,
            op: { kind: "set_text_color", color: fontColor.trim().toUpperCase() },
          });
        } else {
          throw new Error(`Invalid font colour: ${fontColor} (use #RRGGBB or empty)`);
        }
        dirty = true;
      }

      // 4. Fill color.
      if (fillColor.trim() !== (initial.fill_color ?? "")) {
        if (fillColor.trim() === "") {
          await invoke("set_range_style", {
            sheet, r1: rangeR1, c1: rangeC1, r2: rangeR2, c2: rangeC2,
            op: { kind: "clear_fill_color" },
          });
        } else if (isHexColor(fillColor)) {
          await invoke("set_range_style", {
            sheet, r1: rangeR1, c1: rangeC1, r2: rangeR2, c2: rangeC2,
            op: { kind: "set_fill_color", color: fillColor.trim().toUpperCase() },
          });
        } else {
          throw new Error(`Invalid fill colour: ${fillColor} (use #RRGGBB or empty)`);
        }
        dirty = true;
      }

      // 5. Borders. Compose into a single SetBorder call by inferring
      // the side token from which checkbox changed. For simplicity:
      // if any border state differs, clear then re-apply each side
      // that's checked — using the existing per-side tokens.
      const borderChanged =
        borderTop !== initial.border_top ||
        borderRight !== initial.border_right ||
        borderBottom !== initial.border_bottom ||
        borderLeft !== initial.border_left;
      if (borderChanged) {
        await invoke("set_range_style", {
          sheet, r1: rangeR1, c1: rangeC1, r2: rangeR2, c2: rangeC2,
          op: { kind: "set_border", sides: "none" },
        });
        const sides: Array<["top" | "right" | "bottom" | "left", boolean]> = [
          ["top", borderTop],
          ["right", borderRight],
          ["bottom", borderBottom],
          ["left", borderLeft],
        ];
        for (const [side, on] of sides) {
          if (on) {
            await invoke("set_range_style", {
              sheet, r1: rangeR1, c1: rangeC1, r2: rangeR2, c2: rangeC2,
              op: { kind: "set_border", sides: side },
            });
          }
        }
        dirty = true;
      }

      // 6. Alignment.
      if (alignH !== initial.align_h) {
        const map: Record<string, string> = {
          general: "align_general",
          left: "align_left",
          center: "align_center",
          right: "align_right",
          justify: "align_left", // best-effort: justify not supported in StyleOp
        };
        await invoke("set_range_style", {
          sheet, r1: rangeR1, c1: rangeC1, r2: rangeR2, c2: rangeC2,
          op: { kind: map[alignH] },
        });
        dirty = true;
      }

      onStatus(dirty ? `Format applied to ${describeRange()}` : "No changes");
      await onApplied();
      onClose();
    } catch (e) {
      onStatus(`Apply failed: ${e}`);
    } finally {
      busy = false;
    }
  }

  function describeRange(): string {
    const { addr } = ((): { addr: (r: number, c: number) => string } => ({
      addr: (r, c) => {
        let n = c;
        let s = "";
        while (n > 0) {
          n--;
          s = String.fromCharCode(65 + (n % 26)) + s;
          n = Math.floor(n / 26);
        }
        return `${s}${r}`;
      },
    }))();
    if (rangeR1 === rangeR2 && rangeC1 === rangeC2) return addr(rangeR1, rangeC1);
    return `${addr(rangeR1, rangeC1)}:${addr(rangeR2, rangeC2)}`;
  }

  function onKey(e: KeyboardEvent) {
    if (e.key === "Escape") {
      e.preventDefault();
      e.stopPropagation();
      onClose();
      return;
    }
    if (e.key === "Enter" && !(e.target instanceof HTMLTextAreaElement)) {
      // Don't intercept Enter inside text inputs — the user might be
      // typing a custom format string. Apply only when the focus is
      // on a non-input element.
      const t = e.target as HTMLElement | null;
      if (t && (t.tagName === "INPUT" || t.tagName === "SELECT" || t.tagName === "TEXTAREA")) {
        return;
      }
      e.preventDefault();
      e.stopPropagation();
      applyAll();
      return;
    }
    // Ctrl+1..5 to switch tabs.
    if (e.ctrlKey && !e.shiftKey && !e.altKey) {
      const tabs: Tab[] = ["number", "font", "border", "fill", "alignment"];
      const idx = parseInt(e.key, 10) - 1;
      if (idx >= 0 && idx < tabs.length) {
        e.preventDefault();
        e.stopPropagation();
        activeTab = tabs[idx];
      }
    }
  }
</script>

<div
  class="overlay"
  role="dialog"
  aria-modal="true"
  aria-label="Format Cells"
  tabindex="-1"
  onkeydown={onKey}
>
  <div
    class="card"
    bind:this={cardEl}
    tabindex="-1"
  >
    <header class="header">
      <h2>Format Cells — <span class="range">{describeRange()}</span></h2>
      <button class="close" onclick={onClose} aria-label="Close">×</button>
    </header>

    <div class="tabs" role="tablist">
      {#each [
        { id: "number" as const, label: "Number" },
        { id: "font" as const, label: "Font" },
        { id: "border" as const, label: "Border" },
        { id: "fill" as const, label: "Fill" },
        { id: "alignment" as const, label: "Alignment" },
      ] as t (t.id)}
        <button
          class="tab"
          class:active={activeTab === t.id}
          role="tab"
          aria-selected={activeTab === t.id}
          onclick={() => (activeTab = t.id)}
        >
          {t.label}
        </button>
      {/each}
    </div>

    <section class="body">
      {#if !initial}
        <p class="loading">Loading…</p>
      {:else if activeTab === "number"}
        <div class="grid-2">
          <div class="categories">
            {#each [
              { id: "general", label: "General" },
              { id: "number", label: "Number" },
              { id: "currency", label: "Currency" },
              { id: "percent", label: "Percent" },
              { id: "scientific", label: "Scientific" },
              { id: "date", label: "Date" },
              { id: "time", label: "Time" },
              { id: "custom", label: "Custom" },
            ] as cat (cat.id)}
              <button
                class="cat"
                class:active={numCategory === cat.id}
                onclick={() => (numCategory = cat.id as typeof numCategory)}
              >
                {cat.label}
              </button>
            {/each}
          </div>
          <div class="cat-detail">
            {#if ["number", "currency", "percent", "scientific"].includes(numCategory)}
              <label>
                Decimals
                <input
                  type="number"
                  min="0"
                  max="15"
                  bind:value={numDecimals}
                />
              </label>
              {#if numCategory === "number" || numCategory === "currency"}
                <label class="check">
                  <input type="checkbox" bind:checked={numUseSep} />
                  Use thousands separator (1,000)
                </label>
              {/if}
            {:else if numCategory === "custom"}
              <label>
                Format string
                <input
                  type="text"
                  bind:value={numCustom}
                  placeholder="e.g. 0.00 or [Red]#,##0;[Black]-#,##0"
                />
              </label>
              <p class="hint">
                Excel-style format codes. Reset to General by leaving empty.
              </p>
            {:else if numCategory === "date"}
              <p class="hint">Locked to <code>yyyy-mm-dd</code> for now. Use Custom for other date formats.</p>
            {:else if numCategory === "time"}
              <p class="hint">Locked to <code>h:mm:ss</code>. Use Custom for other time formats.</p>
            {:else}
              <p class="hint">No options for General. Cell displays values as-typed.</p>
            {/if}
            <div class="preview">
              <span class="preview-label">Preview</span>
              <span class="preview-value">{previewSample()}</span>
            </div>
          </div>
        </div>
      {:else if activeTab === "font"}
        <div class="font-grid">
          <label class="check">
            <input type="checkbox" bind:checked={bold} /> <strong>Bold</strong>
          </label>
          <label class="check">
            <input type="checkbox" bind:checked={italic} /> <em>Italic</em>
          </label>
          <label class="check">
            <input type="checkbox" bind:checked={underline} /> <span style="text-decoration: underline">Underline</span>
          </label>
          <label class="check">
            <input type="checkbox" bind:checked={strike} /> <span style="text-decoration: line-through">Strike</span>
          </label>
          <label class="color-row">
            Text colour
            <input type="text" placeholder="#RRGGBB or empty" bind:value={fontColor} />
            <span class="swatch" style:background={isHexColor(fontColor) ? fontColor : "transparent"}></span>
          </label>
          <p class="hint">
            Font face + size are sourced from the file's defaults. Use the
            colour field to override per-cell text colour.
          </p>
        </div>
      {:else if activeTab === "border"}
        <div class="border-grid">
          <label class="check"><input type="checkbox" bind:checked={borderTop} /> Top</label>
          <label class="check"><input type="checkbox" bind:checked={borderRight} /> Right</label>
          <label class="check"><input type="checkbox" bind:checked={borderBottom} /> Bottom</label>
          <label class="check"><input type="checkbox" bind:checked={borderLeft} /> Left</label>
          <div class="quick">
            <button onclick={() => { borderTop = borderRight = borderBottom = borderLeft = true; }}>All</button>
            <button onclick={() => { borderTop = borderRight = borderBottom = borderLeft = false; }}>None</button>
          </div>
          <p class="hint">
            All borders are thin black. The /R F R menu has finer controls
            (per-side and outline-only). Outline mode applies to the
            selection's perimeter — not adjustable from this dialog.
          </p>
        </div>
      {:else if activeTab === "fill"}
        <div class="fill-grid">
          <label class="color-row">
            Fill colour
            <input type="text" placeholder="#RRGGBB or empty" bind:value={fillColor} />
            <span class="swatch" style:background={isHexColor(fillColor) ? fillColor : "transparent"}></span>
          </label>
          <div class="swatches">
            {#each ["#FFD966", "#F4B084", "#A9D08E", "#9BC2E6", "#FFC0CB", "#D9D9D9", "#000000", "#FFFFFF"] as c (c)}
              <button
                class="sw"
                style:background={c}
                aria-label={`Set fill to ${c}`}
                onclick={() => (fillColor = c)}
              ></button>
            {/each}
          </div>
          <p class="hint">Leave empty to clear fill.</p>
        </div>
      {:else if activeTab === "alignment"}
        <div class="align-grid">
          <fieldset>
            <legend>Horizontal</legend>
            {#each ["general", "left", "center", "right"] as h (h)}
              <label class="radio">
                <input type="radio" name="align-h" value={h} bind:group={alignH} />
                {h.charAt(0).toUpperCase() + h.slice(1)}
              </label>
            {/each}
          </fieldset>
          <fieldset>
            <legend>Vertical (display only)</legend>
            <p class="hint">Not yet writable — IronCalc preserves vertical alignment from the source file but the toolbar's vertical setter isn't wired.</p>
          </fieldset>
          <label class="check">
            <input type="checkbox" bind:checked={wrap} disabled />
            Wrap text (read-only for now)
          </label>
        </div>
      {/if}
    </section>

    <footer class="footer">
      <span class="hint-keys">Ctrl+1..5 switch tabs · Enter applies · Esc cancels</span>
      <div class="actions">
        <button onclick={onClose}>Cancel</button>
        <button class="primary" onclick={applyAll} disabled={busy || !initial}>
          {busy ? "Applying…" : "Apply"}
        </button>
      </div>
    </footer>
  </div>
</div>

<style>
  .overlay {
    position: fixed;
    inset: 0;
    background: rgba(0, 0, 0, 0.4);
    z-index: 1000;
    display: flex;
    align-items: center;
    justify-content: center;
  }
  .card {
    background: #fff;
    border: 1px solid #888;
    box-shadow: 0 8px 32px rgba(0, 0, 0, 0.25);
    width: min(640px, 92vw);
    max-height: 90vh;
    display: flex;
    flex-direction: column;
    font-size: 13px;
    color: #222;
    outline: none;
  }
  .header {
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 8px 12px;
    border-bottom: 1px solid #ddd;
    background: #f5f5f5;
  }
  .header h2 {
    font-size: 14px;
    font-weight: 600;
    margin: 0;
  }
  .range {
    font-family: ui-monospace, monospace;
    color: #555;
    font-weight: normal;
  }
  .close {
    background: none;
    border: none;
    font-size: 20px;
    line-height: 1;
    cursor: pointer;
    color: #555;
    padding: 0 6px;
  }
  .close:hover { color: #000; }
  .tabs {
    display: flex;
    border-bottom: 1px solid #ddd;
    background: #fafafa;
  }
  .tab {
    background: none;
    border: none;
    padding: 8px 14px;
    cursor: pointer;
    font-size: 13px;
    border-bottom: 2px solid transparent;
    color: #555;
  }
  .tab:hover { color: #000; }
  .tab.active {
    color: #000;
    border-bottom-color: #2c6cb0;
    font-weight: 500;
  }
  .body {
    flex: 1;
    overflow: auto;
    padding: 14px 16px;
  }
  .loading { color: #888; }
  .grid-2 {
    display: grid;
    grid-template-columns: 160px 1fr;
    gap: 16px;
  }
  .categories {
    display: flex;
    flex-direction: column;
    border: 1px solid #ddd;
    border-radius: 2px;
    overflow: hidden;
  }
  .cat {
    background: #fff;
    border: none;
    border-bottom: 1px solid #eee;
    padding: 6px 10px;
    text-align: left;
    cursor: pointer;
    font-size: 13px;
  }
  .cat:last-child { border-bottom: none; }
  .cat:hover { background: #f0f0f0; }
  .cat.active { background: #2c6cb0; color: #fff; }
  .cat-detail label {
    display: block;
    margin-bottom: 8px;
  }
  .cat-detail input[type="number"], .cat-detail input[type="text"] {
    width: 100%;
    padding: 4px 6px;
    margin-top: 2px;
    border: 1px solid #ccc;
    border-radius: 2px;
    font: inherit;
  }
  .check {
    display: flex !important;
    align-items: center;
    gap: 6px;
    margin-bottom: 6px !important;
  }
  .check input[type="checkbox"] { margin: 0; }
  .hint {
    color: #666;
    font-size: 12px;
    margin-top: 8px;
  }
  .hint code {
    background: #f0f0f0;
    padding: 0 4px;
    border-radius: 2px;
    font-family: ui-monospace, monospace;
  }
  .preview {
    margin-top: 14px;
    padding: 8px 10px;
    border: 1px dashed #ccc;
    border-radius: 2px;
    background: #fafafa;
  }
  .preview-label {
    color: #888;
    font-size: 11px;
    text-transform: uppercase;
    letter-spacing: 0.5px;
    display: block;
  }
  .preview-value {
    font-family: ui-monospace, monospace;
    font-size: 14px;
  }
  .font-grid, .border-grid, .fill-grid, .align-grid {
    display: flex;
    flex-direction: column;
    gap: 8px;
  }
  .border-grid {
    display: grid;
    grid-template-columns: 1fr 1fr;
    column-gap: 24px;
  }
  .border-grid .quick {
    grid-column: span 2;
    display: flex;
    gap: 8px;
    margin-top: 4px;
  }
  .border-grid .quick button {
    padding: 4px 12px;
    border: 1px solid #ccc;
    background: #f5f5f5;
    cursor: pointer;
  }
  .border-grid .hint { grid-column: span 2; }
  .color-row {
    display: flex !important;
    align-items: center;
    gap: 8px;
    margin-bottom: 8px;
  }
  .color-row input[type="text"] {
    flex: 1;
    padding: 4px 6px;
    border: 1px solid #ccc;
    border-radius: 2px;
    font: inherit;
  }
  .swatch {
    display: inline-block;
    width: 22px;
    height: 22px;
    border: 1px solid #888;
  }
  .swatches {
    display: flex;
    gap: 4px;
    flex-wrap: wrap;
  }
  .sw {
    width: 22px;
    height: 22px;
    border: 1px solid #888;
    cursor: pointer;
    padding: 0;
  }
  fieldset {
    border: 1px solid #ddd;
    padding: 8px 12px;
    border-radius: 2px;
  }
  fieldset legend {
    padding: 0 6px;
    font-size: 12px;
    color: #666;
  }
  .radio {
    display: inline-flex;
    align-items: center;
    gap: 4px;
    margin-right: 12px;
    font-size: 13px;
  }
  .footer {
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 10px 14px;
    border-top: 1px solid #ddd;
    background: #f5f5f5;
  }
  .hint-keys {
    font-size: 11px;
    color: #777;
  }
  .actions {
    display: flex;
    gap: 8px;
  }
  .actions button {
    padding: 5px 16px;
    border: 1px solid #ccc;
    background: #fff;
    cursor: pointer;
    font: inherit;
  }
  .actions button.primary {
    background: #2c6cb0;
    color: #fff;
    border-color: #1f5290;
  }
  .actions button.primary:disabled {
    background: #aac3df;
    border-color: #aac3df;
    cursor: not-allowed;
  }
</style>
