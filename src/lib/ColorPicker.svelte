<script lang="ts">
  import { onMount, tick } from "svelte";

  type Props = {
    /// Initial color (e.g. "#FFD966") — empty string means "no
    /// existing color" (we'll default the highlight to whatever
    /// matches the typed filter, or to the first recent).
    initial?: string;
    /// Distinct hex colors already used in the workbook. Surfaced as
    /// the top row so the user can match an existing palette without
    /// hunting through 140 named colors.
    recents?: string[];
    /// Friendly title — "Fill colour", "Text colour", etc.
    title?: string;
    /// Whether the picker should offer a Clear button (sets to no
    /// color). Fill always wants this; text often does too. When
    /// `false` the Clear option is hidden.
    allowClear?: boolean;
    onSelect: (hex: string) => void | Promise<void>;
    onClear?: () => void | Promise<void>;
    onCancel: () => void;
  };

  let {
    initial = "",
    recents = [],
    title = "Pick a colour",
    allowClear = true,
    onSelect,
    onClear,
    onCancel,
  }: Props = $props();

  // ---------------------------------------------------------------------
  // Named CSS color list — the standard 140 named colors give the user
  // a recognisable spectrum to type-and-pick from. We pre-extract HSL
  // so filtering by name + previewing in the custom editor share the
  // same source of truth.
  // ---------------------------------------------------------------------
  type NamedColor = { name: string; hex: string };
  const NAMED: NamedColor[] = [
    { name: "AliceBlue", hex: "#F0F8FF" }, { name: "AntiqueWhite", hex: "#FAEBD7" },
    { name: "Aqua", hex: "#00FFFF" }, { name: "Aquamarine", hex: "#7FFFD4" },
    { name: "Azure", hex: "#F0FFFF" }, { name: "Beige", hex: "#F5F5DC" },
    { name: "Bisque", hex: "#FFE4C4" }, { name: "Black", hex: "#000000" },
    { name: "BlanchedAlmond", hex: "#FFEBCD" }, { name: "Blue", hex: "#0000FF" },
    { name: "BlueViolet", hex: "#8A2BE2" }, { name: "Brown", hex: "#A52A2A" },
    { name: "BurlyWood", hex: "#DEB887" }, { name: "CadetBlue", hex: "#5F9EA0" },
    { name: "Chartreuse", hex: "#7FFF00" }, { name: "Chocolate", hex: "#D2691E" },
    { name: "Coral", hex: "#FF7F50" }, { name: "CornflowerBlue", hex: "#6495ED" },
    { name: "Cornsilk", hex: "#FFF8DC" }, { name: "Crimson", hex: "#DC143C" },
    { name: "Cyan", hex: "#00FFFF" }, { name: "DarkBlue", hex: "#00008B" },
    { name: "DarkCyan", hex: "#008B8B" }, { name: "DarkGoldenRod", hex: "#B8860B" },
    { name: "DarkGray", hex: "#A9A9A9" }, { name: "DarkGreen", hex: "#006400" },
    { name: "DarkKhaki", hex: "#BDB76B" }, { name: "DarkMagenta", hex: "#8B008B" },
    { name: "DarkOliveGreen", hex: "#556B2F" }, { name: "DarkOrange", hex: "#FF8C00" },
    { name: "DarkOrchid", hex: "#9932CC" }, { name: "DarkRed", hex: "#8B0000" },
    { name: "DarkSalmon", hex: "#E9967A" }, { name: "DarkSeaGreen", hex: "#8FBC8F" },
    { name: "DarkSlateBlue", hex: "#483D8B" }, { name: "DarkSlateGray", hex: "#2F4F4F" },
    { name: "DarkTurquoise", hex: "#00CED1" }, { name: "DarkViolet", hex: "#9400D3" },
    { name: "DeepPink", hex: "#FF1493" }, { name: "DeepSkyBlue", hex: "#00BFFF" },
    { name: "DimGray", hex: "#696969" }, { name: "DodgerBlue", hex: "#1E90FF" },
    { name: "FireBrick", hex: "#B22222" }, { name: "FloralWhite", hex: "#FFFAF0" },
    { name: "ForestGreen", hex: "#228B22" }, { name: "Fuchsia", hex: "#FF00FF" },
    { name: "Gainsboro", hex: "#DCDCDC" }, { name: "GhostWhite", hex: "#F8F8FF" },
    { name: "Gold", hex: "#FFD700" }, { name: "GoldenRod", hex: "#DAA520" },
    { name: "Gray", hex: "#808080" }, { name: "Green", hex: "#008000" },
    { name: "GreenYellow", hex: "#ADFF2F" }, { name: "HoneyDew", hex: "#F0FFF0" },
    { name: "HotPink", hex: "#FF69B4" }, { name: "IndianRed", hex: "#CD5C5C" },
    { name: "Indigo", hex: "#4B0082" }, { name: "Ivory", hex: "#FFFFF0" },
    { name: "Khaki", hex: "#F0E68C" }, { name: "Lavender", hex: "#E6E6FA" },
    { name: "LavenderBlush", hex: "#FFF0F5" }, { name: "LawnGreen", hex: "#7CFC00" },
    { name: "LemonChiffon", hex: "#FFFACD" }, { name: "LightBlue", hex: "#ADD8E6" },
    { name: "LightCoral", hex: "#F08080" }, { name: "LightCyan", hex: "#E0FFFF" },
    { name: "LightGoldenRodYellow", hex: "#FAFAD2" }, { name: "LightGray", hex: "#D3D3D3" },
    { name: "LightGreen", hex: "#90EE90" }, { name: "LightPink", hex: "#FFB6C1" },
    { name: "LightSalmon", hex: "#FFA07A" }, { name: "LightSeaGreen", hex: "#20B2AA" },
    { name: "LightSkyBlue", hex: "#87CEFA" }, { name: "LightSlateGray", hex: "#778899" },
    { name: "LightSteelBlue", hex: "#B0C4DE" }, { name: "LightYellow", hex: "#FFFFE0" },
    { name: "Lime", hex: "#00FF00" }, { name: "LimeGreen", hex: "#32CD32" },
    { name: "Linen", hex: "#FAF0E6" }, { name: "Magenta", hex: "#FF00FF" },
    { name: "Maroon", hex: "#800000" }, { name: "MediumAquaMarine", hex: "#66CDAA" },
    { name: "MediumBlue", hex: "#0000CD" }, { name: "MediumOrchid", hex: "#BA55D3" },
    { name: "MediumPurple", hex: "#9370DB" }, { name: "MediumSeaGreen", hex: "#3CB371" },
    { name: "MediumSlateBlue", hex: "#7B68EE" }, { name: "MediumSpringGreen", hex: "#00FA9A" },
    { name: "MediumTurquoise", hex: "#48D1CC" }, { name: "MediumVioletRed", hex: "#C71585" },
    { name: "MidnightBlue", hex: "#191970" }, { name: "MintCream", hex: "#F5FFFA" },
    { name: "MistyRose", hex: "#FFE4E1" }, { name: "Moccasin", hex: "#FFE4B5" },
    { name: "NavajoWhite", hex: "#FFDEAD" }, { name: "Navy", hex: "#000080" },
    { name: "OldLace", hex: "#FDF5E6" }, { name: "Olive", hex: "#808000" },
    { name: "OliveDrab", hex: "#6B8E23" }, { name: "Orange", hex: "#FFA500" },
    { name: "OrangeRed", hex: "#FF4500" }, { name: "Orchid", hex: "#DA70D6" },
    { name: "PaleGoldenRod", hex: "#EEE8AA" }, { name: "PaleGreen", hex: "#98FB98" },
    { name: "PaleTurquoise", hex: "#AFEEEE" }, { name: "PaleVioletRed", hex: "#DB7093" },
    { name: "PapayaWhip", hex: "#FFEFD5" }, { name: "PeachPuff", hex: "#FFDAB9" },
    { name: "Peru", hex: "#CD853F" }, { name: "Pink", hex: "#FFC0CB" },
    { name: "Plum", hex: "#DDA0DD" }, { name: "PowderBlue", hex: "#B0E0E6" },
    { name: "Purple", hex: "#800080" }, { name: "RebeccaPurple", hex: "#663399" },
    { name: "Red", hex: "#FF0000" }, { name: "RosyBrown", hex: "#BC8F8F" },
    { name: "RoyalBlue", hex: "#4169E1" }, { name: "SaddleBrown", hex: "#8B4513" },
    { name: "Salmon", hex: "#FA8072" }, { name: "SandyBrown", hex: "#F4A460" },
    { name: "SeaGreen", hex: "#2E8B57" }, { name: "SeaShell", hex: "#FFF5EE" },
    { name: "Sienna", hex: "#A0522D" }, { name: "Silver", hex: "#C0C0C0" },
    { name: "SkyBlue", hex: "#87CEEB" }, { name: "SlateBlue", hex: "#6A5ACD" },
    { name: "SlateGray", hex: "#708090" }, { name: "Snow", hex: "#FFFAFA" },
    { name: "SpringGreen", hex: "#00FF7F" }, { name: "SteelBlue", hex: "#4682B4" },
    { name: "Tan", hex: "#D2B48C" }, { name: "Teal", hex: "#008080" },
    { name: "Thistle", hex: "#D8BFD8" }, { name: "Tomato", hex: "#FF6347" },
    { name: "Turquoise", hex: "#40E0D0" }, { name: "Violet", hex: "#EE82EE" },
    { name: "Wheat", hex: "#F5DEB3" }, { name: "White", hex: "#FFFFFF" },
    { name: "WhiteSmoke", hex: "#F5F5F5" }, { name: "Yellow", hex: "#FFFF00" },
    { name: "YellowGreen", hex: "#9ACD32" },
  ];

  // ---------------------------------------------------------------------
  // Color math — sRGB ↔ HSL with the standard formulas. Custom-mode
  // sliders run on HSL (lightness vertical, saturation horizontal); the
  // hue is locked to whatever the user landed on, so the spectrum
  // remains coherent during the refine step.
  // ---------------------------------------------------------------------
  function hexToRgb(hex: string): [number, number, number] {
    const s = hex.replace("#", "");
    return [
      parseInt(s.slice(0, 2), 16),
      parseInt(s.slice(2, 4), 16),
      parseInt(s.slice(4, 6), 16),
    ];
  }
  function rgbToHex(r: number, g: number, b: number): string {
    const c = (n: number) => Math.round(Math.max(0, Math.min(255, n))).toString(16).padStart(2, "0");
    return ("#" + c(r) + c(g) + c(b)).toUpperCase();
  }
  function rgbToHsl(r: number, g: number, b: number): [number, number, number] {
    r /= 255; g /= 255; b /= 255;
    const max = Math.max(r, g, b), min = Math.min(r, g, b);
    let h = 0, s = 0;
    const l = (max + min) / 2;
    if (max !== min) {
      const d = max - min;
      s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
      switch (max) {
        case r: h = ((g - b) / d + (g < b ? 6 : 0)); break;
        case g: h = ((b - r) / d + 2); break;
        case b: h = ((r - g) / d + 4); break;
      }
      h *= 60;
    }
    return [h, s, l];
  }
  function hslToRgb(h: number, s: number, l: number): [number, number, number] {
    h = ((h % 360) + 360) % 360;
    const c = (1 - Math.abs(2 * l - 1)) * s;
    const x = c * (1 - Math.abs(((h / 60) % 2) - 1));
    const m = l - c / 2;
    let r = 0, g = 0, b = 0;
    if (h < 60)       { r = c; g = x; b = 0; }
    else if (h < 120) { r = x; g = c; b = 0; }
    else if (h < 180) { r = 0; g = c; b = x; }
    else if (h < 240) { r = 0; g = x; b = c; }
    else if (h < 300) { r = x; g = 0; b = c; }
    else              { r = c; g = 0; b = x; }
    return [(r + m) * 255, (g + m) * 255, (b + m) * 255];
  }
  function hexToHsl(hex: string): [number, number, number] {
    const [r, g, b] = hexToRgb(hex);
    return rgbToHsl(r, g, b);
  }
  function hslToHex(h: number, s: number, l: number): string {
    const [r, g, b] = hslToRgb(h, s, l);
    return rgbToHex(r, g, b);
  }

  // ---------------------------------------------------------------------
  // Modes + filtering
  // ---------------------------------------------------------------------
  type Mode = "list" | "custom";
  let mode = $state<Mode>("list");
  let filter = $state("");
  let inputEl: HTMLInputElement | null = $state(null);
  /// Card ref. Receives keyboard focus when the user enters custom
  /// mode — without this, the click on Custom… would leave focus on
  /// the (now-hidden) Custom button, the browser would drop focus to
  /// `body`, and the onkeydown handler on the overlay would never
  /// fire because no descendant has focus.
  let cardEl: HTMLDivElement | null = $state(null);

  // Custom-mode HSL values. Seeded from the highlighted swatch when
  // the user enters custom mode so editing starts at "the colour they
  // were just looking at" rather than a fixed value.
  let customH = $state(0);
  let customS = $state(0.5);
  let customL = $state(0.5);
  /// Last hex the user navigated to that was a real colour (named or
  /// recent — not Custom… / Clear). Used to seed customH/S/L when
  /// they enter custom mode by clicking the Custom… tile, since at
  /// that exact moment `highlight` points at the Custom… entry which
  /// has no colour of its own.
  let lastColorHex = $state<string>("");

  // Build the candidate list (recents + named, with optional filter).
  // Each entry has a stable index used by keyboard navigation. The
  // synthetic "Custom…" tile and the optional "Clear" tile sit at the
  // end so they're reachable but never consume a search-rank slot.
  type Entry =
    | { kind: "recent"; hex: string }
    | { kind: "named"; name: string; hex: string }
    | { kind: "custom" }
    | { kind: "clear" };

  let entries = $derived.by<Entry[]>(() => {
    const f = filter.trim().toLowerCase();
    const recentEntries: Entry[] = recents
      .filter((h) => !f || h.toLowerCase().includes(f))
      .map((hex) => ({ kind: "recent", hex }));
    const namedEntries: Entry[] = NAMED
      .filter((c) => !f || c.name.toLowerCase().includes(f) || c.hex.toLowerCase().includes(f))
      .map((c) => ({ kind: "named", name: c.name, hex: c.hex }));
    const tail: Entry[] = [];
    tail.push({ kind: "custom" });
    if (allowClear) tail.push({ kind: "clear" });
    return [...recentEntries, ...namedEntries, ...tail];
  });

  // Highlight (selection cursor) tracks the focused entry. Reset to
  // the first match whenever the filter changes — the user typed
  // something specific, so the front of the list is what they want
  // first.
  let highlight = $state(0);
  $effect(() => {
    filter;
    highlight = 0;
  });
  // Clamp when entries shrink past the current highlight.
  $effect(() => {
    if (highlight >= entries.length) {
      highlight = Math.max(0, entries.length - 1);
    }
  });
  // Remember the last colour-bearing entry so Custom… seeds from
  // wherever the user was looking before they clicked into custom
  // mode (rather than from the Custom… tile itself, which has no
  // colour).
  $effect(() => {
    const e = entries[highlight];
    if (e && (e.kind === "named" || e.kind === "recent")) {
      lastColorHex = e.hex;
    }
  });

  // ---------------------------------------------------------------------
  // Layout — the swatch grid is a fixed-column grid. We need the
  // column count to support left/right arrow nav. CSS keeps the
  // per-row width responsive, but we read it back via offsetWidth
  // for the keyboard math.
  // ---------------------------------------------------------------------
  let gridEl: HTMLDivElement | null = $state(null);
  let columns = $state(8);
  function recomputeColumns() {
    if (!gridEl) return;
    const styles = window.getComputedStyle(gridEl);
    const cols = styles.getPropertyValue("grid-template-columns").trim().split(/\s+/).length;
    columns = Math.max(1, cols || 8);
  }

  onMount(() => {
    if (initial && /^#[0-9A-Fa-f]{6}$/.test(initial)) {
      const [h, s, l] = hexToHsl(initial);
      customH = h; customS = s; customL = l;
      lastColorHex = initial.toUpperCase();
      // Pre-select the closest matching entry so the keyboard cursor
      // lands on a familiar swatch instead of the top-left tile.
      const idx = entries.findIndex(
        (e) => (e.kind === "recent" || e.kind === "named") && e.hex.toUpperCase() === initial.toUpperCase(),
      );
      if (idx >= 0) highlight = idx;
    }
    // Track the live column count via ResizeObserver on the grid
    // itself. `auto-fill` reflows on every dialog width change (the
    // .card uses min(640px, 92vw), so window resizing rewrites the
    // track count mid-frame). Without this, ArrowDown still uses the
    // initial column count after a resize and lands a few cells off
    // — looks like diagonal movement instead of straight down.
    tick().then(() => {
      inputEl?.focus();
      recomputeColumns();
      if (gridEl && typeof ResizeObserver !== "undefined") {
        gridObserver = new ResizeObserver(() => recomputeColumns());
        gridObserver.observe(gridEl);
      }
    });
    return () => {
      gridObserver?.disconnect();
      gridObserver = null;
    };
  });

  let gridObserver: ResizeObserver | null = null;

  // ---------------------------------------------------------------------
  // Commit / cancel
  // ---------------------------------------------------------------------
  function commitEntry(e: Entry) {
    if (e.kind === "named" || e.kind === "recent") {
      onSelect(e.hex.toUpperCase());
    } else if (e.kind === "custom") {
      enterCustom();
    } else if (e.kind === "clear") {
      if (onClear) onClear();
      else onCancel();
    }
  }
  function commitHighlight() {
    const e = entries[highlight];
    if (e) commitEntry(e);
  }
  function enterCustom() {
    // Seed HSL. Priority order:
    //   1. The currently-highlighted entry, if it carries a colour
    //      (e.g. user pressed Enter directly on a named swatch — but
    //      via mouse click on Custom… the highlight is on Custom
    //      itself, so this branch usually skips).
    //   2. The last colour-bearing entry the user navigated to (so
    //      "type 'blue' → arrow to CornflowerBlue → click Custom…"
    //      seeds at CornflowerBlue, not at the default red).
    //   3. The `initial` prop loaded onMount.
    //   4. Default 0/0.5/0.5.
    const e = entries[highlight];
    let seed = "";
    if (e && (e.kind === "named" || e.kind === "recent")) seed = e.hex;
    else if (lastColorHex) seed = lastColorHex;
    else if (initial && /^#[0-9A-Fa-f]{6}$/.test(initial)) seed = initial;
    if (seed) {
      const [h, s, l] = hexToHsl(seed);
      customH = h; customS = s; customL = l;
    }
    mode = "custom";
    // Refocus the card so onkeydown still fires — the input that had
    // focus is hidden by the mode switch and would otherwise drop
    // focus to body.
    tick().then(() => cardEl?.focus());
  }

  // ---------------------------------------------------------------------
  // Keyboard
  // ---------------------------------------------------------------------
  function onKey(e: KeyboardEvent) {
    if (mode === "custom") {
      // Custom mode: arrows nudge HSL, Enter commits, Esc returns to
      // list. The original swatches stay rendered (greyed, behind a
      // dimmer) so the user can compare while sliding.
      switch (e.key) {
        case "ArrowUp":
          e.preventDefault();
          customL = Math.min(1, customL + (e.shiftKey ? 0.005 : 0.02));
          return;
        case "ArrowDown":
          e.preventDefault();
          customL = Math.max(0, customL - (e.shiftKey ? 0.005 : 0.02));
          return;
        case "ArrowLeft":
          e.preventDefault();
          customS = Math.max(0, customS - (e.shiftKey ? 0.005 : 0.02));
          return;
        case "ArrowRight":
          e.preventDefault();
          customS = Math.min(1, customS + (e.shiftKey ? 0.005 : 0.02));
          return;
        case "[":
        case "{":
          // Hue nudge — not in the original spec but cheap to add and
          // useful when the named hue isn't quite right.
          e.preventDefault();
          customH = (customH - (e.shiftKey ? 0.5 : 5) + 360) % 360;
          return;
        case "]":
        case "}":
          e.preventDefault();
          customH = (customH + (e.shiftKey ? 0.5 : 5)) % 360;
          return;
        case "Enter":
          e.preventDefault();
          onSelect(hslToHex(customH, customS, customL));
          return;
        case "Escape":
          e.preventDefault();
          mode = "list";
          tick().then(() => inputEl?.focus());
          return;
      }
      return;
    }

    // List mode.
    switch (e.key) {
      case "ArrowDown":
        e.preventDefault();
        highlight = Math.min(entries.length - 1, highlight + columns);
        return;
      case "ArrowUp":
        e.preventDefault();
        highlight = Math.max(0, highlight - columns);
        return;
      case "ArrowRight":
        e.preventDefault();
        highlight = Math.min(entries.length - 1, highlight + 1);
        return;
      case "ArrowLeft":
        e.preventDefault();
        highlight = Math.max(0, highlight - 1);
        return;
      case "Home":
        e.preventDefault();
        highlight = 0;
        return;
      case "End":
        e.preventDefault();
        highlight = entries.length - 1;
        return;
      case "Enter":
        e.preventDefault();
        commitHighlight();
        return;
      case "Escape":
        e.preventDefault();
        onCancel();
        return;
    }
  }

  // Scroll the highlighted swatch into view when the cursor leaves
  // the visible window.
  $effect(() => {
    if (mode !== "list") return;
    highlight;
    tick().then(() => {
      const el = gridEl?.querySelector<HTMLElement>(`[data-idx="${highlight}"]`);
      if (el) el.scrollIntoView({ block: "nearest" });
    });
  });

  function entryHex(e: Entry): string | null {
    if (e.kind === "named" || e.kind === "recent") return e.hex;
    return null;
  }
  function entryLabel(e: Entry): string {
    if (e.kind === "named") return e.name;
    if (e.kind === "recent") return e.hex;
    if (e.kind === "custom") return "Custom…";
    return "Clear";
  }

  let customHex = $derived(hslToHex(customH, customS, customL));
</script>

<div
  class="overlay"
  role="dialog"
  aria-modal="true"
  aria-label={title}
  tabindex="-1"
  onkeydown={onKey}
>
  <div class="card" bind:this={cardEl} tabindex="-1">
    <header class="header">
      <h2>{title}</h2>
      <span class="hint">
        {#if mode === "custom"}
          ↑↓ lighter/darker · ←→ saturate · [ ] hue · Enter pick · Esc back
        {:else}
          type to filter · ←→↑↓ Enter pick · Esc cancel
        {/if}
      </span>
    </header>

    {#if mode === "list"}
      <div class="filter-row">
        <input
          class="filter-input"
          bind:this={inputEl}
          bind:value={filter}
          placeholder="Type a colour name (red, blue, sky, …)"
          spellcheck="false"
          autocomplete="off"
        />
      </div>
    {/if}

    <div class="grid-wrap">
      <div class="grid" class:dim={mode === "custom"} bind:this={gridEl}>
        {#each entries as e, i (i)}
          {@const hex = entryHex(e)}
          <button
            type="button"
            class="cell"
            class:hl={i === highlight && mode === "list"}
            class:special={e.kind === "custom" || e.kind === "clear"}
            data-idx={i}
            tabindex="-1"
            onclick={() => { highlight = i; if (mode === "list") commitEntry(e); }}
            onmouseenter={() => { if (mode === "list") highlight = i; }}
            title={entryLabel(e)}
          >
            {#if hex}
              <span class="swatch" style:background={hex}></span>
            {:else if e.kind === "custom"}
              <span class="swatch swatch-custom"></span>
            {:else}
              <span class="swatch swatch-clear"></span>
            {/if}
            <span class="label">{entryLabel(e)}</span>
          </button>
        {/each}
      </div>
    </div>

    {#if mode === "custom"}
      <div class="custom-bar">
        <span class="custom-swatch" style:background={customHex}></span>
        <div class="custom-meta">
          <span class="custom-hex">{customHex}</span>
          <span class="custom-hsl">
            H {Math.round(customH)}° · S {Math.round(customS * 100)}% · L {Math.round(customL * 100)}%
          </span>
        </div>
      </div>
    {/if}
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
    max-height: 80vh;
    display: flex;
    flex-direction: column;
    font-size: 13px;
    color: #222;
    outline: none;
  }
  .header {
    display: flex;
    align-items: baseline;
    justify-content: space-between;
    padding: 8px 12px;
    border-bottom: 1px solid #ddd;
    background: #f5f5f5;
    gap: 12px;
  }
  .header h2 {
    font-size: 14px;
    font-weight: 600;
    margin: 0;
  }
  .hint {
    font-size: 11px;
    color: #666;
  }
  .filter-row {
    padding: 8px 12px;
    border-bottom: 1px solid #eee;
  }
  .filter-input {
    width: 100%;
    padding: 5px 8px;
    border: 1px solid #ccc;
    border-radius: 2px;
    font: inherit;
  }
  .grid-wrap {
    flex: 1;
    overflow: auto;
    padding: 8px 12px;
  }
  .grid {
    display: grid;
    grid-template-columns: repeat(auto-fill, minmax(120px, 1fr));
    gap: 4px;
  }
  /* Dim the swatches in custom mode so the live preview is the
     visual anchor — but keep them legible for comparison. */
  .grid.dim {
    opacity: 0.55;
    pointer-events: none;
  }
  .cell {
    display: flex;
    align-items: center;
    gap: 6px;
    padding: 4px 6px;
    border: 1px solid transparent;
    background: #fff;
    cursor: pointer;
    text-align: left;
    font: inherit;
    color: inherit;
  }
  .cell:hover { background: #f0f0f0; }
  .cell.hl {
    border-color: #2c6cb0;
    background: #e6f0fa;
  }
  .cell.special {
    font-style: italic;
    color: #555;
  }
  .swatch {
    display: inline-block;
    width: 18px;
    height: 18px;
    border: 1px solid #888;
    flex: 0 0 auto;
  }
  /* Custom-mode preview: rainbow gradient hint. */
  .swatch-custom {
    background: linear-gradient(135deg, #ff5757, #ffd166, #06d6a0, #118ab2, #b5179e);
  }
  /* Clear: diagonal red strike on white, the universal "no fill" mark. */
  .swatch-clear {
    background:
      linear-gradient(45deg, transparent calc(50% - 1px), #c00 calc(50% - 1px), #c00 calc(50% + 1px), transparent calc(50% + 1px)),
      #fff;
  }
  .label {
    flex: 1;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
    font-variant-numeric: tabular-nums;
  }
  .custom-bar {
    display: flex;
    align-items: center;
    gap: 14px;
    padding: 10px 12px;
    border-top: 1px solid #ddd;
    background: #fafafa;
  }
  .custom-swatch {
    display: inline-block;
    width: 64px;
    height: 36px;
    border: 1px solid #888;
  }
  .custom-meta {
    display: flex;
    flex-direction: column;
  }
  .custom-hex {
    font-family: ui-monospace, monospace;
    font-size: 14px;
    font-weight: 600;
  }
  .custom-hsl {
    font-size: 11px;
    color: #666;
    font-variant-numeric: tabular-nums;
  }
</style>
