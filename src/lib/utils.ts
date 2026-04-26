import type { CellView } from "./types";

export function key(r: number, c: number): string {
  return `${r}:${c}`;
}

export function colLetter(col: number): string {
  let s = "";
  let n = col;
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

export function addr(r: number, c: number): string {
  return `${colLetter(c)}${r}`;
}

/// Parse "AB12" / "$AB$12" → { row, col }. Returns null for malformed input.
export function parseA1(s: string): { row: number; col: number } | null {
  const m = s.replace(/\$/g, "").match(/^([A-Za-z]+)(\d+)$/);
  if (!m) return null;
  let col = 0;
  for (const ch of m[1].toUpperCase()) col = col * 26 + (ch.charCodeAt(0) - 64);
  const row = parseInt(m[2], 10);
  if (col < 1 || row < 1) return null;
  return { row, col };
}

/// IronCalc returns widths as Excel-chars × 12 and heights as points × 2,
/// which renders wider/taller than Excel's own layout. Excel uses
/// Maximum Digit Width (MDW=7 px for Calibri 11) for columns and
/// 96/72 px per point for rows. Apply correction factors so what
/// fastsheet shows matches what Excel shows for the same file.
export function colWidthPx(ironcalcPx: number): number {
  return Math.round((ironcalcPx * 7) / 12);
}

export function rowHeightPx(ironcalcPx: number): number {
  return Math.round((ironcalcPx * (96 / 72)) / 2);
}

/// Auto-fit a column to its widest visible cell. Uses a cached canvas 2d
/// context to measure text width in the cell's font; returns the target
/// pixel width (with padding). Floor of 30 px so empty cols stay clickable.
let measureCtx: CanvasRenderingContext2D | null = null;
function getMeasureCtx(): CanvasRenderingContext2D | null {
  if (!measureCtx) {
    const c = document.createElement("canvas");
    measureCtx = c.getContext("2d");
  }
  return measureCtx;
}

export function autoFitColumnPx(
  cells: Map<string, CellView>,
  col: number,
): number {
  const ctx = getMeasureCtx();
  if (!ctx) return 73;
  let maxW = 30;
  for (const cv of cells.values()) {
    if (cv.col !== col || !cv.text) continue;
    const sz = cv.style?.size_pt ?? 11;
    const family = cv.style?.family ?? "Calibri";
    const weight = cv.style?.bold ? "700" : "400";
    const italic = cv.style?.italic ? "italic" : "normal";
    ctx.font = `${italic} ${weight} ${sz}pt ${family}, sans-serif`;
    const w = ctx.measureText(cv.text).width;
    if (w > maxW) maxW = w;
  }
  return Math.ceil(maxW) + 12;
}

/// Auto-fit a row to its tallest font. We don't lay out wrapped text, so
/// this is a heuristic: take the largest `size_pt` in the row, convert
/// pt → px (× 96/72) with a 1.2 line-height factor, plus padding.
export function autoFitRowPx(
  cells: Map<string, CellView>,
  row: number,
): number {
  let maxH = 19;
  for (const cv of cells.values()) {
    if (cv.row !== row) continue;
    const sz = cv.style?.size_pt ?? 11;
    const h = Math.ceil((sz * 1.2 * 96) / 72);
    if (h > maxH) maxH = h;
  }
  return maxH + 4;
}

/// Build the inline `style="..."` string for a cell from its CellStyleView.
/// Numeric cells default to right-aligned and string cells to left when
/// the style doesn't override.
export function cellStyle(cell: CellView | undefined): string {
  if (!cell) return "";
  const s = cell.style;
  const parts: string[] = [];
  if (s) {
    if (s.bold) parts.push("font-weight:700");
    if (s.italic) parts.push("font-style:italic");
    if (s.underline || s.strike) {
      const decos: string[] = [];
      if (s.underline) decos.push("underline");
      if (s.strike) decos.push("line-through");
      parts.push(`text-decoration:${decos.join(" ")}`);
    }
    if (s.size_pt) parts.push(`font-size:${s.size_pt}pt`);
    if (s.family)
      parts.push(
        `font-family:"${s.family.replace(/"/g, "")}", ui-monospace, monospace`,
      );
    if (s.color) parts.push(`color:${s.color}`);
    if (s.bg) parts.push(`background:${s.bg}`);
    if (s.align_h) parts.push(`text-align:${s.align_h}`);
    if (s.align_v) parts.push(`vertical-align:${s.align_v}`);
    if (s.wrap) parts.push("white-space:normal");
    if (s.border_top) parts.push("border-top:1px solid #000");
    if (s.border_bottom) parts.push("border-bottom:1px solid #000");
    if (s.border_left) parts.push("border-left:1px solid #000");
    if (s.border_right) parts.push("border-right:1px solid #000");
  }
  if (!s?.align_h) {
    const numeric =
      cell.text !== "" && !cell.is_formula && !isNaN(Number(cell.text));
    if (numeric || (cell.is_formula && !isNaN(Number(cell.text)))) {
      parts.push("text-align:right");
    }
  }
  return parts.join(";");
}
