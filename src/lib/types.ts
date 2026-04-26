export type CellStyleView = {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  size_pt?: number;
  family?: string;
  color?: string;
  bg?: string;
  align_h?: "left" | "center" | "right" | "justify";
  align_v?: "top" | "middle" | "bottom";
  wrap?: boolean;
  border_top?: boolean;
  border_bottom?: boolean;
  border_left?: boolean;
  border_right?: boolean;
};

export type CellView = {
  row: number;
  col: number;
  text: string;
  input: string;
  is_formula: boolean;
  style?: CellStyleView;
};

export type LayoutData = {
  col_widths: [number, number][];
  row_heights: [number, number][];
  frozen_rows: number;
  frozen_cols: number;
  merged_ranges: string[];
};

export type WorkbookInfo = { sheet_names: string[]; active_sheet: number };

export type DirEntry = {
  name: string;
  is_dir: boolean;
  modified: number | null;
  size: number | null;
};

export type DirListing = {
  dir: string;
  parent: string | null;
  entries: DirEntry[];
};

export type RecentEntry = {
  path: string;
  basename: string;
  dir: string;
  hits: number;
  opened_at: number;
};

/// A unified entry shown in the navigator list — either a recent file
/// from the index (with `kind: "recent"`) or a directory entry from the
/// filesystem listing.
export type NavRow =
  | { kind: "recent"; recent: RecentEntry }
  | { kind: "entry"; entry: DirEntry };

export type MenuItem = {
  letter: string; // single uppercase char used for keyboard descent
  label: string;
  description: string; // shown on the description line under the menu
  children?: MenuItem[];
  action?: () => void | Promise<void>;
};

export type SaveResult = {
  path: string;
  mode: "preserved" | "ironcalc" | "xls";
  cells_patched: number;
};

export type BackupResult = { save: SaveResult; backup_path: string };
