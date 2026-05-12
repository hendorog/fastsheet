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
  show_grid_lines: boolean;
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

export type RecentDir = {
  dir: string;
  opened_at: number;
};

/// A unified entry shown in the navigator list — a recent file from
/// the index (`kind: "recent"`), a recent directory (`kind: "recent_dir"`),
/// or a directory entry from the current filesystem listing
/// (`kind: "entry"`). Both recent kinds are hidden after the user
/// navigates away from the start dir; entries are always shown.
export type NavRow =
  | { kind: "recent"; recent: RecentEntry }
  | { kind: "recent_dir"; recent_dir: RecentDir }
  | { kind: "entry"; entry: DirEntry }
  /// Synthetic row only used in fileKind="directory" pickers — Enter
  /// commits the navigator's current listing.dir via onSelectDir.
  | { kind: "select_current"; dir: string };

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
  /// Set when the save would have lost features that existed in the
  /// file being overwritten. Names the .bak (or .bak.N) copy the
  /// backend made before the save.
  backup_path?: string;
  /// True when the .xls writer round-tripped VBA / macro storages
  /// from the source. Lets the UI swap "macros not preserved" for
  /// "macros preserved" on the post-save status line.
  vba_preserved?: boolean;
};

export type WorkbookRange = {
  sheet_name: string;
  rows: string[][];
  source_rows: number;
  source_cols: number;
  cells_read: number;
};

export type BackupResult = { save: SaveResult; backup_path: string };

export type TraceNode = {
  address: string;
  kind: "cell" | "range" | "name" | "literal" | "error";
  sheet: number | null;
  row: number | null;
  col: number | null;
  formula: string | null;
  value: string;
  note: string | null;
  is_error: boolean;
  cycle: boolean;
  truncated: boolean;
  /// Right-side formatted value when a compare session is active.
  /// null otherwise.
  compare_value: string | null;
  /// True iff `compare_value` differs from `value`.
  compare_differs: boolean;
  deps: TraceNode[];
};

export type CompareDiff = {
  sheet: string;
  sheet_idx: number | null;
  row: number;
  col: number;
  address: string;
  left_value: string;
  right_value: string;
  left_formula: string | null;
  right_formula: string | null;
  kind: "value" | "formula" | "missing-left" | "missing-right";
  /// Filter bucket:
  ///   "formula" — formula text differs (Some/None counts as differ)
  ///   "value"   — neither side has a formula; literals differ
  ///   "other"   — same formula text on both sides, value differs
  category: "value" | "formula" | "other";
};

export type CompareSheetMissing = {
  sheet: string;
  side: "left" | "right";
};

export type CompareResult = {
  right_path: string;
  diffs: CompareDiff[];
  missing_sheets: CompareSheetMissing[];
  total_diffs: number;
  diffs_capped: boolean;
};

export type NamedRangeInfo = {
  name: string;
  formula: string;
  scope: string;
  jump_sheet: number | null;
  jump_row: number | null;
  jump_col: number | null;
};
