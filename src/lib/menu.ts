import type { MenuItem } from "./types";

export type MenuCallbacks = {
  newWorkbook: () => void | Promise<void>;
  eraseCurrentCell: () => void | Promise<void>;
  openRetrieveNavigator: () => void | Promise<void>;
  fileSaveFlow: () => void | Promise<void>;
  quitApp: () => void | Promise<void>;
  setStatus: (msg: string) => void;
  // Column ops
  hideColumn: () => void | Promise<void>;
  showAllColumns: () => void | Promise<void>;
  setColumnWidth: () => void | Promise<void>;
  autoColumnWidth: () => void | Promise<void>;
  // Row ops
  hideRow: () => void | Promise<void>;
  showAllRows: () => void | Promise<void>;
  setRowHeight: () => void | Promise<void>;
  autoRowHeight: () => void | Promise<void>;
  // Titles (frozen panes). Each variant pulls the freeze line from the
  // current cursor position, like Lotus and Excel.
  setTitles: (kind: "both" | "horizontal" | "vertical" | "clear") => void | Promise<void>;
  // Structural ops — counts derive from the active selection size on
  // the matching axis (e.g. selecting rows 5-9 inserts 5 rows).
  insertRows: () => void | Promise<void>;
  deleteRows: () => void | Promise<void>;
  insertColumns: () => void | Promise<void>;
  deleteColumns: () => void | Promise<void>;
  // Range/Format. Variants that take decimals raise the inline prompt;
  // others apply directly to the current selection.
  formatRange: (kind: FormatKind) => void | Promise<void>;
  // Range/Search — opens the find-then-replace inline prompt chain.
  searchRange: () => void | Promise<void>;
  // Range/Label — text alignment.
  alignRange: (h: "left" | "center" | "right" | "general") => void | Promise<void>;
  // Range/Format/Color — fill / text color (prompted hex).
  setFillColor: () => void | Promise<void>;
  setTextColor: () => void | Promise<void>;
  // Range/Value — convert formulas in selection to their literal values.
  rangeValue: () => void | Promise<void>;
  // Range/Trans — transpose selection in place.
  rangeTrans: () => void | Promise<void>;
  // /Copy and /Move — Lotus convention: source = current selection,
  // prompt for destination top-left, paste (and clear source for /Move).
  copyRange: () => void | Promise<void>;
  moveRange: () => void | Promise<void>;
  // /Data/Fill — fill selection with arithmetic progression.
  dataFill: () => void | Promise<void>;
  // /Data/Sort — sort selection rows by column.
  dataSort: () => void | Promise<void>;
  // /Range/Name — define / delete / list named ranges.
  nameCreate: () => void | Promise<void>;
  nameDelete: () => void | Promise<void>;
  nameList: () => void | Promise<void>;
  // /Range/Format/Border — apply thin black borders to the selection.
  setBorder: (sides: "all" | "outline" | "top" | "bottom" | "left" | "right" | "none") => void | Promise<void>;
  // /Worksheet/Sheet — sheet management mirrors the tab-bar context menu.
  sheetNew: () => void | Promise<void>;
  sheetDelete: () => void | Promise<void>;
  sheetRename: () => void | Promise<void>;
  // /Trace — formula dependency tools.
  traceFormula: () => void | Promise<void>;
  traceGoto: () => void | Promise<void>;
  traceNames: () => void | Promise<void>;
};

export type FormatKind =
  | "general"
  | "fixed"     // 0..15 decimals, prompted
  | "currency"  // 0..15 decimals, prompted
  | "percent"   // 0..15 decimals, prompted
  | "comma"     // 0..15 decimals, prompted (thousands sep)
  | "date"
  | "time";

function stub(setStatus: (s: string) => void, path: string): () => void {
  return () => setStatus(`Not yet implemented: /${path}`);
}

/// Lotus 1-2-3 menu tree. Letters match the first capital letter of each
/// label (Lotus convention). Some destructive actions have Yes/No confirms.
export function buildMenu(cb: MenuCallbacks): MenuItem[] {
  const stb = (p: string) => stub(cb.setStatus, p);
  return [
    {
      letter: "W",
      label: "Worksheet",
      description:
        "Global settings, columns, rows, titles, windows, page, status",
      children: [
        { letter: "G", label: "Global", description: "Worksheet global settings", action: stb("Worksheet/Global") },
        {
          letter: "I", label: "Insert", description: "Insert columns or rows",
          children: [
            { letter: "C", label: "Column", description: "Insert columns at the cursor (selection width = count)", action: cb.insertColumns },
            { letter: "R", label: "Row", description: "Insert rows at the cursor (selection height = count)", action: cb.insertRows },
          ],
        },
        {
          letter: "D", label: "Delete", description: "Delete columns or rows",
          children: [
            { letter: "C", label: "Column", description: "Delete the selected columns", action: cb.deleteColumns },
            { letter: "R", label: "Row", description: "Delete the selected rows", action: cb.deleteRows },
          ],
        },
        {
          letter: "C", label: "Column",
          description: "Column width, hide, display",
          children: [
            { letter: "S", label: "Set-Width", description: "Pick columns then enter width (px or auto)", action: cb.setColumnWidth },
            { letter: "A", label: "Auto", description: "Pick columns then auto-fit to widest cell", action: cb.autoColumnWidth },
            { letter: "H", label: "Hide", description: "Hide the current column", action: cb.hideColumn },
            { letter: "D", label: "Display", description: "Show every hidden column on this sheet", action: cb.showAllColumns },
          ],
        },
        {
          letter: "R", label: "Row",
          description: "Row height, hide, display",
          children: [
            { letter: "S", label: "Set-Height", description: "Pick rows then enter height (px or auto)", action: cb.setRowHeight },
            { letter: "A", label: "Auto", description: "Pick rows then auto-fit to tallest font", action: cb.autoRowHeight },
            { letter: "H", label: "Hide", description: "Hide the current row", action: cb.hideRow },
            { letter: "D", label: "Display", description: "Show every hidden row on this sheet", action: cb.showAllRows },
          ],
        },
        {
          letter: "E", label: "Erase",
          description: "Erase the entire worksheet",
          children: [
            { letter: "N", label: "No", description: "Do not erase", action: () => cb.setStatus("Erase cancelled") },
            { letter: "Y", label: "Yes", description: "Erase the entire worksheet", action: cb.newWorkbook },
          ],
        },
        {
          letter: "T", label: "Titles",
          description: "Freeze rows above / cols left of the cursor as titles",
          children: [
            { letter: "B", label: "Both", description: "Freeze both rows above and cols left of the cursor", action: () => cb.setTitles("both") },
            { letter: "H", label: "Horizontal", description: "Freeze rows above the cursor", action: () => cb.setTitles("horizontal") },
            { letter: "V", label: "Vertical", description: "Freeze cols left of the cursor", action: () => cb.setTitles("vertical") },
            { letter: "C", label: "Clear", description: "Unfreeze all titles", action: () => cb.setTitles("clear") },
          ],
        },
        { letter: "W", label: "Window", description: "Split or unsplit the window", action: stb("Worksheet/Window") },
        {
          letter: "S", label: "Sheet",
          description: "Add / delete / rename worksheets",
          children: [
            { letter: "N", label: "New", description: "Append a new sheet", action: cb.sheetNew },
            { letter: "D", label: "Delete", description: "Delete the current sheet (with confirm)", action: cb.sheetDelete },
            { letter: "R", label: "Rename", description: "Rename the current sheet", action: cb.sheetRename },
          ],
        },
        { letter: "P", label: "Page", description: "Page break settings", action: stb("Worksheet/Page") },
      ],
    },
    {
      letter: "R", label: "Range",
      description: "Format, name, erase, fill or search a range of cells",
      children: [
        {
          letter: "F", label: "Format",
          description: "Set the display format of the selected range",
          children: [
            { letter: "F", label: "Fixed", description: "Fixed decimal (e.g. 0.00)", action: () => cb.formatRange("fixed") },
            { letter: "C", label: "Currency", description: "Currency with $ prefix and N decimals", action: () => cb.formatRange("currency") },
            { letter: ",", label: ",", description: "Comma (thousands separator) with N decimals", action: () => cb.formatRange("comma") },
            { letter: "P", label: "Percent", description: "Percent with N decimals", action: () => cb.formatRange("percent") },
            { letter: "D", label: "Date", description: "yyyy-mm-dd", action: () => cb.formatRange("date") },
            { letter: "T", label: "Time", description: "h:mm:ss", action: () => cb.formatRange("time") },
            { letter: "G", label: "General", description: "Reset to General format", action: () => cb.formatRange("general") },
            { letter: "B", label: "Background", description: "Fill colour (#RRGGBB or empty to clear)", action: cb.setFillColor },
            { letter: "X", label: "Text", description: "Text colour (#RRGGBB or empty to clear)", action: cb.setTextColor },
            {
              letter: "R", label: "Border",
              description: "Thin black borders around the selection",
              children: [
                { letter: "A", label: "All", description: "Border on every side of every cell", action: () => cb.setBorder("all") },
                { letter: "O", label: "Outline", description: "Border only on the outer edges of the selection", action: () => cb.setBorder("outline") },
                { letter: "T", label: "Top", description: "Top side only", action: () => cb.setBorder("top") },
                { letter: "B", label: "Bottom", description: "Bottom side only", action: () => cb.setBorder("bottom") },
                { letter: "L", label: "Left", description: "Left side only", action: () => cb.setBorder("left") },
                { letter: "R", label: "Right", description: "Right side only", action: () => cb.setBorder("right") },
                { letter: "N", label: "None", description: "Remove all borders", action: () => cb.setBorder("none") },
              ],
            },
          ],
        },
        {
          letter: "L", label: "Label",
          description: "Set text alignment for the selected range",
          children: [
            { letter: "L", label: "Left", description: "Align left", action: () => cb.alignRange("left") },
            { letter: "C", label: "Center", description: "Center", action: () => cb.alignRange("center") },
            { letter: "R", label: "Right", description: "Align right", action: () => cb.alignRange("right") },
            { letter: "G", label: "General", description: "Reset to general (numbers right, text left)", action: () => cb.alignRange("general") },
          ],
        },
        { letter: "E", label: "Erase", description: "Erase the contents of the current cell", action: cb.eraseCurrentCell },
        {
          letter: "N", label: "Name",
          description: "Create / delete / list named ranges",
          children: [
            { letter: "C", label: "Create", description: "Define a name for the current selection", action: cb.nameCreate },
            { letter: "D", label: "Delete", description: "Delete a defined name", action: cb.nameDelete },
            { letter: "L", label: "List", description: "List all defined names in the status bar", action: cb.nameList },
          ],
        },
        { letter: "J", label: "Justify", description: "Justify a range of labels", action: stb("Range/Justify") },
        { letter: "P", label: "Prot", description: "Protect a range from changes", action: stb("Range/Prot") },
        { letter: "U", label: "Unprot", description: "Unprotect a range", action: stb("Range/Unprot") },
        { letter: "I", label: "Input", description: "Restrict input to unprotected cells", action: stb("Range/Input") },
        { letter: "V", label: "Value", description: "Convert formulas in the selection to their literal values", action: cb.rangeValue },
        { letter: "T", label: "Trans", description: "Transpose the selection (rows ↔ cols, in place)", action: cb.rangeTrans },
        { letter: "S", label: "Search", description: "Find and (optionally) replace within the active sheet", action: cb.searchRange },
      ],
    },
    { letter: "C", label: "Copy", description: "Copy the selected range to another location", action: cb.copyRange },
    { letter: "M", label: "Move", description: "Move the selected range to another location", action: cb.moveRange },
    {
      letter: "F", label: "File",
      description: "Retrieve, save, combine, list, import, or change directory",
      children: [
        { letter: "R", label: "Retrieve", description: "Retrieve (open) a worksheet file from disk", action: cb.openRetrieveNavigator },
        { letter: "S", label: "Save", description: "Save the current worksheet", action: cb.fileSaveFlow },
        { letter: "C", label: "Combine", description: "Combine another file into the current one", action: stb("File/Combine") },
        { letter: "X", label: "Xtract", description: "Extract part of the worksheet to a new file", action: stb("File/Xtract") },
        { letter: "E", label: "Erase", description: "Erase a file from disk", action: stb("File/Erase") },
        { letter: "L", label: "List", description: "List worksheet files in the directory", action: stb("File/List") },
        { letter: "I", label: "Import", description: "Import a text file as cells", action: stb("File/Import") },
        { letter: "D", label: "Directory", description: "Change the current directory", action: stb("File/Directory") },
        { letter: "A", label: "Admin", description: "File admin operations", action: stb("File/Admin") },
      ],
    },
    {
      letter: "P", label: "Print",
      description: "Send the worksheet to a printer or a file",
      children: [
        { letter: "P", label: "Printer", description: "Print to the system printer", action: stb("Print/Printer") },
        { letter: "F", label: "File", description: "Print to a file", action: stb("Print/File") },
        { letter: "B", label: "Background", description: "Print in the background", action: stb("Print/Background") },
        { letter: "E", label: "Encoded", description: "Print encoded for typesetters", action: stb("Print/Encoded") },
      ],
    },
    {
      letter: "G", label: "Graph",
      description: "Create or modify graphs from worksheet ranges",
      children: [
        { letter: "T", label: "Type", description: "Choose a graph type", action: stb("Graph/Type") },
        { letter: "X", label: "X", description: "Set the X-axis range", action: stb("Graph/X") },
        { letter: "A", label: "A", description: "Set the A range", action: stb("Graph/A") },
        { letter: "R", label: "Reset", description: "Reset graph settings", action: stb("Graph/Reset") },
        { letter: "V", label: "View", description: "View the current graph", action: stb("Graph/View") },
        { letter: "S", label: "Save", description: "Save the graph to a file", action: stb("Graph/Save") },
        { letter: "O", label: "Options", description: "Graph options", action: stb("Graph/Options") },
        { letter: "N", label: "Name", description: "Manage named graphs", action: stb("Graph/Name") },
      ],
    },
    {
      letter: "D", label: "Data",
      description: "Fill, sort, query, distribute, regress or parse data",
      children: [
        { letter: "F", label: "Fill", description: "Fill the selection with an arithmetic progression", action: cb.dataFill },
        { letter: "T", label: "Table", description: "What-if data tables", action: stb("Data/Table") },
        { letter: "S", label: "Sort", description: "Sort the selected rows by a column", action: cb.dataSort },
        { letter: "Q", label: "Query", description: "Query a database range", action: stb("Data/Query") },
        { letter: "D", label: "Distribution", description: "Frequency distribution", action: stb("Data/Distribution") },
        { letter: "M", label: "Matrix", description: "Matrix operations", action: stb("Data/Matrix") },
        { letter: "R", label: "Regression", description: "Linear regression", action: stb("Data/Regression") },
        { letter: "P", label: "Parse", description: "Parse a column of labels into cells", action: stb("Data/Parse") },
      ],
    },
    { letter: "S", label: "System", description: "Temporarily exit to operating system", action: stb("System") },
    {
      letter: "T", label: "Trace",
      description: "Formula dependency tools — trace, jump to dependency, browse named ranges",
      children: [
        {
          letter: "T", label: "Trace",
          description: "Show the dependency chain of the current cell's formula in a popup",
          action: cb.traceFormula,
        },
        {
          letter: "G", label: "Goto",
          description: "Jump to a top-level dependency of the current cell's formula",
          action: cb.traceGoto,
        },
        {
          letter: "N", label: "Names",
          description: "Browse the workbook's named ranges and jump to one",
          action: cb.traceNames,
        },
      ],
    },
    {
      letter: "Q", label: "Quit",
      description: "End the spreadsheet session",
      children: [
        { letter: "N", label: "No", description: "Continue working", action: () => cb.setStatus("Quit cancelled") },
        { letter: "Y", label: "Yes", description: "Quit fastsheet", action: cb.quitApp },
      ],
    },
  ];
}

export type SaveMenuCallbacks = {
  replace: () => void | Promise<void>;
  saveAs: () => void | Promise<void>;
  backup: () => void | Promise<void>;
  cancel: () => void;
};

/// Lotus /F S picker shown when saving to an existing file.
export function saveMenuItems(cb: SaveMenuCallbacks): MenuItem[] {
  return [
    {
      letter: "R",
      label: "Replace",
      description: "Overwrite the existing file (⚠ unsupported features lost)",
      action: cb.replace,
    },
    {
      letter: "A",
      label: "Save As",
      description: "Save to a different filename via the file navigator",
      action: cb.saveAs,
    },
    {
      letter: "B",
      label: "Backup",
      description: "Rename existing to .bak then save",
      action: cb.backup,
    },
    {
      letter: "C",
      label: "Cancel",
      description: "Don't save",
      action: cb.cancel,
    },
  ];
}

/// Walk the menu tree using the current descent path to find the level
/// being rendered. If `dynamicLevel` is set, it takes precedence (ad-hoc
/// menus like the save picker have no parent in the static tree).
export function currentLevel(
  menu: MenuItem[],
  path: number[],
  dynamicLevel: MenuItem[] | null,
): MenuItem[] {
  if (dynamicLevel) return dynamicLevel;
  let level = menu;
  for (const i of path) {
    const item = level[i];
    if (!item?.children) return [];
    level = item.children;
  }
  return level;
}

export function breadcrumb(
  menu: MenuItem[],
  path: number[],
  dynamicLevel: MenuItem[] | null,
  dynamicTitle: string,
): string {
  if (dynamicLevel) return dynamicTitle || "/";
  const parts: string[] = [];
  let level = menu;
  for (const i of path) {
    const item = level[i];
    if (!item) break;
    parts.push(item.label);
    level = item.children ?? [];
  }
  return parts.length ? "/" + parts.join("/") : "/";
}
