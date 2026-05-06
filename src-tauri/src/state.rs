use std::collections::{HashMap, HashSet};
use std::sync::Mutex;

use ironcalc::base::Model;
use rusqlite::Connection;

/// Snapshot of the source .xlsx kept in memory so we can patch it in-place
/// on save and preserve features IronCalc doesn't understand
/// (charts/pivots/drawings/comments/conditional formatting).
#[derive(Clone)]
pub(crate) struct LoadedFile {
    pub(crate) path: String,
    pub(crate) bytes: Vec<u8>,
    /// sheet_idx → zip entry path (e.g. "xl/worksheets/sheet1.xml")
    pub(crate) sheet_paths: Vec<String>,
}

pub(crate) struct AppState {
    pub(crate) model: Mutex<Option<Model<'static>>>,
    /// Lazily-opened SQLite handle for the file index. None until the first
    /// command needs it (so app startup doesn't pay the cost on workbooks
    /// that never open the navigator).
    pub(crate) index: Mutex<Option<Connection>>,
    /// In-memory snapshot of the loaded file's bytes + sheet path mapping.
    /// None when no file is open (e.g. /Worksheet/Erase produced a blank).
    pub(crate) loaded: Mutex<Option<LoadedFile>>,
    /// Cells the user has changed since the last save. Key: (sheet_idx, row, col).
    /// Value: the user-typed input string ("=SUM(A1:A5)", "42", "hello", "" for clear).
    pub(crate) dirty: Mutex<HashMap<(u32, i32, i32), String>>,
    /// Authoritative hidden-column state, sheet_idx → set of 1-based col
    /// indices. IronCalc's `Col` struct has no `hidden` field, so we
    /// shadow it: populated from the original xlsx's `<col hidden="1">`
    /// markers on load, mutated by set_column_hidden, queried by get_layout.
    pub(crate) hidden_cols: Mutex<HashMap<u32, HashSet<i32>>>,
    /// Sheets with cell-style changes (number format, font, fill, etc.)
    /// that have NOT been persisted by the in-place save_preserving path.
    /// Non-empty → save_workbook routes to save_to_xlsx so styles flow
    /// through IronCalc's full serialiser. Trade-off: charts / pivots /
    /// drawings get dropped by save_to_xlsx — files without those keep
    /// styles correctly; files with them choose styles over preservation.
    pub(crate) style_dirty: Mutex<HashSet<u32>>,
    /// True when an insert/delete row/col, sheet add/delete, or other
    /// structural edit has run since load. The xlsx preservation path
    /// patches sheet XML by absolute (row, col) coordinates and would
    /// silently desync from the underlying data on a structural shift,
    /// so we route past it through `save_to_xlsx` (which loses
    /// unsupported features but keeps cell coordinates correct). Cleared
    /// on save, new_workbook, and successful open.
    pub(crate) structural_dirty: Mutex<bool>,
    /// True when the open workbook has user-visible changes that have
    /// not been saved. This is deliberately separate from `dirty`,
    /// `style_dirty`, `structural_dirty`, and manual recalc-pending UI:
    /// those drive save strategy or stale formula display, while this
    /// drives data-loss prompts and the title/status dirty marker.
    pub(crate) workbook_dirty: Mutex<bool>,
    /// Recalculation mode (Lotus 1-2-3 `/W G R` setting). When `true`,
    /// every successful `set_cell` triggers `model.evaluate()` so
    /// formula cells transition out of the un-evaluated `CellFormula`
    /// variant (which displays as `#ERROR!`) into a real
    /// `CellFormulaNumber` / `CellFormulaString` / etc. with a cached
    /// value. When `false`, only F9 (or `recalc`) evaluates — useful
    /// for very large workbooks where each evaluate is multi-second.
    /// Defaults to `true` to match Excel + Lotus's automatic mode.
    pub(crate) auto_recalc: Mutex<bool>,
    /// VBA / macro storages captured from the source .xls on load. The
    /// .xls writer (`save_xls`) replays these into the new compound
    /// file so macros survive a save+reload through Excel. Cleared on
    /// new_workbook and on any non-.xls open. Set to None when the
    /// source had no macros — most files.
    pub(crate) xls_preserved: Mutex<Option<crate::xls_preserve::PreservedXlsData>>,
    /// Active compare session: a right-side workbook loaded for diff
    /// purposes. Held in state so the trace command can enrich each
    /// node with the right-side value, and so the GUI can survive
    /// arbitrary edits without re-loading the right model. Cleared
    /// on `compare_close`, on `new_workbook`, and on any `open_workbook`
    /// (the new workbook becomes the new "left" — a stale comparison
    /// against a different file would just confuse the user).
    pub(crate) compare: Mutex<Option<crate::compare::CompareSession>>,
    /// File path passed on the command line (e.g. when Windows
    /// Explorer launches fastsheet via "Open with"). Captured once
    /// in `run()` from `std::env::args` and consumed by the frontend
    /// on mount via `take_startup_path` — taking it clears the slot
    /// so a hot reload doesn't reopen the file.
    pub(crate) startup_path: Mutex<Option<String>>,
}

impl AppState {
    pub(crate) fn new() -> Self {
        Self {
            model: Mutex::new(None),
            index: Mutex::new(None),
            loaded: Mutex::new(None),
            dirty: Mutex::new(HashMap::new()),
            hidden_cols: Mutex::new(HashMap::new()),
            style_dirty: Mutex::new(HashSet::new()),
            structural_dirty: Mutex::new(false),
            workbook_dirty: Mutex::new(false),
            auto_recalc: Mutex::new(true),
            xls_preserved: Mutex::new(None),
            compare: Mutex::new(None),
            startup_path: Mutex::new(None),
        }
    }
}
