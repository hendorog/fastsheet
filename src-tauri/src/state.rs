use std::collections::{HashMap, HashSet};
use std::sync::Mutex;

use ironcalc::base::Model;
use rusqlite::Connection;

/// Snapshot of the source .xlsx kept in memory so we can patch it in-place
/// on save and preserve features IronCalc doesn't understand
/// (charts/pivots/drawings/comments/conditional formatting).
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
        }
    }
}
