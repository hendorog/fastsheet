use std::path::PathBuf;

use rusqlite::Connection;
use serde::Serialize;
use tauri::State;

use crate::navigator::home_dir;
use crate::state::AppState;

/// Per-OS data directory for fastsheet. Windows: %APPDATA%\fastsheet,
/// macOS: ~/Library/Application Support/fastsheet, Linux: ~/.local/share/fastsheet.
fn app_data_dir() -> Result<PathBuf, String> {
    #[cfg(windows)]
    {
        if let Ok(appdata) = std::env::var("APPDATA") {
            return Ok(PathBuf::from(appdata).join("fastsheet"));
        }
    }
    #[cfg(target_os = "macos")]
    {
        if let Some(home) = home_dir() {
            return Ok(home
                .join("Library")
                .join("Application Support")
                .join("fastsheet"));
        }
    }
    if let Ok(xdg) = std::env::var("XDG_DATA_HOME") {
        return Ok(PathBuf::from(xdg).join("fastsheet"));
    }
    if let Some(home) = home_dir() {
        return Ok(home.join(".local").join("share").join("fastsheet"));
    }
    Err("could not determine app data dir".into())
}

fn open_index_db(state: &State<'_, AppState>) -> Result<(), String> {
    let mut guard = state.index.lock().unwrap();
    if guard.is_some() {
        return Ok(());
    }
    let dir = app_data_dir()?;
    std::fs::create_dir_all(&dir).map_err(|e| e.to_string())?;
    let path = dir.join("index.db");
    let conn = Connection::open(&path).map_err(|e| e.to_string())?;
    conn.execute_batch(
        "CREATE TABLE IF NOT EXISTS files (
             path       TEXT PRIMARY KEY,
             basename   TEXT NOT NULL,
             dir        TEXT NOT NULL,
             opened_at  INTEGER NOT NULL,
             hits       INTEGER NOT NULL DEFAULT 1
         );
         CREATE INDEX IF NOT EXISTS files_basename_idx ON files(basename);",
    )
    .map_err(|e| e.to_string())?;
    *guard = Some(conn);
    Ok(())
}

pub(crate) fn record_open_internal(state: &State<'_, AppState>, path: &str) -> Result<(), String> {
    open_index_db(state)?;
    let guard = state.index.lock().unwrap();
    let conn = guard.as_ref().ok_or("index not open")?;
    let p = std::fs::canonicalize(path)
        .map(|p| p.to_string_lossy().into_owned())
        .unwrap_or_else(|_| path.to_string());
    let pb = PathBuf::from(&p);
    let basename = pb
        .file_name()
        .map(|s| s.to_string_lossy().into_owned())
        .unwrap_or_else(|| p.clone());
    let dir = pb
        .parent()
        .map(|s| s.to_string_lossy().into_owned())
        .unwrap_or_default();
    let now = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .map(|d| d.as_secs() as i64)
        .unwrap_or(0);
    conn.execute(
        "INSERT INTO files(path, basename, dir, opened_at, hits)
         VALUES (?1, ?2, ?3, ?4, 1)
         ON CONFLICT(path) DO UPDATE SET
             opened_at = excluded.opened_at,
             hits      = files.hits + 1",
        rusqlite::params![p, basename, dir, now],
    )
    .map_err(|e| e.to_string())?;
    Ok(())
}

#[derive(Serialize)]
pub(crate) struct RecentEntry {
    path: String,
    basename: String,
    dir: String,
    hits: i64,
    opened_at: i64,
}

#[tauri::command]
pub(crate) fn query_recents(
    query: String,
    limit: u32,
    state: State<'_, AppState>,
) -> Result<Vec<RecentEntry>, String> {
    open_index_db(&state)?;
    let guard = state.index.lock().unwrap();
    let conn = guard.as_ref().ok_or("index not open")?;
    let q = query.trim().to_lowercase();
    let mut stmt = if q.is_empty() {
        conn.prepare(
            "SELECT path, basename, dir, hits, opened_at
             FROM files
             ORDER BY hits DESC, opened_at DESC
             LIMIT ?1",
        )
        .map_err(|e| e.to_string())?
    } else {
        conn.prepare(
            "SELECT path, basename, dir, hits, opened_at
             FROM files
             WHERE LOWER(basename) LIKE ?1 OR LOWER(path) LIKE ?1
             ORDER BY hits DESC, opened_at DESC
             LIMIT ?2",
        )
        .map_err(|e| e.to_string())?
    };
    let rows: Vec<RecentEntry> = if q.is_empty() {
        stmt.query_map([limit as i64], |row| {
            Ok(RecentEntry {
                path: row.get(0)?,
                basename: row.get(1)?,
                dir: row.get(2)?,
                hits: row.get(3)?,
                opened_at: row.get(4)?,
            })
        })
        .map_err(|e| e.to_string())?
        .collect::<Result<_, _>>()
        .map_err(|e| e.to_string())?
    } else {
        let pattern = format!("%{q}%");
        stmt.query_map(rusqlite::params![pattern, limit as i64], |row| {
            Ok(RecentEntry {
                path: row.get(0)?,
                basename: row.get(1)?,
                dir: row.get(2)?,
                hits: row.get(3)?,
                opened_at: row.get(4)?,
            })
        })
        .map_err(|e| e.to_string())?
        .collect::<Result<_, _>>()
        .map_err(|e| e.to_string())?
    };
    Ok(rows)
}
