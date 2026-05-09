mod atomic;
mod cells;
mod data_analysis;
pub mod compare;
mod hidden;
mod index;
mod navigator;
mod state;
pub mod trace;
mod util;
mod workbook;
mod wsl;
pub mod xls_biff;
pub mod xls_load;
pub mod xls_preserve;
pub mod xls_save;
mod xlsx_load;
mod xlsx_save;

// Re-exports for the probe binary, which uses the same xlsx loader and
// hidden-col scraper as the GUI.
pub use hidden::{extract_default_row_height, extract_hidden_col_ranges};
pub use xls_load::load_xls;
pub use xls_save::{save_xls, save_xls_with_preserved};
pub use xlsx_load::{load_xlsx_with_fallback, replicate_my_array_formulas};

use state::AppState;

/// Pluck the first non-flag argv entry — Windows passes the file
/// path as a single positional arg when launching via "Open with"
/// or shell association. We intentionally ignore anything starting
/// with `-` so future CLI flags don't accidentally get treated as
/// paths. Returns None when nothing was passed.
fn capture_startup_path() -> Option<String> {
    std::env::args()
        .skip(1)
        .find(|a| !a.starts_with('-'))
        .filter(|a| !a.is_empty())
}

#[tauri::command]
fn take_startup_path(state: tauri::State<'_, AppState>) -> Option<String> {
    state.startup_path.lock().unwrap().take()
}

/// Frontend-callable: log a timestamp + label to the profile log,
/// reporting elapsed-since-process-start so we can measure boot
/// latency from launch to first interactive frame. Active only when
/// FASTSHEET_PROFILE_LOAD is set; otherwise a no-op.
#[tauri::command]
fn profile_mark(label: String) {
    let elapsed_ms = util::app_start_instant().elapsed().as_secs_f64() * 1000.0;
    util::profile_log(&format!("[boot] {:>20} {:>7.1}ms (since process start)", label, elapsed_ms));
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    let _ = util::app_start_instant();
    util::profile_log("[boot] === process_start");
    wsl::apply_wsl_webkit_workaround();
    util::profile_log(&format!(
        "[boot] {:>20} {:>7.1}ms (since process start)",
        "wsl_workaround",
        util::app_start_instant().elapsed().as_secs_f64() * 1000.0
    ));
    let startup = capture_startup_path();
    let app_state = AppState::new();
    if let Some(p) = startup {
        *app_state.startup_path.lock().unwrap() = Some(p);
    }
    tauri::Builder::default()
        .plugin(tauri_plugin_opener::init())
        .manage(app_state)
        .setup(|_app| {
            let elapsed_ms = util::app_start_instant().elapsed().as_secs_f64() * 1000.0;
            util::profile_log(&format!(
                "[boot] {:>20} {:>7.1}ms (since process start)",
                "tauri_setup",
                elapsed_ms
            ));
            Ok(())
        })
        .invoke_handler(tauri::generate_handler![
            profile_mark,
            take_startup_path,
            workbook::open_workbook,
            workbook::new_workbook,
            workbook::save_workbook,
            workbook::extract_cells_to_workbook,
            workbook::read_workbook_first_sheet,
            workbook::workbook_has_unsaved_changes,
            workbook::file_exists,
            workbook::erase_file,
            workbook::read_text_file,
            workbook::backup_and_save,
            workbook::compare_open,
            workbook::compare_close,
            workbook::compare_value_at,
            workbook::add_sheet,
            workbook::add_sheet_named,
            workbook::rename_sheet,
            workbook::delete_sheet,
            workbook::workbook_info,
            cells::get_cells,
            cells::set_cell,
            cells::protect_range,
            cells::unprotect_range,
            cells::restrict_input_range,
            cells::clear_input_restriction,
            cells::set_show_grid_lines,
            cells::get_layout,
            cells::get_sheet_dim,
            cells::get_used_range,
            cells::set_row_hidden,
            cells::set_column_hidden,
            cells::show_all_rows,
            cells::show_all_cols,
            cells::set_frozen_panes,
            cells::set_row_height,
            cells::set_column_width,
            cells::insert_rows,
            cells::delete_rows,
            cells::insert_columns,
            cells::delete_columns,
            cells::insert_cells_shift_right,
            cells::insert_cells_shift_down,
            cells::delete_cells_shift_left,
            cells::delete_cells_shift_up,
            cells::merge_cells,
            cells::unmerge_cells,
            cells::set_range_number_format,
            cells::set_range_style,
            cells::apply_style_indices,
            cells::get_cell_format,
            cells::list_workbook_colors,
            cells::define_name,
            cells::delete_name,
            cells::list_names,
            cells::recalc,
            cells::get_auto_recalc,
            cells::set_auto_recalc,
            cells::cell_addr,
            cells::jump_edge,
            cells::trace_formula,
            cells::list_named_ranges,
            data_analysis::data_summary,
            data_analysis::data_filter,
            data_analysis::data_distribution,
            data_analysis::data_regression,
            data_analysis::data_parse,
            navigator::start_dir,
            navigator::home_dir_path,
            navigator::list_dir,
            index::query_recents,
            index::query_recent_dirs,
            hidden::debug_hidden_cols,
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
