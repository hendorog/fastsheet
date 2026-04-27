mod cells;
mod hidden;
mod index;
mod navigator;
mod state;
mod util;
mod workbook;
mod wsl;
pub mod xls_biff;
mod xls_load;
mod xls_save;
mod xlsx_load;
mod xlsx_save;

// Re-exports for the probe binary, which uses the same xlsx loader and
// hidden-col scraper as the GUI.
pub use hidden::extract_hidden_col_ranges;
pub use xls_load::load_xls;
pub use xls_save::save_xls;
pub use xlsx_load::{load_xlsx_with_fallback, replicate_my_array_formulas};

use state::AppState;

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
    tauri::Builder::default()
        .plugin(tauri_plugin_opener::init())
        .manage(AppState::new())
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
            workbook::open_workbook,
            workbook::new_workbook,
            workbook::save_workbook,
            workbook::file_exists,
            workbook::backup_and_save,
            workbook::add_sheet,
            workbook::add_sheet_named,
            workbook::rename_sheet,
            workbook::delete_sheet,
            workbook::workbook_info,
            cells::get_cells,
            cells::set_cell,
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
            cells::set_range_number_format,
            cells::set_range_style,
            cells::apply_style_indices,
            cells::define_name,
            cells::delete_name,
            cells::list_names,
            cells::recalc,
            cells::cell_addr,
            cells::jump_edge,
            navigator::start_dir,
            navigator::home_dir_path,
            navigator::list_dir,
            index::query_recents,
            hidden::debug_hidden_cols,
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
