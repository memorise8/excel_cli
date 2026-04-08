pub mod calamine_read;
pub mod conditional;
pub mod export;
pub mod file;
pub mod format;
pub mod formula;
pub mod named_range;
pub mod range;
pub mod safe_io;
pub mod sheet;
pub mod table;

#[cfg(test)]
mod tests;

use crate::models::*;
use crate::services::ExcelService;
use std::path::Path;

/// Local file-based Excel service using umya-spreadsheet and calamine
pub struct LocalService;

impl LocalService {
    pub fn new() -> Self {
        Self
    }
}

impl Default for LocalService {
    fn default() -> Self {
        Self::new()
    }
}

impl ExcelService for LocalService {
    fn file_create(&self, path: &Path, sheets: Option<Vec<String>>) -> ExcelResult<WorkbookInfo> {
        file::create(path, sheets)
    }

    fn file_info(&self, path: &Path) -> ExcelResult<WorkbookInfo> {
        file::info(path)
    }

    fn file_save(&self, path: &Path, output: &Path) -> ExcelResult<()> {
        file::save(path, output)
    }

    fn sheet_list(&self, path: &Path) -> ExcelResult<Vec<SheetInfo>> {
        sheet::list(path)
    }

    fn sheet_add(&self, path: &Path, name: &str, position: Option<usize>) -> ExcelResult<SheetInfo> {
        sheet::add(path, name, position)
    }

    fn sheet_rename(&self, path: &Path, old_name: &str, new_name: &str) -> ExcelResult<()> {
        sheet::rename(path, old_name, new_name)
    }

    fn sheet_delete(&self, path: &Path, name: &str) -> ExcelResult<()> {
        sheet::delete(path, name)
    }

    fn sheet_copy(&self, path: &Path, name: &str, new_name: Option<&str>) -> ExcelResult<SheetInfo> {
        sheet::copy(path, name, new_name)
    }

    fn range_read(&self, path: &Path, range_str: &str) -> ExcelResult<RangeData> {
        range::read(path, range_str)
    }

    fn range_write(&self, path: &Path, range_str: &str, data: Vec<Vec<CellValue>>) -> ExcelResult<()> {
        range::write(path, range_str, data)
    }

    fn range_clear(&self, path: &Path, range_str: &str, values_only: bool) -> ExcelResult<()> {
        range::clear(path, range_str, values_only)
    }
}
