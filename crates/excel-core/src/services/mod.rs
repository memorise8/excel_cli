pub mod graph;
pub mod local;

use crate::models::*;
use std::path::Path;

/// Core trait for Excel operations — implemented by Local and Graph backends
pub trait ExcelService {
    // File operations
    fn file_create(&self, path: &Path, sheets: Option<Vec<String>>) -> ExcelResult<WorkbookInfo>;
    fn file_info(&self, path: &Path) -> ExcelResult<WorkbookInfo>;
    fn file_save(&self, path: &Path, output: &Path) -> ExcelResult<()>;

    // Sheet operations
    fn sheet_list(&self, path: &Path) -> ExcelResult<Vec<SheetInfo>>;
    fn sheet_add(&self, path: &Path, name: &str, position: Option<usize>) -> ExcelResult<SheetInfo>;
    fn sheet_rename(&self, path: &Path, old_name: &str, new_name: &str) -> ExcelResult<()>;
    fn sheet_delete(&self, path: &Path, name: &str) -> ExcelResult<()>;
    fn sheet_copy(&self, path: &Path, name: &str, new_name: Option<&str>) -> ExcelResult<SheetInfo>;

    // Range operations
    fn range_read(&self, path: &Path, range: &str) -> ExcelResult<RangeData>;
    fn range_write(&self, path: &Path, range: &str, data: Vec<Vec<CellValue>>) -> ExcelResult<()>;
    fn range_clear(&self, path: &Path, range: &str, values_only: bool) -> ExcelResult<()>;
}
