use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct WorkbookInfo {
    pub path: String,
    pub file_name: String,
    pub file_size: u64,
    pub sheet_count: usize,
    pub sheets: Vec<SheetInfo>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SheetInfo {
    pub name: String,
    pub index: usize,
    pub visible: bool,
    pub color: Option<String>,
    pub row_count: Option<usize>,
    pub col_count: Option<usize>,
}
