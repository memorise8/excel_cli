use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SheetDetail {
    pub name: String,
    pub index: usize,
    pub visible: bool,
    pub color: Option<String>,
    pub protected: bool,
    pub row_count: usize,
    pub col_count: usize,
    pub merged_cells: Vec<String>,
}
