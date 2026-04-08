use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct TableInfo {
    pub name: String,
    pub sheet: String,
    pub range: String,
    pub columns: Vec<String>,
    pub row_count: usize,
    pub has_total_row: bool,
    pub style: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct TableData {
    pub info: TableInfo,
    pub headers: Vec<String>,
    pub rows: Vec<Vec<serde_json::Value>>,
}
