pub mod csv;
pub mod json;
pub mod table;

use serde::Serialize;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum OutputFormat {
    Json,
    Table,
    Csv,
}

impl OutputFormat {
    pub fn from_str(s: &str) -> Self {
        match s.to_lowercase().as_str() {
            "table" => Self::Table,
            "csv" => Self::Csv,
            _ => Self::Json,
        }
    }
}

/// Format any serializable value according to the output format
pub fn format_output<T: Serialize>(value: &T, format: OutputFormat) -> String {
    match format {
        OutputFormat::Json => json::format(value),
        OutputFormat::Table => table::format(value),
        OutputFormat::Csv => csv::format(value),
    }
}
