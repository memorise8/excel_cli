use thiserror::Error;

#[derive(Error, Debug)]
pub enum ExcelError {
    #[error("File not found: {0}")]
    FileNotFound(String),

    #[error("Invalid file format: {0}")]
    InvalidFormat(String),

    #[error("Sheet not found: {0}")]
    SheetNotFound(String),

    #[error("Sheet already exists: {0}")]
    SheetAlreadyExists(String),

    #[error("Invalid range: {0}")]
    InvalidRange(String),

    #[error("Table not found: {0}")]
    TableNotFound(String),

    #[error("Named range not found: {0}")]
    NamedRangeNotFound(String),

    #[error("Write error: {0}")]
    WriteError(String),

    #[error("Authentication required: {0}")]
    AuthRequired(String),

    #[error("Cloud API error: {0}")]
    CloudApiError(String),

    #[error("IO error: {0}")]
    Io(#[from] std::io::Error),

    #[error("Serialization error: {0}")]
    Serialization(#[from] serde_json::Error),

    #[error("Spreadsheet error: {0}")]
    Spreadsheet(String),

    #[error("Calamine error: {0}")]
    Calamine(String),

    #[error("Excel engine panic: {0}")]
    EnginePanic(String),

    #[error("Unsupported operation: {0}")]
    Unsupported(String),

    #[error("{0}")]
    Other(String),
}

pub type ExcelResult<T> = Result<T, ExcelError>;
