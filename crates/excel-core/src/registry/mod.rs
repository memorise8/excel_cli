pub mod calc;
pub mod chart;
pub mod conditional;
pub mod connection;
pub mod export;
pub mod file;
pub mod format;
pub mod formula;
pub mod named_range;
pub mod pivot;
pub mod range;
pub mod sheet;
pub mod slicer;
pub mod table;

use serde::{Deserialize, Serialize};

/// Execution layer for an operation
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum ExecutionLayer {
    /// Local file manipulation (umya-spreadsheet/calamine)
    Local,
    /// Microsoft Graph API (requires OAuth2 auth)
    Graph,
    /// Works on both, prefers local
    Any,
}

/// Argument type for dynamic CLI generation
#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum ArgType {
    String,
    Int,
    Float,
    Bool,
    Json,
    FilePath,
}

/// Argument definition for an operation
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ArgDef {
    pub name: &'static str,
    pub description: &'static str,
    pub arg_type: ArgType,
    pub required: bool,
    pub default: Option<&'static str>,
}

/// Flag definition for an operation
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct FlagDef {
    pub name: &'static str,
    pub short: Option<char>,
    pub description: &'static str,
    pub takes_value: bool,
    pub default: Option<&'static str>,
}

/// A single operation definition (drives CLI generation)
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct OperationDef {
    pub service: &'static str,
    pub verb: &'static str,
    pub description: &'static str,
    pub long_description: Option<&'static str>,
    pub args: Vec<ArgDef>,
    pub flags: Vec<FlagDef>,
    pub layer: ExecutionLayer,
    pub auth_required: bool,
}

/// Service definition grouping operations
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ServiceDef {
    pub name: &'static str,
    pub description: &'static str,
    pub operations: Vec<OperationDef>,
}

/// Common flags shared across many operations
pub fn file_arg() -> ArgDef {
    ArgDef {
        name: "file",
        description: "Path to the Excel file (.xlsx)",
        arg_type: ArgType::FilePath,
        required: true,
        default: None,
    }
}

pub fn range_arg() -> ArgDef {
    ArgDef {
        name: "range",
        description: "Cell range (e.g., 'Sheet1!A1:C3' or 'A1:C3')",
        arg_type: ArgType::String,
        required: true,
        default: None,
    }
}

pub fn sheet_arg() -> ArgDef {
    ArgDef {
        name: "sheet",
        description: "Sheet name",
        arg_type: ArgType::String,
        required: true,
        default: None,
    }
}

pub fn format_flag() -> FlagDef {
    FlagDef {
        name: "format",
        short: Some('f'),
        description: "Output format: json (default), table, csv",
        takes_value: true,
        default: Some("json"),
    }
}

pub fn cloud_flag() -> FlagDef {
    FlagDef {
        name: "cloud",
        short: None,
        description: "Use Microsoft Graph API (requires auth)",
        takes_value: false,
        default: None,
    }
}

/// Get all service definitions (the full registry)
pub fn all_services() -> Vec<ServiceDef> {
    vec![
        file::service(),
        sheet::service(),
        range::service(),
        formula::service(),
        format::service(),
        conditional::service(),
        table::service(),
        named_range::service(),
        pivot::service(),
        chart::service(),
        calc::service(),
        connection::service(),
        slicer::service(),
        export::service(),
    ]
}
