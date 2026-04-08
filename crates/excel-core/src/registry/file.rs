use super::*;

pub fn service() -> ServiceDef {
    ServiceDef {
        name: "file",
        description: "Workbook file management",
        operations: vec![
            OperationDef {
                service: "file",
                verb: "create",
                description: "Create a new Excel workbook",
                long_description: Some("Creates a new .xlsx file with a default sheet"),
                args: vec![file_arg()],
                flags: vec![
                    FlagDef {
                        name: "sheets",
                        short: Some('s'),
                        description: "Comma-separated sheet names to create",
                        takes_value: true,
                        default: None,
                    },
                ],
                layer: ExecutionLayer::Local,
                auth_required: false,
            },
            OperationDef {
                service: "file",
                verb: "info",
                description: "Show workbook information",
                long_description: Some("Display metadata: sheets, size, tables, named ranges"),
                args: vec![file_arg()],
                flags: vec![format_flag()],
                layer: ExecutionLayer::Local,
                auth_required: false,
            },
            OperationDef {
                service: "file",
                verb: "save",
                description: "Save workbook to a different path",
                long_description: None,
                args: vec![
                    file_arg(),
                    ArgDef {
                        name: "output",
                        description: "Output file path",
                        arg_type: ArgType::FilePath,
                        required: true,
                        default: None,
                    },
                ],
                flags: vec![],
                layer: ExecutionLayer::Local,
                auth_required: false,
            },
            OperationDef {
                service: "file",
                verb: "convert",
                description: "Convert file format (xlsx, csv, json)",
                long_description: None,
                args: vec![file_arg()],
                flags: vec![
                    FlagDef {
                        name: "to",
                        short: Some('t'),
                        description: "Target format: csv, json, xlsx",
                        takes_value: true,
                        default: None,
                    },
                    FlagDef {
                        name: "output",
                        short: Some('o'),
                        description: "Output file path",
                        takes_value: true,
                        default: None,
                    },
                ],
                layer: ExecutionLayer::Local,
                auth_required: false,
            },
            OperationDef {
                service: "file",
                verb: "upload",
                description: "Upload workbook to OneDrive",
                long_description: Some("Requires Microsoft Graph API authentication"),
                args: vec![file_arg()],
                flags: vec![
                    FlagDef {
                        name: "folder",
                        short: None,
                        description: "OneDrive folder path",
                        takes_value: true,
                        default: None,
                    },
                    cloud_flag(),
                ],
                layer: ExecutionLayer::Graph,
                auth_required: true,
            },
            OperationDef {
                service: "file",
                verb: "download",
                description: "Download workbook from OneDrive",
                long_description: Some("Requires Microsoft Graph API authentication"),
                args: vec![
                    ArgDef {
                        name: "remote-path",
                        description: "OneDrive file path",
                        arg_type: ArgType::String,
                        required: true,
                        default: None,
                    },
                ],
                flags: vec![
                    FlagDef {
                        name: "output",
                        short: Some('o'),
                        description: "Local output path",
                        takes_value: true,
                        default: None,
                    },
                    cloud_flag(),
                ],
                layer: ExecutionLayer::Graph,
                auth_required: true,
            },
        ],
    }
}
