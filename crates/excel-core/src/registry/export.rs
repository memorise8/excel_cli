use super::*;

pub fn service() -> ServiceDef {
    ServiceDef {
        name: "export",
        description: "Export operations",
        operations: vec![
            op("csv", "Export sheet to CSV", vec![file_arg()], vec![
                FlagDef { name: "sheet", short: Some('s'), description: "Sheet name", takes_value: true, default: None },
                FlagDef { name: "output", short: Some('o'), description: "Output file path", takes_value: true, default: None },
                FlagDef { name: "delimiter", short: Some('d'), description: "Delimiter character", takes_value: true, default: Some(",") },
            ], false),
            op("json", "Export sheet to JSON", vec![file_arg()], vec![
                FlagDef { name: "sheet", short: Some('s'), description: "Sheet name", takes_value: true, default: None },
                FlagDef { name: "output", short: Some('o'), description: "Output file path", takes_value: true, default: None },
                FlagDef { name: "orient", short: None, description: "JSON orientation: records, columns, values", takes_value: true, default: Some("records") },
            ], false),
            op("html", "Export sheet to HTML table", vec![file_arg()], vec![
                FlagDef { name: "sheet", short: Some('s'), description: "Sheet name", takes_value: true, default: None },
                FlagDef { name: "output", short: Some('o'), description: "Output file path", takes_value: true, default: None },
            ], false),
            op("pdf", "Export to PDF (requires --cloud or LibreOffice)", vec![file_arg()], vec![
                FlagDef { name: "output", short: Some('o'), description: "Output PDF path", takes_value: true, default: None },
                FlagDef { name: "sheets", short: Some('s'), description: "Sheet names (comma-separated, all if omitted)", takes_value: true, default: None },
                cloud_flag(),
            ], true),
            op("screenshot", "Capture range as PNG (requires --cloud)", vec![file_arg(), range_arg()], vec![
                FlagDef { name: "output", short: Some('o'), description: "Output PNG path", takes_value: true, default: None },
                cloud_flag(),
            ], true),
        ],
    }
}

fn op(verb: &'static str, desc: &'static str, args: Vec<ArgDef>, flags: Vec<FlagDef>, auth: bool) -> OperationDef {
    OperationDef {
        service: "export",
        verb,
        description: desc,
        long_description: None,
        args,
        flags,
        layer: if auth { ExecutionLayer::Graph } else { ExecutionLayer::Local },
        auth_required: auth,
    }
}
