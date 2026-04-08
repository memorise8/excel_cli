use super::*;

pub fn service() -> ServiceDef {
    ServiceDef {
        name: "sheet",
        description: "Worksheet management",
        operations: vec![
            op("list", "List all sheets in workbook", vec![file_arg()], vec![format_flag()], false),
            op("add", "Add a new sheet", vec![file_arg(), sheet_arg()], vec![
                FlagDef { name: "position", short: Some('p'), description: "Insert position (0-based index)", takes_value: true, default: None },
            ], false),
            op("rename", "Rename a sheet", vec![file_arg(), sheet_arg(), ArgDef { name: "new-name", description: "New sheet name", arg_type: ArgType::String, required: true, default: None }], vec![], false),
            op("delete", "Delete a sheet", vec![file_arg(), sheet_arg()], vec![], false),
            op("copy", "Copy a sheet", vec![file_arg(), sheet_arg()], vec![
                FlagDef { name: "new-name", short: Some('n'), description: "Name for the copy", takes_value: true, default: None },
            ], false),
            op("move", "Move sheet to a different position", vec![file_arg(), sheet_arg()], vec![
                FlagDef { name: "position", short: Some('p'), description: "Target position (0-based)", takes_value: true, default: None },
            ], false),
            op("hide", "Hide a sheet", vec![file_arg(), sheet_arg()], vec![], false),
            op("unhide", "Unhide a sheet", vec![file_arg(), sheet_arg()], vec![], false),
            op("color", "Set sheet tab color", vec![file_arg(), sheet_arg()], vec![
                FlagDef { name: "color", short: Some('c'), description: "Tab color (hex, e.g., FF0000)", takes_value: true, default: None },
            ], false),
            op("protect", "Protect a sheet", vec![file_arg(), sheet_arg()], vec![
                FlagDef { name: "password", short: None, description: "Protection password", takes_value: true, default: None },
            ], false),
            op("unprotect", "Unprotect a sheet", vec![file_arg(), sheet_arg()], vec![
                FlagDef { name: "password", short: None, description: "Protection password", takes_value: true, default: None },
            ], false),
        ],
    }
}

fn op(verb: &'static str, desc: &'static str, args: Vec<ArgDef>, flags: Vec<FlagDef>, auth: bool) -> OperationDef {
    OperationDef {
        service: "sheet",
        verb,
        description: desc,
        long_description: None,
        args,
        flags,
        layer: if auth { ExecutionLayer::Graph } else { ExecutionLayer::Local },
        auth_required: auth,
    }
}
