use super::*;

pub fn service() -> ServiceDef {
    ServiceDef {
        name: "named-range",
        description: "Named range (defined names) management",
        operations: vec![
            op("list", "List all named ranges", vec![file_arg()], vec![format_flag()]),
            op("create", "Create a named range", vec![file_arg()], vec![
                FlagDef { name: "name", short: Some('n'), description: "Range name", takes_value: true, default: None },
                FlagDef { name: "refers-to", short: Some('r'), description: "Range reference (e.g., Sheet1!$A$1:$C$10)", takes_value: true, default: None },
                FlagDef { name: "scope", short: None, description: "Scope: workbook (default) or sheet name", takes_value: true, default: Some("workbook") },
            ]),
            op("delete", "Delete a named range", vec![file_arg()], vec![
                FlagDef { name: "name", short: Some('n'), description: "Range name to delete", takes_value: true, default: None },
            ]),
            op("update", "Update a named range reference", vec![file_arg()], vec![
                FlagDef { name: "name", short: Some('n'), description: "Range name", takes_value: true, default: None },
                FlagDef { name: "refers-to", short: Some('r'), description: "New range reference", takes_value: true, default: None },
            ]),
            op("resolve", "Resolve name to range address", vec![file_arg()], vec![
                FlagDef { name: "name", short: Some('n'), description: "Range name to resolve", takes_value: true, default: None },
                format_flag(),
            ]),
            op("read", "Read values from a named range", vec![file_arg()], vec![
                FlagDef { name: "name", short: Some('n'), description: "Range name", takes_value: true, default: None },
                format_flag(),
            ]),
        ],
    }
}

fn op(verb: &'static str, desc: &'static str, args: Vec<ArgDef>, flags: Vec<FlagDef>) -> OperationDef {
    OperationDef {
        service: "named-range",
        verb,
        description: desc,
        long_description: None,
        args,
        flags,
        layer: ExecutionLayer::Local,
        auth_required: false,
    }
}
