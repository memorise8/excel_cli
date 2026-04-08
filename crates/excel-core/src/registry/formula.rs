use super::*;

pub fn service() -> ServiceDef {
    ServiceDef {
        name: "formula",
        description: "Formula operations",
        operations: vec![
            op("read", "Read formula text from a cell", vec![file_arg(), range_arg()], vec![format_flag()], false),
            op("write", "Write a formula to a cell", vec![file_arg(), range_arg()], vec![
                FlagDef { name: "formula", short: None, description: "Formula text (e.g., =SUM(A1:A10))", takes_value: true, default: None },
            ], false),
            op("list", "List all formulas in a sheet", vec![file_arg()], vec![
                FlagDef { name: "sheet", short: Some('s'), description: "Sheet name", takes_value: true, default: None },
                format_flag(),
            ], false),
            op("evaluate", "Evaluate formula (requires cloud)", vec![file_arg(), range_arg()], vec![
                cloud_flag(),
                format_flag(),
            ], true),
            op("audit", "Trace formula precedents/dependents", vec![file_arg(), range_arg()], vec![
                FlagDef { name: "direction", short: Some('d'), description: "Direction: precedents or dependents", takes_value: true, default: Some("precedents") },
                format_flag(),
            ], false),
        ],
    }
}

fn op(verb: &'static str, desc: &'static str, args: Vec<ArgDef>, flags: Vec<FlagDef>, auth: bool) -> OperationDef {
    OperationDef {
        service: "formula",
        verb,
        description: desc,
        long_description: None,
        args,
        flags,
        layer: if auth { ExecutionLayer::Graph } else { ExecutionLayer::Local },
        auth_required: auth,
    }
}
