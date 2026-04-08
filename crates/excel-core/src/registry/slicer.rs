use super::*;

pub fn service() -> ServiceDef {
    ServiceDef {
        name: "slicer",
        description: "Slicer operations (require --cloud)",
        operations: vec![
            op("list", "List slicers", vec![file_arg()], vec![format_flag()], false),
            op("create", "Create a slicer", vec![file_arg()], vec![
                FlagDef { name: "source", short: Some('s'), description: "Source table or pivot table name", takes_value: true, default: None },
                FlagDef { name: "field", short: None, description: "Field name for slicer", takes_value: true, default: None },
                cloud_flag(),
            ], true),
            op("delete", "Delete a slicer", vec![file_arg()], vec![
                FlagDef { name: "name", short: Some('n'), description: "Slicer name", takes_value: true, default: None },
                cloud_flag(),
            ], true),
            op("select", "Select slicer items", vec![file_arg()], vec![
                FlagDef { name: "name", short: Some('n'), description: "Slicer name", takes_value: true, default: None },
                FlagDef { name: "items", short: Some('i'), description: "Items to select (comma-separated)", takes_value: true, default: None },
                cloud_flag(),
            ], true),
            op("clear", "Clear slicer selection", vec![file_arg()], vec![
                FlagDef { name: "name", short: Some('n'), description: "Slicer name", takes_value: true, default: None },
                cloud_flag(),
            ], true),
        ],
    }
}

fn op(verb: &'static str, desc: &'static str, args: Vec<ArgDef>, flags: Vec<FlagDef>, auth: bool) -> OperationDef {
    OperationDef {
        service: "slicer",
        verb,
        description: desc,
        long_description: None,
        args,
        flags,
        layer: if auth { ExecutionLayer::Graph } else { ExecutionLayer::Local },
        auth_required: auth,
    }
}
