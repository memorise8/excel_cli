use super::*;

pub fn service() -> ServiceDef {
    ServiceDef {
        name: "connection",
        description: "Data connection operations (require --cloud)",
        operations: vec![
            op("list", "List data connections", vec![file_arg()], vec![format_flag()], false),
            op("create", "Create a data connection", vec![file_arg()], vec![
                FlagDef { name: "type", short: Some('t'), description: "Connection type: oledb, odbc", takes_value: true, default: None },
                FlagDef { name: "connection-string", short: Some('c'), description: "Connection string", takes_value: true, default: None },
                FlagDef { name: "name", short: Some('n'), description: "Connection name", takes_value: true, default: None },
                cloud_flag(),
            ], true),
            op("delete", "Delete a data connection", vec![file_arg()], vec![
                FlagDef { name: "name", short: Some('n'), description: "Connection name", takes_value: true, default: None },
                cloud_flag(),
            ], true),
            op("refresh", "Refresh a data connection", vec![file_arg()], vec![
                FlagDef { name: "name", short: Some('n'), description: "Connection name", takes_value: true, default: None },
                cloud_flag(),
            ], true),
        ],
    }
}

fn op(verb: &'static str, desc: &'static str, args: Vec<ArgDef>, flags: Vec<FlagDef>, auth: bool) -> OperationDef {
    OperationDef {
        service: "connection",
        verb,
        description: desc,
        long_description: None,
        args,
        flags,
        layer: if auth { ExecutionLayer::Graph } else { ExecutionLayer::Local },
        auth_required: auth,
    }
}
