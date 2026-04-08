use super::*;

pub fn service() -> ServiceDef {
    ServiceDef {
        name: "pivot",
        description: "PivotTable operations (most require --cloud)",
        operations: vec![
            op("list", "List pivot tables", vec![file_arg()], vec![format_flag()], false),
            op("create", "Create a pivot table", vec![file_arg()], vec![
                FlagDef { name: "source", short: Some('s'), description: "Source data range (e.g., Sheet1!A1:D100)", takes_value: true, default: None },
                FlagDef { name: "dest", short: Some('d'), description: "Destination (e.g., Sheet2!A1)", takes_value: true, default: None },
                FlagDef { name: "rows", short: None, description: "Row field names (comma-separated)", takes_value: true, default: None },
                FlagDef { name: "cols", short: None, description: "Column field names (comma-separated)", takes_value: true, default: None },
                FlagDef { name: "values", short: None, description: "Value fields (name:aggregation, comma-separated)", takes_value: true, default: None },
                cloud_flag(),
            ], true),
            op("refresh", "Refresh pivot table data", vec![file_arg()], vec![
                FlagDef { name: "name", short: Some('n'), description: "Pivot table name", takes_value: true, default: None },
                cloud_flag(),
            ], true),
            op("field-add", "Add a field to pivot table", vec![file_arg()], vec![
                FlagDef { name: "name", short: Some('n'), description: "Pivot table name", takes_value: true, default: None },
                FlagDef { name: "field", short: None, description: "Field name", takes_value: true, default: None },
                FlagDef { name: "area", short: None, description: "Area: row, column, value, filter", takes_value: true, default: None },
                cloud_flag(),
            ], true),
            op("field-remove", "Remove a field from pivot table", vec![file_arg()], vec![
                FlagDef { name: "name", short: Some('n'), description: "Pivot table name", takes_value: true, default: None },
                FlagDef { name: "field", short: None, description: "Field name", takes_value: true, default: None },
                cloud_flag(),
            ], true),
            op("filter", "Set pivot table filter", vec![file_arg()], vec![
                FlagDef { name: "name", short: Some('n'), description: "Pivot table name", takes_value: true, default: None },
                FlagDef { name: "field", short: None, description: "Filter field", takes_value: true, default: None },
                FlagDef { name: "values", short: Some('v'), description: "Filter values (comma-separated)", takes_value: true, default: None },
                cloud_flag(),
            ], true),
            op("group", "Group pivot table items", vec![file_arg()], vec![
                FlagDef { name: "name", short: Some('n'), description: "Pivot table name", takes_value: true, default: None },
                FlagDef { name: "field", short: None, description: "Field to group", takes_value: true, default: None },
                FlagDef { name: "by", short: Some('b'), description: "Group by: days, months, quarters, years", takes_value: true, default: None },
                cloud_flag(),
            ], true),
            op("style", "Apply pivot table style", vec![file_arg()], vec![
                FlagDef { name: "name", short: Some('n'), description: "Pivot table name", takes_value: true, default: None },
                FlagDef { name: "style", short: Some('s'), description: "Style name", takes_value: true, default: None },
                cloud_flag(),
            ], true),
        ],
    }
}

fn op(verb: &'static str, desc: &'static str, args: Vec<ArgDef>, flags: Vec<FlagDef>, auth: bool) -> OperationDef {
    OperationDef {
        service: "pivot",
        verb,
        description: desc,
        long_description: None,
        args,
        flags,
        layer: if auth { ExecutionLayer::Graph } else { ExecutionLayer::Local },
        auth_required: auth,
    }
}
