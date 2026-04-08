use super::*;

pub fn service() -> ServiceDef {
    ServiceDef {
        name: "table",
        description: "Excel table operations",
        operations: vec![
            op("list", "List all tables in workbook", vec![file_arg()], vec![format_flag()]),
            op("create", "Create an Excel table", vec![file_arg(), range_arg()], vec![
                FlagDef { name: "name", short: Some('n'), description: "Table name", takes_value: true, default: None },
                FlagDef { name: "style", short: None, description: "Table style (e.g., TableStyleMedium2)", takes_value: true, default: None },
                FlagDef { name: "has-headers", short: None, description: "First row contains headers", takes_value: false, default: None },
            ]),
            op("delete", "Delete a table (keeps data)", vec![file_arg()], vec![
                FlagDef { name: "name", short: Some('n'), description: "Table name", takes_value: true, default: None },
            ]),
            op("read", "Read table data", vec![file_arg()], vec![
                FlagDef { name: "name", short: Some('n'), description: "Table name", takes_value: true, default: None },
                format_flag(),
            ]),
            op("append", "Append rows to a table", vec![file_arg()], vec![
                FlagDef { name: "name", short: Some('n'), description: "Table name", takes_value: true, default: None },
                FlagDef { name: "data", short: Some('d'), description: "JSON array of rows", takes_value: true, default: None },
            ]),
            op("resize", "Resize table range", vec![file_arg()], vec![
                FlagDef { name: "name", short: Some('n'), description: "Table name", takes_value: true, default: None },
                FlagDef { name: "range", short: Some('r'), description: "New range", takes_value: true, default: None },
            ]),
            op("rename", "Rename a table", vec![file_arg()], vec![
                FlagDef { name: "name", short: Some('n'), description: "Current table name", takes_value: true, default: None },
                FlagDef { name: "new-name", short: None, description: "New table name", takes_value: true, default: None },
            ]),
            op("sort", "Sort table by column", vec![file_arg()], vec![
                FlagDef { name: "name", short: Some('n'), description: "Table name", takes_value: true, default: None },
                FlagDef { name: "by", short: Some('b'), description: "Column name to sort by", takes_value: true, default: None },
                FlagDef { name: "desc", short: None, description: "Sort descending", takes_value: false, default: None },
            ]),
            op("filter", "Filter table data", vec![file_arg()], vec![
                FlagDef { name: "name", short: Some('n'), description: "Table name", takes_value: true, default: None },
                FlagDef { name: "column", short: Some('c'), description: "Column name", takes_value: true, default: None },
                FlagDef { name: "value", short: Some('v'), description: "Filter value", takes_value: true, default: None },
            ]),
            op("style", "Apply table style", vec![file_arg()], vec![
                FlagDef { name: "name", short: Some('n'), description: "Table name", takes_value: true, default: None },
                FlagDef { name: "style", short: Some('s'), description: "Style name (e.g., TableStyleMedium2)", takes_value: true, default: None },
            ]),
            op("total-row", "Toggle total row", vec![file_arg()], vec![
                FlagDef { name: "name", short: Some('n'), description: "Table name", takes_value: true, default: None },
                FlagDef { name: "enable", short: None, description: "Enable total row", takes_value: false, default: None },
                FlagDef { name: "disable", short: None, description: "Disable total row", takes_value: false, default: None },
            ]),
            op("column-add", "Add a column to table", vec![file_arg()], vec![
                FlagDef { name: "name", short: Some('n'), description: "Table name", takes_value: true, default: None },
                FlagDef { name: "header", short: None, description: "Column header name", takes_value: true, default: None },
            ]),
            op("column-delete", "Delete a table column", vec![file_arg()], vec![
                FlagDef { name: "name", short: Some('n'), description: "Table name", takes_value: true, default: None },
                FlagDef { name: "column", short: Some('c'), description: "Column name to delete", takes_value: true, default: None },
            ]),
            op("to-range", "Convert table to plain range", vec![file_arg()], vec![
                FlagDef { name: "name", short: Some('n'), description: "Table name", takes_value: true, default: None },
            ]),
        ],
    }
}

fn op(verb: &'static str, desc: &'static str, args: Vec<ArgDef>, flags: Vec<FlagDef>) -> OperationDef {
    OperationDef {
        service: "table",
        verb,
        description: desc,
        long_description: None,
        args,
        flags,
        layer: ExecutionLayer::Local,
        auth_required: false,
    }
}
