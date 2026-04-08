use super::*;

pub fn service() -> ServiceDef {
    ServiceDef {
        name: "range",
        description: "Cell range operations",
        operations: vec![
            op("read", "Read values from a range", vec![file_arg(), range_arg()], vec![
                format_flag(),
                FlagDef { name: "with-format", short: None, description: "Include font, fill, and border formatting in output (cloud only)", takes_value: false, default: None },
            ]),
            op("write", "Write values to a range", vec![file_arg(), range_arg()], vec![
                FlagDef { name: "data", short: Some('d'), description: "JSON array of values (e.g., [[1,2],[3,4]])", takes_value: true, default: None },
                FlagDef { name: "value", short: Some('v'), description: "Single value to write to all cells", takes_value: true, default: None },
            ]),
            op("clear", "Clear a range", vec![file_arg(), range_arg()], vec![
                FlagDef { name: "values-only", short: None, description: "Clear only values, keep formatting", takes_value: false, default: None },
            ]),
            op("copy", "Copy range to another location", vec![file_arg(), range_arg()], vec![
                FlagDef { name: "to", short: Some('t'), description: "Destination range (e.g., Sheet2!A1)", takes_value: true, default: None },
            ]),
            op("move", "Move range to another location", vec![file_arg(), range_arg()], vec![
                FlagDef { name: "to", short: Some('t'), description: "Destination range", takes_value: true, default: None },
            ]),
            op("insert", "Insert rows or columns", vec![file_arg()], vec![
                FlagDef { name: "sheet", short: Some('s'), description: "Sheet name", takes_value: true, default: None },
                FlagDef { name: "row", short: Some('r'), description: "Row number to insert before", takes_value: true, default: None },
                FlagDef { name: "col", short: Some('c'), description: "Column (letter) to insert before", takes_value: true, default: None },
                FlagDef { name: "count", short: Some('n'), description: "Number of rows/cols to insert", takes_value: true, default: Some("1") },
            ]),
            op("delete", "Delete rows or columns", vec![file_arg()], vec![
                FlagDef { name: "sheet", short: Some('s'), description: "Sheet name", takes_value: true, default: None },
                FlagDef { name: "row", short: Some('r'), description: "Row number to delete", takes_value: true, default: None },
                FlagDef { name: "col", short: Some('c'), description: "Column (letter) to delete", takes_value: true, default: None },
                FlagDef { name: "count", short: Some('n'), description: "Number of rows/cols to delete", takes_value: true, default: Some("1") },
            ]),
            op("merge", "Merge cells in a range", vec![file_arg(), range_arg()], vec![]),
            op("unmerge", "Unmerge cells in a range", vec![file_arg(), range_arg()], vec![]),
            op("sort", "Sort a range", vec![file_arg(), range_arg()], vec![
                FlagDef { name: "by", short: Some('b'), description: "Column to sort by (letter)", takes_value: true, default: None },
                FlagDef { name: "desc", short: None, description: "Sort descending", takes_value: false, default: None },
            ]),
            op("filter", "Set auto-filter on a range", vec![file_arg(), range_arg()], vec![]),
            op("find", "Find values in a range or sheet", vec![file_arg()], vec![
                FlagDef { name: "query", short: Some('q'), description: "Search query", takes_value: true, default: None },
                FlagDef { name: "sheet", short: Some('s'), description: "Sheet name (searches all if omitted)", takes_value: true, default: None },
                FlagDef { name: "regex", short: None, description: "Treat query as regex", takes_value: false, default: None },
                format_flag(),
            ]),
            op("replace", "Replace values in a range or sheet", vec![file_arg()], vec![
                FlagDef { name: "find", short: None, description: "Value to find", takes_value: true, default: None },
                FlagDef { name: "replace", short: None, description: "Replacement value", takes_value: true, default: None },
                FlagDef { name: "sheet", short: Some('s'), description: "Sheet name", takes_value: true, default: None },
            ]),
            op("validate", "Set data validation on a range", vec![file_arg(), range_arg()], vec![
                FlagDef { name: "type", short: Some('t'), description: "Validation type: list, number, date, text-length", takes_value: true, default: None },
                FlagDef { name: "values", short: Some('v'), description: "Allowed values (comma-separated for list)", takes_value: true, default: None },
                FlagDef { name: "min", short: None, description: "Minimum value", takes_value: true, default: None },
                FlagDef { name: "max", short: None, description: "Maximum value", takes_value: true, default: None },
            ]),
        ],
    }
}

fn op(verb: &'static str, desc: &'static str, args: Vec<ArgDef>, flags: Vec<FlagDef>) -> OperationDef {
    OperationDef {
        service: "range",
        verb,
        description: desc,
        long_description: None,
        args,
        flags,
        layer: ExecutionLayer::Local,
        auth_required: false,
    }
}
