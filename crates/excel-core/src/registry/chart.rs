use super::*;

pub fn service() -> ServiceDef {
    ServiceDef {
        name: "chart",
        description: "Chart operations (require --cloud)",
        operations: vec![
            op("list", "List charts in workbook", vec![file_arg()], vec![format_flag()], false),
            op("create", "Create a chart", vec![file_arg()], vec![
                FlagDef { name: "type", short: Some('t'), description: "Chart type: bar, column, line, pie, scatter, area", takes_value: true, default: None },
                FlagDef { name: "source", short: Some('s'), description: "Data source range", takes_value: true, default: None },
                FlagDef { name: "title", short: None, description: "Chart title", takes_value: true, default: None },
                FlagDef { name: "sheet", short: None, description: "Target sheet for chart", takes_value: true, default: None },
                cloud_flag(),
            ], true),
            op("delete", "Delete a chart", vec![file_arg()], vec![
                FlagDef { name: "name", short: Some('n'), description: "Chart name", takes_value: true, default: None },
                cloud_flag(),
            ], true),
            op("update", "Update chart properties", vec![file_arg()], vec![
                FlagDef { name: "name", short: Some('n'), description: "Chart name", takes_value: true, default: None },
                FlagDef { name: "title", short: None, description: "New title", takes_value: true, default: None },
                FlagDef { name: "type", short: Some('t'), description: "New chart type", takes_value: true, default: None },
                cloud_flag(),
            ], true),
            op("export", "Export chart as image", vec![file_arg()], vec![
                FlagDef { name: "name", short: Some('n'), description: "Chart name", takes_value: true, default: None },
                FlagDef { name: "output", short: Some('o'), description: "Output file path", takes_value: true, default: None },
                FlagDef { name: "format", short: Some('f'), description: "Image format: png, svg", takes_value: true, default: Some("png") },
                cloud_flag(),
            ], true),
            op("series-add", "Add data series to chart", vec![file_arg()], vec![
                FlagDef { name: "name", short: Some('n'), description: "Chart name", takes_value: true, default: None },
                FlagDef { name: "values", short: Some('v'), description: "Values range", takes_value: true, default: None },
                FlagDef { name: "label", short: Some('l'), description: "Series label", takes_value: true, default: None },
                cloud_flag(),
            ], true),
            op("series-remove", "Remove data series", vec![file_arg()], vec![
                FlagDef { name: "name", short: Some('n'), description: "Chart name", takes_value: true, default: None },
                FlagDef { name: "index", short: Some('i'), description: "Series index", takes_value: true, default: None },
                cloud_flag(),
            ], true),
            op("style", "Apply chart style", vec![file_arg()], vec![
                FlagDef { name: "name", short: Some('n'), description: "Chart name", takes_value: true, default: None },
                FlagDef { name: "style", short: Some('s'), description: "Style index or name", takes_value: true, default: None },
                cloud_flag(),
            ], true),
        ],
    }
}

fn op(verb: &'static str, desc: &'static str, args: Vec<ArgDef>, flags: Vec<FlagDef>, auth: bool) -> OperationDef {
    OperationDef {
        service: "chart",
        verb,
        description: desc,
        long_description: None,
        args,
        flags,
        layer: if auth { ExecutionLayer::Graph } else { ExecutionLayer::Local },
        auth_required: auth,
    }
}
