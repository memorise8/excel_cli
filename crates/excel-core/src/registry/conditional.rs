use super::*;

pub fn service() -> ServiceDef {
    ServiceDef {
        name: "conditional",
        description: "Conditional formatting rules",
        operations: vec![
            op("add", "Add a conditional formatting rule", vec![file_arg(), range_arg()], vec![
                FlagDef { name: "type", short: Some('t'), description: "Rule type: cell-value, color-scale, data-bar, icon-set, formula", takes_value: true, default: None },
                FlagDef { name: "operator", short: None, description: "Operator: greater-than, less-than, between, equal, etc.", takes_value: true, default: None },
                FlagDef { name: "value", short: Some('v'), description: "Comparison value", takes_value: true, default: None },
                FlagDef { name: "format-json", short: None, description: "Format to apply (JSON style)", takes_value: true, default: None },
            ]),
            op("list", "List conditional formatting rules", vec![file_arg()], vec![
                FlagDef { name: "sheet", short: Some('s'), description: "Sheet name", takes_value: true, default: None },
                format_flag(),
            ]),
            op("delete", "Delete a conditional formatting rule", vec![file_arg()], vec![
                FlagDef { name: "sheet", short: Some('s'), description: "Sheet name", takes_value: true, default: None },
                FlagDef { name: "index", short: Some('i'), description: "Rule index to delete", takes_value: true, default: None },
            ]),
            op("clear", "Clear all rules from a range", vec![file_arg(), range_arg()], vec![]),
        ],
    }
}

fn op(verb: &'static str, desc: &'static str, args: Vec<ArgDef>, flags: Vec<FlagDef>) -> OperationDef {
    OperationDef {
        service: "conditional",
        verb,
        description: desc,
        long_description: None,
        args,
        flags,
        layer: ExecutionLayer::Local,
        auth_required: false,
    }
}
