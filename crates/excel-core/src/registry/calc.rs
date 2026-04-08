use super::*;

pub fn service() -> ServiceDef {
    ServiceDef {
        name: "calc",
        description: "Calculation mode operations",
        operations: vec![
            OperationDef {
                service: "calc",
                verb: "mode",
                description: "Get or set calculation mode",
                long_description: Some("View or change automatic/manual calculation mode"),
                args: vec![file_arg()],
                flags: vec![
                    FlagDef { name: "set", short: Some('s'), description: "Set mode: automatic, manual", takes_value: true, default: None },
                    format_flag(),
                ],
                layer: ExecutionLayer::Local,
                auth_required: false,
            },
            OperationDef {
                service: "calc",
                verb: "now",
                description: "Recalculate all formulas (requires --cloud)",
                long_description: Some("Triggers full workbook recalculation via Graph API"),
                args: vec![file_arg()],
                flags: vec![cloud_flag()],
                layer: ExecutionLayer::Graph,
                auth_required: true,
            },
            OperationDef {
                service: "calc",
                verb: "sheet",
                description: "Recalculate a specific sheet (requires --cloud)",
                long_description: None,
                args: vec![file_arg()],
                flags: vec![
                    FlagDef { name: "sheet", short: Some('s'), description: "Sheet name", takes_value: true, default: None },
                    cloud_flag(),
                ],
                layer: ExecutionLayer::Graph,
                auth_required: true,
            },
        ],
    }
}
