use super::*;

pub fn service() -> ServiceDef {
    ServiceDef {
        name: "format",
        description: "Cell formatting operations",
        operations: vec![
            op("font", "Set font properties", vec![file_arg(), range_arg()], vec![
                FlagDef { name: "name", short: Some('n'), description: "Font name (e.g., Arial)", takes_value: true, default: None },
                FlagDef { name: "size", short: Some('s'), description: "Font size", takes_value: true, default: None },
                FlagDef { name: "color", short: Some('c'), description: "Font color (hex)", takes_value: true, default: None },
                FlagDef { name: "bold", short: Some('b'), description: "Bold", takes_value: false, default: None },
                FlagDef { name: "italic", short: Some('i'), description: "Italic", takes_value: false, default: None },
                FlagDef { name: "underline", short: Some('u'), description: "Underline", takes_value: false, default: None },
            ]),
            op("fill", "Set background fill", vec![file_arg(), range_arg()], vec![
                FlagDef { name: "color", short: Some('c'), description: "Fill color (hex)", takes_value: true, default: None },
                FlagDef { name: "pattern", short: None, description: "Fill pattern", takes_value: true, default: Some("solid") },
            ]),
            op("border", "Set borders", vec![file_arg(), range_arg()], vec![
                FlagDef { name: "style", short: Some('s'), description: "Border style: thin, medium, thick, double, dashed", takes_value: true, default: Some("thin") },
                FlagDef { name: "color", short: Some('c'), description: "Border color (hex)", takes_value: true, default: None },
                FlagDef { name: "sides", short: None, description: "Sides: all, top, bottom, left, right, outline", takes_value: true, default: Some("all") },
            ]),
            op("align", "Set alignment", vec![file_arg(), range_arg()], vec![
                FlagDef { name: "horizontal", short: Some('h'), description: "Horizontal: left, center, right", takes_value: true, default: None },
                FlagDef { name: "vertical", short: Some('v'), description: "Vertical: top, center, bottom", takes_value: true, default: None },
                FlagDef { name: "wrap", short: Some('w'), description: "Wrap text", takes_value: false, default: None },
            ]),
            op("number", "Set number format", vec![file_arg(), range_arg()], vec![
                FlagDef { name: "format", short: None, description: "Format code (e.g., #,##0.00, 0%, yyyy-mm-dd)", takes_value: true, default: None },
                FlagDef { name: "preset", short: Some('p'), description: "Preset: number, currency, percent, date, time, scientific", takes_value: true, default: None },
            ]),
            op("width", "Set column width", vec![file_arg()], vec![
                FlagDef { name: "sheet", short: Some('s'), description: "Sheet name", takes_value: true, default: None },
                FlagDef { name: "col", short: Some('c'), description: "Column letter", takes_value: true, default: None },
                FlagDef { name: "width", short: Some('w'), description: "Width value", takes_value: true, default: None },
            ]),
            op("height", "Set row height", vec![file_arg()], vec![
                FlagDef { name: "sheet", short: Some('s'), description: "Sheet name", takes_value: true, default: None },
                FlagDef { name: "row", short: Some('r'), description: "Row number", takes_value: true, default: None },
                FlagDef { name: "height", short: None, description: "Height value", takes_value: true, default: None },
            ]),
            op("autofit", "Auto-fit column widths", vec![file_arg()], vec![
                FlagDef { name: "sheet", short: Some('s'), description: "Sheet name", takes_value: true, default: None },
                FlagDef { name: "cols", short: Some('c'), description: "Column range (e.g., A:D)", takes_value: true, default: None },
            ]),
            op("style", "Apply compound style", vec![file_arg(), range_arg()], vec![
                FlagDef { name: "json", short: Some('j'), description: "JSON style definition", takes_value: true, default: None },
            ]),
        ],
    }
}

fn op(verb: &'static str, desc: &'static str, args: Vec<ArgDef>, flags: Vec<FlagDef>) -> OperationDef {
    OperationDef {
        service: "format",
        verb,
        description: desc,
        long_description: None,
        args,
        flags,
        layer: ExecutionLayer::Local,
        auth_required: false,
    }
}
