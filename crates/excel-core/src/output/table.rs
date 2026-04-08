use serde::Serialize;

pub fn format<T: Serialize>(value: &T) -> String {
    let json_value = serde_json::to_value(value).unwrap_or_default();

    match &json_value {
        serde_json::Value::Array(arr) => format_array(arr),
        serde_json::Value::Object(obj) => format_object(obj),
        other => other.to_string(),
    }
}

fn format_array(arr: &[serde_json::Value]) -> String {
    if arr.is_empty() {
        return "(empty)".to_string();
    }

    // Collect all keys from first object
    let keys: Vec<String> = match &arr[0] {
        serde_json::Value::Object(obj) => obj.keys().cloned().collect(),
        _ => return arr.iter().map(|v| v.to_string()).collect::<Vec<_>>().join("\n"),
    };

    let mut rows: Vec<Vec<String>> = Vec::new();
    rows.push(keys.clone());

    for item in arr {
        if let serde_json::Value::Object(obj) = item {
            let row: Vec<String> = keys
                .iter()
                .map(|k| value_to_string(obj.get(k).unwrap_or(&serde_json::Value::Null)))
                .collect();
            rows.push(row);
        }
    }

    render_table(&rows)
}

fn format_object(obj: &serde_json::Map<String, serde_json::Value>) -> String {
    let mut rows: Vec<Vec<String>> = Vec::new();
    rows.push(vec!["Key".to_string(), "Value".to_string()]);

    for (k, v) in obj {
        rows.push(vec![k.clone(), value_to_string(v)]);
    }

    render_table(&rows)
}

fn value_to_string(v: &serde_json::Value) -> String {
    match v {
        serde_json::Value::String(s) => s.clone(),
        serde_json::Value::Null => "".to_string(),
        serde_json::Value::Bool(b) => b.to_string(),
        serde_json::Value::Number(n) => n.to_string(),
        other => other.to_string(),
    }
}

fn render_table(rows: &[Vec<String>]) -> String {
    if rows.is_empty() {
        return String::new();
    }

    let col_count = rows[0].len();
    let mut widths = vec![0usize; col_count];

    for row in rows {
        for (i, cell) in row.iter().enumerate() {
            if i < col_count {
                widths[i] = widths[i].max(cell.len());
            }
        }
    }

    let mut output = String::new();

    // Header
    let header: Vec<String> = rows[0]
        .iter()
        .enumerate()
        .map(|(i, cell)| format!("{:<width$}", cell, width = widths[i]))
        .collect();
    output.push_str(&header.join(" | "));
    output.push('\n');

    // Separator
    let sep: Vec<String> = widths.iter().map(|w| "-".repeat(*w)).collect();
    output.push_str(&sep.join("-+-"));
    output.push('\n');

    // Data rows
    for row in rows.iter().skip(1) {
        let line: Vec<String> = row
            .iter()
            .enumerate()
            .map(|(i, cell)| {
                let w = if i < widths.len() { widths[i] } else { cell.len() };
                format!("{:<width$}", cell, width = w)
            })
            .collect();
        output.push_str(&line.join(" | "));
        output.push('\n');
    }

    output
}
