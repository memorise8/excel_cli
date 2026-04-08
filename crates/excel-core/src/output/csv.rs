use serde::Serialize;

pub fn format<T: Serialize>(value: &T) -> String {
    let json_value = serde_json::to_value(value).unwrap_or_default();

    match &json_value {
        serde_json::Value::Array(arr) => format_array(arr),
        serde_json::Value::Object(obj) => {
            let keys: Vec<&str> = obj.keys().map(|k| k.as_str()).collect();
            let values: Vec<String> = obj.values().map(value_to_csv_field).collect();
            format!("{}\n{}", keys.join(","), values.join(","))
        }
        other => other.to_string(),
    }
}

fn format_array(arr: &[serde_json::Value]) -> String {
    if arr.is_empty() {
        return String::new();
    }

    let keys: Vec<String> = match &arr[0] {
        serde_json::Value::Object(obj) => obj.keys().cloned().collect(),
        _ => {
            return arr
                .iter()
                .map(|v| value_to_csv_field(v))
                .collect::<Vec<_>>()
                .join("\n");
        }
    };

    let mut output = keys.join(",");
    output.push('\n');

    for item in arr {
        if let serde_json::Value::Object(obj) = item {
            let row: Vec<String> = keys
                .iter()
                .map(|k| value_to_csv_field(obj.get(k).unwrap_or(&serde_json::Value::Null)))
                .collect();
            output.push_str(&row.join(","));
            output.push('\n');
        }
    }

    output
}

fn value_to_csv_field(v: &serde_json::Value) -> String {
    match v {
        serde_json::Value::String(s) => {
            if s.contains(',') || s.contains('"') || s.contains('\n') {
                format!("\"{}\"", s.replace('"', "\"\""))
            } else {
                s.clone()
            }
        }
        serde_json::Value::Null => String::new(),
        serde_json::Value::Bool(b) => b.to_string(),
        serde_json::Value::Number(n) => n.to_string(),
        other => {
            let s = other.to_string();
            format!("\"{}\"", s.replace('"', "\"\""))
        }
    }
}
