use crate::models::error::*;
use crate::models::range::col_index_to_letter;
use std::path::Path;

fn get_sheet_data(
    book: &mut umya_spreadsheet::Spreadsheet,
    sheet_name: &str,
) -> ExcelResult<(Vec<Vec<String>>, usize, usize)> {
    let idx = super::safe_io::find_sheet_index(book, sheet_name)?;
    super::safe_io::safe_read_sheet(book, idx)?;

    let sheet = book
        .get_sheet(&idx)
        .ok_or_else(|| ExcelError::SheetNotFound(sheet_name.to_string()))?;

    // Determine dimensions
    let cells = sheet.get_cell_collection();
    if cells.is_empty() {
        return Ok((vec![], 0, 0));
    }

    let mut max_row: usize = 0;
    let mut max_col: usize = 0;
    for cell in cells {
        let r = *cell.get_coordinate().get_row_num() as usize;
        let c = *cell.get_coordinate().get_col_num() as usize;
        if r > max_row { max_row = r; }
        if c > max_col { max_col = c; }
    }

    // Read all cells into a grid
    let mut grid: Vec<Vec<String>> = vec![vec![String::new(); max_col]; max_row];
    for cell in sheet.get_cell_collection() {
        let r = (*cell.get_coordinate().get_row_num() as usize) - 1;
        let c = (*cell.get_coordinate().get_col_num() as usize) - 1;
        let val = cell.get_value().to_string();
        grid[r][c] = val;
    }

    Ok((grid, max_row, max_col))
}

/// Export a sheet to CSV, returning the CSV content as a string
pub fn to_csv(path: &Path, sheet_name: &str, delimiter: char) -> ExcelResult<String> {
    let mut book = super::safe_io::safe_full_read(path)?;

    let (grid, _rows, _cols) = get_sheet_data(&mut book, sheet_name)?;

    let mut csv = String::new();
    for row in &grid {
        let line: Vec<String> = row
            .iter()
            .map(|cell| {
                // Quote cells that contain the delimiter, quotes, or newlines
                if cell.contains(delimiter) || cell.contains('"') || cell.contains('\n') {
                    format!("\"{}\"", cell.replace('"', "\"\""))
                } else {
                    cell.clone()
                }
            })
            .collect();
        csv.push_str(&line.join(&delimiter.to_string()));
        csv.push('\n');
    }

    Ok(csv)
}

/// Export a sheet to JSON
pub fn to_json(path: &Path, sheet_name: &str, orient: &str) -> ExcelResult<String> {
    let mut book = super::safe_io::safe_full_read(path)?;

    let (grid, rows, _cols) = get_sheet_data(&mut book, sheet_name)?;
    if rows == 0 {
        return Ok("[]".to_string());
    }

    let result = match orient {
        "records" => {
            // First row is headers
            let headers = grid[0].clone();
            let records: Vec<serde_json::Value> = grid[1..]
                .iter()
                .map(|row| {
                    let obj: serde_json::Map<String, serde_json::Value> = headers
                        .iter()
                        .enumerate()
                        .map(|(i, h)| {
                            let val = row.get(i).cloned().unwrap_or_default();
                            let json_val = if let Ok(n) = val.parse::<f64>() {
                                serde_json::Value::Number(
                                    serde_json::Number::from_f64(n)
                                        .unwrap_or(serde_json::Number::from(0)),
                                )
                            } else if val == "TRUE" || val == "true" {
                                serde_json::Value::Bool(true)
                            } else if val == "FALSE" || val == "false" {
                                serde_json::Value::Bool(false)
                            } else if val.is_empty() {
                                serde_json::Value::Null
                            } else {
                                serde_json::Value::String(val)
                            };
                            (h.clone(), json_val)
                        })
                        .collect();
                    serde_json::Value::Object(obj)
                })
                .collect();
            serde_json::to_string_pretty(&records)?
        }
        "values" => {
            let values: Vec<Vec<serde_json::Value>> = grid
                .iter()
                .map(|row| {
                    row.iter()
                        .map(|v| serde_json::Value::String(v.clone()))
                        .collect()
                })
                .collect();
            serde_json::to_string_pretty(&values)?
        }
        "columns" => {
            // First row is headers, columns as keys with arrays of values
            if grid.is_empty() {
                return Ok("{}".to_string());
            }
            let headers = &grid[0];
            let mut col_map: serde_json::Map<String, serde_json::Value> =
                serde_json::Map::new();
            for (i, h) in headers.iter().enumerate() {
                let col_vals: Vec<serde_json::Value> = grid[1..]
                    .iter()
                    .map(|row| {
                        let v = row.get(i).cloned().unwrap_or_default();
                        serde_json::Value::String(v)
                    })
                    .collect();
                col_map.insert(h.clone(), serde_json::Value::Array(col_vals));
            }
            serde_json::to_string_pretty(&serde_json::Value::Object(col_map))?
        }
        _ => {
            return Err(ExcelError::Other(format!("Unknown orient: {orient}")));
        }
    };

    Ok(result)
}

/// Export a sheet to an HTML table
pub fn to_html(path: &Path, sheet_name: &str) -> ExcelResult<String> {
    let mut book = super::safe_io::safe_full_read(path)?;

    let (grid, rows, cols) = get_sheet_data(&mut book, sheet_name)?;

    let mut html = String::new();
    html.push_str(&format!(
        "<!DOCTYPE html>\n<html>\n<head><meta charset=\"UTF-8\"><title>{sheet_name}</title></head>\n<body>\n"
    ));
    html.push_str(&format!(
        "<table border=\"1\" cellpadding=\"4\" cellspacing=\"0\">\n"
    ));

    // Header row with column letters
    html.push_str("  <thead>\n    <tr>\n      <th></th>\n");
    for c in 1..=cols {
        html.push_str(&format!(
            "      <th>{}</th>\n",
            col_index_to_letter(c as u32)
        ));
    }
    html.push_str("    </tr>\n  </thead>\n");

    html.push_str("  <tbody>\n");
    for (ri, row) in grid.iter().enumerate() {
        html.push_str(&format!("    <tr>\n      <th>{}</th>\n", ri + 1));
        for cell in row {
            let escaped = cell
                .replace('&', "&amp;")
                .replace('<', "&lt;")
                .replace('>', "&gt;")
                .replace('"', "&quot;");
            html.push_str(&format!("      <td>{escaped}</td>\n"));
        }
        html.push_str("    </tr>\n");
    }
    html.push_str("  </tbody>\n");
    html.push_str("</table>\n</body>\n</html>\n");

    let _ = rows; // suppress unused warning
    Ok(html)
}

/// Save content to a file, or return as string if no output path
pub fn save_or_return(content: String, output: Option<&str>) -> ExcelResult<String> {
    if let Some(out_path) = output {
        std::fs::write(out_path, &content)?;
        Ok(serde_json::json!({
            "status": "ok",
            "saved_to": out_path,
            "bytes": content.len(),
        })
        .to_string())
    } else {
        Ok(content)
    }
}
