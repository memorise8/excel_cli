use crate::models::error::*;
use crate::models::range::{col_index_to_letter, parse_range_ref};
use crate::models::table::{TableData, TableInfo};
use std::path::Path;

/// List all tables in a workbook (across all sheets)
pub fn list(path: &Path) -> ExcelResult<Vec<TableInfo>> {
    let mut book = super::safe_io::safe_full_read(path)?;

    let sheet_count = book.get_sheet_collection().len();
    let mut result = Vec::new();

    for idx in 0..sheet_count {
        super::safe_io::safe_read_sheet(&mut book, idx)?;
        let sheet = match book.get_sheet(&idx) {
            Some(s) => s,
            None => continue,
        };
        let sheet_name = sheet.get_name().to_string();

        for table in sheet.get_tables() {
            let (start, end) = table.get_area();
            let sc = *start.get_col_num();
            let sr = *start.get_row_num();
            let ec = *end.get_col_num();
            let er = *end.get_row_num();

            let range_str = format!(
                "{}{}:{}{}",
                col_index_to_letter(sc),
                sr,
                col_index_to_letter(ec),
                er
            );

            let columns: Vec<String> = table
                .get_columns()
                .iter()
                .map(|c| c.get_name().to_string())
                .collect();

            // Data rows = total rows minus header row
            let row_count = if er >= sr + 1 { (er - sr) as usize } else { 0 };

            let style = table
                .get_style_info()
                .map(|s| s.get_name().to_string());

            result.push(TableInfo {
                name: table.get_name().to_string(),
                sheet: sheet_name.clone(),
                range: format!("{sheet_name}!{range_str}"),
                columns,
                row_count,
                has_total_row: *table.get_totals_row_count() > 0,
                style,
            });
        }
    }

    Ok(result)
}

/// Create an Excel table
pub fn create(
    path: &Path,
    range_str: &str,
    name: &str,
    style: Option<&str>,
    _has_headers: bool,
) -> ExcelResult<TableInfo> {
    let mut book = super::safe_io::safe_full_read(path)?;

    let addr = parse_range_ref(range_str).map_err(ExcelError::InvalidRange)?;
    let sheet_name = addr.sheet.as_deref().unwrap_or("Sheet1");
    let idx = super::safe_io::find_sheet_index(&book, sheet_name)?;
    super::safe_io::safe_read_sheet(&mut book, idx)?;

    let sheet = book
        .get_sheet_mut(&idx)
        .ok_or_else(|| ExcelError::SheetNotFound(sheet_name.to_string()))?;

    let mut table = umya_spreadsheet::Table::new(
        name,
        (
            (addr.start_col, addr.start_row),
            (addr.end_col, addr.end_row),
        ),
    );

    if let Some(s) = style {
        let style_info = umya_spreadsheet::TableStyleInfo::new(s, true, false, true, false);
        table.set_style_info(Some(style_info));
    }

    // Add default columns based on first row content (header row)
    // Since we don't read cell content here, generate column names Col1, Col2...
    let col_count = (addr.end_col - addr.start_col + 1) as usize;
    for i in 0..col_count {
        let col_name = col_index_to_letter(addr.start_col + i as u32);
        table.add_column(umya_spreadsheet::TableColumn::new(&col_name));
    }

    let row_count = if addr.end_row >= addr.start_row + 1 {
        (addr.end_row - addr.start_row) as usize
    } else {
        0
    };

    let range_display = format!(
        "{sheet_name}!{}{}:{}{}",
        col_index_to_letter(addr.start_col),
        addr.start_row,
        col_index_to_letter(addr.end_col),
        addr.end_row
    );

    let columns: Vec<String> = (0..col_count)
        .map(|i| col_index_to_letter(addr.start_col + i as u32))
        .collect();

    let table_style = style.map(|s| s.to_string());

    sheet.get_tables_mut().push(table);

    super::safe_io::safe_write(&mut book, path)?;

    Ok(TableInfo {
        name: name.to_string(),
        sheet: sheet_name.to_string(),
        range: range_display,
        columns,
        row_count,
        has_total_row: false,
        style: table_style,
    })
}

/// Read table data
pub fn read(path: &Path, table_name: &str) -> ExcelResult<TableData> {
    let mut book = super::safe_io::safe_full_read(path)?;

    let sheet_count = book.get_sheet_collection().len();

    for idx in 0..sheet_count {
        super::safe_io::safe_read_sheet(&mut book, idx)?;
        let sheet = match book.get_sheet(&idx) {
            Some(s) => s,
            None => continue,
        };
        let sheet_name = sheet.get_name().to_string();

        // Find the table
        let maybe_table = sheet.get_tables().iter().find(|t| t.get_name() == table_name);
        if let Some(table) = maybe_table {
            let (start, end) = table.get_area();
            let sc = *start.get_col_num();
            let sr = *start.get_row_num();
            let ec = *end.get_col_num();
            let er = *end.get_row_num();

            let range_str = format!(
                "{sheet_name}!{}{}:{}{}",
                col_index_to_letter(sc),
                sr,
                col_index_to_letter(ec),
                er
            );

            let columns: Vec<String> = table
                .get_columns()
                .iter()
                .map(|c| c.get_name().to_string())
                .collect();

            let row_count = if er >= sr + 1 { (er - sr) as usize } else { 0 };
            let style = table.get_style_info().map(|s| s.get_name().to_string());

            // Read actual cell data starting from row sr+1 (data rows, skip header)
            let mut data_rows: Vec<Vec<serde_json::Value>> = Vec::new();
            for row in (sr + 1)..=er {
                let mut row_data = Vec::new();
                for col in sc..=ec {
                    let val = sheet
                        .get_cell((col, row))
                        .map(|c| {
                            let v = c.get_value().to_string();
                            if let Ok(n) = v.parse::<f64>() {
                                serde_json::Value::Number(
                                    serde_json::Number::from_f64(n).unwrap_or(serde_json::Number::from(0)),
                                )
                            } else {
                                serde_json::Value::String(v)
                            }
                        })
                        .unwrap_or(serde_json::Value::Null);
                    row_data.push(val);
                }
                data_rows.push(row_data);
            }

            let info = TableInfo {
                name: table_name.to_string(),
                sheet: sheet_name.clone(),
                range: range_str,
                columns: columns.clone(),
                row_count,
                has_total_row: false,
                style,
            };

            return Ok(TableData {
                info,
                headers: columns,
                rows: data_rows,
            });
        }
    }

    Err(ExcelError::TableNotFound(table_name.to_string()))
}

/// Append rows to a table
pub fn append(path: &Path, table_name: &str, rows: Vec<Vec<serde_json::Value>>) -> ExcelResult<()> {
    let mut book = super::safe_io::safe_full_read(path)?;

    let sheet_count = book.get_sheet_collection().len();

    // Find which sheet has the table and its area
    let mut found: Option<(usize, u32, u32, u32, u32)> = None;
    for idx in 0..sheet_count {
        super::safe_io::safe_read_sheet(&mut book, idx)?;
        let sheet = match book.get_sheet(&idx) {
            Some(s) => s,
            None => continue,
        };
        if let Some(table) = sheet.get_tables().iter().find(|t| t.get_name() == table_name) {
            let (start, end) = table.get_area();
            found = Some((
                idx,
                *start.get_col_num(),
                *start.get_row_num(),
                *end.get_col_num(),
                *end.get_row_num(),
            ));
            break;
        }
    }

    let (sheet_idx, sc, _sr, ec, er) =
        found.ok_or_else(|| ExcelError::TableNotFound(table_name.to_string()))?;

    let sheet = book
        .get_sheet_mut(&sheet_idx)
        .ok_or_else(|| ExcelError::SheetNotFound("unknown".to_string()))?;

    // Write rows starting at er+1
    for (ri, row_data) in rows.iter().enumerate() {
        let row = er + 1 + ri as u32;
        for (ci, val) in row_data.iter().enumerate() {
            let col = sc + ci as u32;
            if col > ec {
                break;
            }
            let cell = sheet.get_cell_mut((col, row));
            match val {
                serde_json::Value::String(s) => { cell.set_value(s); }
                serde_json::Value::Number(n) => { cell.set_value(n.to_string()); }
                serde_json::Value::Bool(b) => { cell.set_value(if *b { "TRUE" } else { "FALSE" }); }
                serde_json::Value::Null => { cell.set_value(""); }
                other => { cell.set_value(other.to_string()); }
            }
        }
    }

    // Update table area to include new rows
    let new_end_row = er + rows.len() as u32;
    if let Some(table) = sheet
        .get_tables_mut()
        .iter_mut()
        .find(|t| t.get_name() == table_name)
    {
        table.set_area(((sc, _sr), (ec, new_end_row)));
    }

    super::safe_io::safe_write(&mut book, path)?;
    Ok(())
}
