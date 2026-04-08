//! Calamine-based read-only operations — safe fallback for files that
//! cause umya-spreadsheet to panic on shared formula parsing.

use crate::models::error::*;
use crate::models::range::*;
use crate::models::workbook::*;
use calamine::{Reader, Xlsx, Data};
use std::path::Path;

fn open_xlsx(path: &Path) -> ExcelResult<Xlsx<std::io::BufReader<std::fs::File>>> {
    let wb: Result<Xlsx<std::io::BufReader<std::fs::File>>, calamine::XlsxError> =
        calamine::open_workbook(path);
    wb.map_err(|e| ExcelError::Calamine(e.to_string()))
}

pub fn info(path: &Path) -> ExcelResult<WorkbookInfo> {
    let mut workbook = open_xlsx(path)?;

    let file_size = std::fs::metadata(path)?.len();
    let file_name = path
        .file_name()
        .map(|n| n.to_string_lossy().to_string())
        .unwrap_or_default();

    let sheet_names = workbook.sheet_names().to_vec();
    let mut sheets_info = Vec::new();

    for (i, name) in sheet_names.iter().enumerate() {
        let (row_count, col_count) = if let Ok(range) = workbook.worksheet_range(name) {
            range.get_size()
        } else {
            (0, 0)
        };

        sheets_info.push(SheetInfo {
            name: name.clone(),
            index: i,
            visible: true,
            color: None,
            row_count: Some(row_count),
            col_count: Some(col_count),
        });
    }

    Ok(WorkbookInfo {
        path: path.display().to_string(),
        file_name,
        file_size,
        sheet_count: sheets_info.len(),
        sheets: sheets_info,
    })
}

pub fn range_read(path: &Path, range_str: &str) -> ExcelResult<RangeData> {
    let mut workbook = open_xlsx(path)?;

    let addr = parse_range_ref(range_str)
        .map_err(ExcelError::InvalidRange)?;

    let sheet_name = addr.sheet.as_deref().unwrap_or("Sheet1");

    let range = workbook
        .worksheet_range(sheet_name)
        .map_err(|e| ExcelError::Calamine(e.to_string()))?;

    // calamine Range::get() uses relative position from range.start()
    // range.start() returns 0-based (row, col) of the first cell with data
    let (start_r, start_c) = range.start().unwrap_or((0, 0));

    let mut rows = Vec::new();

    for row_idx in addr.start_row..=addr.end_row {
        let mut row_data = Vec::new();
        for col_idx in addr.start_col..=addr.end_col {
            // Convert 1-based absolute coords to 0-based relative to range start
            let abs_r = (row_idx - 1) as u32;
            let abs_c = (col_idx - 1) as u32;

            let cell_value = if abs_r >= start_r && abs_c >= start_c {
                let rel_r = (abs_r - start_r) as usize;
                let rel_c = (abs_c - start_c) as usize;
                match range.get((rel_r, rel_c)) {
                    Some(data) => data_to_cell_value(data),
                    None => CellValue::Empty,
                }
            } else {
                CellValue::Empty
            };
            row_data.push(cell_value);
        }
        rows.push(row_data);
    }

    Ok(RangeData {
        range: range_str.to_string(),
        sheet: sheet_name.to_string(),
        rows,
        row_count: addr.row_count() as usize,
        col_count: addr.col_count() as usize,
    })
}

fn data_to_cell_value(data: &Data) -> CellValue {
    match data {
        Data::Empty => CellValue::Empty,
        Data::Bool(b) => CellValue::Bool(*b),
        Data::Int(i) => CellValue::Int(*i),
        Data::Float(f) => CellValue::Float(*f),
        Data::String(s) => CellValue::String(s.clone()),
        Data::DateTime(dt) => CellValue::String(format!("{dt:?}")),
        Data::DateTimeIso(s) => CellValue::String(s.clone()),
        Data::DurationIso(s) => CellValue::String(s.clone()),
        Data::Error(e) => CellValue::Error(format!("{e:?}")),
    }
}
