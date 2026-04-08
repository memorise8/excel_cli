use crate::models::range::*;
use crate::models::error::*;
use std::path::Path;
use std::panic;

/// Find sheet index by name
fn find_sheet_index(book: &umya_spreadsheet::Spreadsheet, name: &str) -> ExcelResult<usize> {
    let sheets: Vec<(usize, String)> = book.get_sheet_collection()
        .iter()
        .enumerate()
        .map(|(i, s)| (i, s.get_name().to_string()))
        .collect();
    for (i, sname) in &sheets {
        if sname == name {
            return Ok(*i);
        }
    }
    Err(ExcelError::SheetNotFound(format!(
        "'{}' (available: {:?})",
        name,
        sheets.iter().map(|(_, n)| n.as_str()).collect::<Vec<_>>()
    )))
}

/// Panic-safe wrapper for umya-spreadsheet lazy_read
fn _safe_lazy_read_unused(path: &Path) -> ExcelResult<umya_spreadsheet::Spreadsheet> {
    let path_buf = path.to_path_buf();
    let result = panic::catch_unwind(panic::AssertUnwindSafe(|| {
        umya_spreadsheet::reader::xlsx::lazy_read(&path_buf)
    }));
    match result {
        Ok(Ok(book)) => Ok(book),
        Ok(Err(e)) => Err(ExcelError::Spreadsheet(e.to_string())),
        Err(panic_info) => {
            let msg = if let Some(s) = panic_info.downcast_ref::<String>() {
                s.clone()
            } else if let Some(s) = panic_info.downcast_ref::<&str>() {
                s.to_string()
            } else {
                "Unknown panic during file read".to_string()
            };
            Err(ExcelError::EnginePanic(format!("Failed to open file: {msg}")))
        }
    }
}

/// Panic-safe wrapper for read_sheet (deserializes a single sheet)
fn safe_read_sheet(book: &mut umya_spreadsheet::Spreadsheet, idx: usize) -> ExcelResult<()> {
    let result = panic::catch_unwind(panic::AssertUnwindSafe(|| {
        book.read_sheet(idx);
    }));
    match result {
        Ok(()) => Ok(()),
        Err(panic_info) => {
            let msg = if let Some(s) = panic_info.downcast_ref::<String>() {
                s.clone()
            } else if let Some(s) = panic_info.downcast_ref::<&str>() {
                s.to_string()
            } else {
                "Unknown panic during sheet deserialization".to_string()
            };
            Err(ExcelError::EnginePanic(format!(
                "Sheet deserialization failed (likely complex shared formula): {msg}"
            )))
        }
    }
}

/// Panic-safe wrapper for umya-spreadsheet write
fn safe_write(book: &umya_spreadsheet::Spreadsheet, path: &Path) -> ExcelResult<()> {
    let path_buf = path.to_path_buf();
    let result = panic::catch_unwind(panic::AssertUnwindSafe(|| {
        umya_spreadsheet::writer::xlsx::write(book, &path_buf)
    }));
    match result {
        Ok(Ok(())) => Ok(()),
        Ok(Err(e)) => Err(ExcelError::Spreadsheet(format!("Save failed: {e}"))),
        Err(panic_info) => {
            let msg = if let Some(s) = panic_info.downcast_ref::<String>() {
                s.clone()
            } else if let Some(s) = panic_info.downcast_ref::<&str>() {
                s.to_string()
            } else {
                "Unknown panic during file save".to_string()
            };
            Err(ExcelError::EnginePanic(format!("Save panicked: {msg}")))
        }
    }
}

pub fn read(path: &Path, range_str: &str) -> ExcelResult<RangeData> {
    let mut book = match super::safe_io::safe_full_read(path) {
        Ok(b) => b,
        Err(ExcelError::EnginePanic(_)) => {
            // umya panicked — use calamine fallback for read
            return super::calamine_read::range_read(path, range_str);
        }
        Err(e) => return Err(e),
    };

    let addr = parse_range_ref(range_str)
        .map_err(ExcelError::InvalidRange)?;

    let sheet_name = addr.sheet.as_deref().unwrap_or("Sheet1");
    let idx = find_sheet_index(&book, sheet_name)?;

    super::safe_io::safe_read_sheet(&mut book, idx)?;

    let sheet = book.get_sheet(&idx)
        .ok_or_else(|| ExcelError::SheetNotFound(sheet_name.to_string()))?;

    let mut rows = Vec::new();

    for row in addr.start_row..=addr.end_row {
        let mut row_data = Vec::new();
        for col in addr.start_col..=addr.end_col {
            let cell_value = match sheet.get_cell((col, row)) {
                Some(cell) => {
                    // Panic-safe formula access
                    let formula = panic::catch_unwind(panic::AssertUnwindSafe(|| {
                        cell.get_formula().to_string()
                    })).unwrap_or_default();

                    if !formula.is_empty() {
                        CellValue::Formula(FormulaValue {
                            formula,
                            cached_value: Some(Box::new(raw_value_to_cell_value(cell))),
                        })
                    } else {
                        raw_value_to_cell_value(cell)
                    }
                }
                None => CellValue::Empty,
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

/// Value-only write path — uses lazy_read to avoid deserializing unrelated sheets.
/// Only the target sheet is deserialized, minimizing exposure to formula parsing panics.
pub fn write(path: &Path, range_str: &str, data: Vec<Vec<CellValue>>) -> ExcelResult<()> {
    let mut book = super::safe_io::safe_full_read(path)?;

    let addr = parse_range_ref(range_str)
        .map_err(ExcelError::InvalidRange)?;

    let sheet_name = addr.sheet.as_deref().unwrap_or("Sheet1");
    let idx = find_sheet_index(&book, sheet_name)?;

    super::safe_io::safe_read_sheet(&mut book, idx)?;

    let sheet = book.get_sheet_mut(&idx)
        .ok_or_else(|| ExcelError::SheetNotFound(sheet_name.to_string()))?;

    for (ri, row_data) in data.iter().enumerate() {
        for (ci, cell_value) in row_data.iter().enumerate() {
            let row = addr.start_row + ri as u32;
            let col = addr.start_col + ci as u32;
            write_cell_value(sheet, col, row, cell_value);
        }
    }

    super::safe_io::safe_write(&book, path)?;
    Ok(())
}

pub fn clear(path: &Path, range_str: &str, _values_only: bool) -> ExcelResult<()> {
    let mut book = super::safe_io::safe_full_read(path)?;

    let addr = parse_range_ref(range_str)
        .map_err(ExcelError::InvalidRange)?;

    let sheet_name = addr.sheet.as_deref().unwrap_or("Sheet1");
    let idx = find_sheet_index(&book, sheet_name)?;

    super::safe_io::safe_read_sheet(&mut book, idx)?;

    let sheet = book.get_sheet_mut(&idx)
        .ok_or_else(|| ExcelError::SheetNotFound(sheet_name.to_string()))?;

    for row in addr.start_row..=addr.end_row {
        for col in addr.start_col..=addr.end_col {
            sheet.get_cell_mut((col, row)).set_value("");
        }
    }

    super::safe_io::safe_write(&book, path)?;
    Ok(())
}

fn raw_value_to_cell_value(cell: &umya_spreadsheet::Cell) -> CellValue {
    let value = cell.get_value();
    if value.is_empty() {
        return CellValue::Empty;
    }

    if let Ok(i) = value.parse::<i64>() {
        return CellValue::Int(i);
    }
    if let Ok(f) = value.parse::<f64>() {
        return CellValue::Float(f);
    }

    match value.to_lowercase().as_str() {
        "true" => CellValue::Bool(true),
        "false" => CellValue::Bool(false),
        _ => CellValue::String(value.to_string()),
    }
}

fn write_cell_value(sheet: &mut umya_spreadsheet::Worksheet, col: u32, row: u32, value: &CellValue) {
    let cell = sheet.get_cell_mut((col, row));
    match value {
        CellValue::Empty => { cell.set_value(""); }
        CellValue::Bool(b) => { cell.set_value(if *b { "TRUE" } else { "FALSE" }); }
        CellValue::Int(i) => { cell.set_value(i.to_string()); }
        CellValue::Float(f) => { cell.set_value(f.to_string()); }
        CellValue::String(s) => { cell.set_value(s); }
        CellValue::Formula(fv) => { cell.set_formula(&fv.formula); }
        CellValue::Error(e) => { cell.set_value(e); }
    }
}
