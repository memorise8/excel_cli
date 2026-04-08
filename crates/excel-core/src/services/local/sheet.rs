use crate::models::*;
use std::path::Path;

pub fn list(path: &Path) -> ExcelResult<Vec<SheetInfo>> {
    let book = match super::safe_io::safe_full_read(path) {
        Ok(b) => b,
        Err(ExcelError::EnginePanic(_)) => {
            // Fallback to calamine for sheet list
            let info = super::calamine_read::info(path)?;
            return Ok(info.sheets);
        }
        Err(e) => return Err(e),
    };

    let sheets = book
        .get_sheet_collection()
        .iter()
        .enumerate()
        .map(|(i, sheet)| SheetInfo {
            name: sheet.get_name().to_string(),
            index: i,
            visible: sheet.get_sheet_state() == "visible" || sheet.get_sheet_state().is_empty(),
            color: sheet.get_tab_color().map(|c| c.get_argb().to_string()),
            row_count: None,
            col_count: None,
        })
        .collect();

    Ok(sheets)
}

pub fn add(path: &Path, name: &str, _position: Option<usize>) -> ExcelResult<SheetInfo> {
    let mut book = super::safe_io::safe_full_read(path)?;

    book.new_sheet(name)
        .map_err(|e| ExcelError::SheetAlreadyExists(e.to_string()))?;

    let index = book.get_sheet_collection().len() - 1;

    super::safe_io::safe_write(&mut book, path)?;

    Ok(SheetInfo {
        name: name.to_string(),
        index,
        visible: true,
        color: None,
        row_count: Some(0),
        col_count: Some(0),
    })
}

pub fn rename(path: &Path, old_name: &str, new_name: &str) -> ExcelResult<()> {
    let mut book = super::safe_io::safe_full_read(path)?;

    let idx = super::safe_io::find_sheet_index(&book, old_name)?;
    let sheet = book.get_sheet_mut(&idx)
        .ok_or_else(|| ExcelError::SheetNotFound(old_name.to_string()))?;

    sheet.set_name(new_name);

    super::safe_io::safe_write(&mut book, path)?;
    Ok(())
}

pub fn delete(path: &Path, name: &str) -> ExcelResult<()> {
    let mut book = super::safe_io::safe_full_read(path)?;

    book.remove_sheet_by_name(name)
        .map_err(|e| ExcelError::Spreadsheet(e.to_string()))?;

    super::safe_io::safe_write(&mut book, path)?;
    Ok(())
}

pub fn copy(path: &Path, name: &str, new_name: Option<&str>) -> ExcelResult<SheetInfo> {
    let mut book = super::safe_io::safe_full_read(path)?;

    let copy_name = new_name.map(|s| s.to_string())
        .unwrap_or_else(|| format!("{name} (Copy)"));

    // Verify source exists
    let _source_idx = super::safe_io::find_sheet_index(&book, name)?;

    book.new_sheet(&copy_name)
        .map_err(|e| ExcelError::Spreadsheet(e.to_string()))?;

    let index = book.get_sheet_collection().len() - 1;

    super::safe_io::safe_write(&mut book, path)?;

    Ok(SheetInfo {
        name: copy_name,
        index,
        visible: true,
        color: None,
        row_count: Some(0),
        col_count: Some(0),
    })
}
