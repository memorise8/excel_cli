use crate::models::*;
use std::panic;
use std::path::Path;

pub fn create(path: &Path, sheets: Option<Vec<String>>) -> ExcelResult<WorkbookInfo> {
    let mut book = umya_spreadsheet::new_file();

    if let Some(sheet_names) = &sheets {
        if let Some(first_sheet) = book.get_sheet_mut(&0) {
            first_sheet.set_name(sheet_names[0].clone());
        }
        for name in sheet_names.iter().skip(1) {
            book.new_sheet(name)
                .map_err(|e| ExcelError::Spreadsheet(e.to_string()))?;
        }
    }

    umya_spreadsheet::writer::xlsx::write(&book, path)
        .map_err(|e| ExcelError::Spreadsheet(e.to_string()))?;

    // Return info from in-memory book directly
    let file_size = std::fs::metadata(path).map(|m| m.len()).unwrap_or(0);
    let file_name = path
        .file_name()
        .map(|n| n.to_string_lossy().to_string())
        .unwrap_or_default();

    let sheets_info: Vec<SheetInfo> = book
        .get_sheet_collection()
        .iter()
        .enumerate()
        .map(|(i, sheet)| SheetInfo {
            name: sheet.get_name().to_string(),
            index: i,
            visible: true,
            color: None,
            row_count: Some(0),
            col_count: Some(0),
        })
        .collect();

    Ok(WorkbookInfo {
        path: path.display().to_string(),
        file_name,
        file_size,
        sheet_count: sheets_info.len(),
        sheets: sheets_info,
    })
}

pub fn info(path: &Path) -> ExcelResult<WorkbookInfo> {
    if !path.exists() {
        return Err(ExcelError::FileNotFound(path.display().to_string()));
    }

    let file_size = std::fs::metadata(path)?.len();
    let file_name = path
        .file_name()
        .map(|n| n.to_string_lossy().to_string())
        .unwrap_or_default();

    // Try umya full read first, fall back to calamine on panic
    let book = match super::safe_io::safe_full_read(path) {
        Ok(b) => b,
        Err(ExcelError::EnginePanic(_)) => {
            // umya panicked — use calamine fallback for info
            return super::calamine_read::info(path);
        }
        Err(e) => return Err(e),
    };

    let mut sheets_info = Vec::new();
    for (i, sheet) in book.get_sheet_collection().iter().enumerate() {
        let name = sheet.get_name().to_string();
        let (row_count, col_count) = get_sheet_dimensions_safe(sheet);
        let visible = panic::catch_unwind(panic::AssertUnwindSafe(|| {
            let state = sheet.get_sheet_state();
            state == "visible" || state.is_empty()
        }))
        .unwrap_or(true);
        let color = panic::catch_unwind(panic::AssertUnwindSafe(|| {
            sheet.get_tab_color().map(|c| c.get_argb().to_string())
        }))
        .unwrap_or(None);

        sheets_info.push(SheetInfo {
            name,
            index: i,
            visible,
            color,
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

pub fn save(path: &Path, output: &Path) -> ExcelResult<()> {
    if !path.exists() {
        return Err(ExcelError::FileNotFound(path.display().to_string()));
    }

    let book = super::safe_io::safe_lazy_read(path)?;
    super::safe_io::safe_write(&book, output)?;
    Ok(())
}

fn get_sheet_dimensions_safe(sheet: &umya_spreadsheet::Worksheet) -> (usize, usize) {
    panic::catch_unwind(panic::AssertUnwindSafe(|| {
        let cells = sheet.get_cell_collection();
        if cells.len() == 0 {
            return (0, 0);
        }
        let mut max_row: usize = 0;
        let mut max_col: usize = 0;
        for cell in cells {
            let r = *cell.get_coordinate().get_row_num() as usize;
            let c = *cell.get_coordinate().get_col_num() as usize;
            if r > max_row { max_row = r; }
            if c > max_col { max_col = c; }
        }
        (max_row, max_col)
    }))
    .unwrap_or((0, 0))
}
