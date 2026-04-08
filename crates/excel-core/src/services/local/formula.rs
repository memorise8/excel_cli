use crate::models::error::*;
use crate::models::range::{col_index_to_letter, parse_range_ref};
use std::panic;
use std::path::Path;

fn find_sheet_index(book: &umya_spreadsheet::Spreadsheet, name: &str) -> ExcelResult<usize> {
    book.get_sheet_collection()
        .iter()
        .position(|s| s.get_name() == name)
        .ok_or_else(|| ExcelError::SheetNotFound(name.to_string()))
}

/// Panic-safe lazy_read
fn _safe_lazy_read_unused(path: &Path) -> ExcelResult<umya_spreadsheet::Spreadsheet> {
    let path_buf = path.to_path_buf();
    let result = panic::catch_unwind(panic::AssertUnwindSafe(|| {
        umya_spreadsheet::reader::xlsx::lazy_read(&path_buf)
    }));
    match result {
        Ok(Ok(book)) => Ok(book),
        Ok(Err(e)) => Err(ExcelError::Spreadsheet(e.to_string())),
        Err(p) => {
            let msg = panic_msg(p);
            Err(ExcelError::EnginePanic(format!("File open failed: {msg}")))
        }
    }
}

fn safe_read_sheet(book: &mut umya_spreadsheet::Spreadsheet, idx: usize) -> ExcelResult<()> {
    let result = panic::catch_unwind(panic::AssertUnwindSafe(|| {
        book.read_sheet(idx);
    }));
    match result {
        Ok(()) => Ok(()),
        Err(p) => {
            let msg = panic_msg(p);
            Err(ExcelError::EnginePanic(format!(
                "Sheet deserialization failed (complex shared formula): {msg}"
            )))
        }
    }
}

fn safe_write(book: &umya_spreadsheet::Spreadsheet, path: &Path) -> ExcelResult<()> {
    let path_buf = path.to_path_buf();
    let result = panic::catch_unwind(panic::AssertUnwindSafe(|| {
        umya_spreadsheet::writer::xlsx::write(book, &path_buf)
    }));
    match result {
        Ok(Ok(())) => Ok(()),
        Ok(Err(e)) => Err(ExcelError::Spreadsheet(format!("Save failed: {e}"))),
        Err(p) => Err(ExcelError::EnginePanic(format!("Save panicked: {}", panic_msg(p)))),
    }
}

fn panic_msg(p: Box<dyn std::any::Any + Send>) -> String {
    if let Some(s) = p.downcast_ref::<String>() {
        s.clone()
    } else if let Some(s) = p.downcast_ref::<&str>() {
        s.to_string()
    } else {
        "Unknown panic".to_string()
    }
}

/// Read formula text from a cell — panic-safe
pub fn read(path: &Path, range_str: &str) -> ExcelResult<serde_json::Value> {
    let mut book = super::safe_io::safe_full_read(path)?;

    let addr = parse_range_ref(range_str).map_err(ExcelError::InvalidRange)?;
    let sheet_name = addr.sheet.as_deref().unwrap_or("Sheet1");
    let idx = find_sheet_index(&book, sheet_name)?;
    super::safe_io::safe_read_sheet(&mut book, idx)?;

    let sheet = book
        .get_sheet(&idx)
        .ok_or_else(|| ExcelError::SheetNotFound(sheet_name.to_string()))?;

    let cell = sheet.get_cell((addr.start_col, addr.start_row));

    // Panic-safe formula access
    let formula = cell
        .and_then(|c| {
            panic::catch_unwind(panic::AssertUnwindSafe(|| c.get_formula().to_string())).ok()
        })
        .unwrap_or_default();

    let value = cell
        .map(|c| c.get_value().to_string())
        .unwrap_or_default();

    Ok(serde_json::json!({
        "cell": range_str,
        "formula": formula,
        "value": value,
        "has_formula": !formula.is_empty(),
    }))
}

/// Write a formula to a cell — panic-safe with lazy_read
pub fn write(path: &Path, range_str: &str, formula: &str) -> ExcelResult<()> {
    let mut book = super::safe_io::safe_full_read(path)?;

    let addr = parse_range_ref(range_str).map_err(ExcelError::InvalidRange)?;
    let sheet_name = addr.sheet.as_deref().unwrap_or("Sheet1");
    let idx = find_sheet_index(&book, sheet_name)?;
    super::safe_io::safe_read_sheet(&mut book, idx)?;

    let sheet = book
        .get_sheet_mut(&idx)
        .ok_or_else(|| ExcelError::SheetNotFound(sheet_name.to_string()))?;

    sheet
        .get_cell_mut((addr.start_col, addr.start_row))
        .set_formula(formula);

    super::safe_io::safe_write(&book, path)?;
    Ok(())
}

/// List all formulas in a sheet — best-effort, skips cells that panic
pub fn list(path: &Path, sheet_name: &str) -> ExcelResult<serde_json::Value> {
    let mut book = super::safe_io::safe_full_read(path)?;

    let idx = find_sheet_index(&book, sheet_name)?;
    super::safe_io::safe_read_sheet(&mut book, idx)?;

    let sheet = book
        .get_sheet(&idx)
        .ok_or_else(|| ExcelError::SheetNotFound(sheet_name.to_string()))?;

    let mut formulas = Vec::new();
    let mut skipped = 0u32;

    for cell in sheet.get_cell_collection() {
        // Panic-safe per-cell formula access
        let formula_result = panic::catch_unwind(panic::AssertUnwindSafe(|| {
            cell.get_formula().to_string()
        }));

        match formula_result {
            Ok(formula) if !formula.is_empty() => {
                let col = *cell.get_coordinate().get_col_num();
                let row = *cell.get_coordinate().get_row_num();
                let col_letter = col_index_to_letter(col);
                formulas.push(serde_json::json!({
                    "cell": format!("{col_letter}{row}"),
                    "col": col,
                    "row": row,
                    "formula": formula,
                    "value": cell.get_value().to_string(),
                }));
            }
            Ok(_) => {} // empty formula, skip
            Err(_) => {
                skipped += 1;
                // Cell formula caused panic — skip it, continue with others
            }
        }
    }

    Ok(serde_json::json!({
        "sheet": sheet_name,
        "count": formulas.len(),
        "skipped_unsupported": skipped,
        "formulas": formulas,
    }))
}

/// Audit formula precedents — panic-safe
pub fn audit(path: &Path, range_str: &str, direction: &str) -> ExcelResult<serde_json::Value> {
    let mut book = super::safe_io::safe_full_read(path)?;

    let addr = parse_range_ref(range_str).map_err(ExcelError::InvalidRange)?;
    let sheet_name = addr.sheet.as_deref().unwrap_or("Sheet1");
    let idx = find_sheet_index(&book, sheet_name)?;
    super::safe_io::safe_read_sheet(&mut book, idx)?;

    let sheet = book
        .get_sheet(&idx)
        .ok_or_else(|| ExcelError::SheetNotFound(sheet_name.to_string()))?;

    let col_letter = col_index_to_letter(addr.start_col);
    let cell_ref = format!("{col_letter}{}", addr.start_row);

    let formula = sheet
        .get_cell((addr.start_col, addr.start_row))
        .and_then(|c| {
            panic::catch_unwind(panic::AssertUnwindSafe(|| c.get_formula().to_string())).ok()
        })
        .unwrap_or_default();

    let precedents = if direction == "precedents" && !formula.is_empty() {
        extract_cell_refs(&formula)
    } else {
        let target = cell_ref.to_uppercase();
        let mut deps = Vec::new();
        for cell in sheet.get_cell_collection() {
            // Panic-safe per-cell
            let f = panic::catch_unwind(panic::AssertUnwindSafe(|| {
                cell.get_formula().to_string()
            }))
            .unwrap_or_default();

            if !f.is_empty() && f.to_uppercase().contains(&target) {
                let c = *cell.get_coordinate().get_col_num();
                let r = *cell.get_coordinate().get_row_num();
                deps.push(format!("{}{r}", col_index_to_letter(c)));
            }
        }
        deps
    };

    Ok(serde_json::json!({
        "cell": range_str,
        "formula": formula,
        "direction": direction,
        "references": precedents,
        "count": precedents.len(),
    }))
}

/// Extract cell references like A1, B2, AA10 from a formula string
fn extract_cell_refs(formula: &str) -> Vec<String> {
    let mut refs = Vec::new();
    let chars: Vec<char> = formula.chars().collect();
    let mut i = 0;
    while i < chars.len() {
        if chars[i] == '"' {
            i += 1;
            while i < chars.len() && chars[i] != '"' {
                i += 1;
            }
            i += 1;
            continue;
        }
        if chars[i].is_ascii_alphabetic() {
            let start = i;
            while i < chars.len() && chars[i].is_ascii_alphabetic() {
                i += 1;
            }
            if i < chars.len() && chars[i].is_ascii_digit() {
                let col_part: String = chars[start..i].iter().collect();
                let row_start = i;
                while i < chars.len() && chars[i].is_ascii_digit() {
                    i += 1;
                }
                let row_part: String = chars[row_start..i].iter().collect();
                if i < chars.len() && chars[i] == '(' {
                    // function name, skip
                } else {
                    refs.push(format!("{col_part}{row_part}").to_uppercase());
                }
            }
        } else {
            i += 1;
        }
    }
    refs.sort();
    refs.dedup();
    refs
}
