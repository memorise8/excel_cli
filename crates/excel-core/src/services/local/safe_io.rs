//! Panic-safe wrappers for umya-spreadsheet I/O operations.
//!
//! All umya-spreadsheet calls that can panic (especially shared formula parsing)
//! are wrapped with `catch_unwind` and converted to structured errors.

use crate::models::error::*;
use std::panic;
use std::path::Path;

/// Panic-safe lazy_read — only loads workbook structure, NOT sheet contents.
/// Sheets must be individually deserialized via `safe_read_sheet()`.
pub fn safe_lazy_read(path: &Path) -> ExcelResult<umya_spreadsheet::Spreadsheet> {
    let path_buf = path.to_path_buf();
    let result = panic::catch_unwind(panic::AssertUnwindSafe(|| {
        umya_spreadsheet::reader::xlsx::lazy_read(&path_buf)
    }));
    match result {
        Ok(Ok(book)) => Ok(book),
        Ok(Err(e)) => Err(ExcelError::Spreadsheet(e.to_string())),
        Err(p) => Err(ExcelError::EnginePanic(format!("File open failed: {}", panic_msg(p)))),
    }
}

/// Panic-safe full read — loads and deserializes ALL sheets.
/// Use this when writer needs fully deserialized workbook (e.g. for write operations).
/// Falls back to lazy_read + target sheet only if full read panics.
pub fn safe_full_read(path: &Path) -> ExcelResult<umya_spreadsheet::Spreadsheet> {
    let path_buf = path.to_path_buf();
    let result = panic::catch_unwind(panic::AssertUnwindSafe(|| {
        umya_spreadsheet::reader::xlsx::read(&path_buf)
    }));
    match result {
        Ok(Ok(book)) => Ok(book),
        Ok(Err(e)) => Err(ExcelError::Spreadsheet(e.to_string())),
        Err(p) => Err(ExcelError::EnginePanic(format!(
            "File read failed (complex shared formula in workbook): {}. \
             Try using read-only operations instead.",
            panic_msg(p)
        ))),
    }
}

/// Panic-safe sheet deserialization — deserializes a single sheet by index.
pub fn safe_read_sheet(book: &mut umya_spreadsheet::Spreadsheet, idx: usize) -> ExcelResult<()> {
    let result = panic::catch_unwind(panic::AssertUnwindSafe(|| {
        book.read_sheet(idx);
    }));
    match result {
        Ok(()) => Ok(()),
        Err(p) => Err(ExcelError::EnginePanic(format!(
            "Sheet deserialization failed (complex shared formula): {}",
            panic_msg(p)
        ))),
    }
}

/// Panic-safe write — saves workbook to disk.
/// Does NOT deserialize unread sheets — writer copies them as raw data.
/// This avoids triggering shared formula panics in unrelated sheets.
pub fn safe_write(book: &umya_spreadsheet::Spreadsheet, path: &Path) -> ExcelResult<()> {
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

/// Find sheet index by name
pub fn find_sheet_index(book: &umya_spreadsheet::Spreadsheet, name: &str) -> ExcelResult<usize> {
    let sheets: Vec<(usize, String)> = book
        .get_sheet_collection()
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

fn panic_msg(p: Box<dyn std::any::Any + Send>) -> String {
    if let Some(s) = p.downcast_ref::<String>() {
        s.clone()
    } else if let Some(s) = p.downcast_ref::<&str>() {
        s.to_string()
    } else {
        "Unknown panic".to_string()
    }
}
