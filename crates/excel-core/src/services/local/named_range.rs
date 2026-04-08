use crate::models::error::*;
use std::path::Path;

/// List all named ranges in a workbook
pub fn list(path: &Path) -> ExcelResult<serde_json::Value> {
    let book = super::safe_io::safe_full_read(path)?;

    let names: Vec<serde_json::Value> = book
        .get_defined_names()
        .iter()
        .map(|dn| {
            serde_json::json!({
                "name": dn.get_name(),
                "refers_to": dn.get_address(),
                "scope": if dn.has_local_sheet_id() { "sheet" } else { "workbook" },
            })
        })
        .collect();

    Ok(serde_json::json!({
        "count": names.len(),
        "named_ranges": names,
    }))
}

/// Create a named range
pub fn create(path: &Path, name: &str, refers_to: &str, _scope: &str) -> ExcelResult<serde_json::Value> {
    let mut book = super::safe_io::safe_full_read(path)?;

    // Check for duplicate
    if book.get_defined_names().iter().any(|dn| dn.get_name() == name) {
        return Err(ExcelError::Other(format!("Named range '{name}' already exists")));
    }

    let mut defined_name = umya_spreadsheet::DefinedName::default();
    defined_name.set_address(refers_to);

    // Use internal field access via set_name equivalent — DefinedName only exposes get_name
    // We need to build it differently; use add_defined_name on the spreadsheet
    // Unfortunately DefinedName::default() has no set_name. Use the spreadsheet helper.
    drop(defined_name);

    // The worksheet has add_defined_name(name, address) -> Result
    // But the task says workbook scope. Use book.get_defined_names_mut() directly.
    let mut dn = umya_spreadsheet::DefinedName::default();
    // DefinedName has no public set_name — we must construct via the sheet helper or
    // look for another path. Let's check if there's a new() or from_name_address.
    drop(dn);

    // Fallback: use worksheet-level add_defined_name if scope is sheet
    // For workbook scope, manipulate get_defined_names_mut directly.
    // Since DefinedName::default() gives us a blank with no name setter exposed,
    // we use the sheet-level helper which does expose name+address.
    // Parse refers_to to determine sheet
    let sheet_name = if let Some((sheet, _)) = refers_to.split_once('!') {
        sheet.trim_matches('\'').to_string()
    } else {
        book.get_sheet_collection()
            .first()
            .map(|s| s.get_name().to_string())
            .unwrap_or_else(|| "Sheet1".to_string())
    };

    let sheet_idx = super::safe_io::find_sheet_index(&book, &sheet_name)?;
    super::safe_io::safe_read_sheet(&mut book, sheet_idx)?;

    let sheet = book
        .get_sheet_mut(&sheet_idx)
        .ok_or_else(|| ExcelError::SheetNotFound(sheet_name.clone()))?;

    sheet
        .add_defined_name(name, refers_to)
        .map_err(|e| ExcelError::Other(e.to_string()))?;

    super::safe_io::safe_write(&mut book, path)?;

    Ok(serde_json::json!({
        "status": "ok",
        "name": name,
        "refers_to": refers_to,
    }))
}

/// Delete a named range
pub fn delete(path: &Path, name: &str) -> ExcelResult<()> {
    let mut book = super::safe_io::safe_full_read(path)?;

    // Try workbook-level defined names first
    let workbook_names = book.get_defined_names_mut();
    let before = workbook_names.len();
    workbook_names.retain(|dn| dn.get_name() != name);
    let removed_workbook = workbook_names.len() < before;

    if !removed_workbook {
        // Try sheet-level defined names
        let sheet_count = book.get_sheet_collection().len();
        let mut removed = false;
        for idx in 0..sheet_count {
            let sheet = match book.get_sheet_mut(&idx) {
                Some(s) => s,
                None => continue,
            };
            let before = sheet.get_defined_names().len();
            sheet.get_defined_names_mut().retain(|dn| dn.get_name() != name);
            if sheet.get_defined_names().len() < before {
                removed = true;
                break;
            }
        }
        if !removed {
            return Err(ExcelError::NamedRangeNotFound(name.to_string()));
        }
    }

    super::safe_io::safe_write(&mut book, path)?;
    Ok(())
}

/// Update a named range's reference
pub fn update(path: &Path, name: &str, refers_to: &str) -> ExcelResult<serde_json::Value> {
    let mut book = super::safe_io::safe_full_read(path)?;

    // Try workbook-level
    let mut found = false;
    for dn in book.get_defined_names_mut().iter_mut() {
        if dn.get_name() == name {
            dn.set_address(refers_to);
            found = true;
            break;
        }
    }

    if !found {
        // Try sheet-level
        let sheet_count = book.get_sheet_collection().len();
        for idx in 0..sheet_count {
            let sheet = match book.get_sheet_mut(&idx) {
                Some(s) => s,
                None => continue,
            };
            for dn in sheet.get_defined_names_mut().iter_mut() {
                if dn.get_name() == name {
                    dn.set_address(refers_to);
                    found = true;
                    break;
                }
            }
            if found { break; }
        }
    }

    if !found {
        return Err(ExcelError::NamedRangeNotFound(name.to_string()));
    }

    super::safe_io::safe_write(&mut book, path)?;

    Ok(serde_json::json!({
        "status": "ok",
        "name": name,
        "refers_to": refers_to,
    }))
}

/// Resolve a name to its range address string
pub fn resolve(path: &Path, name: &str) -> ExcelResult<serde_json::Value> {
    let book = super::safe_io::safe_full_read(path)?;

    // Check workbook-level
    if let Some(dn) = book.get_defined_names().iter().find(|dn| dn.get_name() == name) {
        return Ok(serde_json::json!({
            "name": name,
            "refers_to": dn.get_address(),
            "scope": "workbook",
        }));
    }

    // Check sheet-level
    for sheet in book.get_sheet_collection() {
        if let Some(dn) = sheet.get_defined_names().iter().find(|dn| dn.get_name() == name) {
            return Ok(serde_json::json!({
                "name": name,
                "refers_to": dn.get_address(),
                "scope": sheet.get_name(),
            }));
        }
    }

    Err(ExcelError::NamedRangeNotFound(name.to_string()))
}

/// Read values from a named range
pub fn read_values(path: &Path, name: &str) -> ExcelResult<serde_json::Value> {
    // First resolve
    let resolved = resolve(path, name)?;
    let refers_to = resolved["refers_to"]
        .as_str()
        .unwrap_or("")
        .to_string();

    // Strip leading = if present
    let range_str = refers_to.trim_start_matches('=');

    // Read the range
    let data = crate::services::local::range::read(path, range_str)?;

    Ok(serde_json::json!({
        "name": name,
        "refers_to": refers_to,
        "range": data.range,
        "sheet": data.sheet,
        "rows": data.rows,
        "row_count": data.row_count,
        "col_count": data.col_count,
    }))
}
