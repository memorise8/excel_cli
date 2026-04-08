use crate::models::error::*;
use crate::models::range::parse_range_ref;
use std::path::Path;

/// Add a conditional formatting rule to a range
pub fn add(
    path: &Path,
    range_str: &str,
    rule_type: &str,
    operator: Option<&str>,
    value: Option<&str>,
    format_json: Option<&str>,
) -> ExcelResult<serde_json::Value> {
    let mut book = super::safe_io::safe_full_read(path)?;

    let addr = parse_range_ref(range_str).map_err(ExcelError::InvalidRange)?;
    let sheet_name = addr.sheet.as_deref().unwrap_or("Sheet1");
    let idx = super::safe_io::find_sheet_index(&book, sheet_name)?;
    super::safe_io::safe_read_sheet(&mut book, idx)?;

    let sheet = book
        .get_sheet_mut(&idx)
        .ok_or_else(|| ExcelError::SheetNotFound(sheet_name.to_string()))?;

    // Build the conditional formatting rule
    let mut rule = umya_spreadsheet::ConditionalFormattingRule::default();

    // Set rule type
    let cf_type = match rule_type {
        "cell-value" | "cellValue" => umya_spreadsheet::ConditionalFormatValues::CellIs,
        "color-scale" | "colorScale" => umya_spreadsheet::ConditionalFormatValues::ColorScale,
        "data-bar" | "dataBar" => umya_spreadsheet::ConditionalFormatValues::DataBar,
        "icon-set" | "iconSet" => umya_spreadsheet::ConditionalFormatValues::IconSet,
        "formula" => umya_spreadsheet::ConditionalFormatValues::Expression,
        "above-average" => umya_spreadsheet::ConditionalFormatValues::AboveAverage,
        "top10" => umya_spreadsheet::ConditionalFormatValues::Top10,
        _ => umya_spreadsheet::ConditionalFormatValues::CellIs,
    };
    rule.set_type(cf_type);
    rule.set_priority(1);

    // Set operator if provided
    if let Some(op) = operator {
        let op_val = match op {
            "greater-than" | "greaterThan" => umya_spreadsheet::ConditionalFormattingOperatorValues::GreaterThan,
            "less-than" | "lessThan" => umya_spreadsheet::ConditionalFormattingOperatorValues::LessThan,
            "greater-than-or-equal" | "greaterThanOrEqual" => umya_spreadsheet::ConditionalFormattingOperatorValues::GreaterThanOrEqual,
            "less-than-or-equal" | "lessThanOrEqual" => umya_spreadsheet::ConditionalFormattingOperatorValues::LessThanOrEqual,
            "equal" | "equal-to" => umya_spreadsheet::ConditionalFormattingOperatorValues::Equal,
            "not-equal" | "notEqual" => umya_spreadsheet::ConditionalFormattingOperatorValues::NotEqual,
            "between" => umya_spreadsheet::ConditionalFormattingOperatorValues::Between,
            "not-between" | "notBetween" => umya_spreadsheet::ConditionalFormattingOperatorValues::NotBetween,
            _ => umya_spreadsheet::ConditionalFormattingOperatorValues::Equal,
        };
        rule.set_operator(op_val);
    }

    if let Some(v) = value {
        rule.set_text(v);
    }

    // Apply format from JSON if provided
    if let Some(fj) = format_json {
        if let Ok(style_json) = serde_json::from_str::<serde_json::Value>(fj) {
            let mut style = umya_spreadsheet::Style::default();
            if let Some(bg) = style_json.get("background_color").and_then(|v| v.as_str()) {
                style.set_background_color_solid(bg);
            }
            if let Some(font_obj) = style_json.get("font") {
                if let Some(bold) = font_obj.get("bold").and_then(|v| v.as_bool()) {
                    if bold { style.get_font_mut().set_bold(true); }
                }
                if let Some(color) = font_obj.get("color").and_then(|v| v.as_str()) {
                    style.get_font_mut().get_color_mut().set_argb(color);
                }
            }
            rule.set_style(style);
        }
    }

    // Build the ConditionalFormatting object
    let mut cf = umya_spreadsheet::ConditionalFormatting::default();

    // Set the range reference via SequenceOfReferences
    let seq_ref = cf.get_sequence_of_references_mut();
    // SequenceOfReferences stores range strings - add the range
    let col_start = crate::models::range::col_index_to_letter(addr.start_col);
    let col_end = crate::models::range::col_index_to_letter(addr.end_col);
    let range_ref = format!("{col_start}{}:{col_end}{}", addr.start_row, addr.end_row);
    seq_ref.set_sqref(range_ref.clone());

    cf.add_conditional_collection(rule);

    sheet.add_conditional_formatting_collection(cf);

    super::safe_io::safe_write(&mut book, path)?;

    Ok(serde_json::json!({
        "status": "ok",
        "range": range_str,
        "type": rule_type,
        "operator": operator,
        "value": value,
    }))
}

/// List all conditional formatting rules on a sheet
pub fn list(path: &Path, sheet_name: &str) -> ExcelResult<serde_json::Value> {
    let mut book = super::safe_io::safe_full_read(path)?;

    let idx = super::safe_io::find_sheet_index(&book, sheet_name)?;
    super::safe_io::safe_read_sheet(&mut book, idx)?;

    let sheet = book
        .get_sheet(&idx)
        .ok_or_else(|| ExcelError::SheetNotFound(sheet_name.to_string()))?;

    let mut rules = Vec::new();
    for (i, cf) in sheet.get_conditional_formatting_collection().iter().enumerate() {
        let range_ref = cf.get_sequence_of_references().get_sqref().to_string();
        for (j, rule) in cf.get_conditional_collection().iter().enumerate() {
            rules.push(serde_json::json!({
                "index": i,
                "rule_index": j,
                "range": range_ref,
                "type": format!("{:?}", rule.get_type()),
                "operator": format!("{:?}", rule.get_operator()),
                "text": rule.get_text(),
                "priority": rule.get_priority(),
            }));
        }
    }

    Ok(serde_json::json!({
        "sheet": sheet_name,
        "count": rules.len(),
        "rules": rules,
    }))
}

/// Delete a conditional formatting rule by index
pub fn delete(path: &Path, sheet_name: &str, index: usize) -> ExcelResult<()> {
    let mut book = super::safe_io::safe_full_read(path)?;

    let idx = super::safe_io::find_sheet_index(&book, sheet_name)?;
    super::safe_io::safe_read_sheet(&mut book, idx)?;

    let sheet = book
        .get_sheet_mut(&idx)
        .ok_or_else(|| ExcelError::SheetNotFound(sheet_name.to_string()))?;

    let coll = sheet.get_conditional_formatting_collection().to_vec();
    if index >= coll.len() {
        return Err(ExcelError::Other(format!(
            "Conditional formatting index {index} out of range (total: {})",
            coll.len()
        )));
    }

    let new_coll: Vec<_> = coll
        .into_iter()
        .enumerate()
        .filter(|(i, _)| *i != index)
        .map(|(_, cf)| cf)
        .collect();

    sheet.set_conditional_formatting_collection(new_coll);

    super::safe_io::safe_write(&mut book, path)?;
    Ok(())
}

/// Clear all conditional formatting rules from a range
pub fn clear(path: &Path, range_str: &str) -> ExcelResult<()> {
    let mut book = super::safe_io::safe_full_read(path)?;

    let addr = parse_range_ref(range_str).map_err(ExcelError::InvalidRange)?;
    let sheet_name = addr.sheet.as_deref().unwrap_or("Sheet1");
    let idx = super::safe_io::find_sheet_index(&book, sheet_name)?;
    super::safe_io::safe_read_sheet(&mut book, idx)?;

    let col_start = crate::models::range::col_index_to_letter(addr.start_col);
    let col_end = crate::models::range::col_index_to_letter(addr.end_col);
    let target_range = format!("{col_start}{}:{col_end}{}", addr.start_row, addr.end_row);

    let sheet = book
        .get_sheet_mut(&idx)
        .ok_or_else(|| ExcelError::SheetNotFound(sheet_name.to_string()))?;

    let coll = sheet.get_conditional_formatting_collection().to_vec();
    let new_coll: Vec<_> = coll
        .into_iter()
        .filter(|cf| cf.get_sequence_of_references().get_sqref() != target_range)
        .collect();

    sheet.set_conditional_formatting_collection(new_coll);

    super::safe_io::safe_write(&mut book, path)?;
    Ok(())
}
