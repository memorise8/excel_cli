use crate::models::error::*;
use crate::models::range::{col_letter_to_index, parse_range_ref};
use std::path::Path;

/// Set font properties on a range
pub fn font(
    path: &Path,
    range_str: &str,
    name: Option<&str>,
    size: Option<f64>,
    color: Option<&str>,
    bold: bool,
    italic: bool,
    underline: bool,
) -> ExcelResult<()> {
    let mut book = super::safe_io::safe_full_read(path)?;

    let addr = parse_range_ref(range_str).map_err(ExcelError::InvalidRange)?;
    let sheet_name = addr.sheet.as_deref().unwrap_or("Sheet1");
    let idx = super::safe_io::find_sheet_index(&book, sheet_name)?;
    super::safe_io::safe_read_sheet(&mut book, idx)?;

    let sheet = book
        .get_sheet_mut(&idx)
        .ok_or_else(|| ExcelError::SheetNotFound(sheet_name.to_string()))?;

    for row in addr.start_row..=addr.end_row {
        for col in addr.start_col..=addr.end_col {
            let cell = sheet.get_cell_mut((col, row));
            let font = cell.get_style_mut().get_font_mut();
            if let Some(n) = name {
                font.set_name(n);
            }
            if let Some(s) = size {
                font.set_size(s);
            }
            if let Some(c) = color {
                font.get_color_mut().set_argb(c);
            }
            if bold {
                font.set_bold(true);
            }
            if italic {
                font.set_italic(true);
            }
            if underline {
                font.set_underline("single");
            }
        }
    }

    super::safe_io::safe_write(&mut book, path)?;
    Ok(())
}

/// Set background fill on a range
pub fn fill(path: &Path, range_str: &str, color: &str, _pattern: &str) -> ExcelResult<()> {
    let mut book = super::safe_io::safe_full_read(path)?;

    let addr = parse_range_ref(range_str).map_err(ExcelError::InvalidRange)?;
    let sheet_name = addr.sheet.as_deref().unwrap_or("Sheet1");
    let idx = super::safe_io::find_sheet_index(&book, sheet_name)?;
    super::safe_io::safe_read_sheet(&mut book, idx)?;

    let sheet = book
        .get_sheet_mut(&idx)
        .ok_or_else(|| ExcelError::SheetNotFound(sheet_name.to_string()))?;

    for row in addr.start_row..=addr.end_row {
        for col in addr.start_col..=addr.end_col {
            let cell = sheet.get_cell_mut((col, row));
            cell.get_style_mut().set_background_color_solid(color);
        }
    }

    super::safe_io::safe_write(&mut book, path)?;
    Ok(())
}

/// Set borders on a range
pub fn border(
    path: &Path,
    range_str: &str,
    style: &str,
    color: Option<&str>,
    sides: &str,
) -> ExcelResult<()> {
    let mut book = super::safe_io::safe_full_read(path)?;

    let addr = parse_range_ref(range_str).map_err(ExcelError::InvalidRange)?;
    let sheet_name = addr.sheet.as_deref().unwrap_or("Sheet1");
    let idx = super::safe_io::find_sheet_index(&book, sheet_name)?;
    super::safe_io::safe_read_sheet(&mut book, idx)?;

    let sheet = book
        .get_sheet_mut(&idx)
        .ok_or_else(|| ExcelError::SheetNotFound(sheet_name.to_string()))?;

    let apply_top = matches!(sides, "all" | "top" | "outline");
    let apply_bottom = matches!(sides, "all" | "bottom" | "outline");
    let apply_left = matches!(sides, "all" | "left" | "outline");
    let apply_right = matches!(sides, "all" | "right" | "outline");

    for row in addr.start_row..=addr.end_row {
        for col in addr.start_col..=addr.end_col {
            let cell = sheet.get_cell_mut((col, row));
            let borders = cell.get_style_mut().get_borders_mut();

            let do_top = apply_top || (sides == "outline" && row == addr.start_row);
            let do_bottom = apply_bottom || (sides == "outline" && row == addr.end_row);
            let do_left = apply_left || (sides == "outline" && col == addr.start_col);
            let do_right = apply_right || (sides == "outline" && col == addr.end_col);

            if do_top {
                borders.get_top_mut().set_border_style(style);
                if let Some(c) = color {
                    borders.get_top_mut().get_color_mut().set_argb(c);
                }
            }
            if do_bottom {
                borders.get_bottom_mut().set_border_style(style);
                if let Some(c) = color {
                    borders.get_bottom_mut().get_color_mut().set_argb(c);
                }
            }
            if do_left {
                borders.get_left_mut().set_border_style(style);
                if let Some(c) = color {
                    borders.get_left_mut().get_color_mut().set_argb(c);
                }
            }
            if do_right {
                borders.get_right_mut().set_border_style(style);
                if let Some(c) = color {
                    borders.get_right_mut().get_color_mut().set_argb(c);
                }
            }
        }
    }

    super::safe_io::safe_write(&mut book, path)?;
    Ok(())
}

/// Set alignment on a range
pub fn align(
    path: &Path,
    range_str: &str,
    horizontal: Option<&str>,
    vertical: Option<&str>,
    wrap: bool,
) -> ExcelResult<()> {
    let mut book = super::safe_io::safe_full_read(path)?;

    let addr = parse_range_ref(range_str).map_err(ExcelError::InvalidRange)?;
    let sheet_name = addr.sheet.as_deref().unwrap_or("Sheet1");
    let idx = super::safe_io::find_sheet_index(&book, sheet_name)?;
    super::safe_io::safe_read_sheet(&mut book, idx)?;

    let sheet = book
        .get_sheet_mut(&idx)
        .ok_or_else(|| ExcelError::SheetNotFound(sheet_name.to_string()))?;

    for row in addr.start_row..=addr.end_row {
        for col in addr.start_col..=addr.end_col {
            let cell = sheet.get_cell_mut((col, row));
            let alignment = cell.get_style_mut().get_alignment_mut();

            if let Some(h) = horizontal {
                let h_val = match h {
                    "left" => umya_spreadsheet::HorizontalAlignmentValues::Left,
                    "center" => umya_spreadsheet::HorizontalAlignmentValues::Center,
                    "right" => umya_spreadsheet::HorizontalAlignmentValues::Right,
                    "justify" => umya_spreadsheet::HorizontalAlignmentValues::Justify,
                    _ => umya_spreadsheet::HorizontalAlignmentValues::General,
                };
                alignment.set_horizontal(h_val);
            }
            if let Some(v) = vertical {
                let v_val = match v {
                    "top" => umya_spreadsheet::VerticalAlignmentValues::Top,
                    "center" => umya_spreadsheet::VerticalAlignmentValues::Center,
                    "bottom" => umya_spreadsheet::VerticalAlignmentValues::Bottom,
                    _ => umya_spreadsheet::VerticalAlignmentValues::Bottom,
                };
                alignment.set_vertical(v_val);
            }
            if wrap {
                alignment.set_wrap_text(true);
            }
        }
    }

    super::safe_io::safe_write(&mut book, path)?;
    Ok(())
}

/// Set number format on a range
pub fn number_format(
    path: &Path,
    range_str: &str,
    format_code: Option<&str>,
    preset: Option<&str>,
) -> ExcelResult<()> {
    let code = if let Some(fc) = format_code {
        fc.to_string()
    } else if let Some(p) = preset {
        match p {
            "number" => "#,##0.00".to_string(),
            "currency" => "$#,##0.00".to_string(),
            "percent" => "0.00%".to_string(),
            "date" => "yyyy-mm-dd".to_string(),
            "time" => "hh:mm:ss".to_string(),
            "scientific" => "0.00E+00".to_string(),
            "integer" => "#,##0".to_string(),
            _ => "General".to_string(),
        }
    } else {
        return Err(ExcelError::Other(
            "Either --format or --preset is required".to_string(),
        ));
    };

    let mut book = super::safe_io::safe_full_read(path)?;

    let addr = parse_range_ref(range_str).map_err(ExcelError::InvalidRange)?;
    let sheet_name = addr.sheet.as_deref().unwrap_or("Sheet1");
    let idx = super::safe_io::find_sheet_index(&book, sheet_name)?;
    super::safe_io::safe_read_sheet(&mut book, idx)?;

    let sheet = book
        .get_sheet_mut(&idx)
        .ok_or_else(|| ExcelError::SheetNotFound(sheet_name.to_string()))?;

    for row in addr.start_row..=addr.end_row {
        for col in addr.start_col..=addr.end_col {
            let cell = sheet.get_cell_mut((col, row));
            cell.get_style_mut()
                .get_numbering_format_mut()
                .set_format_code(code.as_str());
        }
    }

    super::safe_io::safe_write(&mut book, path)?;
    Ok(())
}

/// Set column width
pub fn column_width(path: &Path, sheet_name: &str, col_letter: &str, width: f64) -> ExcelResult<()> {
    let mut book = super::safe_io::safe_full_read(path)?;

    let idx = super::safe_io::find_sheet_index(&book, sheet_name)?;
    super::safe_io::safe_read_sheet(&mut book, idx)?;

    let sheet = book
        .get_sheet_mut(&idx)
        .ok_or_else(|| ExcelError::SheetNotFound(sheet_name.to_string()))?;

    let col_num = col_letter_to_index(col_letter);
    sheet.get_column_dimension_by_number_mut(&col_num).set_width(width);

    super::safe_io::safe_write(&mut book, path)?;
    Ok(())
}

/// Set row height
pub fn row_height(path: &Path, sheet_name: &str, row_num: u32, height: f64) -> ExcelResult<()> {
    let mut book = super::safe_io::safe_full_read(path)?;

    let idx = super::safe_io::find_sheet_index(&book, sheet_name)?;
    super::safe_io::safe_read_sheet(&mut book, idx)?;

    let sheet = book
        .get_sheet_mut(&idx)
        .ok_or_else(|| ExcelError::SheetNotFound(sheet_name.to_string()))?;

    sheet.get_row_dimension_mut(&row_num).set_height(height);

    super::safe_io::safe_write(&mut book, path)?;
    Ok(())
}

/// Auto-fit column widths (sets best_fit flag)
pub fn autofit(path: &Path, sheet_name: &str, cols: Option<&str>) -> ExcelResult<()> {
    let mut book = super::safe_io::safe_full_read(path)?;

    let idx = super::safe_io::find_sheet_index(&book, sheet_name)?;
    super::safe_io::safe_read_sheet(&mut book, idx)?;

    let sheet = book
        .get_sheet_mut(&idx)
        .ok_or_else(|| ExcelError::SheetNotFound(sheet_name.to_string()))?;

    if let Some(col_range) = cols {
        // Parse "A:D" style range
        let parts: Vec<&str> = col_range.split(':').collect();
        let start = col_letter_to_index(parts[0]);
        let end = if parts.len() > 1 {
            col_letter_to_index(parts[1])
        } else {
            start
        };
        for col in start..=end {
            sheet.get_column_dimension_by_number_mut(&col).set_auto_width(true);
        }
    } else {
        sheet.calculation_auto_width();
    }

    super::safe_io::safe_write(&mut book, path)?;
    Ok(())
}
