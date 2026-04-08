use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq)]
#[serde(untagged)]
pub enum CellValue {
    Empty,
    Bool(bool),
    Int(i64),
    Float(f64),
    String(String),
    Formula(FormulaValue),
    Error(String),
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq)]
pub struct FormulaValue {
    pub formula: String,
    pub cached_value: Option<Box<CellValue>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct RangeData {
    pub range: String,
    pub sheet: String,
    pub rows: Vec<Vec<CellValue>>,
    pub row_count: usize,
    pub col_count: usize,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CellAddress {
    pub sheet: Option<String>,
    pub col: u32,
    pub row: u32,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct RangeAddress {
    pub sheet: Option<String>,
    pub start_col: u32,
    pub start_row: u32,
    pub end_col: u32,
    pub end_row: u32,
}

impl RangeAddress {
    pub fn row_count(&self) -> u32 {
        self.end_row - self.start_row + 1
    }

    pub fn col_count(&self) -> u32 {
        self.end_col - self.start_col + 1
    }
}

/// Parse "Sheet1!A1:C3" or "A1:C3" into RangeAddress
pub fn parse_range(input: &str) -> Result<(Option<String>, String), String> {
    if let Some((sheet, range)) = input.split_once('!') {
        Ok((Some(sheet.to_string()), range.to_string()))
    } else {
        Ok((None, input.to_string()))
    }
}

/// Convert column letter to 1-based index: A=1, B=2, ..., Z=26, AA=27
pub fn col_letter_to_index(col: &str) -> u32 {
    col.chars().fold(0u32, |acc, c| {
        acc * 26 + (c.to_ascii_uppercase() as u32 - 'A' as u32 + 1)
    })
}

/// Convert 1-based column index to letter: 1=A, 2=B, ..., 26=Z, 27=AA
pub fn col_index_to_letter(mut idx: u32) -> String {
    let mut result = String::new();
    while idx > 0 {
        idx -= 1;
        result.insert(0, (b'A' + (idx % 26) as u8) as char);
        idx /= 26;
    }
    result
}

/// Parse "A1" into (col_index, row_index) both 1-based
pub fn parse_cell_ref(cell: &str) -> Result<(u32, u32), String> {
    let col_end = cell
        .find(|c: char| c.is_ascii_digit())
        .ok_or_else(|| format!("Invalid cell reference: {cell}"))?;
    let col_str = &cell[..col_end];
    let row_str = &cell[col_end..];
    let col = col_letter_to_index(col_str);
    let row: u32 = row_str
        .parse()
        .map_err(|_| format!("Invalid row number in: {cell}"))?;
    Ok((col, row))
}

/// Parse "A1:C3" into RangeAddress
pub fn parse_range_ref(range: &str) -> Result<RangeAddress, String> {
    let (sheet, range_str) = parse_range(range)?;
    let parts: Vec<&str> = range_str.split(':').collect();
    match parts.len() {
        1 => {
            let (col, row) = parse_cell_ref(parts[0])?;
            Ok(RangeAddress {
                sheet,
                start_col: col,
                start_row: row,
                end_col: col,
                end_row: row,
            })
        }
        2 => {
            let (sc, sr) = parse_cell_ref(parts[0])?;
            let (ec, er) = parse_cell_ref(parts[1])?;
            Ok(RangeAddress {
                sheet,
                start_col: sc,
                start_row: sr,
                end_col: ec,
                end_row: er,
            })
        }
        _ => Err(format!("Invalid range format: {range}")),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_col_letter_to_index() {
        assert_eq!(col_letter_to_index("A"), 1);
        assert_eq!(col_letter_to_index("Z"), 26);
        assert_eq!(col_letter_to_index("AA"), 27);
        assert_eq!(col_letter_to_index("AZ"), 52);
    }

    #[test]
    fn test_col_index_to_letter() {
        assert_eq!(col_index_to_letter(1), "A");
        assert_eq!(col_index_to_letter(26), "Z");
        assert_eq!(col_index_to_letter(27), "AA");
        assert_eq!(col_index_to_letter(52), "AZ");
    }

    #[test]
    fn test_parse_cell_ref() {
        assert_eq!(parse_cell_ref("A1").unwrap(), (1, 1));
        assert_eq!(parse_cell_ref("C3").unwrap(), (3, 3));
        assert_eq!(parse_cell_ref("AA100").unwrap(), (27, 100));
    }

    #[test]
    fn test_parse_range_ref() {
        let r = parse_range_ref("Sheet1!A1:C3").unwrap();
        assert_eq!(r.sheet, Some("Sheet1".to_string()));
        assert_eq!(r.start_col, 1);
        assert_eq!(r.start_row, 1);
        assert_eq!(r.end_col, 3);
        assert_eq!(r.end_row, 3);
    }
}
