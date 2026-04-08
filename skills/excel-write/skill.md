---
name: excel-write
description: Write data to Excel workbooks - cell values, formulas, bulk data
tools:
  - name: excel-cli
    description: CLI tool for Excel file manipulation
---

# Write Excel Data

Write values and formulas to Excel files.

## Write Cell Range

```bash
# Write 2D array of values
excel-cli range write file.xlsx 'Sheet1!A1:C3' -d '[[1,2,3],[4,5,6],[7,8,9]]'

# Write strings
excel-cli range write file.xlsx 'Sheet1!A1:B2' -d '[["Name","Age"],["Alice",30]]'

# Write single value
excel-cli range write file.xlsx 'A1' -v 'Hello World'

# Write to specific sheet
excel-cli range write file.xlsx 'Summary!A1:B1' -d '[["Total","Revenue"]]'
```

## Write Formulas

```bash
excel-cli formula write file.xlsx 'B10' --formula '=SUM(B1:B9)'
```

## Clear Data

```bash
# Clear values and formatting
excel-cli range clear file.xlsx 'Sheet1!A1:C3'

# Clear values only, keep formatting
excel-cli range clear file.xlsx 'Sheet1!A1:C3' --values-only
```

## Data Format

The `--data` (`-d`) flag accepts JSON 2D arrays:
- Numbers: `1`, `3.14`
- Strings: `"hello"`
- Booleans: `true`, `false`
- Null (empty): `null`
