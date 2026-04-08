---
name: excel-read
description: Read data from Excel workbooks - cell ranges, sheets, file info
tools:
  - name: excel-cli
    description: CLI tool for Excel file manipulation
---

# Read Excel Data

Read values, metadata, and structure from Excel files.

## Read Cell Range

```bash
# Read specific range (JSON output)
excel-cli range read file.xlsx 'Sheet1!A1:D10'

# Read as table format
excel-cli range read file.xlsx 'Sheet1!A1:D10' --format table

# Read as CSV
excel-cli range read file.xlsx 'Sheet1!A1:D10' --format csv

# Read single cell
excel-cli range read file.xlsx 'A1'
```

## File Information

```bash
# Get workbook metadata (sheets, size, dimensions)
excel-cli file info file.xlsx

# Quick summary with statistics
excel-cli +summarize file.xlsx
```

## Sheet Operations

```bash
# List all sheets
excel-cli sheet list file.xlsx

# List as table
excel-cli sheet list file.xlsx --format table
```

## Output Formats

All read commands support `--format`:
- `json` (default) — structured JSON, ideal for AI processing
- `table` — human-readable terminal table
- `csv` — comma-separated values
