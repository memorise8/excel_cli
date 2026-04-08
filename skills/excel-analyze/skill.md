---
name: excel-analyze
description: Analyze and compare Excel workbooks - summarize, diff, validate
tools:
  - name: excel-cli
    description: CLI tool for Excel file manipulation
---

# Analyze Excel Workbooks

Inspect, compare, and validate Excel files.

## Summarize Workbook

```bash
# Get full structure overview
excel-cli +summarize report.xlsx
```

Output includes: file name, size, sheet count, per-sheet row/col counts.

## Compare Two Workbooks

```bash
excel-cli +diff file1.xlsx file2.xlsx
```

Output includes: sheets only in file1, only in file2, common sheets, size difference.

## Validate Against Schema

```bash
excel-cli +validate schema.json data.xlsx
```

Schema format (JSON):

```json
{
  "required_sheets": ["Sales", "Summary"],
  "sheets": {
    "Sales": {
      "required_columns": ["Date", "Product", "Amount"]
    }
  }
}
```

## File Information

```bash
# Detailed file metadata
excel-cli file info report.xlsx

# Sheet list with visibility
excel-cli sheet list report.xlsx --format table
```

## Find Data

```bash
# Search for a value across sheets
excel-cli range find report.xlsx --query 'Revenue'

# Search specific sheet
excel-cli range find report.xlsx --query 'Total' --sheet 'Summary'
```
