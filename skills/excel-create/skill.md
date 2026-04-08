---
name: excel-create
description: Create new Excel workbooks with optional sheet names and templates
tools:
  - name: excel-cli
    description: CLI tool for Excel file manipulation
---

# Create Excel Workbook

Create a new Excel workbook file (.xlsx).

## Basic Usage

```bash
# Create with default Sheet1
excel-cli file create output.xlsx

# Create with custom sheets
excel-cli file create output.xlsx --sheets 'Sales,Inventory,Summary'
```

## Using Templates

Pre-built templates with headers:

```bash
# Budget tracker (Income, Expenses, Summary sheets)
excel-cli +template budget budget.xlsx

# Task tracker (Tasks, Done sheets)
excel-cli +template tracker tasks.xlsx

# Sales report (Sales, Products, Customers sheets)
excel-cli +template sales report.xlsx

# Copy existing file as template
excel-cli +template /path/to/existing.xlsx new_file.xlsx
```

## After Creation

Write data to the workbook:

```bash
excel-cli range write output.xlsx 'Sheet1!A1:C1' -d '[["Name","Age","City"]]'
excel-cli range write output.xlsx 'Sheet1!A2:C3' -d '[["Alice",30,"Seoul"],["Bob",25,"Busan"]]'
```

## Output

Returns JSON with file info:

```json
{
  "path": "output.xlsx",
  "file_name": "output.xlsx",
  "file_size": 5544,
  "sheet_count": 2,
  "sheets": [...]
}
```
