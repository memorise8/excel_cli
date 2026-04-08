---
name: excel-merge
description: Merge sheets from one workbook into another
tools:
  - name: excel-cli
    description: CLI tool for Excel file manipulation
---

# Merge Excel Workbooks

Add sheets from one workbook into another.

## Usage

```bash
# Adds all sheets from file2 into file1
excel-cli +merge base.xlsx additional.xlsx
```

## Behavior

- Sheets from the second file are added to the first file
- If a sheet name already exists, it gets a `_merged` suffix
- The first file is modified in-place
- The second file is not modified

## Example

```bash
# Create two workbooks
excel-cli +template budget budget.xlsx
excel-cli +template sales sales.xlsx

# Merge sales sheets into budget workbook
excel-cli +merge budget.xlsx sales.xlsx

# Verify result
excel-cli sheet list budget.xlsx --format table
# Shows: Income, Expenses, Summary, Sales, Products, Customers
```
