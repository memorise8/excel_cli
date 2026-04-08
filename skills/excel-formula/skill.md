---
name: excel-formula
description: Work with Excel formulas - read, write, list, audit
tools:
  - name: excel-cli
    description: CLI tool for Excel file manipulation
---

# Excel Formulas

Read, write, and audit formulas in Excel files.

## Write Formula

```bash
excel-cli formula write file.xlsx 'B10' --formula '=SUM(B1:B9)'
excel-cli formula write file.xlsx 'C10' --formula '=AVERAGE(C1:C9)'
```

## Read Formula

```bash
# Shows formula text and cached value
excel-cli formula read file.xlsx 'B10'
```

## List All Formulas

```bash
# List all formulas in a sheet
excel-cli formula list file.xlsx --sheet 'Summary'
```

## Audit (Trace References)

```bash
# Find what cells a formula depends on
excel-cli formula audit file.xlsx 'B10' --direction precedents

# Find what cells depend on this cell
excel-cli formula audit file.xlsx 'B1' --direction dependents
```

## Evaluate Formula (Cloud Only)

```bash
# Requires Microsoft Graph API auth
excel-cli formula evaluate file.xlsx 'B10' --cloud
```

## Important Notes

- **Local mode**: Formulas are stored as text. Cached values from last Excel save are returned.
- **Cloud mode** (`--cloud`): Real-time formula evaluation via Microsoft Graph API.
- Use `excel-cli auth login` to set up cloud access.
