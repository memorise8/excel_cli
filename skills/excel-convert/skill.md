---
name: excel-convert
description: Convert Excel files to CSV, JSON, or other formats
tools:
  - name: excel-cli
    description: CLI tool for Excel file manipulation
---

# Convert Excel Files

Convert between Excel and other formats.

## Excel to CSV

```bash
excel-cli +convert report.xlsx --to csv
excel-cli +convert report.xlsx --to csv -o custom_output.csv
```

## Excel to JSON

```bash
excel-cli +convert report.xlsx --to json
excel-cli +convert report.xlsx --to json -o data.json
```

## Export Specific Sheet

```bash
# Export as CSV
excel-cli export csv report.xlsx --sheet 'Sales' -o sales.csv

# Export as JSON
excel-cli export json report.xlsx --sheet 'Sales' -o sales.json

# Export as HTML
excel-cli export html report.xlsx --sheet 'Sales' -o report.html
```

## Notes

- Default output filename is derived from input (e.g., `report.xlsx` → `report.csv`)
- CSV conversion exports the first sheet by default
- JSON export preserves data types (numbers, strings, booleans)
- PDF export requires `--cloud` flag (Microsoft Graph API)
