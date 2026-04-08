---
name: excel-export
description: Export Excel data to CSV, JSON, HTML, PDF formats
tools:
  - name: excel-cli
    description: CLI tool for Excel file manipulation
---

# Export Excel Data

Export workbook data to various formats.

## CSV

```bash
excel-cli export csv file.xlsx --sheet 'Sales' -o sales.csv
excel-cli export csv file.xlsx --sheet 'Sales' -o sales.tsv --delimiter '\t'
```

## JSON

```bash
# Records orientation (array of objects)
excel-cli export json file.xlsx --sheet 'Sales' -o sales.json --orient records

# Values orientation (2D array)
excel-cli export json file.xlsx --sheet 'Sales' -o sales.json --orient values
```

## HTML

```bash
excel-cli export html file.xlsx --sheet 'Sales' -o report.html
```

## PDF (Cloud Only)

```bash
# Requires Microsoft Graph API authentication
excel-cli export pdf file.xlsx -o report.pdf --cloud
excel-cli export pdf file.xlsx -o report.pdf --sheets 'Sales,Summary' --cloud
```

## Quick Convert (Helper)

```bash
# Auto-names output file
excel-cli +convert file.xlsx --to csv
excel-cli +convert file.xlsx --to json
```

## Notes

- Local exports (CSV, JSON, HTML) work offline
- PDF and screenshot exports require `--cloud` flag
- CSV/JSON export the first sheet by default; use `--sheet` to specify
