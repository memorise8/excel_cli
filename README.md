# excel-cli

One CLI for Excel — built for humans and AI agents.

Inspired by [googleworkspace/cli (gws)](https://github.com/googleworkspace/cli).

## Install

```bash
cargo install --path crates/excel-cli
```

## Usage

```bash
excel-cli <service> <verb> [args] [flags]
```

### File Management

```bash
excel-cli file create report.xlsx --sheets 'Sales,Summary'
excel-cli file info report.xlsx
excel-cli file save report.xlsx backup.xlsx
```

### Read & Write Data

```bash
excel-cli range write report.xlsx 'Sales!A1:C3' -d '[[1,2,3],[4,5,6],[7,8,9]]'
excel-cli range read report.xlsx 'Sales!A1:C3'
excel-cli range read report.xlsx 'Sales!A1:C3' --format table
```

### Sheet Management

```bash
excel-cli sheet list report.xlsx --format table
excel-cli sheet add report.xlsx 'NewSheet'
excel-cli sheet rename report.xlsx 'OldName' 'NewName'
excel-cli sheet delete report.xlsx 'SheetName'
```

### Formatting

```bash
excel-cli format font report.xlsx 'A1:D1' --bold --size 14 --color 'FFFFFF'
excel-cli format fill report.xlsx 'A1:D1' --color '4472C4'
excel-cli format number report.xlsx 'C2:C100' --preset currency
```

### Tables

```bash
excel-cli table create report.xlsx 'Sheet1!A1:D10' --name 'SalesTable' --has-headers
excel-cli table read report.xlsx --name 'SalesTable'
excel-cli table append report.xlsx --name 'SalesTable' -d '[["2024-03-15","Widget",5,99.99]]'
```

### Formulas

```bash
excel-cli formula write report.xlsx 'B10' --formula '=SUM(B1:B9)'
excel-cli formula read report.xlsx 'B10'
excel-cli formula list report.xlsx --sheet 'Summary'
```

### Helper Commands

```bash
excel-cli +summarize report.xlsx              # Workbook overview
excel-cli +diff file1.xlsx file2.xlsx         # Compare workbooks
excel-cli +validate schema.json data.xlsx     # Validate against schema
excel-cli +convert report.xlsx --to csv       # Format conversion
excel-cli +merge base.xlsx extra.xlsx         # Merge workbooks
excel-cli +template sales output.xlsx         # Create from template
```

### Export

```bash
excel-cli export csv report.xlsx --sheet 'Sales' -o sales.csv
excel-cli export json report.xlsx --sheet 'Sales' -o data.json
excel-cli export html report.xlsx --sheet 'Sales' -o report.html
```

## Output Formats

All commands support `--format`:

| Format | Flag | Description |
|--------|------|-------------|
| JSON | `--format json` | Default. Structured output for AI agents |
| Table | `--format table` | Human-readable terminal table |
| CSV | `--format csv` | Comma-separated values |

## Cloud Operations (Optional)

Some features require Microsoft Graph API authentication:

```bash
# One-time setup
export EXCEL_CLI_CLIENT_ID='your-azure-app-client-id'
excel-cli auth login

# Use cloud features
excel-cli formula evaluate report.xlsx 'B10' --cloud
excel-cli pivot create report.xlsx --source 'A1:D100' --cloud
excel-cli export pdf report.xlsx -o report.pdf --cloud
```

Cloud-required features: formula evaluation, pivot tables, charts, slicers, PDF export, file upload/download.

## Services

| Service | Description | Auth |
|---------|-------------|------|
| file | Workbook management | upload/download only |
| sheet | Worksheet management | No |
| range | Cell range operations | No |
| formula | Formula operations | evaluate only |
| format | Cell formatting | No |
| conditional | Conditional formatting | No |
| table | Excel table operations | No |
| named-range | Defined names | No |
| pivot | PivotTable operations | Yes |
| chart | Chart operations | Yes |
| calc | Calculation mode | recalc only |
| connection | Data connections | Yes |
| slicer | Slicer operations | Yes |
| export | Export (csv, json, html, pdf) | pdf/screenshot only |

## Architecture

```
excel-cli (binary)
    └── excel-core (library)
            ├── Operation Registry — ~80 commands defined as data
            ├── Local Backend — umya-spreadsheet/calamine (.xlsx read/write)
            └── Graph Backend — Microsoft Graph API (cloud operations)
```

Built with Rust. Single binary, no runtime dependencies.

## Skills

The `skills/` directory contains 14 AI agent skill definitions for use with Claude Code, GitHub Copilot, and other AI coding assistants.

## License

MIT
