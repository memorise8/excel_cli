---
name: excel-table
description: Manage Excel tables - create, read, append, sort, filter
tools:
  - name: excel-cli
    description: CLI tool for Excel file manipulation
---

# Excel Tables

Create and manage structured Excel tables.

## List Tables

```bash
excel-cli table list file.xlsx
```

## Create Table

```bash
excel-cli table create file.xlsx 'Sheet1!A1:D10' --name 'SalesTable' --has-headers
```

## Read Table Data

```bash
excel-cli table read file.xlsx --name 'SalesTable'
excel-cli table read file.xlsx --name 'SalesTable' --format table
```

## Append Rows

```bash
excel-cli table append file.xlsx --name 'SalesTable' -d '[["2024-03-15","Widget",5,99.99]]'
```

## Sort

```bash
excel-cli table sort file.xlsx --name 'SalesTable' --by 'Amount' --desc
```

## Filter

```bash
excel-cli table filter file.xlsx --name 'SalesTable' --column 'Product' --value 'Widget'
```

## Other Operations

```bash
# Rename table
excel-cli table rename file.xlsx --name 'SalesTable' --new-name 'Q1Sales'

# Apply style
excel-cli table style file.xlsx --name 'SalesTable' --style 'TableStyleMedium2'

# Toggle total row
excel-cli table total-row file.xlsx --name 'SalesTable' --enable

# Add column
excel-cli table column-add file.xlsx --name 'SalesTable' --header 'Notes'

# Convert back to plain range
excel-cli table to-range file.xlsx --name 'SalesTable'
```
