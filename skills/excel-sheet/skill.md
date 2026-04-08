---
name: excel-sheet
description: Manage worksheets - add, rename, delete, copy, reorder, hide/show
tools:
  - name: excel-cli
    description: CLI tool for Excel file manipulation
---

# Worksheet Management

Manage sheets within an Excel workbook.

## List Sheets

```bash
excel-cli sheet list file.xlsx
excel-cli sheet list file.xlsx --format table
```

## Add Sheet

```bash
excel-cli sheet add file.xlsx 'NewSheet'
excel-cli sheet add file.xlsx 'NewSheet' --position 0  # Insert at beginning
```

## Rename Sheet

```bash
excel-cli sheet rename file.xlsx 'OldName' 'NewName'
```

## Delete Sheet

```bash
excel-cli sheet delete file.xlsx 'SheetName'
```

## Copy Sheet

```bash
excel-cli sheet copy file.xlsx 'Source' --new-name 'Source Copy'
```

## Hide / Unhide

```bash
excel-cli sheet hide file.xlsx 'SheetName'
excel-cli sheet unhide file.xlsx 'SheetName'
```

## Tab Color

```bash
excel-cli sheet color file.xlsx 'SheetName' --color 'FF0000'  # Red
```

## Protect / Unprotect

```bash
excel-cli sheet protect file.xlsx 'SheetName' --password 'secret'
excel-cli sheet unprotect file.xlsx 'SheetName' --password 'secret'
```
