---
name: excel-format
description: Format cells - fonts, colors, borders, alignment, number formats
tools:
  - name: excel-cli
    description: CLI tool for Excel file manipulation
---

# Cell Formatting

Apply visual formatting to Excel cells.

## Font

```bash
excel-cli format font file.xlsx 'A1:D1' --name 'Arial' --size 14 --bold --color 'FFFFFF'
```

## Fill (Background Color)

```bash
excel-cli format fill file.xlsx 'A1:D1' --color '4472C4'
```

## Borders

```bash
excel-cli format border file.xlsx 'A1:D10' --style thin --color '000000' --sides all
```

## Alignment

```bash
excel-cli format align file.xlsx 'A1:D1' --horizontal center --vertical center --wrap
```

## Number Format

```bash
# Currency
excel-cli format number file.xlsx 'C2:C100' --preset currency

# Percentage
excel-cli format number file.xlsx 'D2:D100' --preset percent

# Custom format
excel-cli format number file.xlsx 'B2:B100' --format '#,##0.00'

# Date format
excel-cli format number file.xlsx 'A2:A100' --preset date
```

## Column Width / Row Height

```bash
excel-cli format width file.xlsx --sheet 'Sheet1' --col A --width 20
excel-cli format height file.xlsx --sheet 'Sheet1' --row 1 --height 30
excel-cli format autofit file.xlsx --sheet 'Sheet1' --cols 'A:D'
```

## Compound Style (JSON)

```bash
excel-cli format style file.xlsx 'A1:D1' -j '{"font":{"bold":true,"size":12},"fill":{"color":"4472C4"},"alignment":{"horizontal":"center"}}'
```
