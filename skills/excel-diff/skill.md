---
name: excel-diff
description: Compare two Excel workbooks for structural differences
tools:
  - name: excel-cli
    description: CLI tool for Excel file manipulation
---

# Compare Excel Workbooks

Compare the structure of two Excel files.

## Usage

```bash
excel-cli +diff original.xlsx modified.xlsx
```

## Output

```json
{
  "file1": "original.xlsx",
  "file2": "modified.xlsx",
  "sheets": {
    "only_in_file1": ["OldSheet"],
    "only_in_file2": ["NewSheet"],
    "common": ["Sales", "Summary"]
  },
  "size_diff": 1234
}
```

## Workflow Example

```bash
# Make a backup before changes
excel-cli file save report.xlsx backup.xlsx

# ... make changes ...

# Compare what changed
excel-cli +diff backup.xlsx report.xlsx
```
