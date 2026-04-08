---
name: excel-template
description: Create workbooks from built-in or custom templates
tools:
  - name: excel-cli
    description: CLI tool for Excel file manipulation
---

# Excel Templates

Create pre-structured workbooks from templates.

## Built-in Templates

```bash
# Budget tracker — Income, Expenses, Summary sheets with headers
excel-cli +template budget budget_2024.xlsx

# Task tracker — Tasks, Done sheets with ID/Task/Status/Due Date columns
excel-cli +template tracker project_tasks.xlsx

# Sales report — Sales, Products, Customers sheets with headers
excel-cli +template sales quarterly_report.xlsx

# Blank workbook
excel-cli +template blank empty.xlsx
```

## Custom Template (from existing file)

```bash
# Use any .xlsx file as a template
excel-cli +template /path/to/my_template.xlsx new_workbook.xlsx
```

## After Template Creation

```bash
# Verify structure
excel-cli +summarize budget_2024.xlsx

# Start adding data
excel-cli range write budget_2024.xlsx 'Income!A2:C2' -d '[["2024-01-15","Salary",5000]]'
```

## Template Details

| Template | Sheets | Headers |
|----------|--------|---------|
| budget | Income, Expenses, Summary | Date, Description, Amount (+Category for Expenses) |
| tracker | Tasks, Done | ID, Task, Status, Due Date |
| sales | Sales, Products, Customers | Date, Product, Customer, Quantity, Amount |
| blank | Sheet1 | (none) |
