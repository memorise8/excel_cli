# CONTEXT.md — Runtime Agent Operational Guide

## Rules of Engagement

1. **Summarize first**: Always use `excel_summarize` before reading data to understand structure.
2. **Limit ranges**: For large files, read max 50 rows at a time. Use `excel_range_read` with specific ranges.
3. **Confirm writes**: Never write data without user confirmation.
4. **Use search**: Use `excel_find` to locate specific values before reading large ranges.
5. **Language match**: Respond in the same language the user uses.

## Tool Usage Patterns

### Understanding a file
```
1. excel_summarize → get sheet names, dimensions
2. excel_range_read "Sheet1!A1:J1" → read headers
3. excel_range_read "Sheet1!A2:J20" → read first data rows
```

### Finding specific data
```
1. excel_find query="keyword" → locate cells
2. excel_cell_read "Sheet1!B5" → read specific cell
```

### Analyzing data
```
1. excel_export_csv sheet="Sheet1" → get full data as CSV
2. Analyze in response → summarize findings
```

### Modifying data (CONFIRM FIRST)
```
1. Explain what will change to user
2. Wait for confirmation
3. excel_range_write range="Sheet1!A1:B2" data='[["a","b"],["c","d"]]'
```

## Common Mistakes to Avoid

- Reading entire sheet without range limit (causes timeout on large files)
- Writing without user confirmation
- Assuming sheet names (always check with excel_summarize first)
- Reading formulas without specifying that values may be cached

## Cloud Operations (Microsoft Graph API)

### Setup
```
1. Register an Azure AD app at https://portal.azure.com
2. Add API permissions: Files.ReadWrite, Sites.ReadWrite.All
3. excel-cli auth login --client-id YOUR_CLIENT_ID
```

### Cloud Workflow
```
1. excel_cloud_upload → get item_id
2. excel_cloud_range_read item_id range with_format=true → read data + formatting
3. excel_cloud_range_write item_id range data → write data
4. excel_cloud_calc item_id → recalculate formulas
5. excel_cloud_export_pdf item_id output → export as PDF
6. excel_cloud_download item_id output → download modified file
```

### Template Cloning Workflow
```
1. Upload reference file → excel_cloud_upload → ref_item_id
2. Read structure → excel_summarize (local)
3. Read data + format → excel_cloud_range_read ref_item_id 'Sheet1!A1:Z100' with_format=true
4. Create new file → excel file create new.xlsx
5. Upload new file → excel_cloud_upload → new_item_id
6. Write data → excel_cloud_range_write new_item_id 'Sheet1!A1:Z100' data
7. Apply formatting (font/fill/border via format tools)
8. Recalculate → excel_cloud_calc new_item_id
9. Download → excel_cloud_download new_item_id result.xlsx
```
