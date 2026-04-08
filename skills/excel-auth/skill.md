---
name: excel-auth
description: Authenticate with Microsoft Graph API for cloud Excel operations
tools:
  - name: excel-cli
    description: CLI tool for Excel file manipulation
---

# Authentication (Microsoft Graph API)

Some features require Microsoft Graph API access (pivot tables, charts, formula evaluation, PDF export).

## Login

```bash
# Device code flow — opens browser for authentication
excel-cli auth login
```

Set credentials via environment variables:
```bash
export EXCEL_CLI_CLIENT_ID='your-azure-app-client-id'
export EXCEL_CLI_TENANT_ID='your-tenant-id'  # or 'common' for multi-tenant
```

## Check Status

```bash
excel-cli auth status
```

## Logout

```bash
excel-cli auth logout
```

## Azure AD Setup (One-time)

1. Go to [Azure Portal](https://portal.azure.com) > App registrations > New registration
2. Set redirect URI to `http://localhost`
3. Under API permissions, add:
   - `Files.ReadWrite`
   - `Sites.ReadWrite.All`
4. Under Authentication, enable "Allow public client flows"
5. Copy the Application (client) ID

## Features Requiring Auth

| Feature | Command Example |
|---------|----------------|
| Formula evaluation | `excel-cli formula evaluate file.xlsx 'A1' --cloud` |
| Pivot tables | `excel-cli pivot create file.xlsx --source 'A1:D100' --cloud` |
| Charts | `excel-cli chart create file.xlsx --type bar --cloud` |
| PDF export | `excel-cli export pdf file.xlsx -o report.pdf --cloud` |
| File upload/download | `excel-cli file upload file.xlsx --cloud` |
