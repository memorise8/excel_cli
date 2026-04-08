# AGENTS.md — AI Agent Development Guide for excel-cli

## Project Overview

**excel-cli** is a Rust-based CLI tool for Excel file manipulation, inspired by [googleworkspace/cli (gws)](https://github.com/googleworkspace/cli). It provides a comprehensive set of commands for working with .xlsx files locally, with optional Microsoft Graph API integration for cloud operations.

## Architecture

```
excel-cli <service> <verb> [args] [flags]

┌─────────────────────┐
│   excel-cli (bin)   │     Binary crate — CLI parsing, dispatch, helpers
└─────────┬───────────┘
          ▼
┌─────────────────────────┐
│   excel-core (library)  │     Library crate — registry, models, services
│  Operation Registry     │
│  + Service Backends     │
└───┬──────────┬──────────┘
    ▼          ▼
┌──────────┐ ┌──────────────┐
│ Local    │ │ Graph API    │
│ (umya/   │ │ (reqwest/    │
│ calamine)│ │  oauth2)     │
└──────────┘ └──────────────┘
```

### Key Design Patterns

1. **Operation Registry** (`crates/excel-core/src/registry/`): All ~80 commands are defined as `OperationDef` structs. The CLI crate reads these at startup and builds the `clap` command tree dynamically. This mirrors gws's Discovery API pattern.

2. **Library/Binary Separation**: `excel-core` has zero CLI dependencies. Both future `excel-mcp` and `excel-cli` share `excel-core`.

3. **Hybrid Backend**: Local operations use `umya-spreadsheet`/`calamine`. Cloud operations (marked `auth_required: true`) use Microsoft Graph API.

4. **AI-First Output**: JSON is the default output format. `--format table` and `--format csv` are available for human consumption.

## Workspace Structure

```
crates/
├── excel-core/           # Library crate
│   └── src/
│       ├── registry/     # Operation definitions (14 services)
│       ├── services/     # Execution backends
│       │   ├── local/    # umya-spreadsheet (offline)
│       │   └── graph/    # Microsoft Graph API (cloud)
│       ├── models/       # Domain types (CellValue, RangeData, etc.)
│       └── output/       # JSON, table, CSV formatters
│
├── excel-cli/            # Binary crate
│   └── src/
│       ├── main.rs       # Entry point
│       ├── builder.rs    # Registry → clap Command tree
│       ├── dispatch.rs   # Route parsed args → service calls
│       └── helpers/      # +verb helper commands
│
skills/                   # AI agent skill definitions (14 skills)
tests/                    # Integration tests
```

## Services

| Service | Commands | Auth Required |
|---------|----------|---------------|
| file | create, info, save, convert, upload*, download* | upload/download only |
| sheet | list, add, rename, delete, copy, move, hide, unhide, color, protect, unprotect | No |
| range | read, write, clear, copy, move, insert, delete, merge, unmerge, sort, filter, find, replace, validate | No |
| formula | read, write, list, evaluate*, audit | evaluate only |
| format | font, fill, border, align, number, width, height, autofit, style | No |
| conditional | add, list, delete, clear | No |
| table | list, create, delete, read, append, resize, rename, sort, filter, style, total-row, column-add, column-delete, to-range | No |
| named-range | list, create, delete, update, resolve, read | No |
| pivot | list, create*, refresh*, field-add*, field-remove*, filter*, group*, style* | Most require cloud |
| chart | list, create*, delete*, update*, export*, series-add*, series-remove*, style* | Most require cloud |
| calc | mode, now*, sheet* | now/sheet require cloud |
| connection | list, create*, delete*, refresh* | Most require cloud |
| slicer | list, create*, delete*, select*, clear* | Most require cloud |
| export | csv, json, html, pdf*, screenshot* | pdf/screenshot require cloud |

\* = Requires `--cloud` flag and Microsoft Graph API authentication

## Helper Commands

| Command | Description |
|---------|-------------|
| `+summarize <file>` | Workbook structure and statistics |
| `+diff <file1> <file2>` | Compare two workbooks |
| `+validate <schema> <file>` | Validate against JSON schema |
| `+convert <file> --to <fmt>` | Format conversion |
| `+merge <file1> <file2>` | Merge workbooks |
| `+template <name> <output>` | Create from template (blank, budget, tracker, sales) |

## Development Commands

```bash
# Build
source ~/.cargo/env
cargo build --workspace

# Test
cargo test --workspace

# Lint
cargo clippy --workspace

# Run
./target/debug/excel-cli --help
```

## Important Notes for AI Agents

1. **Name conflict**: `umya_spreadsheet` has its own `CellValue` type. Never use glob imports (`use umya_spreadsheet::*`). Always use fully qualified paths.

2. **Sheet access**: Use index-based access for mutable operations:
   ```rust
   let idx = find_sheet_index(&book, sheet_name)?;
   book.read_sheet(idx);  // Ensure deserialized
   let sheet = book.get_sheet_mut(&idx).unwrap();
   ```

3. **File I/O**: Always use `umya_spreadsheet::reader::xlsx::read()` and `umya_spreadsheet::writer::xlsx::write()`.

4. **Output**: Return data structures, let `format_output()` handle JSON/table/CSV formatting.

5. **Auth boundary**: Local operations must never require auth. Cloud operations must check for token and return clear error if missing.
