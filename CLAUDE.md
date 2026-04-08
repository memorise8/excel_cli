# CLAUDE.md — Instructions for Claude Code

Read AGENTS.md for full project context. This file adds Claude-specific guidance.

## Quick Reference

```bash
# Build
source ~/.cargo/env && cargo build --workspace

# Test
cargo test --workspace

# Run
./target/debug/excel-cli <service> <verb> [args] [flags]
```

## Project Structure

- `crates/excel-core/` — Library crate (models, registry, services)
- `crates/excel-cli/` — Binary crate (CLI parsing, dispatch)
- `skills/` — AI agent skill definitions
- `tests/` — Integration tests

## Code Conventions

- Rust 2021 edition
- Use `thiserror` for error types
- Use `serde` for all serialization
- JSON is default output format
- Never use `umya_spreadsheet::*` glob imports (name conflicts with our CellValue)
- Use index-based sheet access for mutable operations

## Adding a New Command

1. Add `OperationDef` in `crates/excel-core/src/registry/<service>.rs`
2. Implement in `crates/excel-core/src/services/local/<service>.rs`
3. Wire dispatch in `crates/excel-cli/src/dispatch.rs`
4. Update skill in `skills/excel-<service>/skill.md`

## Testing

```bash
# Unit tests
cargo test --workspace

# Manual E2E test
excel-cli file create /tmp/test.xlsx --sheets 'A,B'
excel-cli range write /tmp/test.xlsx 'A!A1:C1' -d '[[1,2,3]]'
excel-cli range read /tmp/test.xlsx 'A!A1:C1'
```
