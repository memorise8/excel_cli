"""
Tool definitions and execution for the Excel CLI Orchestrator.
Each tool wraps an excel-cli command for the AI to use.
"""

import json
import subprocess
from typing import Any

TOOL_DEFINITIONS = [
    {
        "name": "excel_summarize",
        "description": "Get a summary of an Excel workbook: sheets, row/column counts, file size. Use this first to understand the file structure.",
        "parameters": {
            "type": "object",
            "properties": {},
            "required": [],
        },
    },
    {
        "name": "excel_sheet_list",
        "description": "List all sheets in the workbook with their names and visibility.",
        "parameters": {
            "type": "object",
            "properties": {},
            "required": [],
        },
    },
    {
        "name": "excel_range_read",
        "description": "Read cell values from a specific range. Use Excel-style range notation like 'Sheet1!A1:D10' or 'A1:B5' (defaults to first sheet).",
        "parameters": {
            "type": "object",
            "properties": {
                "range": {
                    "type": "string",
                    "description": "Cell range to read, e.g. 'Sheet1!A1:D10', 'A1:Z1' for headers, 'B5' for single cell",
                },
            },
            "required": ["range"],
        },
    },
    {
        "name": "excel_range_write",
        "description": "Write values to a cell range. Data is a JSON 2D array. Confirm with user before using this.",
        "parameters": {
            "type": "object",
            "properties": {
                "range": {
                    "type": "string",
                    "description": "Target range, e.g. 'Sheet1!A1:C3'",
                },
                "data": {
                    "type": "string",
                    "description": "JSON 2D array of values, e.g. '[[1,2,3],[4,5,6]]'",
                },
            },
            "required": ["range", "data"],
        },
    },
    {
        "name": "excel_find",
        "description": "Search for a value across the workbook or in a specific sheet.",
        "parameters": {
            "type": "object",
            "properties": {
                "query": {
                    "type": "string",
                    "description": "Value to search for",
                },
                "sheet": {
                    "type": "string",
                    "description": "Optional: limit search to this sheet name",
                },
            },
            "required": ["query"],
        },
    },
    {
        "name": "excel_formula_read",
        "description": "Read a formula from a specific cell, including the formula text and cached value.",
        "parameters": {
            "type": "object",
            "properties": {
                "range": {
                    "type": "string",
                    "description": "Cell reference, e.g. 'Sheet1!B10'",
                },
            },
            "required": ["range"],
        },
    },
    {
        "name": "excel_formula_write",
        "description": "Write a formula to a cell. Confirm with user before using this.",
        "parameters": {
            "type": "object",
            "properties": {
                "range": {
                    "type": "string",
                    "description": "Target cell, e.g. 'Sheet1!B10'",
                },
                "formula": {
                    "type": "string",
                    "description": "Excel formula, e.g. '=SUM(B1:B9)'",
                },
            },
            "required": ["range", "formula"],
        },
    },
    {
        "name": "excel_formula_list",
        "description": "List all formulas in a specific sheet.",
        "parameters": {
            "type": "object",
            "properties": {
                "sheet": {
                    "type": "string",
                    "description": "Sheet name to scan for formulas",
                },
            },
            "required": ["sheet"],
        },
    },
    {
        "name": "excel_file_info",
        "description": "Get detailed file information: path, size, sheet details with dimensions.",
        "parameters": {
            "type": "object",
            "properties": {},
            "required": [],
        },
    },
    {
        "name": "excel_export_csv",
        "description": "Export a sheet to CSV format and return the content.",
        "parameters": {
            "type": "object",
            "properties": {
                "sheet": {
                    "type": "string",
                    "description": "Sheet name to export",
                },
            },
            "required": ["sheet"],
        },
    },
    {
        "name": "excel_sheet_add",
        "description": "Add a new sheet to the workbook. Confirm with user before using this.",
        "parameters": {
            "type": "object",
            "properties": {
                "name": {
                    "type": "string",
                    "description": "Name for the new sheet",
                },
            },
            "required": ["name"],
        },
    },
    {
        "name": "excel_format_font",
        "description": "Apply font formatting to a range. Confirm with user before using this.",
        "parameters": {
            "type": "object",
            "properties": {
                "range": {
                    "type": "string",
                    "description": "Target range, e.g. 'Sheet1!A1:D1'",
                },
                "bold": {
                    "type": "string",
                    "description": "Set bold: 'true' or 'false'",
                },
                "size": {
                    "type": "string",
                    "description": "Font size, e.g. '14'",
                },
                "color": {
                    "type": "string",
                    "description": "Font color hex, e.g. 'FF0000'",
                },
            },
            "required": ["range"],
        },
    },
]


def execute_tool(name: str, args: dict, file_path: str, cli_path: str) -> str:
    """Execute an excel-cli tool and return the result."""
    try:
        cmd = build_command(name, args, file_path, cli_path)
        if not cmd:
            return json.dumps({"error": f"Unknown tool: {name}"})

        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=30,
        )

        output = result.stdout.strip()
        if result.returncode != 0:
            error = result.stderr.strip() or output
            return json.dumps({"error": error, "exit_code": result.returncode})

        # Try to parse as JSON for cleaner output
        try:
            parsed = json.loads(output)
            # Truncate very large outputs
            output_str = json.dumps(parsed, ensure_ascii=False)
            if len(output_str) > 8000:
                return json.dumps({
                    "truncated": True,
                    "message": f"Output too large ({len(output_str)} chars). Showing first part.",
                    "data": output_str[:8000],
                }, ensure_ascii=False)
            return output_str
        except json.JSONDecodeError:
            if len(output) > 8000:
                return output[:8000] + "\n...(truncated)"
            return output

    except subprocess.TimeoutExpired:
        return json.dumps({"error": "Command timed out after 30 seconds"})
    except Exception as e:
        return json.dumps({"error": str(e)})


def build_command(name: str, args: dict, file_path: str, cli_path: str) -> list[str] | None:
    """Build the excel-cli command for a given tool call."""

    match name:
        case "excel_summarize":
            return [cli_path, "+summarize", file_path]

        case "excel_sheet_list":
            return [cli_path, "sheet", "list", file_path]

        case "excel_range_read":
            return [cli_path, "range", "read", file_path, args["range"]]

        case "excel_range_write":
            return [cli_path, "range", "write", file_path, args["range"], "-d", args["data"]]

        case "excel_find":
            cmd = [cli_path, "range", "find", file_path, "--query", args["query"]]
            if args.get("sheet"):
                cmd.extend(["--sheet", args["sheet"]])
            return cmd

        case "excel_formula_read":
            return [cli_path, "formula", "read", file_path, args["range"]]

        case "excel_formula_write":
            return [cli_path, "formula", "write", file_path, args["range"], "--formula", args["formula"]]

        case "excel_formula_list":
            return [cli_path, "formula", "list", file_path, "--sheet", args["sheet"]]

        case "excel_file_info":
            return [cli_path, "file", "info", file_path]

        case "excel_export_csv":
            return [cli_path, "export", "csv", file_path, "--sheet", args["sheet"]]

        case "excel_sheet_add":
            return [cli_path, "sheet", "add", file_path, args["name"]]

        case "excel_format_font":
            cmd = [cli_path, "format", "font", file_path, args["range"]]
            if args.get("bold") == "true":
                cmd.append("--bold")
            if args.get("size"):
                cmd.extend(["--size", args["size"]])
            if args.get("color"):
                cmd.extend(["--color", args["color"]])
            return cmd

        case _:
            return None
