#!/usr/bin/env python3
"""
Harness MCP Server — exposes harness plugin tools via Model Context Protocol.

Connects to OpenClaw, Claude Code, Cursor, or any MCP client.

Usage (stdio):
    python3 orchestrator/mcp_server.py --plugin excel --file report.xlsx

Usage (SSE):
    python3 orchestrator/mcp_server.py --plugin excel --file report.xlsx --transport sse --port 8100

Register with OpenClaw:
    openclaw mcp set excel-harness '{"command":"python3","args":["orchestrator/mcp_server.py","--plugin","excel","--file","report.xlsx"]}'
"""

from __future__ import annotations

import argparse
import json
import os
import sys
from pathlib import Path

# Ensure orchestrator/ is in sys.path for local imports
sys.path.insert(0, str(Path(__file__).parent))

from dotenv import load_dotenv

load_dotenv(Path(__file__).parent / ".env")

from mcp.server.fastmcp import FastMCP

from harness.tools import ToolRegistry, ToolDef
from harness.hooks import HookEngine, HookEvent, HookDecision
from plugins.excel import ExcelPlugin

# Plugin registry
PLUGINS = {
    "excel": ExcelPlugin,
}


def build_server(plugin_name: str, file_path: str) -> FastMCP:
    """Build an MCP server with all tools from the given harness plugin."""

    plugin = PLUGINS[plugin_name]()

    # Validate target file
    ok, err = plugin.validate_file(file_path)
    if not ok:
        print(f"Error: {err}", file=sys.stderr)
        sys.exit(1)

    # Find CLI binary
    project_root = Path(__file__).parent.parent
    cli_path = plugin.find_cli(project_root)

    # Build tool registry
    registry = ToolRegistry(context_path=file_path, cli_path=cli_path)
    plugin.register_tools(registry)

    # Build hook engine with plugin defaults
    hooks = HookEngine()
    for hook_def in plugin.get_default_hooks():
        hooks.register_fn(
            event=HookEvent(hook_def["event"]),
            fn=hook_def["fn"],
            matcher=hook_def.get("matcher"),
            description=hook_def.get("description", ""),
        )

    # Get file summary for server description
    summary = plugin.get_file_summary(file_path)

    # Create MCP server
    server = FastMCP(
        name=f"harness-{plugin.name}",
        instructions=plugin.build_system_prompt(
            file_path=file_path,
            file_summary=summary,
            skill_names=[],
            tool_names=registry.tool_names(),
        ),
    )

    # Register each harness tool as an MCP tool
    for tool_def in registry.list_tools():
        _register_mcp_tool(server, tool_def, registry, hooks, file_path)

    # Add a session info resource
    @server.resource(f"harness://{plugin.name}/info")
    def session_info() -> str:
        return json.dumps({
            "plugin": plugin.name,
            "file": file_path,
            "summary": summary,
            "tools": registry.tool_names(),
            "cli_available": bool(cli_path),
        }, ensure_ascii=False, indent=2)

    return server


def _register_mcp_tool(
    server: FastMCP,
    tool_def: ToolDef,
    registry: ToolRegistry,
    hooks: HookEngine,
    file_path: str,
) -> None:
    """Register a single harness ToolDef as an MCP tool with hook integration."""

    # Extract parameter properties for the MCP function signature
    props = tool_def.parameters.get("properties", {})
    required = tool_def.parameters.get("required", [])

    # Build a dynamic function that MCP will introspect
    # We use a factory to capture tool_def in the closure
    def make_handler(td: ToolDef):
        async def handler(**kwargs) -> str:
            # PreToolUse hook
            pre_result = hooks.execute(
                HookEvent.PRE_TOOL_USE,
                tool_name=td.name,
                tool_input=kwargs,
            )

            if pre_result.decision == HookDecision.DENY:
                return json.dumps({
                    "error": f"Blocked by hook: {pre_result.reason}"
                })

            # Use modified input if hook provided one
            args = pre_result.modified_input or kwargs

            # Execute via registry
            execution = registry.execute(td.name, args)

            result = execution.output if execution.success else json.dumps({"error": execution.error})

            # PostToolUse hook
            hooks.execute(
                HookEvent.POST_TOOL_USE,
                tool_name=td.name,
                tool_input=args,
                tool_output=result,
            )

            # Prepend hook context if any
            if pre_result.extra_context:
                result = f"{pre_result.extra_context}\n{result}"

            return result

        # Set function metadata for FastMCP introspection
        handler.__name__ = td.name
        handler.__doc__ = td.description

        # Build proper annotations so FastMCP generates correct schema
        annotations = {}
        defaults = {}
        for param_name, param_def in props.items():
            annotations[param_name] = str
            if param_name not in required:
                defaults[param_name] = param_def.get("default", "")

        handler.__annotations__ = {**annotations, "return": str}
        if defaults:
            handler.__defaults__ = tuple(defaults.values())
            handler.__kwdefaults__ = defaults

        return handler

    fn = make_handler(tool_def)
    server.tool(name=tool_def.name, description=tool_def.description)(fn)


def main():
    parser = argparse.ArgumentParser(
        description="Harness MCP Server — expose plugin tools via MCP",
    )
    parser.add_argument("--plugin", choices=list(PLUGINS.keys()), default="excel",
                        help="Plugin to use (default: excel)")
    parser.add_argument("--file", required=True,
                        help="Target file path")
    parser.add_argument("--transport", choices=["stdio", "sse"], default="stdio",
                        help="MCP transport (default: stdio)")
    parser.add_argument("--port", type=int, default=8100,
                        help="Port for SSE transport (default: 8100)")
    args = parser.parse_args()

    file_path = os.path.abspath(args.file)
    server = build_server(args.plugin, file_path)

    if args.transport == "sse":
        server.run(transport="sse", port=args.port)
    else:
        server.run(transport="stdio")


if __name__ == "__main__":
    main()
