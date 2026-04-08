"""Generic tool registry and execution for the harness framework.

Tools are registered by plugins — this module contains no domain-specific logic.
"""

from __future__ import annotations

import json
import subprocess
from dataclasses import dataclass
from typing import Callable


@dataclass(frozen=True)
class ToolDef:
    """Tool definition for the LLM."""
    name: str
    description: str
    parameters: dict
    executor: Callable  # fn(args, context_path) -> str


@dataclass(frozen=True)
class ToolExecution:
    """Result of executing a tool."""
    name: str
    args: dict
    output: str
    success: bool
    error: str = ""


class ToolRegistry:
    """Generic tool registry — plugins register tools, engine executes them.

    Provides a CLI runner utility that any plugin can use to call
    its backing CLI binary via subprocess.
    """

    def __init__(self, context_path: str = "", cli_path: str = ""):
        self._tools: dict[str, ToolDef] = {}
        self.context_path = context_path  # Primary file/resource being operated on
        self.cli_path = cli_path

    def register(self, tool: ToolDef) -> None:
        self._tools[tool.name] = tool

    def get_tool(self, name: str) -> ToolDef | None:
        return self._tools.get(name)

    def list_tools(self) -> list[ToolDef]:
        return list(self._tools.values())

    def tool_names(self) -> list[str]:
        return list(self._tools.keys())

    def get_llm_tool_definitions(self) -> list[dict]:
        """Get tool definitions in OpenAI function calling format."""
        return [
            {
                "type": "function",
                "function": {
                    "name": t.name,
                    "description": t.description,
                    "parameters": t.parameters,
                },
            }
            for t in self._tools.values()
        ]

    def execute(self, name: str, args: dict) -> ToolExecution:
        """Execute a tool by name."""
        tool = self._tools.get(name)
        if not tool:
            return ToolExecution(
                name=name, args=args, output="", success=False,
                error=f"Unknown tool: {name}",
            )
        try:
            output = tool.executor(args, self.context_path)
            return ToolExecution(name=name, args=args, output=output, success=True)
        except Exception as e:
            return ToolExecution(
                name=name, args=args, output="", success=False, error=str(e),
            )

    # ── CLI runner utility (used by plugins) ──

    def run_cli(self, *args: str, timeout: int = 60) -> str:
        """Run the backing CLI binary as a subprocess. Returns stdout or JSON error."""
        if not self.cli_path:
            return json.dumps({"error": "CLI binary not available"})
        cmd = [self.cli_path] + list(args)
        try:
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout)
            if result.returncode == 0:
                return result.stdout.strip()
            error = result.stderr.strip() or result.stdout.strip()
            return json.dumps({"error": error})
        except subprocess.TimeoutExpired:
            return json.dumps({"error": f"Command timed out after {timeout} seconds"})
        except Exception as e:
            return json.dumps({"error": str(e)})
