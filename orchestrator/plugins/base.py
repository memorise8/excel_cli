"""Abstract base class for harness plugins.

A plugin provides domain-specific tools, context, and file handling
for the generic harness framework. Implement this to support any CLI tool.
"""

from __future__ import annotations

from abc import ABC, abstractmethod
from pathlib import Path

from harness.tools import ToolRegistry


class HarnessPlugin(ABC):
    """Base class for all harness plugins.

    Each plugin adapts a specific CLI tool (excel-cli, kubectl, gh, etc.)
    to the harness framework by providing:
    - Tool definitions and executors
    - System prompt and context for the LLM
    - File validation and summarization
    """

    @property
    @abstractmethod
    def name(self) -> str:
        """Plugin identifier, e.g. 'excel', 'k8s', 'git'."""

    @property
    @abstractmethod
    def description(self) -> str:
        """One-line description of what this plugin does."""

    @abstractmethod
    def register_tools(self, registry: ToolRegistry) -> None:
        """Register domain-specific tools into the generic registry."""

    @abstractmethod
    def build_system_prompt(
        self,
        file_path: str,
        file_summary: dict,
        skill_names: list[str],
        tool_names: list[str],
    ) -> str:
        """Build the LLM system prompt with domain-specific instructions."""

    @abstractmethod
    def get_file_summary(self, file_path: str) -> dict:
        """Analyze the target file and return a summary dict."""

    @abstractmethod
    def validate_file(self, file_path: str) -> tuple[bool, str]:
        """Validate that the file is supported. Returns (ok, error_message)."""

    def find_cli(self, project_root: Path) -> str:
        """Locate the CLI binary. Override if needed. Default returns empty."""
        return ""

    def get_default_hooks(self) -> list[dict]:
        """Return default hook definitions for this plugin. Override if needed."""
        return []
