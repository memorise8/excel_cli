"""Generic context loader — builds runtime context from plugin + CONTEXT.md."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class RuntimeContext:
    """Loaded context for an AI session."""
    system_prompt: str
    context_path: str
    file_summary: dict
    skill_names: list[str]
    tool_names: list[str]


class ContextLoader:
    """Loads CONTEXT.md and builds runtime context.

    The system prompt comes from the plugin — this class only appends
    project-level CONTEXT.md and tool/skill metadata.
    """

    def __init__(self, project_root: Path | None = None):
        self.project_root = project_root or Path(__file__).parent.parent.parent
        self.context_path = self.project_root / "CONTEXT.md"

    def load_context_file(self) -> str:
        """Load CONTEXT.md content."""
        if self.context_path.exists():
            return self.context_path.read_text(encoding="utf-8")
        return ""

    def build_runtime_context(
        self,
        context_path: str,
        file_summary: dict,
        skill_names: list[str],
        tool_names: list[str],
        plugin_system_prompt: str = "",
    ) -> RuntimeContext:
        """Build complete runtime context for a session."""
        system_prompt = self._build_system_prompt(
            plugin_system_prompt, context_path, file_summary, skill_names, tool_names
        )

        return RuntimeContext(
            system_prompt=system_prompt,
            context_path=context_path,
            file_summary=file_summary,
            skill_names=skill_names,
            tool_names=tool_names,
        )

    def _build_system_prompt(
        self,
        plugin_prompt: str,
        context_path: str,
        file_summary: dict,
        skill_names: list[str],
        tool_names: list[str],
    ) -> str:
        parts = []

        # Plugin provides the domain-specific system prompt
        if plugin_prompt:
            parts.append(plugin_prompt)

        # Append project-level CONTEXT.md
        context_md = self.load_context_file()
        if context_md:
            parts.append(f"\n## Operational Context\n{context_md}")

        # Current target
        if context_path:
            parts.append(f"\n## Current Target\nPath: {context_path}")

        # Available tools
        if tool_names:
            parts.append(f"\n## Available Tools\n{', '.join(tool_names)}")

        # Common guidelines
        parts.append(
            "\n## General Guidelines\n"
            "- Respond in the same language the user uses.\n"
            "- For destructive operations, confirm with the user before proceeding."
        )

        return "\n".join(parts)
