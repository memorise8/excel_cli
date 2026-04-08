"""Generic runtime session — manages the full lifecycle of an AI conversation.

All domain-specific logic comes from the plugin.
"""

from __future__ import annotations

import uuid
from dataclasses import dataclass, field
from pathlib import Path
from typing import TYPE_CHECKING

from .context import ContextLoader, RuntimeContext
from .hooks import HookEngine, HookEvent
from .skills import SkillLoader
from .tools import ToolRegistry

if TYPE_CHECKING:
    from plugins.base import HarnessPlugin


@dataclass
class HarnessSession:
    """A complete session: target + tools + hooks + LLM conversation."""
    context_path: str  # Primary file/resource being operated on
    plugin_name: str = ""
    provider: str = "openai"
    model: str = ""
    messages: list[dict] = field(default_factory=list)
    context: RuntimeContext | None = None
    tools: ToolRegistry | None = None
    hooks: HookEngine = field(default_factory=HookEngine)
    skills: SkillLoader = field(default_factory=SkillLoader)
    session_id: str = ""

    @classmethod
    def create(
        cls,
        plugin: HarnessPlugin,
        context_path: str,
        provider: str = "openai",
        model: str = "",
        cli_path: str = "",
        project_root: Path | None = None,
        hooks_config: Path | None = None,
    ) -> "HarnessSession":
        """Create a fully initialized session using the given plugin."""
        root = project_root or Path(__file__).parent.parent.parent

        # Locate CLI binary — plugin override or explicit path
        if not cli_path:
            cli_path = plugin.find_cli(root)

        # Load skills
        skill_loader = SkillLoader(root / "skills")
        skill_loader.discover_skills()

        # Build tool registry (empty — plugin fills it)
        tool_registry = ToolRegistry(context_path=context_path, cli_path=cli_path)
        plugin.register_tools(tool_registry)

        # Load hooks
        hook_engine = HookEngine()
        if hooks_config and hooks_config.exists():
            hook_engine.load_from_config(hooks_config)

        # Register plugin-specific default hooks
        for hook_def in plugin.get_default_hooks():
            hook_engine.register_fn(
                event=HookEvent(hook_def["event"]),
                fn=hook_def["fn"],
                matcher=hook_def.get("matcher"),
                description=hook_def.get("description", ""),
            )

        # Get file summary from plugin
        file_summary = plugin.get_file_summary(context_path)

        # Build system prompt from plugin
        plugin_prompt = plugin.build_system_prompt(
            file_path=context_path,
            file_summary=file_summary,
            skill_names=skill_loader.get_skill_names(),
            tool_names=tool_registry.tool_names(),
        )

        # Build runtime context
        ctx_loader = ContextLoader(root)
        context = ctx_loader.build_runtime_context(
            context_path=context_path,
            file_summary=file_summary,
            skill_names=skill_loader.get_skill_names(),
            tool_names=tool_registry.tool_names(),
            plugin_system_prompt=plugin_prompt,
        )

        return cls(
            context_path=context_path,
            plugin_name=plugin.name,
            provider=provider,
            model=model,
            context=context,
            tools=tool_registry,
            hooks=hook_engine,
            skills=skill_loader,
            session_id=uuid.uuid4().hex[:12],
        )
