"""Hook engine — PreToolUse / PostToolUse lifecycle hooks."""

from __future__ import annotations

import json
import subprocess
from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path
from typing import Callable


class HookEvent(str, Enum):
    PRE_TOOL_USE = "PreToolUse"
    POST_TOOL_USE = "PostToolUse"
    SESSION_START = "SessionStart"
    SESSION_END = "SessionEnd"


class HookDecision(str, Enum):
    ALLOW = "allow"
    DENY = "deny"
    ASK = "ask"


@dataclass
class HookResult:
    decision: HookDecision = HookDecision.ALLOW
    reason: str = ""
    modified_input: dict | None = None
    extra_context: str = ""


@dataclass
class HookDef:
    """A single hook definition."""
    event: HookEvent
    matcher: str | None = None  # Tool name pattern to match
    hook_type: str = "command"  # "command", "python", "prompt"
    command: str = ""  # Shell command to run
    python_fn: Callable | None = None  # Python function hook
    description: str = ""


class HookEngine:
    """Manages and executes lifecycle hooks, inspired by Claude Code hooks."""

    def __init__(self):
        self._hooks: list[HookDef] = []

    def register(self, hook: HookDef) -> None:
        """Register a hook."""
        self._hooks.append(hook)

    def register_fn(
        self,
        event: HookEvent,
        fn: Callable,
        matcher: str | None = None,
        description: str = "",
    ) -> None:
        """Register a Python function as a hook."""
        self._hooks.append(
            HookDef(
                event=event,
                matcher=matcher,
                hook_type="python",
                python_fn=fn,
                description=description,
            )
        )

    def register_command(
        self,
        event: HookEvent,
        command: str,
        matcher: str | None = None,
        description: str = "",
    ) -> None:
        """Register a shell command as a hook."""
        self._hooks.append(
            HookDef(
                event=event,
                matcher=matcher,
                hook_type="command",
                command=command,
                description=description,
            )
        )

    def execute(
        self,
        event: HookEvent,
        tool_name: str = "",
        tool_input: dict | None = None,
        tool_output: str = "",
    ) -> HookResult:
        """Execute all hooks for an event. Returns combined result."""
        matching_hooks = [
            h
            for h in self._hooks
            if h.event == event and self._matches(h.matcher, tool_name)
        ]

        if not matching_hooks:
            return HookResult()

        for hook in matching_hooks:
            result = self._execute_hook(hook, tool_name, tool_input, tool_output)
            if result.decision == HookDecision.DENY:
                return result  # Stop on first deny

        return HookResult()

    def load_from_config(self, config_path: Path) -> None:
        """Load hooks from a JSON config file (Claude Code settings.json format)."""
        if not config_path.exists():
            return

        config = json.loads(config_path.read_text())
        hooks_config = config.get("hooks", {})

        for event_name, hook_list in hooks_config.items():
            try:
                event = HookEvent(event_name)
            except ValueError:
                continue

            for hook_def in hook_list:
                matcher = hook_def.get("matcher")
                for h in hook_def.get("hooks", []):
                    self.register(
                        HookDef(
                            event=event,
                            matcher=matcher,
                            hook_type=h.get("type", "command"),
                            command=h.get("command", ""),
                            description=h.get("description", ""),
                        )
                    )

    def _matches(self, matcher: str | None, tool_name: str) -> bool:
        """Check if a hook matcher matches a tool name."""
        if matcher is None:
            return True
        # Support simple patterns: "Bash", "excel_*", "excel_range_read"
        if "*" in matcher:
            prefix = matcher.replace("*", "")
            return tool_name.startswith(prefix)
        return matcher == tool_name or matcher in tool_name

    def _execute_hook(
        self,
        hook: HookDef,
        tool_name: str,
        tool_input: dict | None,
        tool_output: str,
    ) -> HookResult:
        """Execute a single hook."""
        if hook.hook_type == "python" and hook.python_fn:
            try:
                return hook.python_fn(tool_name, tool_input, tool_output)
            except Exception as e:
                return HookResult(
                    decision=HookDecision.ALLOW,
                    reason=f"Hook error: {e}",
                )

        if hook.hook_type == "command" and hook.command:
            return self._execute_command_hook(
                hook.command, tool_name, tool_input, tool_output
            )

        return HookResult()

    def _execute_command_hook(
        self,
        command: str,
        tool_name: str,
        tool_input: dict | None,
        tool_output: str,
    ) -> HookResult:
        """Execute a shell command hook, passing context via stdin."""
        stdin_data = json.dumps(
            {
                "tool_name": tool_name,
                "tool_input": tool_input or {},
                "tool_output": tool_output,
            }
        )

        try:
            result = subprocess.run(
                command,
                shell=True,
                input=stdin_data,
                capture_output=True,
                text=True,
                timeout=10,
            )

            if result.returncode == 2:
                return HookResult(
                    decision=HookDecision.DENY,
                    reason=result.stderr.strip() or "Blocked by hook",
                )

            extra = result.stdout.strip()
            return HookResult(
                decision=HookDecision.ALLOW,
                extra_context=extra,
            )

        except (subprocess.TimeoutExpired, Exception) as e:
            return HookResult(
                decision=HookDecision.ALLOW,
                reason=f"Hook timeout/error: {e}",
            )
