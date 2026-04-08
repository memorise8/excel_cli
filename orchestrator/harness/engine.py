"""Harness engine — manages LLM turns with tool calling and hook execution."""

from __future__ import annotations

import json
from dataclasses import dataclass

from .hooks import HookEvent, HookDecision
from .runtime import HarnessSession
from .tools import ToolExecution


@dataclass(frozen=True)
class TurnResult:
    """Result of one LLM turn."""
    response: str
    tool_calls: list[ToolExecution]
    turn_count: int
    stop_reason: str


class HarnessEngine:
    """Orchestrates LLM conversation with tool calling and hooks."""

    def __init__(self, session: HarnessSession):
        self.session = session
        self._provider = None
        self._init_provider()

    def _init_provider(self):
        if self.session.provider == "openai":
            from providers.openai import OpenAIProvider
            self._provider = OpenAIProvider(self.session.model or "gpt-5.4")
        elif self.session.provider == "gemini":
            from providers.gemini import GeminiProvider
            self._provider = GeminiProvider(self.session.model or "gemini-3.1-pro-preview")
        else:
            raise ValueError(f"Unknown provider: {self.session.provider}")

    def chat(self, user_message: str) -> str:
        """Process a user message through the full harness pipeline."""
        self.session.messages.append({"role": "user", "content": user_message})

        # SessionStart hook (first message only)
        if len(self.session.messages) == 1:
            self.session.hooks.execute(HookEvent.SESSION_START)

        max_iterations = 10
        tool_calls_made = []

        for _ in range(max_iterations):
            # Call LLM
            response = self._provider.chat(
                messages=self.session.messages,
                system_prompt=self.session.context.system_prompt,
                tools=self.session.tools.get_llm_tool_definitions(),
            )

            # Check if LLM wants to call tools
            if response.tool_calls:
                # Add assistant message with tool calls
                self.session.messages.append(response.to_message_dict())

                for tc in response.tool_calls:
                    # PreToolUse hook
                    pre_result = self.session.hooks.execute(
                        HookEvent.PRE_TOOL_USE,
                        tool_name=tc.name,
                        tool_input=tc.args,
                    )

                    if pre_result.decision == HookDecision.DENY:
                        tool_result = json.dumps({
                            "error": f"Blocked by hook: {pre_result.reason}"
                        })
                    else:
                        # Use modified input if hook provided one
                        args = pre_result.modified_input or tc.args

                        # Execute tool
                        execution = self.session.tools.execute(tc.name, args)
                        tool_calls_made.append(execution)

                        tool_result = execution.output if execution.success else json.dumps({"error": execution.error})

                        # PostToolUse hook
                        self.session.hooks.execute(
                            HookEvent.POST_TOOL_USE,
                            tool_name=tc.name,
                            tool_input=args,
                            tool_output=tool_result,
                        )

                        # Prepend hook context if any
                        if pre_result.extra_context:
                            tool_result = f"{pre_result.extra_context}\n{tool_result}"

                    # Add tool result to messages
                    self.session.messages.append(
                        self._provider.format_tool_result(tc.id, tool_result)
                    )
            else:
                # No tool calls — final response
                content = response.content
                self.session.messages.append({"role": "assistant", "content": content})
                return content

        return "Maximum tool call iterations reached."


@dataclass
class LLMToolCall:
    """A tool call from the LLM."""
    id: str
    name: str
    args: dict


@dataclass
class LLMResponse:
    """Response from an LLM provider."""
    content: str | None
    tool_calls: list[LLMToolCall]

    def to_message_dict(self) -> dict:
        """Convert to OpenAI message format."""
        msg = {"role": "assistant", "content": self.content}
        if self.tool_calls:
            msg["tool_calls"] = [
                {
                    "id": tc.id,
                    "type": "function",
                    "function": {"name": tc.name, "arguments": json.dumps(tc.args)},
                }
                for tc in self.tool_calls
            ]
        return msg
