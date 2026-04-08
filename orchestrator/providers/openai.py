"""OpenAI provider for the harness engine."""

from __future__ import annotations

import json
from openai import OpenAI

from harness.engine import LLMResponse, LLMToolCall


class OpenAIProvider:
    """OpenAI GPT provider."""

    def __init__(self, model: str = "gpt-5.4"):
        self.model = model
        self.client = OpenAI()

    def chat(
        self,
        messages: list[dict],
        system_prompt: str,
        tools: list[dict],
    ) -> LLMResponse:
        full_messages = [{"role": "system", "content": system_prompt}] + messages

        response = self.client.chat.completions.create(
            model=self.model,
            messages=full_messages,
            tools=tools if tools else None,
            tool_choice="auto" if tools else None,
        )

        msg = response.choices[0].message

        tool_calls = []
        if msg.tool_calls:
            for tc in msg.tool_calls:
                tool_calls.append(
                    LLMToolCall(
                        id=tc.id,
                        name=tc.function.name,
                        args=json.loads(tc.function.arguments),
                    )
                )

        return LLMResponse(content=msg.content, tool_calls=tool_calls)

    def format_tool_result(self, tool_call_id: str, result: str) -> dict:
        return {"role": "tool", "tool_call_id": tool_call_id, "content": result}
