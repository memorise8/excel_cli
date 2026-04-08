"""Gemini provider for the harness engine."""

from __future__ import annotations

import json
import os
import uuid

import google.generativeai as genai

from harness.engine import LLMResponse, LLMToolCall


class GeminiProvider:
    """Google Gemini provider."""

    def __init__(self, model: str = "gemini-3.1-pro-preview"):
        self.model = model
        genai.configure(api_key=os.environ.get("GEMINI_API_KEY"))
        self._chat_session = None
        self._gemini_model = None

    def _ensure_model(self, tools: list[dict]):
        """Initialize or reinitialize the model with tool definitions."""
        gemini_tools = self._convert_tools(tools)
        self._gemini_model = genai.GenerativeModel(
            model_name=self.model,
            tools=gemini_tools if gemini_tools else None,
        )
        self._chat_session = None  # Reset chat session

    def chat(
        self,
        messages: list[dict],
        system_prompt: str,
        tools: list[dict],
    ) -> LLMResponse:
        self._ensure_model(tools)

        # Build history from messages (excluding last)
        history = []
        for msg in messages[:-1]:
            if msg["role"] == "user":
                history.append({"role": "user", "parts": [msg["content"]]})
            elif msg["role"] == "assistant" and msg.get("content"):
                history.append({"role": "model", "parts": [msg["content"]]})

        chat = self._gemini_model.start_chat(history=history)

        last_msg = messages[-1]["content"]
        prompt = f"{system_prompt}\n\nUser: {last_msg}" if not history else last_msg

        response = chat.send_message(prompt)

        # Check for function calls
        tool_calls = []
        for part in response.parts:
            if hasattr(part, "function_call") and part.function_call.name:
                fc = part.function_call
                args = dict(fc.args) if fc.args else {}
                tool_calls.append(
                    LLMToolCall(id=uuid.uuid4().hex[:8], name=fc.name, args=args)
                )

        if tool_calls:
            # Store chat session for follow-up
            self._chat_session = chat
            return LLMResponse(content=None, tool_calls=tool_calls)

        return LLMResponse(content=response.text, tool_calls=[])

    def format_tool_result(self, tool_call_id: str, result: str) -> dict:
        """For Gemini, we handle tool results differently in the chat flow."""
        # Store for next chat turn
        return {"role": "tool", "tool_call_id": tool_call_id, "content": result}

    def _convert_tools(self, tools: list[dict]) -> list:
        """Convert OpenAI tool format to Gemini format."""
        if not tools:
            return []

        gemini_tools = []
        for tool in tools:
            func = tool.get("function", {})
            params = func.get("parameters", {})
            props = {}
            for pname, pdef in params.get("properties", {}).items():
                type_str = pdef.get("type", "string").upper()
                gemini_type = getattr(genai.protos.Type, type_str, genai.protos.Type.STRING)
                props[pname] = genai.protos.Schema(
                    type_=gemini_type,
                    description=pdef.get("description", ""),
                )

            gemini_tools.append(
                genai.protos.Tool(
                    function_declarations=[
                        genai.protos.FunctionDeclaration(
                            name=func.get("name", ""),
                            description=func.get("description", ""),
                            parameters=genai.protos.Schema(
                                type_=genai.protos.Type.OBJECT,
                                properties=props,
                                required=params.get("required", []),
                            ),
                        )
                    ]
                )
            )

        return gemini_tools
