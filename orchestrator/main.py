#!/usr/bin/env python3
"""
CLI Orchestrator — AI-powered assistant with plugin architecture.

Usage:
    python3 orchestrator/main.py --plugin excel --provider openai report.xlsx
    python3 orchestrator/main.py --plugin excel --provider gemini report.xlsx
"""

import argparse
import os
import sys
from pathlib import Path

from dotenv import load_dotenv

load_dotenv(Path(__file__).parent / ".env")

from harness import HarnessSession
from harness.engine import HarnessEngine
from plugins.excel import ExcelPlugin
from plugins.excel_com import ExcelComPlugin


# Plugin registry — add new plugins here
PLUGINS = {
    "excel": ExcelPlugin,
    "excel_com": ExcelComPlugin,
}


def main():
    parser = argparse.ArgumentParser(
        description="CLI Orchestrator — AI-powered assistant with plugin architecture",
    )
    parser.add_argument("file", help="Target file to analyze")
    parser.add_argument("--plugin", choices=list(PLUGINS.keys()), default="excel",
                        help="Plugin to use (default: excel)")
    parser.add_argument("--provider", "-p", choices=["openai", "gemini"], default="openai")
    parser.add_argument("--model", "-m", default=None)
    args = parser.parse_args()

    file_path = os.path.abspath(args.file)

    # Initialize plugin
    plugin = PLUGINS[args.plugin]()

    # Validate file
    ok, err = plugin.validate_file(file_path)
    if not ok:
        print(f"Error: {err}")
        sys.exit(1)

    provider = args.provider
    model = args.model or ("gpt-5.4" if provider == "openai" else "gemini-3.1-pro-preview")

    if provider == "openai" and not os.environ.get("OPENAI_API_KEY"):
        print("Error: OPENAI_API_KEY not set. Check orchestrator/.env")
        sys.exit(1)
    if provider == "gemini" and not os.environ.get("GEMINI_API_KEY"):
        print("Error: GEMINI_API_KEY not set. Check orchestrator/.env")
        sys.exit(1)

    # Create harness session with plugin
    session = HarnessSession.create(
        plugin=plugin,
        context_path=file_path,
        provider=provider,
        model=model,
        project_root=Path(__file__).parent.parent,
    )

    engine = HarnessEngine(session)

    # Print session info
    summary = session.context.file_summary
    print(f"Harness ({plugin.name}/{provider}/{model})")
    print(f"Target: {file_path}")
    print(f"Session: {session.session_id}")

    if summary and summary.get("sheets"):
        sheets_desc = ", ".join(
            f"{s.get('name', '?')}({s.get('rows', 0)}x{s.get('cols', 0)})"
            for s in summary.get("sheets", [])
        )
        print(f"Sheets: {sheets_desc}")
        print(f"Size: {summary.get('size_bytes', 0):,} bytes")

    print(f"Tools: {len(session.tools.tool_names())} registered")
    print(f"Skills: {len(session.skills.get_skill_names())} loaded")
    print("Type 'quit' to exit, 'reset' to clear history.\n")

    while True:
        try:
            user_input = input("You: ").strip()
        except (EOFError, KeyboardInterrupt):
            print("\nBye!")
            break

        if not user_input:
            continue
        if user_input.lower() == "quit":
            print("Bye!")
            break
        if user_input.lower() == "reset":
            session.messages.clear()
            print("History cleared.\n")
            continue

        try:
            response = engine.chat(user_input)
            print(f"\nAssistant: {response}\n")
        except Exception as e:
            print(f"\nError: {e}\n")


if __name__ == "__main__":
    main()
