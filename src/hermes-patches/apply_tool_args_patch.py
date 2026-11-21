#!/usr/bin/env python3
"""Pass each tool call's short arguments to the edge on Hermes' tool progress event.

The pinned Hermes Chat Completions stream reports a tool start with ``label``, which is
``build_tool_preview`` output: a short, already-truncated preview ("What is the..."). Teams step
names must be full and specific ("Searching the web for this week's biggest news"), so the edge
needs the real arguments. Only short scalar values are passed, never code, file content or message
bodies, and the stream stays on the container's loopback between Hermes, the gateway and the edge.
"""

from __future__ import annotations

import os
from pathlib import Path


ROOT = Path(os.environ.get("HERMES_PATCH_ROOT", "/opt/hermes"))
MARKER = "edge step names need the real tool arguments"

OLD = '''                label = build_tool_preview(function_name, function_args) or function_name
                _stream_q.put(("__tool_progress__", {
                    "tool": function_name,
                    "emoji": get_tool_emoji(function_name),
                    "label": label,
                    "toolCallId": tool_call_id,
                    "status": "running",
                }))
'''

NEW = '''                label = build_tool_preview(function_name, function_args) or function_name
                # Hermes Foundry: edge step names need the real tool arguments. Short scalars only.
                _step_args = {}
                if isinstance(function_args, dict):
                    for _k, _v in list(function_args.items())[:12]:
                        if str(_k).lower() in {"code", "content", "body", "csv_content",
                                               "append_content", "find_text", "replace_text",
                                               "message", "text", "html", "script"}:
                            continue
                        if isinstance(_v, (str, int, float, bool)):
                            _step_args[str(_k)[:40]] = str(_v)[:300]
                _stream_q.put(("__tool_progress__", {
                    "tool": function_name,
                    "emoji": get_tool_emoji(function_name),
                    "label": label,
                    "args": _step_args,
                    "toolCallId": tool_call_id,
                    "status": "running",
                }))
'''


def main() -> None:
    path = ROOT / "gateway" / "platforms" / "api_server.py"
    text = path.read_text(encoding="utf-8")
    if MARKER in text:
        return
    count = text.count(OLD)
    if count != 1:
        raise RuntimeError(f"tool progress args: expected one source match, found {count}")
    path.write_text(text.replace(OLD, NEW, 1), encoding="utf-8")


if __name__ == "__main__":
    main()
