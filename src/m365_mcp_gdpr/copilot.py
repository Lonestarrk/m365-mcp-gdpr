"""Optional second redaction layer: Microsoft 365 Copilot reviews the text.

Enabled with M365_MCP_REDACT_ENGINE=copilot. Uses the Microsoft 365 Copilot
Chat API (Graph beta), which needs a Microsoft 365 Copilot licence and seven
delegated scopes (`SCOPES`). Copilot gets editable instructions
(`copilot_instructions.md`, or M365_MCP_COPILOT_INSTRUCTIONS) and answers with
JSON: what to redact, and whether a text holds special-category data
(GDPR art. 9: health, allergies, diet, religion ...) and must be blocked.

The layer fails closed: if Copilot errors or answers in an unexpected shape,
every text in the batch is blocked rather than passed on unreviewed.
"""

from __future__ import annotations

import json
import os
import re
from dataclasses import dataclass, field
from importlib import resources
from pathlib import Path
from typing import Any

COPILOT = "https://graph.microsoft.com/beta/copilot/conversations"
# Required together by the Copilot Chat API; application permissions aren't supported.
SCOPES = [
    "Sites.Read.All", "Mail.Read", "People.Read.All", "OnlineMeetingTranscript.Read.All",
    "Chat.Read", "ChannelMessage.Read.All", "ExternalItem.Read.All",
]
KINDS = {"PERSON", "INDIRECT", "OTHER"}


def enabled() -> bool:
    engine = os.environ.get("M365_MCP_REDACT_ENGINE", "local").strip().lower()
    if engine not in ("local", "copilot"):
        raise ValueError(f"Okänd M365_MCP_REDACT_ENGINE: {engine}. Giltiga: local, copilot")
    return engine == "copilot"


def instructions() -> str:
    path = os.environ.get("M365_MCP_COPILOT_INSTRUCTIONS")
    if path:
        return Path(path).expanduser().read_text(encoding="utf-8")
    return resources.files(__package__).joinpath("copilot_instructions.md").read_text(encoding="utf-8")


@dataclass
class Verdict:
    block: bool = False
    reason: str = ""
    items: list[tuple[str, str]] = field(default_factory=list)  # (text, kind)


def _prompt(texts: list[str]) -> str:
    parts = [instructions(), "", "## Texter att granska"]
    for i, t in enumerate(texts, 1):
        parts += ["", f"<<<TEXT {i}", t, f"TEXT {i}>>>"]
    return "\n".join(parts)


def _parse(answer: str, count: int) -> list[Verdict]:
    match = re.search(r"\{.*\}", answer, re.DOTALL)
    if not match:
        raise ValueError("Copilot svarade inte med JSON")
    data = json.loads(match.group(0))
    by_index = {r["index"]: r for r in data["results"]}
    verdicts = []
    for i in range(1, count + 1):
        r = by_index[i]  # a missing index is an error: fail closed
        items = [
            (str(it["text"]), it.get("type", "OTHER") if it.get("type") in KINDS else "OTHER")
            for it in r.get("items") or []
            if str(it.get("text", "")).strip()
        ]
        verdicts.append(Verdict(bool(r.get("block")), str(r.get("reason") or ""), items))
    return verdicts


class CopilotScreener:
    def __init__(self, graph: Any) -> None:
        self._graph = graph

    def screen(self, texts: list[str]) -> list[Verdict]:
        """One Copilot call for the whole batch. Never raises: fails closed."""
        if not texts:
            return []
        try:
            conv = self._graph.request("POST", COPILOT, json={})
            body = {
                "message": {"text": _prompt(texts)},
                "locationHint": {"timeZone": os.environ.get("M365_MCP_TIMEZONE") or "UTC"},
                "contextualResources": {"webContext": {"isWebEnabled": False}},
            }
            reply = self._graph.request("POST", f"{COPILOT}/{conv['id']}/chat", json=body)
            answers = [m.get("text", "") for m in reply.get("messages", [])[1:]]
            return _parse(answers[-1] if answers else "", len(texts))
        except Exception as exc:  # noqa: BLE001 - any failure blocks the batch
            reason = f"Copilot-granskningen misslyckades ({type(exc).__name__}), så texten blockerades"
            return [Verdict(block=True, reason=reason) for _ in texts]
