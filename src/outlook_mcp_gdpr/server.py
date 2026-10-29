"""MCP server exposing an Outlook mailbox with GDPR pseudonymisation.

Rule of thumb: every string returned to the model goes through
`R.redact`, every string coming *from* the model that will be written to
Graph goes through `R.restore`.
"""

from __future__ import annotations

import logging
import os
import sys
from pathlib import Path
from typing import Any

from mcp.server.mcpserver import MCPServer

from .graph import Auth, Graph
from .redact import Redactor

INSTRUCTIONS = """\
Den här servern ger åtkomst till användarens Outlook-brevlåda via Microsoft Graph.
Personuppgifter är pseudonymiserade: namn, e-postadresser, telefonnummer,
personnummer, adresser m.m. visas som tokens, t.ex. [[PERSON_1]], [[PERSON_1_FIRST]]
eller [[EMAIL_2]]. Försök aldrig gissa vad som döljs bakom en token.
När du skriver utkast eller söker: använd tokens ordagrant, så ersätts de med
riktiga värden innan utkastet sparas i Outlook. Servern kan inte skicka mejl –
den skapar endast utkast som användaren själv granskar och skickar.
"""

mcp = MCPServer("outlook-gdpr", instructions=INSTRUCTIONS)


def _load_list(env: str) -> list[str]:
    raw = os.environ.get(env, "")
    path = os.environ.get(env + "_FILE")
    items = [x.strip() for x in raw.split(",")]
    if path:
        items += [line.strip() for line in Path(path).read_text().splitlines()]
    return [x for x in items if x and not x.startswith("#")]


R = Redactor(
    allowlist=_load_list("REDACT_ALLOWLIST"),
    extra_names=_load_list("REDACT_NAMES"),
    use_ner=os.environ.get("REDACT_USE_NER") == "1",
)
_graph: Graph | None = None


def graph() -> Graph:
    global _graph
    if _graph is None:
        _graph = Graph()
    return _graph


# ---- helpers ------------------------------------------------------------

def _learn(msg: dict) -> None:
    """Register people from headers before redacting the body."""
    addrs = [msg.get("from")] + msg.get("toRecipients", []) + msg.get("ccRecipients", [])
    for a in addrs:
        ea = (a or {}).get("emailAddress") or {}
        R.register_person(ea.get("name"), ea.get("address"))


def _addr(a: dict | None) -> str:
    ea = (a or {}).get("emailAddress") or {}
    name, address = ea.get("name") or "", ea.get("address") or ""
    if name and name != address:
        return R.redact(f"{name} <{address}>")
    return R.redact(address)


def _summary(m: dict) -> dict[str, Any]:
    _learn(m)
    return {
        "id": m["id"],
        "received": m.get("receivedDateTime"),
        "from": _addr(m.get("from")),
        "subject": R.redact(m.get("subject")),
        "preview": R.redact(m.get("bodyPreview")),
        "is_read": m.get("isRead"),
        "is_draft": m.get("isDraft"),
        "has_attachments": m.get("hasAttachments"),
        "flagged": (m.get("flag") or {}).get("flagStatus") == "flagged",
    }


def _recipients(addresses: list[str] | None) -> list[dict]:
    return [{"emailAddress": {"address": R.restore(a).strip()}} for a in addresses or []]


# ---- tools --------------------------------------------------------------

@mcp.tool()
def list_folders() -> list[dict]:
    """Lista mappar i brevlådan med antal (olästa) meddelanden."""
    return [
        {"id": f["id"], "name": R.redact(f["displayName"]),
         "unread": f.get("unreadItemCount"), "total": f.get("totalItemCount")}
        for f in graph().list_folders()
    ]


@mcp.tool()
def list_messages(folder: str = "inbox", top: int = 20, skip: int = 0, unread_only: bool = False) -> list[dict]:
    """Lista meddelanden i en mapp, nyast först.

    folder: mapp-id eller välkänt namn (inbox, drafts, sentitems, archive, deleteditems, junkemail).
    """
    top = max(1, min(top, 50))
    return [_summary(m) for m in graph().list_messages(folder, top, skip, unread_only)]


@mcp.tool()
def search_messages(query: str, top: int = 20) -> list[dict]:
    """Fritextsök i hela brevlådan. Tokens som [[PERSON_1]] får användas i söksträngen."""
    top = max(1, min(top, 50))
    return [_summary(m) for m in graph().search(R.restore(query), top)]


@mcp.tool()
def read_message(message_id: str) -> dict:
    """Läs ett helt meddelande (pseudonymiserat). Bilagornas innehåll hämtas inte."""
    m = graph().get_message(message_id)
    _learn(m)
    return {
        "id": m["id"],
        "conversation_id": m.get("conversationId"),
        "received": m.get("receivedDateTime"),
        "from": _addr(m.get("from")),
        "to": [_addr(a) for a in m.get("toRecipients", [])],
        "cc": [_addr(a) for a in m.get("ccRecipients", [])],
        "subject": R.redact(m.get("subject")),
        "body": R.redact((m.get("body") or {}).get("content")),
        "is_read": m.get("isRead"),
        "is_draft": m.get("isDraft"),
        "importance": m.get("importance"),
        "categories": m.get("categories"),
        "attachments": [
            {"name": R.redact(a.get("name")), "type": a.get("contentType"), "size": a.get("size")}
            for a in m.get("attachments", [])
        ],
    }


@mcp.tool()
def mark_read(message_id: str, is_read: bool = True) -> str:
    """Markera ett meddelande som läst eller oläst."""
    graph().update_message(message_id, {"isRead": is_read})
    return "ok"


@mcp.tool()
def flag_message(message_id: str, flagged: bool = True) -> str:
    """Flagga eller avflagga ett meddelande för uppföljning."""
    graph().update_message(message_id, {"flag": {"flagStatus": "flagged" if flagged else "notFlagged"}})
    return "ok"


@mcp.tool()
def set_categories(message_id: str, categories: list[str]) -> str:
    """Sätt Outlook-kategorier på ett meddelande (ersätter befintliga)."""
    graph().update_message(message_id, {"categories": categories})
    return "ok"


@mcp.tool()
def move_message(message_id: str, destination: str) -> dict:
    """Flytta ett meddelande. destination: mapp-id eller välkänt namn (archive, deleteditems, inbox …)."""
    m = graph().move(message_id, destination)
    return {"id": m["id"]}


@mcp.tool()
def create_draft(subject: str, body: str, to: list[str] | None = None, cc: list[str] | None = None) -> dict:
    """Skapa ett nytt utkast i Utkast-mappen. Skickas INTE.

    Använd tokens (t.ex. [[EMAIL_1]], [[PERSON_2_FIRST]]) ordagrant – de ersätts med riktiga värden.
    """
    payload = {
        "subject": R.restore(subject),
        "body": {"contentType": "Text", "content": R.restore(body)},
        "toRecipients": _recipients(to),
        "ccRecipients": _recipients(cc),
    }
    d = graph().create_draft(payload)
    return {"id": d["id"], "status": "Utkast sparat i Outlook – granska och skicka själv."}


@mcp.tool()
def create_reply_draft(message_id: str, body: str, reply_all: bool = False) -> dict:
    """Skapa ett svarsutkast på ett befintligt meddelande. Skickas INTE."""
    d = graph().create_reply_draft(message_id, R.restore(body), reply_all)
    return {"id": d["id"], "status": "Svarsutkast sparat i Outlook – granska och skicka själv."}


@mcp.tool()
def update_draft(
    draft_id: str,
    subject: str | None = None,
    body: str | None = None,
    to: list[str] | None = None,
    cc: list[str] | None = None,
) -> str:
    """Uppdatera ett befintligt utkast. Fält som utelämnas lämnas orörda."""
    patch: dict[str, Any] = {}
    if subject is not None:
        patch["subject"] = R.restore(subject)
    if body is not None:
        patch["body"] = {"contentType": "Text", "content": R.restore(body)}
    if to is not None:
        patch["toRecipients"] = _recipients(to)
    if cc is not None:
        patch["ccRecipients"] = _recipients(cc)
    _require_draft(draft_id)
    graph().update_message(draft_id, patch)
    return "ok"


@mcp.tool()
def delete_draft(draft_id: str) -> str:
    """Radera ett utkast. Fungerar bara på utkast, aldrig på mottagna mejl."""
    _require_draft(draft_id)
    graph().delete(draft_id)
    return "ok"


def _require_draft(message_id: str) -> None:
    m = graph().request("GET", f"/me/messages/{message_id}", params={"$select": "isDraft"})
    if not m.get("isDraft"):
        raise ValueError("Meddelandet är inte ett utkast.")


def main() -> None:
    # httpx loggar hela URL:er på INFO, inklusive söksträngar med riktiga namn.
    logging.getLogger("httpx").setLevel(logging.WARNING)
    if len(sys.argv) > 1 and sys.argv[1] == "login":
        Auth().device_login()
        print("Inloggad. Token-cache sparad.", file=sys.stderr)
        return
    mcp.run("stdio")


if __name__ == "__main__":
    main()
