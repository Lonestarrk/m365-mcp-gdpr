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

from . import audit, features
from .graph import Auth, Graph
from .redact import Redactor
from .text import html_to_text

INSTRUCTIONS = """\
Den här servern ger åtkomst till delar av användarens Microsoft 365 via Microsoft Graph.
Vilka delar som finns beror på inställningarna; använd bara de verktyg som listas.

Mejl: personuppgifter är pseudonymiserade. Namn, e-postadresser, telefonnummer,
personnummer, adresser m.m. visas som tokens, t.ex. [[PERSON_1]], [[PERSON_1_FIRST]]
eller [[EMAIL_2]]. Försök aldrig gissa vad som döljs bakom en token.
När du skriver utkast eller söker: använd tokens ordagrant, så ersätts de med
riktiga värden innan utkastet sparas i Outlook. Servern kan inte skicka mejl –
den skapar endast utkast som användaren själv granskar och skickar.

Kalender: du kan läsa kalendern och boka tid för användaren själv, men inte skapa
möten med deltagare (det skulle skicka inbjudningar). Tider anges i ISO 8601.

Teams: du kan läsa chattar och kanaler men inte skriva. Namn och innehåll är
pseudonymiserade på samma sätt som mejl.

SharePoint: releasefiler kan laddas upp till en förbestämd mapp
(sharepoint_upload_release) och listas därifrån. Packa gärna en release
som en enda zip-fil innan uppladdning.
"""

mcp = MCPServer("m365-gdpr", instructions=INSTRUCTIONS)


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
AUDIT = audit.AuditLog(R)
_graph: Graph | None = None


TOOL_FEATURE: dict[str, str] = {}


def tool(feature: str) -> Any:
    """Register an MCP tool belonging to a Microsoft 365 feature (see `features`).

    Calls show up on the audit page when it is enabled.
    """

    def register(fn: Any) -> Any:
        TOOL_FEATURE[fn.__name__] = feature
        return mcp.tool()(AUDIT.wrap(fn))

    return register


def apply_features(enabled: tuple[str, ...]) -> None:
    """Remove the tools of every feature that is not enabled."""
    for name, feature in TOOL_FEATURE.items():
        if feature not in enabled:
            mcp.remove_tool(name)


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


_folder_id: str | None = None


def _mail_folder() -> str | None:
    """Id of the only folder the mail tools may touch, if M365_MCP_MAIL_FOLDER is set."""
    global _folder_id
    name = os.environ.get("M365_MCP_MAIL_FOLDER", "").strip()
    if not name:
        return None
    if _folder_id is None:
        folder = graph().find_folder(name)
        if folder is None:
            raise ValueError(
                f"Mappen '{name}' finns inte. Skapa den i Outlook (överst eller under Inkorgen) först."
            )
        _folder_id = folder["id"]
    return _folder_id


def _require_in_folder(message_id: str, parent: str | None = None) -> None:
    allowed = _mail_folder()
    if allowed is None:
        return
    if (parent or graph().parent_folder(message_id)) != allowed:
        raise ValueError("Meddelandet ligger utanför den tillåtna mappen (M365_MCP_MAIL_FOLDER).")


def _recipients(addresses: list[str] | None) -> list[dict]:
    return [{"emailAddress": {"address": R.restore(a).strip()}} for a in addresses or []]


# ---- tools --------------------------------------------------------------

@tool("mail")
def list_folders() -> list[dict]:
    """Lista mappar i brevlådan med antal (olästa) meddelanden."""
    folders = [graph().get_folder(f)] if (f := _mail_folder()) else graph().list_folders()
    return [
        {"id": f["id"], "name": R.redact(f["displayName"]),
         "unread": f.get("unreadItemCount"), "total": f.get("totalItemCount")}
        for f in folders
    ]


@tool("mail")
def list_messages(folder: str | None = None, top: int = 20, skip: int = 0, unread_only: bool = False) -> list[dict]:
    """Lista meddelanden i en mapp, nyast först.

    folder: mapp-id eller välkänt namn (inbox, drafts, sentitems, archive, deleteditems, junkemail).
    Standard är Inkorgen, eller den tillåtna mappen om servern är begränsad till en mapp.
    """
    top = max(1, min(top, 50))
    allowed = _mail_folder()
    if allowed and folder not in (None, allowed, os.environ.get("M365_MCP_MAIL_FOLDER", "").strip()):
        raise ValueError("Servern är begränsad till en mapp (M365_MCP_MAIL_FOLDER); utelämna folder.")
    target = allowed or folder or "inbox"
    return [_summary(m) for m in graph().list_messages(target, top, skip, unread_only)]


@tool("mail")
def search_messages(query: str, top: int = 20) -> list[dict]:
    """Fritextsök i brevlådan (eller i den tillåtna mappen). Tokens som [[PERSON_1]] får användas."""
    top = max(1, min(top, 50))
    return [_summary(m) for m in graph().search(R.restore(query), top, _mail_folder())]


@tool("mail")
def read_message(message_id: str) -> dict:
    """Läs ett helt meddelande (pseudonymiserat). Bilagornas innehåll hämtas inte."""
    m = graph().get_message(message_id)
    _require_in_folder(message_id, m.get("parentFolderId"))
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


@tool("mail")
def mark_read(message_id: str, is_read: bool = True) -> str:
    """Markera ett meddelande som läst eller oläst."""
    _require_in_folder(message_id)
    graph().update_message(message_id, {"isRead": is_read})
    return "ok"


@tool("mail")
def flag_message(message_id: str, flagged: bool = True) -> str:
    """Flagga eller avflagga ett meddelande för uppföljning."""
    _require_in_folder(message_id)
    graph().update_message(message_id, {"flag": {"flagStatus": "flagged" if flagged else "notFlagged"}})
    return "ok"


@tool("mail")
def set_categories(message_id: str, categories: list[str]) -> str:
    """Sätt Outlook-kategorier på ett meddelande (ersätter befintliga)."""
    _require_in_folder(message_id)
    graph().update_message(message_id, {"categories": categories})
    return "ok"


@tool("mail")
def move_message(message_id: str, destination: str) -> dict:
    """Flytta ett meddelande. destination: mapp-id eller välkänt namn (archive, deleteditems, inbox …)."""
    if _mail_folder():
        raise ValueError("Flytt är avstängd när servern är begränsad till en mapp (M365_MCP_MAIL_FOLDER).")
    m = graph().move(message_id, destination)
    return {"id": m["id"]}


@tool("mail")
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


@tool("mail")
def create_reply_draft(message_id: str, body: str, reply_all: bool = False) -> dict:
    """Skapa ett svarsutkast på ett befintligt meddelande. Skickas INTE."""
    _require_in_folder(message_id)
    d = graph().create_reply_draft(message_id, R.restore(body), reply_all)
    return {"id": d["id"], "status": "Svarsutkast sparat i Outlook – granska och skicka själv."}


@tool("mail")
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


@tool("mail")
def delete_draft(draft_id: str) -> str:
    """Radera ett utkast. Fungerar bara på utkast, aldrig på mottagna mejl."""
    _require_draft(draft_id)
    graph().delete(draft_id)
    return "ok"


# ---- calendar -------------------------------------------------------------

def _event(e: dict, full: bool = False) -> dict[str, Any]:
    people = [e.get("organizer")] + (e.get("attendees") or [])
    for p in people:
        ea = (p or {}).get("emailAddress") or {}
        R.register_person(ea.get("name"), ea.get("address"))
    out: dict[str, Any] = {
        "id": e["id"],
        "subject": R.redact(e.get("subject")),
        "start": e.get("start"),
        "end": e.get("end"),
        "all_day": e.get("isAllDay"),
        "cancelled": e.get("isCancelled"),
        "online": e.get("isOnlineMeeting"),
        "show_as": e.get("showAs"),
        "location": R.redact((e.get("location") or {}).get("displayName")),
        "organizer": _addr(e.get("organizer")),
        "i_am_organizer": e.get("isOrganizer"),
        "attendees": [
            {"who": _addr(a), "response": ((a.get("status") or {}).get("response"))}
            for a in e.get("attendees") or []
        ],
    }
    if full:
        out["body"] = R.redact((e.get("body") or {}).get("content"))
    else:
        out["preview"] = R.redact(e.get("bodyPreview"))
    return out


@tool("calendar")
def calendar_list_events(start: str, end: str, top: int = 50) -> list[dict]:
    """Lista kalenderhändelser mellan start och end (ISO 8601, t.ex. 2026-11-02T00:00:00).

    Utan tidszon i tiden tolkas den som UTC, om inte M365_MCP_TIMEZONE är satt.
    """
    return [_event(e) for e in graph().calendar_view(start, end, max(1, min(top, 100)))]


@tool("calendar")
def calendar_get_event(event_id: str) -> dict:
    """Läs en kalenderhändelse med beskrivning (pseudonymiserad)."""
    return _event(graph().get_event(event_id), full=True)


@tool("calendar")
def calendar_create_event(
    subject: str,
    start: str,
    end: str,
    body: str | None = None,
    location: str | None = None,
    show_as: str = "busy",
    time_zone: str | None = None,
) -> dict:
    """Boka tid i användarens egen kalender, t.ex. fokustid. Bjuder aldrig in någon.

    start/end: ISO 8601 utan tidszon, t.ex. 2026-11-02T09:00:00.
    time_zone: t.ex. "Europe/Stockholm" (standard: M365_MCP_TIMEZONE eller UTC).
    show_as: free, tentative, busy, oof eller workingElsewhere.
    """
    tz = time_zone or os.environ.get("M365_MCP_TIMEZONE") or "UTC"
    payload: dict[str, Any] = {
        "subject": R.restore(subject),
        "start": {"dateTime": start, "timeZone": tz},
        "end": {"dateTime": end, "timeZone": tz},
        "showAs": show_as,
        "attendees": [],
    }
    if body:
        payload["body"] = {"contentType": "Text", "content": R.restore(body)}
    if location:
        payload["location"] = {"displayName": R.restore(location)}
    e = graph().create_event(payload)
    return {"id": e["id"], "status": "Bokat i din kalender. Inga inbjudningar skickades."}


@tool("calendar")
def calendar_delete_event(event_id: str) -> str:
    """Ta bort en händelse som användaren själv skapat och som saknar deltagare.

    Möten med deltagare tas inte bort, eftersom det skulle skicka avbokningar.
    """
    e = graph().get_event(event_id)
    if not e.get("isOrganizer"):
        raise ValueError("Du är inte organisatör för händelsen, så den tas inte bort.")
    if e.get("attendees"):
        raise ValueError("Händelsen har deltagare. Ta bort den i Outlook, så att du ser avbokningen som skickas.")
    graph().delete_event(event_id)
    return "ok"


# ---- Teams ----------------------------------------------------------------

def _teams_person(identity: dict | None) -> str:
    user = (identity or {}).get("user") or {}
    name = user.get("displayName")
    if not name:
        return ""
    R.register_person(name)
    return R.redact(name)


def _teams_message(m: dict) -> dict[str, Any]:
    sender = _teams_person(m.get("from")) or ("system" if m.get("messageType") != "message" else "")
    body = m.get("body") or {}
    content = body.get("content")
    text = html_to_text(content) if body.get("contentType") == "html" else (content or "")
    return {
        "id": m.get("id"),
        "time": m.get("createdDateTime"),
        "from": sender,
        "text": R.redact(text),
        "deleted": m.get("deletedDateTime") is not None,
    }


@tool("teams")
def teams_list_chats(top: int = 20) -> list[dict]:
    """Lista dina Teams-chattar, senast aktiva först."""
    chats = []
    for c in graph().list_chats(max(1, min(top, 50))):
        members = []
        for mem in c.get("members") or []:
            name = mem.get("displayName")
            if name:
                R.register_person(name, mem.get("email"))
                members.append(R.redact(name))
        chats.append({
            "id": c["id"],
            "type": c.get("chatType"),
            "topic": R.redact(c.get("topic")),
            "members": members,
            "last_activity": c.get("lastUpdatedDateTime"),
        })
    return chats


@tool("teams")
def teams_read_chat(chat_id: str, top: int = 30) -> list[dict]:
    """Läs de senaste meddelandena i en chatt (pseudonymiserade), nyast först."""
    return [_teams_message(m) for m in graph().chat_messages(chat_id, max(1, min(top, 50)))]


@tool("teams")
def teams_list_teams() -> list[dict]:
    """Lista de team du är med i."""
    return [
        {"id": t["id"], "name": R.redact(t.get("displayName")), "description": R.redact(t.get("description"))}
        for t in graph().joined_teams()
    ]


@tool("teams")
def teams_list_channels(team_id: str) -> list[dict]:
    """Lista kanalerna i ett team."""
    return [
        {"id": c["id"], "name": R.redact(c.get("displayName")), "description": R.redact(c.get("description"))}
        for c in graph().channels(team_id)
    ]


@tool("teams")
def teams_read_channel(team_id: str, channel_id: str, top: int = 30) -> list[dict]:
    """Läs de senaste inläggen i en kanal (pseudonymiserade), nyast först."""
    return [_teams_message(m) for m in graph().channel_messages(team_id, channel_id, max(1, min(top, 50)))]


# ---- SharePoint releases --------------------------------------------------

def _sp_config() -> tuple[str, str, str]:
    site = os.environ.get("SHAREPOINT_SITE", "").strip()
    folder = os.environ.get("SHAREPOINT_FOLDER", "").strip().strip("/")
    if not site or not folder:
        raise RuntimeError(
            "SharePoint är inte konfigurerat. Sätt SHAREPOINT_SITE "
            "(t.ex. 'contoso.sharepoint.com:/sites/Utveckling') och SHAREPOINT_FOLDER."
        )
    return site, os.environ.get("SHAREPOINT_DRIVE", "").strip(), folder


def _safe_subfolder(subfolder: str | None) -> str:
    if not subfolder:
        return ""
    parts = [p for p in subfolder.replace("\\", "/").split("/") if p]
    if any(p in (".", "..") for p in parts):
        raise ValueError("Undermappen får inte innehålla '.' eller '..'.")
    return "/".join(parts)


def _check_local(local_path: str) -> Path:
    path = Path(local_path).expanduser().resolve()
    if not path.is_file():
        raise ValueError(f"Hittar ingen fil: {local_path}")
    roots = [Path(r).expanduser().resolve() for r in _load_list("SHAREPOINT_UPLOAD_ROOTS")]
    if roots and not any(path.is_relative_to(r) for r in roots):
        raise ValueError("Filen ligger utanför de mappar som får laddas upp (SHAREPOINT_UPLOAD_ROOTS).")
    return path


@tool("sharepoint")
def sharepoint_upload_release(local_path: str, subfolder: str | None = None, overwrite: bool = False) -> dict:
    """Ladda upp en lokal fil (t.ex. release-zip) till den förbestämda SharePoint-mappen.

    subfolder: valfri undermapp, t.ex. versionsnummer ("v1.4.0"). Skapas vid behov.
    overwrite: ersätt en befintlig fil med samma namn (annars avbryts uppladdningen).
    """
    site, drive_name, folder = _sp_config()
    path = _check_local(local_path)
    target = "/".join(p for p in (folder, _safe_subfolder(subfolder), path.name) if p)
    g = graph()
    item = g.upload(g.drive_id(site, drive_name), target, path.read_bytes(), overwrite)
    return {"name": item.get("name"), "path": target, "size": item.get("size"), "web_url": item.get("webUrl")}


@tool("sharepoint")
def sharepoint_list_releases(subfolder: str | None = None, top: int = 50) -> list[dict]:
    """Lista filer och mappar i den förbestämda SharePoint-mappen, senast ändrade först."""
    site, drive_name, folder = _sp_config()
    target = "/".join(p for p in (folder, _safe_subfolder(subfolder)) if p)
    g = graph()
    return [
        {"name": i["name"], "type": "folder" if "folder" in i else "file", "size": i.get("size"),
         "modified": i.get("lastModifiedDateTime"), "web_url": i.get("webUrl")}
        for i in g.list_folder(g.drive_id(site, drive_name), target, max(1, min(top, 200)))
    ]


def _require_draft(message_id: str) -> None:
    m = graph().request("GET", f"/me/messages/{message_id}", params={"$select": "isDraft"})
    if not m.get("isDraft"):
        raise ValueError("Meddelandet är inte ett utkast.")


def main() -> None:
    # httpx loggar hela URL:er på INFO, inklusive söksträngar med riktiga namn.
    logging.getLogger("httpx").setLevel(logging.WARNING)
    cmd = sys.argv[1] if len(sys.argv) > 1 else ""
    if cmd == "login":
        auth = Auth()
        result = auth.device_login() if "--device-code" in sys.argv else auth.interactive_login()
        user = (result.get("id_token_claims") or {}).get("preferred_username", "")
        print(f"Inloggad som {user}. Inloggningen sparas och förnyas automatiskt.", file=sys.stderr)
        return
    if cmd == "logout":
        removed = Auth().logout()
        print("Utloggad: " + (", ".join(removed) or "inga konton"), file=sys.stderr)
        return
    enabled = features.enabled()
    apply_features(enabled)
    print(f"Påslagna delar: {', '.join(enabled)}", file=sys.stderr, flush=True)
    if audit.enabled():
        AUDIT.active = True
        page = audit.AuditServer(AUDIT, int(os.environ.get("M365_MCP_AUDIT_PORT", "0")))
        page.start(open_browser=os.environ.get("M365_MCP_AUDIT_OPEN", "1") != "0")
    mcp.run("stdio")


if __name__ == "__main__":
    main()
