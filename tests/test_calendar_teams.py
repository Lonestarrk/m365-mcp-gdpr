import json

import httpx
import pytest

from m365_mcp_gdpr import server
from m365_mcp_gdpr.graph import Graph
from m365_mcp_gdpr.text import html_to_text

ANNA = {"emailAddress": {"name": "Anna Svensson", "address": "anna.svensson@kund.se"}}
ME = {"emailAddress": {"name": "Jimmy Revtorp", "address": "jimmy@example.com"}}

MEETING = {
    "id": "ev1", "subject": "Demo för Anna Svensson", "isOrganizer": True,
    "start": {"dateTime": "2026-11-02T09:00:00", "timeZone": "UTC"},
    "end": {"dateTime": "2026-11-02T10:00:00", "timeZone": "UTC"},
    "location": {"displayName": "Storgatan 12B, 114 51 Stockholm"},
    "organizer": ME,
    "attendees": [{**ANNA, "status": {"response": "accepted"}}],
    "bodyPreview": "Ring 070-123 45 67 om du blir sen",
    "body": {"content": "Ring 070-123 45 67 om du blir sen"},
}
FOCUS = {**MEETING, "id": "ev2", "subject": "Fokustid", "attendees": [], "location": {}}

CHAT = {
    "id": "chat1", "chatType": "oneOnOne", "topic": None, "lastUpdatedDateTime": "2026-10-29T10:00:00Z",
    "members": [{"displayName": "Anna Svensson", "email": "anna.svensson@kund.se"}],
}
CHAT_MSG = {
    "id": "m1", "createdDateTime": "2026-10-29T10:00:00Z", "messageType": "message",
    "from": {"user": {"displayName": "Anna Svensson"}},
    "body": {"contentType": "html", "content": "<p>Hej!</p><p>Ring mig p&aring; 070-123 45 67</p><script>x()</script>"},
}


class FakeAuth:
    def token(self) -> str:
        return "t"


@pytest.fixture
def calls(monkeypatch):
    monkeypatch.delenv("M365_MCP_TIMEZONE", raising=False)
    seen = []

    def handler(req: httpx.Request) -> httpx.Response:
        seen.append(req)
        path = req.url.path
        if path == "/v1.0/me/calendarView":
            return httpx.Response(200, json={"value": [MEETING]})
        if path == "/v1.0/me/events/ev1":
            return httpx.Response(200, json=MEETING)
        if path == "/v1.0/me/events/ev2":
            return httpx.Response(200 if req.method == "GET" else 204, json=FOCUS if req.method == "GET" else None)
        if path == "/v1.0/me/events" and req.method == "POST":
            return httpx.Response(201, json={"id": "new1"})
        if path == "/v1.0/me/chats":
            return httpx.Response(200, json={"value": [CHAT]})
        if path == "/v1.0/me/chats/chat1/messages":
            return httpx.Response(200, json={"value": [CHAT_MSG]})
        if path == "/v1.0/me/joinedTeams":
            return httpx.Response(200, json={"value": [{"id": "t1", "displayName": "Utveckling"}]})
        if path == "/v1.0/teams/t1/channels":
            return httpx.Response(200, json={"value": [{"id": "c1", "displayName": "Releaser"}]})
        if path == "/v1.0/teams/t1/channels/c1/messages":
            return httpx.Response(200, json={"value": [CHAT_MSG]})
        return httpx.Response(404, json={"error": {"code": "NotFound", "message": path}})

    server.R.reset()
    server._graph = Graph(auth=FakeAuth(), transport=httpx.MockTransport(handler))
    yield seen
    server._graph = None


def _leaks(obj) -> list[str]:
    text = json.dumps(obj, ensure_ascii=False)
    return [s for s in ["Anna", "Svensson", "anna.svensson", "070-123", "Storgatan"] if s in text]


def test_list_events_redacted(calls):
    events = server.calendar_list_events("2026-11-02T00:00:00", "2026-11-03T00:00:00")
    assert _leaks(events) == []
    assert events[0]["attendees"][0]["response"] == "accepted"
    assert events[0]["start"]["dateTime"] == "2026-11-02T09:00:00"
    assert "Prefer" not in calls[-1].headers


def test_get_event_redacted_with_timezone(calls, monkeypatch):
    monkeypatch.setenv("M365_MCP_TIMEZONE", "Europe/Stockholm")
    event = server.calendar_get_event("ev1")
    assert _leaks(event) == [] and "[[PHONE_1]]" in event["body"]
    prefer = calls[-1].headers["Prefer"]
    assert 'outlook.timezone="Europe/Stockholm"' in prefer and "body-content-type" in prefer


def test_create_event_never_invites(calls):
    events = server.calendar_list_events("2026-11-02T00:00:00", "2026-11-03T00:00:00")
    anna = events[0]["attendees"][0]["who"].split(" <")[0]  # the token Claude saw
    server.calendar_create_event(
        f"Förbered demo för {anna}", "2026-11-02T08:00:00", "2026-11-02T09:00:00",
        time_zone="Europe/Stockholm",
    )
    payload = json.loads(calls[-1].content)
    assert payload["attendees"] == []
    assert payload["subject"] == "Förbered demo för Anna Svensson"
    assert payload["start"] == {"dateTime": "2026-11-02T08:00:00", "timeZone": "Europe/Stockholm"}


def test_create_event_has_no_attendee_parameter():
    import inspect
    assert "attendees" not in inspect.signature(server.calendar_create_event).parameters


def test_delete_refuses_meetings_with_attendees(calls):
    with pytest.raises(ValueError, match="deltagare"):
        server.calendar_delete_event("ev1")
    assert not any(c.method == "DELETE" for c in calls)


def test_delete_own_focus_time(calls):
    assert server.calendar_delete_event("ev2") == "ok"
    assert calls[-1].method == "DELETE"


def test_teams_chats_and_messages_redacted(calls):
    chats = server.teams_list_chats()
    msgs = server.teams_read_chat("chat1")
    assert _leaks(chats) == [] and _leaks(msgs) == []
    assert chats[0]["members"] == ["[[PERSON_1]]"]
    assert msgs[0]["from"] == "[[PERSON_1]]"
    assert msgs[0]["text"] == "Hej!\n\nRing mig på [[PHONE_1]]"


def test_teams_channels(calls):
    assert server.teams_list_teams() == [{"id": "t1", "name": "Utveckling", "description": ""}]
    assert server.teams_list_channels("t1")[0]["name"] == "Releaser"
    assert _leaks(server.teams_read_channel("t1", "c1")) == []


def test_teams_is_read_only():
    teams_tools = {n for n, f in server.TOOL_FEATURE.items() if f == "teams"}
    assert teams_tools == {"teams_list_chats", "teams_read_chat", "teams_list_teams",
                           "teams_list_channels", "teams_read_channel"}


def test_html_to_text():
    html = "<div>Rad 1<br>Rad 2</div><style>x{}</style>&nbsp;<attachment id='1'></attachment>"
    assert html_to_text(html) == "Rad 1\nRad 2\n[bilaga]"
