import httpx
import pytest

from m365_mcp_gdpr import server
from m365_mcp_gdpr.graph import Graph

INSIDE = {"id": "in1", "parentFolderId": "F-TEST", "subject": "Test", "from": None,
          "toRecipients": [], "ccRecipients": [], "body": {"content": "Hej"}}
OUTSIDE = {**INSIDE, "id": "out1", "parentFolderId": "F-INBOX", "subject": "Kundmejl"}


class FakeAuth:
    def token(self) -> str:
        return "t"


@pytest.fixture
def calls(monkeypatch):
    monkeypatch.setenv("M365_MCP_MAIL_FOLDER", "MCP-test")
    monkeypatch.setattr(server, "_folder_id", None)
    seen = []

    def handler(req: httpx.Request) -> httpx.Response:
        seen.append(req)
        path = req.url.path
        if path in ("/v1.0/me/mailFolders", "/v1.0/me/mailFolders/inbox/childFolders"):
            hit = path.endswith("childFolders") and "MCP-test" in req.url.params.get("$filter", "")
            return httpx.Response(200, json={"value": [{"id": "F-TEST", "displayName": "MCP-test"}] if hit else []})
        if path == "/v1.0/me/mailFolders/F-TEST":
            return httpx.Response(200, json={"id": "F-TEST", "displayName": "MCP-test", "totalItemCount": 1})
        if path == "/v1.0/me/mailFolders/F-TEST/messages":
            return httpx.Response(200, json={"value": [INSIDE]})
        if path.startswith("/v1.0/me/messages/"):
            msg = INSIDE if path.endswith("/in1") else OUTSIDE
            if req.method == "PATCH":
                return httpx.Response(200, json=msg)
            return httpx.Response(200, json=msg)
        return httpx.Response(404, json={"error": {"code": "NotFound", "message": path}})

    server.R.reset()
    server._graph = Graph(auth=FakeAuth(), transport=httpx.MockTransport(handler))
    yield seen
    server._graph = None


def test_finds_folder_under_inbox_and_lists_only_it(calls):
    assert [f["id"] for f in server.list_folders()] == ["F-TEST"]
    assert [m["id"] for m in server.list_messages()] == ["in1"]


def test_other_folders_refused(calls):
    with pytest.raises(ValueError, match="begränsad"):
        server.list_messages(folder="inbox")
    assert [m["id"] for m in server.list_messages(folder="MCP-test")] == ["in1"]


def test_search_scoped_to_folder(calls):
    server.search_messages("faktura")
    assert calls[-1].url.path == "/v1.0/me/mailFolders/F-TEST/messages"


def test_read_inside_ok_outside_refused(calls):
    assert server.read_message("in1")["subject"] == "Test"
    with pytest.raises(ValueError, match="utanför"):
        server.read_message("out1")


@pytest.mark.parametrize("call", [
    lambda: server.mark_read("out1"),
    lambda: server.flag_message("out1"),
    lambda: server.set_categories("out1", ["x"]),
    lambda: server.create_reply_draft("out1", "hej"),
    lambda: server.move_message("in1", "archive"),
])
def test_actions_outside_refused(calls, call):
    with pytest.raises(ValueError):
        call()
    assert not any(c.method in ("PATCH", "POST") for c in calls)


def test_missing_folder_gives_clear_error(calls, monkeypatch):
    monkeypatch.setenv("M365_MCP_MAIL_FOLDER", "Finns-inte")
    with pytest.raises(ValueError, match="Skapa den i Outlook"):
        server.list_messages()


def test_no_restriction_by_default(monkeypatch):
    monkeypatch.delenv("M365_MCP_MAIL_FOLDER", raising=False)
    monkeypatch.setattr(server, "_folder_id", None)
    assert server._mail_folder() is None
