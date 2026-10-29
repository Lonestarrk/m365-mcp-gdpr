import http.client
import json

import httpx
import pytest

from m365_mcp_gdpr import audit, server
from m365_mcp_gdpr.graph import Graph
from m365_mcp_gdpr.redact import Redactor

MESSAGE = {
    "id": "AAMk1",
    "subject": "Hej",
    "from": {"emailAddress": {"name": "Anna Svensson", "address": "anna.svensson@kund.se"}},
    "toRecipients": [],
    "ccRecipients": [],
    "body": {"content": "Ring 070-123 45 67 <script>alert(1)</script>"},
}


class FakeAuth:
    def token(self) -> str:
        return "t"


@pytest.fixture
def mail():
    server.R.reset()
    server.AUDIT.clear()
    server._graph = Graph(
        auth=FakeAuth(), transport=httpx.MockTransport(lambda r: httpx.Response(200, json=MESSAGE))
    )
    yield
    server._graph = None
    server.AUDIT.active = False
    server.AUDIT.clear()


def test_nothing_recorded_when_disabled(mail):
    server.read_message("AAMk1")
    assert server.AUDIT.since(0) == []


def test_records_redacted_and_original(mail):
    server.AUDIT.active = True
    out = server.read_message(message_id="AAMk1")
    assert "070-123" not in json.dumps(out, ensure_ascii=False)

    (event,) = server.AUDIT.since(0)
    assert event["tool"] == "read_message"
    assert event["args"] == [{"text": '{\n  "message_id": "AAMk1"\n}'}]
    tokens = {s["token"]: s["value"] for s in event["result"] if "token" in s}
    assert tokens["[[PHONE_1]]"] == "070-123 45 67"
    assert tokens["[[EMAIL_1]]"] == "anna.svensson@kund.se"
    # The text Claude saw contains only tokens.
    claude_saw = "".join(s.get("text", s.get("token", "")) for s in event["result"])
    assert "Anna" not in claude_saw and "070-123" not in claude_saw


def test_errors_are_recorded(mail):
    server.AUDIT.active = True
    with pytest.raises(RuntimeError, match="SHAREPOINT_SITE"):
        server.sharepoint_list_releases()
    (event,) = server.AUDIT.since(0)
    assert event["error"] is True


@pytest.fixture
def page():
    log = audit.AuditLog(Redactor())
    log.record("list_folders", {}, [{"name": "Inkorg"}])
    srv = audit.AuditServer(log)
    srv.start(open_browser=False)
    yield srv
    srv.stop()


def _get(srv, path, host=None):
    conn = http.client.HTTPConnection("127.0.0.1", srv.port, timeout=5)
    conn.request("GET", path, headers={"Host": host or f"127.0.0.1:{srv.port}"})
    resp = conn.getresponse()
    return resp.status, resp.read(), resp


def test_page_requires_key(page):
    assert _get(page, "/")[0] == 403
    assert _get(page, "/events?k=fel")[0] == 403
    status, body, resp = _get(page, f"/?k={page.key}")
    assert status == 200 and b"Granskning" in body
    assert "default-src 'none'" in resp.getheader("Content-Security-Policy")


def test_page_rejects_foreign_host(page):
    assert _get(page, f"/?k={page.key}", host="evil.example")[0] == 421


def test_events_endpoint(page):
    status, body, _ = _get(page, f"/events?k={page.key}&after=0")
    events = json.loads(body)
    assert status == 200 and events[0]["tool"] == "list_folders"
    assert json.loads(_get(page, f"/events?k={page.key}&after={events[0]['id']}")[1]) == []


def test_listens_on_localhost_only(page):
    assert page._httpd.server_address[0] == "127.0.0.1"


def test_page_never_uses_innerhtml():
    html = (audit.resources.files("m365_mcp_gdpr") / "audit.html").read_text(encoding="utf-8")
    assert "innerHTML" not in html and "insertAdjacentHTML" not in html
