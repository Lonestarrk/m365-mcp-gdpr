import json

import httpx
import pytest

from m365_mcp_gdpr import server
from m365_mcp_gdpr.graph import Graph

MESSAGE = {
    "id": "AAMk1",
    "conversationId": "c1",
    "subject": "Offert till Anna Svensson",
    "from": {"emailAddress": {"name": "Anna Svensson", "address": "anna.svensson@kund.se"}},
    "toRecipients": [{"emailAddress": {"name": "Jimmy", "address": "jimmy@example.com"}}],
    "ccRecipients": [],
    "receivedDateTime": "2026-10-28T09:00:00Z",
    "isRead": False,
    "isDraft": False,
    "hasAttachments": False,
    "body": {"contentType": "text", "content": "Hej! Ring mig på 070-123 45 67.\n/Anna"},
}


class FakeAuth:
    def token(self) -> str:
        return "t"


@pytest.fixture
def sent():
    calls = []

    def handler(req: httpx.Request) -> httpx.Response:
        calls.append(req)
        if req.method == "GET" and req.url.path.endswith("/messages/AAMk1"):
            return httpx.Response(200, json=MESSAGE)
        if req.method == "POST" and req.url.path.endswith("/me/messages"):
            return httpx.Response(201, json={"id": "draft1"})
        return httpx.Response(404, json={"error": {"code": "NotFound", "message": "x"}})

    server.R.vault.clear()
    server._graph = Graph(auth=FakeAuth(), transport=httpx.MockTransport(handler))
    yield calls
    server._graph = None


async def _call(name, args):
    result = await server.mcp.call_tool(name, args)
    return "".join(getattr(c, "text", "") for c in result.content)


@pytest.mark.anyio
async def test_read_then_draft_roundtrip(sent):
    out = await _call("read_message", {"message_id": "AAMk1"})
    for secret in ["Anna", "Svensson", "anna.svensson@kund.se", "070-123"]:
        assert secret not in out, secret
    data = json.loads(out)
    assert data["from"] == "[[PERSON_1]] <[[EMAIL_1]]>"

    await _call("create_draft", {
        "subject": "Re: offert",
        "body": "Hej [[PERSON_1_FIRST]]! Jag ringer [[PHONE_1]].",
        "to": ["[[EMAIL_1]]"],
    })
    payload = json.loads(sent[-1].content)
    assert payload["body"]["content"] == "Hej Anna! Jag ringer 070-123 45 67."
    assert payload["toRecipients"][0]["emailAddress"]["address"] == "anna.svensson@kund.se"


@pytest.fixture
def anyio_backend():
    return "asyncio"
