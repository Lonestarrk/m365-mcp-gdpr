import json

import httpx
import pytest

from m365_mcp_gdpr import copilot, features, server
from m365_mcp_gdpr.graph import Graph

BODY = (
    "Hej! Elias Berg börjar på avdelning Blå. Kökschefen på Solgårdens förskola "
    "vill ha en uppdaterad meny. /Anna"
)
ALLERGY = "Elias har nötallergi, kan ni ta bort nötter i hans mat?"

MESSAGES = {
    "m1": {"id": "m1", "subject": "Ny elev", "from": {"emailAddress": {"name": "Anna Svensson", "address": "anna@kund.se"}},
           "toRecipients": [], "ccRecipients": [], "body": {"content": BODY}, "bodyPreview": BODY[:40]},
    "m2": {"id": "m2", "subject": "Allergi", "from": {"emailAddress": {"name": "Anna Svensson", "address": "anna@kund.se"}},
           "toRecipients": [], "ccRecipients": [], "body": {"content": ALLERGY}, "bodyPreview": ALLERGY},
}


def verdicts_for(prompt: str) -> dict:
    """A stand-in for Copilot that follows the instructions for our two sample mails."""
    results = []
    for i, chunk in enumerate(prompt.split("<<<TEXT ")[1:], 1):
        if "nötallergi" in chunk:
            results.append({"index": i, "block": True, "reason": "allergi", "items": []})
        else:
            items = [{"text": "Elias Berg", "type": "PERSON"}] if "Elias Berg" in chunk else []
            if "Kökschefen på Solgårdens förskola" in chunk:
                items.append({"text": "Kökschefen på Solgårdens förskola", "type": "INDIRECT"})
            results.append({"index": i, "block": False, "reason": "", "items": items})
    return {"results": results}


@pytest.fixture
def env(monkeypatch):
    monkeypatch.setenv("M365_MCP_REDACT_ENGINE", "copilot")
    monkeypatch.delenv("M365_MCP_MAIL_FOLDER", raising=False)
    monkeypatch.setattr(server, "_folder_id", None)
    state = {"calls": [], "mode": "ok"}

    def handler(req: httpx.Request) -> httpx.Response:
        state["calls"].append(req)
        url = str(req.url)
        if url.startswith(copilot.COPILOT):
            if state["mode"] == "error":
                return httpx.Response(500, json={"error": {"code": "x", "message": "nere"}})
            if url == copilot.COPILOT:
                return httpx.Response(201, json={"id": "conv1"})
            prompt = json.loads(req.content)["message"]["text"]
            answer = "Här är svaret: " + json.dumps(verdicts_for(prompt)) if state["mode"] == "ok" else "Jag vet inte."
            return httpx.Response(200, json={"messages": [{"text": prompt}, {"text": answer}]})
        path = req.url.path
        if path == "/v1.0/me/mailFolders/inbox/messages":
            return httpx.Response(200, json={"value": list(MESSAGES.values())})
        if path.startswith("/v1.0/me/messages/"):
            return httpx.Response(200, json=MESSAGES[path.rsplit("/", 1)[1]])
        return httpx.Response(404, json={"error": {"code": "NotFound", "message": path}})

    class FakeAuth:
        def token(self):
            return "t"

    server.R.reset()
    server._graph = Graph(auth=FakeAuth(), transport=httpx.MockTransport(handler))
    yield state
    server._graph = None


def test_names_in_text_and_indirect_identifiers_redacted(env):
    out = json.dumps(server.read_message("m1"), ensure_ascii=False)
    for secret in ["Elias", "Berg", "Kökschefen", "Solgårdens", "Anna"]:
        assert secret not in out, secret
    assert "[[INDIRECT_1]]" in out


def test_special_category_data_blocks_message(env):
    out = server.read_message("m2")
    assert out["blocked"] is True and "allergi" in out["reason"]
    assert "body" not in out and "nötallergi" not in json.dumps(out, ensure_ascii=False)


def test_request_disables_web_and_sends_instructions(env):
    server.read_message("m1")
    chat = json.loads(env["calls"][-1].content)
    assert chat["contextualResources"]["webContext"]["isWebEnabled"] is False
    assert "Svara ENDAST med JSON" in chat["message"]["text"]
    assert "<<<TEXT 1" in chat["message"]["text"]


def test_list_is_one_copilot_call_and_blocks_per_message(env):
    out = server.list_messages()
    chats = [c for c in env["calls"] if str(c.url).endswith("/chat")]
    assert len(chats) == 1
    by_id = {m["id"]: m for m in out}
    assert by_id["m2"]["subject"].startswith("Blockerat") and by_id["m2"]["preview"] == ""
    assert "Elias" not in json.dumps(by_id["m1"], ensure_ascii=False)


@pytest.mark.parametrize("mode", ["error", "garbage"])
def test_fails_closed(env, mode):
    env["mode"] = mode
    out = server.read_message("m1")
    assert out["blocked"] is True
    assert "Elias" not in json.dumps(out, ensure_ascii=False)


def test_missing_index_fails_closed():
    with pytest.raises(KeyError):
        copilot._parse('{"results": [{"index": 1, "block": false, "items": []}]}', 2)


def test_local_engine_by_default(monkeypatch):
    monkeypatch.delenv("M365_MCP_REDACT_ENGINE", raising=False)
    assert copilot.enabled() is False
    assert server._screen(["text"]) == [None]


def test_unknown_engine(monkeypatch):
    monkeypatch.setenv("M365_MCP_REDACT_ENGINE", "chatgpt")
    with pytest.raises(ValueError, match="chatgpt"):
        copilot.enabled()


def test_copilot_scopes_only_when_enabled(monkeypatch):
    monkeypatch.setenv("M365_MCP_FEATURES", "mail")
    monkeypatch.delenv("M365_MCP_REDACT_ENGINE", raising=False)
    assert "ChannelMessage.Read.All" not in features.request_scopes()
    monkeypatch.setenv("M365_MCP_REDACT_ENGINE", "copilot")
    scopes = features.request_scopes()
    assert set(copilot.SCOPES) <= set(scopes) and "Mail.ReadWrite" in scopes
    assert len(scopes) == len(set(scopes))


def test_custom_instructions(monkeypatch, tmp_path):
    f = tmp_path / "egna.md"
    f.write_text("Egna regler", encoding="utf-8")
    monkeypatch.setenv("M365_MCP_COPILOT_INSTRUCTIONS", str(f))
    assert copilot._prompt(["x"]).startswith("Egna regler")
