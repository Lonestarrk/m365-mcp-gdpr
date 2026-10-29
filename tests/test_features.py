import asyncio
import os
import sys

import pytest
from mcp.client.session import ClientSession
from mcp.client.stdio import StdioServerParameters, stdio_client

from m365_mcp_gdpr import features, graph

MAIL_TOOLS = {"list_folders", "read_message", "create_draft"}
SP_TOOLS = {"sharepoint_upload_release", "sharepoint_list_releases"}


def test_default_is_sharepoint_only(monkeypatch):
    monkeypatch.delenv("M365_MCP_FEATURES", raising=False)
    assert features.enabled() == ("sharepoint",)


def test_parse_and_all(monkeypatch):
    monkeypatch.setenv("M365_MCP_FEATURES", " Mail , sharepoint,mail ")
    assert features.enabled() == ("mail", "sharepoint")
    monkeypatch.setenv("M365_MCP_FEATURES", "all")
    assert features.enabled() == ("mail", "sharepoint")


def test_unknown_feature(monkeypatch):
    monkeypatch.setenv("M365_MCP_FEATURES", "mail,calender")
    with pytest.raises(ValueError, match="calender"):
        features.enabled()


def test_scopes_follow_features():
    assert features.scopes(("sharepoint",)) == ["Sites.ReadWrite.All", "User.Read"]
    assert "Mail.ReadWrite" not in features.scopes(("sharepoint",))
    assert features.scopes(("mail",)) == ["Mail.ReadWrite", "User.Read"]


def test_auth_requests_only_enabled_scopes(monkeypatch, tmp_path):
    seen = {}

    class FakeApp:
        def __init__(self, *a, **kw):
            self.accounts = []

        def get_accounts(self, username=None):
            return []

        def acquire_token_interactive(self, scopes, **kw):
            seen["scopes"] = scopes
            return {"access_token": "x"}

    monkeypatch.setenv("M365_MCP_CLIENT_ID", "cid")
    monkeypatch.setenv("M365_MCP_CACHE_DIR", str(tmp_path))
    monkeypatch.setenv("M365_MCP_FEATURES", "sharepoint")
    monkeypatch.setattr(graph.msal, "PublicClientApplication", FakeApp)
    graph.Auth().token()
    assert seen["scopes"] == ["Sites.ReadWrite.All", "User.Read"]


def _listed_tools(feature_env: str | None) -> set[str]:
    env = {**os.environ, "M365_MCP_CLIENT_ID": "x"}
    env.pop("M365_MCP_FEATURES", None)
    if feature_env is not None:
        env["M365_MCP_FEATURES"] = feature_env

    async def run() -> set[str]:
        params = StdioServerParameters(command=sys.executable, args=["-m", "m365_mcp_gdpr.server"], env=env)
        async with stdio_client(params) as (r, w):
            async with ClientSession(r, w) as s:
                await s.initialize()
                return {t.name for t in (await s.list_tools()).tools}

    return asyncio.run(run())


@pytest.mark.parametrize(
    "env, present, absent",
    [
        (None, SP_TOOLS, MAIL_TOOLS),
        ("sharepoint", SP_TOOLS, MAIL_TOOLS),
        ("mail", MAIL_TOOLS, SP_TOOLS),
        ("all", MAIL_TOOLS | SP_TOOLS, set()),
    ],
)
def test_server_exposes_only_enabled_tools(env, present, absent):
    tools = _listed_tools(env)
    assert present <= tools
    assert not (absent & tools)
