import os

import pytest

from m365_mcp_gdpr import graph


class FakeApp:
    CONSOLE_WINDOW_HANDLE = object()

    def __init__(self, client_id, **kw):
        self.kw = kw
        self.accounts = []
        self.interactive_calls = []

    def get_accounts(self, username=None):
        return [a for a in self.accounts if username in (None, a["username"])]

    def acquire_token_silent(self, scopes, account):
        return {"access_token": "silent-token"}

    def acquire_token_interactive(self, scopes, **kw):
        self.interactive_calls.append(kw)
        self.accounts.append({"username": "jimmy@example.com"})
        return {"access_token": "fresh-token", "id_token_claims": {"preferred_username": "jimmy@example.com"}}

    def remove_account(self, account):
        self.accounts.remove(account)


@pytest.fixture
def auth(monkeypatch, tmp_path):
    monkeypatch.setenv("M365_MCP_CLIENT_ID", "cid")
    monkeypatch.setenv("M365_MCP_CACHE_DIR", str(tmp_path))
    monkeypatch.setattr(graph.msal, "PublicClientApplication", FakeApp)
    return graph.Auth()


def test_first_call_opens_login_then_silent(auth):
    assert auth.token() == "fresh-token"
    assert auth.token() == "silent-token"
    assert len(auth._app.interactive_calls) == 1
    assert auth._app.interactive_calls[0]["prompt"] == "select_account"


def test_no_auto_login(monkeypatch, auth):
    auth._auto_login = False
    with pytest.raises(graph.AuthRequired):
        auth.token()


def test_logout(auth):
    auth.token()
    assert auth.logout() == ["jimmy@example.com"]
    assert auth._app.get_accounts() == []


def test_broker_only_when_installed(monkeypatch):
    monkeypatch.setattr(graph.sys, "platform", "linux")
    assert graph._broker_available() is False


def test_cache_falls_back_to_private_file(monkeypatch, tmp_path):
    monkeypatch.setenv("M365_MCP_CACHE_DIR", str(tmp_path))

    def no_secret_store(location):
        raise RuntimeError("libsecret saknas")

    monkeypatch.setattr(graph, "build_encrypted_persistence", no_secret_store)
    cache = graph._token_cache()
    assert cache is not None
    if os.name != "nt":  # Windows has no Unix file modes
        assert (tmp_path / "token_cache.bin").stat().st_mode & 0o777 == 0o600
