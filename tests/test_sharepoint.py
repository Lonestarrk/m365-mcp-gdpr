import json

import httpx
import pytest

from outlook_mcp_gdpr import graph as graph_mod
from outlook_mcp_gdpr import server
from outlook_mcp_gdpr.graph import Graph


class FakeAuth:
    def token(self) -> str:
        return "t"


@pytest.fixture
def sp(monkeypatch, tmp_path):
    monkeypatch.setenv("SHAREPOINT_SITE", "contoso.sharepoint.com:/sites/Dev")
    monkeypatch.setenv("SHAREPOINT_FOLDER", "Releases/MinApp")
    monkeypatch.setenv("SHAREPOINT_UPLOAD_ROOTS", str(tmp_path))
    calls = []

    def handler(req: httpx.Request) -> httpx.Response:
        calls.append(req)
        path = req.url.path
        if path.startswith("/v1.0/sites/") and path.endswith("/drive"):
            return httpx.Response(200, json={"id": "drive1"})
        if path.startswith("/v1.0/sites/"):
            return httpx.Response(200, json={"id": "site1"})
        if path.endswith(":/content"):
            return httpx.Response(201, json={"name": "app.zip", "size": len(req.content), "webUrl": "https://x/app.zip"})
        if path.endswith(":/createUploadSession"):
            return httpx.Response(200, json={"uploadUrl": "https://upload.example/session"})
        if req.url.host == "upload.example":
            start, rest = req.headers["Content-Range"].split(" ")[1].split("-")
            end, total = rest.split("/")
            if int(end) + 1 < int(total):
                return httpx.Response(202, json={"nextExpectedRanges": [f"{int(end) + 1}-"]})
            return httpx.Response(201, json={"name": "big.zip", "size": int(total), "webUrl": "https://x/big.zip"})
        if path.endswith(":/children"):
            return httpx.Response(200, json={"value": [
                {"name": "v1.0", "folder": {}, "lastModifiedDateTime": "2026-10-01T00:00:00Z", "webUrl": "https://x/v1.0"},
            ]})
        return httpx.Response(404, json={"error": {"code": "NotFound", "message": path}})

    server._graph = Graph(auth=FakeAuth(), transport=httpx.MockTransport(handler))
    yield calls, tmp_path
    server._graph = None


def test_small_upload_goes_to_configured_folder(sp):
    calls, tmp = sp
    f = tmp / "app.zip"
    f.write_bytes(b"zipdata")
    out = server.sharepoint_upload_release(str(f), subfolder="v1.4.0")
    put = calls[-1]
    assert put.method == "PUT"
    assert "/drives/drive1/root:/Releases/MinApp/v1.4.0/app.zip:/content" in put.url.path
    assert put.url.params["@microsoft.graph.conflictBehavior"] == "fail"
    assert put.content == b"zipdata"
    assert out["web_url"] == "https://x/app.zip"


def test_large_upload_uses_chunked_session(sp, monkeypatch):
    calls, tmp = sp
    monkeypatch.setattr(graph_mod, "SIMPLE_UPLOAD_MAX", 10)
    monkeypatch.setattr(graph_mod, "CHUNK", 8)
    f = tmp / "big.zip"
    f.write_bytes(b"x" * 20)
    out = server.sharepoint_upload_release(str(f), overwrite=True)
    session = [c for c in calls if c.url.path.endswith("createUploadSession")][0]
    assert json.loads(session.content)["item"]["@microsoft.graph.conflictBehavior"] == "replace"
    chunks = [c for c in calls if c.url.host == "upload.example"]
    assert [c.headers["Content-Range"] for c in chunks] == ["bytes 0-7/20", "bytes 8-15/20", "bytes 16-19/20"]
    assert all("Authorization" not in c.headers for c in chunks)
    assert out["size"] == 20


def test_cannot_escape_folder(sp):
    _, tmp = sp
    f = tmp / "app.zip"
    f.write_bytes(b"x")
    with pytest.raises(ValueError):
        server.sharepoint_upload_release(str(f), subfolder="../../Ekonomi")


def test_upload_roots_enforced(sp, tmp_path_factory):
    other = tmp_path_factory.mktemp("other") / "secret.txt"
    other.write_text("hemligt")
    with pytest.raises(ValueError, match="utanför"):
        server.sharepoint_upload_release(str(other))


def test_list_releases(sp):
    out = server.sharepoint_list_releases()
    assert out == [{"name": "v1.0", "type": "folder", "size": None,
                    "modified": "2026-10-01T00:00:00Z", "web_url": "https://x/v1.0"}]


def test_requires_config(monkeypatch):
    monkeypatch.delenv("SHAREPOINT_SITE", raising=False)
    with pytest.raises(RuntimeError, match="SHAREPOINT_SITE"):
        server.sharepoint_list_releases()
