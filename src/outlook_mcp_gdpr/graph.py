"""Thin Microsoft Graph client for the signed-in user's mailbox and SharePoint.

Uses delegated auth (device code flow) via MSAL. There is deliberately no
Mail.Send scope: the server can read, organise and draft mail, but it can
never send anything. Sites.ReadWrite.All is used to upload release files.
"""

from __future__ import annotations

import os
import sys
from pathlib import Path
from typing import Any
from urllib.parse import quote

import httpx
import msal

GRAPH = "https://graph.microsoft.com/v1.0"
SCOPES = ["Mail.ReadWrite", "Sites.ReadWrite.All", "User.Read"]
# Graph requires upload-session chunks to be multiples of 320 KiB.
CHUNK = 320 * 1024 * 10
SIMPLE_UPLOAD_MAX = 4 * 1024 * 1024
# Ask Graph for plain-text bodies so HTML never reaches the redactor.
TEXT_BODY = {"Prefer": 'outlook.body-content-type="text"'}


class AuthRequired(RuntimeError):
    pass


def _cache_path() -> Path:
    base = Path(os.environ.get("OUTLOOK_MCP_CACHE_DIR", Path.home() / ".outlook-mcp-gdpr"))
    base.mkdir(parents=True, exist_ok=True, mode=0o700)
    return base / "token_cache.json"


class Auth:
    def __init__(self) -> None:
        client_id = os.environ.get("OUTLOOK_MCP_CLIENT_ID")
        if not client_id:
            raise RuntimeError("OUTLOOK_MCP_CLIENT_ID saknas (Entra ID app-registrering)")
        tenant = os.environ.get("OUTLOOK_MCP_TENANT_ID", "organizations")
        self._path = _cache_path()
        self._cache = msal.SerializableTokenCache()
        if self._path.exists():
            self._cache.deserialize(self._path.read_text())
        self._app = msal.PublicClientApplication(
            client_id,
            authority=f"https://login.microsoftonline.com/{tenant}",
            token_cache=self._cache,
        )

    def _save(self) -> None:
        if self._cache.has_state_changed:
            self._path.write_text(self._cache.serialize())
            os.chmod(self._path, 0o600)

    def token(self) -> str:
        accounts = self._app.get_accounts()
        result = self._app.acquire_token_silent(SCOPES, account=accounts[0]) if accounts else None
        self._save()
        if not result or "access_token" not in result:
            raise AuthRequired("Inte inloggad. Kör 'outlook-mcp-gdpr login' i en terminal först.")
        return result["access_token"]

    def device_login(self) -> None:
        flow = self._app.initiate_device_flow(scopes=SCOPES)
        if "user_code" not in flow:
            raise RuntimeError(f"Kunde inte starta inloggning: {flow}")
        print(flow["message"], file=sys.stderr, flush=True)
        result = self._app.acquire_token_by_device_flow(flow)
        if "access_token" not in result:
            raise RuntimeError(f"Inloggning misslyckades: {result.get('error_description')}")
        self._save()


class Graph:
    def __init__(self, auth: Auth | None = None, transport: httpx.BaseTransport | None = None):
        self._auth = auth or Auth()
        self._http = httpx.Client(base_url=GRAPH, timeout=120, transport=transport)
        self._drives: dict[tuple[str, str], str] = {}

    def request(self, method: str, path: str, **kw: Any) -> Any:
        headers = kw.pop("headers", {})
        headers["Authorization"] = f"Bearer {self._auth.token()}"
        resp = self._http.request(method, path, headers=headers, **kw)
        return self._handle(resp)

    @staticmethod
    def _handle(resp: httpx.Response) -> Any:
        if resp.status_code >= 400:
            try:
                err = resp.json()["error"]
                msg = f"{err.get('code')}: {err.get('message')}"
            except Exception:
                msg = resp.text[:300]
            raise RuntimeError(f"Graph {resp.status_code} – {msg}")
        return resp.json() if resp.content else None

    # ---- mail operations ----------------------------------------------

    def list_folders(self) -> list[dict]:
        data = self.request(
            "GET", "/me/mailFolders",
            params={"$top": 100, "$select": "id,displayName,unreadItemCount,totalItemCount"},
        )
        return data["value"]

    def list_messages(self, folder: str, top: int, skip: int, unread_only: bool) -> list[dict]:
        params: dict[str, Any] = {
            "$top": top,
            "$skip": skip,
            "$select": "id,subject,from,receivedDateTime,isRead,hasAttachments,bodyPreview,isDraft,flag",
            "$orderby": "receivedDateTime desc",
        }
        if unread_only:
            params["$filter"] = "isRead eq false"
        return self.request("GET", f"/me/mailFolders/{folder}/messages", params=params)["value"]

    def search(self, query: str, top: int) -> list[dict]:
        params = {
            "$search": f'"{query.replace(chr(34), "")}"',
            "$top": top,
            "$select": "id,subject,from,receivedDateTime,isRead,hasAttachments,bodyPreview",
        }
        return self.request("GET", "/me/messages", params=params)["value"]

    def get_message(self, message_id: str) -> dict:
        select = (
            "id,conversationId,subject,from,toRecipients,ccRecipients,receivedDateTime,"
            "isRead,isDraft,hasAttachments,body,importance,flag,categories"
        )
        msg = self.request(
            "GET", f"/me/messages/{message_id}", params={"$select": select}, headers=dict(TEXT_BODY)
        )
        if msg.get("hasAttachments"):
            att = self.request(
                "GET", f"/me/messages/{message_id}/attachments",
                params={"$select": "id,name,contentType,size"},
            )
            msg["attachments"] = att["value"]
        return msg

    def update_message(self, message_id: str, patch: dict) -> dict:
        return self.request("PATCH", f"/me/messages/{message_id}", json=patch)

    def move(self, message_id: str, destination: str) -> dict:
        return self.request("POST", f"/me/messages/{message_id}/move", json={"destinationId": destination})

    def create_draft(self, payload: dict) -> dict:
        return self.request("POST", "/me/messages", json=payload)

    def create_reply_draft(self, message_id: str, comment: str, reply_all: bool) -> dict:
        action = "createReplyAll" if reply_all else "createReply"
        return self.request("POST", f"/me/messages/{message_id}/{action}", json={"comment": comment})

    def delete(self, message_id: str) -> None:
        self.request("DELETE", f"/me/messages/{message_id}")

    # ---- SharePoint ---------------------------------------------------

    def drive_id(self, site: str, drive_name: str = "") -> str:
        """Resolve a site ("host:/sites/Name") and library name to a drive id."""
        key = (site, drive_name)
        if key not in self._drives:
            site_id = self.request("GET", f"/sites/{site}", params={"$select": "id"})["id"]
            if drive_name:
                drives = self.request("GET", f"/sites/{site_id}/drives", params={"$select": "id,name"})
                match = [d for d in drives["value"] if d["name"] == drive_name]
                if not match:
                    names = ", ".join(d["name"] for d in drives["value"])
                    raise RuntimeError(f"Dokumentbiblioteket '{drive_name}' finns inte. Tillgängliga: {names}")
                self._drives[key] = match[0]["id"]
            else:
                self._drives[key] = self.request("GET", f"/sites/{site_id}/drive", params={"$select": "id"})["id"]
        return self._drives[key]

    def list_folder(self, drive_id: str, folder: str, top: int) -> list[dict]:
        data = self.request(
            "GET", f"/drives/{drive_id}/root:/{quote(folder)}:/children",
            params={"$top": top, "$select": "name,size,lastModifiedDateTime,webUrl,folder,file",
                    "$orderby": "lastModifiedDateTime desc"},
        )
        return data["value"]

    def upload(self, drive_id: str, path: str, data: bytes, overwrite: bool) -> dict:
        conflict = "replace" if overwrite else "fail"
        item = f"/drives/{drive_id}/root:/{quote(path)}"
        if len(data) <= SIMPLE_UPLOAD_MAX:
            return self.request(
                "PUT", f"{item}:/content",
                params={"@microsoft.graph.conflictBehavior": conflict},
                content=data, headers={"Content-Type": "application/octet-stream"},
            )
        session = self.request(
            "POST", f"{item}:/createUploadSession",
            json={"item": {"@microsoft.graph.conflictBehavior": conflict}},
        )
        url = session["uploadUrl"]
        result: Any = None
        for start in range(0, len(data), CHUNK):
            chunk = data[start:start + CHUNK]
            end = start + len(chunk) - 1
            # The upload URL is pre-authenticated; Graph rejects an Authorization header here.
            resp = self._http.put(url, content=chunk, headers={
                "Content-Length": str(len(chunk)),
                "Content-Range": f"bytes {start}-{end}/{len(data)}",
            })
            result = self._handle(resp)
        return result
