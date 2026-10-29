"""Thin Microsoft Graph client for the signed-in user's mailbox.

Uses delegated auth (device code flow) via MSAL. Scopes are deliberately
limited to Mail.ReadWrite: the server can read, organise and draft mail,
but it can never send anything.
"""

from __future__ import annotations

import os
import sys
from pathlib import Path
from typing import Any

import httpx
import msal

GRAPH = "https://graph.microsoft.com/v1.0"
SCOPES = ["Mail.ReadWrite", "User.Read"]


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
        self._http = httpx.Client(base_url=GRAPH, timeout=30, transport=transport)

    def request(self, method: str, path: str, **kw: Any) -> Any:
        headers = kw.pop("headers", {})
        headers["Authorization"] = f"Bearer {self._auth.token()}"
        # Ask Graph for plain-text bodies so HTML never reaches the redactor.
        headers.setdefault("Prefer", 'outlook.body-content-type="text"')
        resp = self._http.request(method, path, headers=headers, **kw)
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
        msg = self.request("GET", f"/me/messages/{message_id}", params={"$select": select})
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
