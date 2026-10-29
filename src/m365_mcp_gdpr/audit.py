"""Optional local web page showing what the server hands to Claude.

Enabled with M365_MCP_AUDIT=1. Every tool call is recorded in memory (never
on disk) and shown side by side: the pseudonymised text Claude received, and
the original text with the replaced parts highlighted. The originals are
reconstructed from the vault, so they only exist in this process.

The page listens on 127.0.0.1 only, rejects other Host headers (DNS
rebinding) and requires a random key in the URL. The URL is opened in the
user's browser and printed to stderr; it is never returned to the model.
"""

from __future__ import annotations

import functools
import inspect
import itertools
import json
import os
import secrets
import sys
import threading
import webbrowser
from collections import deque
from datetime import datetime, timezone
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from importlib import resources
from typing import Any, Callable
from urllib.parse import parse_qs, urlparse

from .redact import TOKEN_RE, Redactor


def _to_text(value: Any) -> str:
    if isinstance(value, str):
        return value
    # Display only: show line breaks in mail bodies as real line breaks.
    return json.dumps(value, ensure_ascii=False, indent=2, default=str).replace("\\n", "\n")


def segments(text: str, redactor: Redactor) -> list[dict]:
    """Split text into plain parts and tokens with their original values."""
    out: list[dict] = []
    pos = 0
    for m in TOKEN_RE.finditer(text):
        if m.start() > pos:
            out.append({"text": text[pos:m.start()]})
        out.append({"token": m.group(0), "kind": m.group(1), "value": redactor.vault.lookup(m.group(0))})
        pos = m.end()
    if pos < len(text):
        out.append({"text": text[pos:]})
    return out


class AuditLog:
    def __init__(self, redactor: Redactor, maxlen: int = 200) -> None:
        self._redactor = redactor
        self._events: deque[dict] = deque(maxlen=maxlen)
        self._ids = itertools.count(1)
        self._lock = threading.Lock()
        self.active = False

    def record(self, tool: str, args: dict, result: Any = None, error: str | None = None) -> None:
        args_text = _to_text(args)
        result_text = _to_text(result) if error is None else error
        event = {
            "id": next(self._ids),
            "time": datetime.now(timezone.utc).isoformat(timespec="seconds"),
            "tool": tool,
            "args": segments(args_text, self._redactor),
            "result": segments(result_text, self._redactor),
            "error": error is not None,
        }
        with self._lock:
            self._events.append(event)

    def since(self, after: int) -> list[dict]:
        with self._lock:
            return [e for e in self._events if e["id"] > after]

    def clear(self) -> None:
        with self._lock:
            self._events.clear()

    def wrap(self, fn: Callable) -> Callable:
        """Record calls to a tool while the audit page is active."""
        sig = inspect.signature(fn)

        @functools.wraps(fn)
        def wrapper(*args: Any, **kwargs: Any) -> Any:
            if not self.active:
                return fn(*args, **kwargs)
            call_args = dict(sig.bind_partial(*args, **kwargs).arguments)
            try:
                result = fn(*args, **kwargs)
            except Exception as exc:
                self.record(fn.__name__, call_args, error=f"{type(exc).__name__}: {exc}")
                raise
            self.record(fn.__name__, call_args, result)
            return result

        return wrapper


class AuditServer:
    def __init__(self, log: AuditLog, port: int = 0) -> None:
        self.log = log
        self.key = secrets.token_urlsafe(24)
        self._httpd = ThreadingHTTPServer(("127.0.0.1", port), self._handler())
        self.port = self._httpd.server_address[1]

    @property
    def url(self) -> str:
        return f"http://127.0.0.1:{self.port}/?k={self.key}"

    def start(self, open_browser: bool = True) -> None:
        threading.Thread(target=self._httpd.serve_forever, daemon=True, name="audit-http").start()
        print(f"Granskningssidan: {self.url}", file=sys.stderr, flush=True)
        if open_browser:
            webbrowser.open(self.url)

    def stop(self) -> None:
        self._httpd.shutdown()
        self._httpd.server_close()

    def _handler(self) -> type[BaseHTTPRequestHandler]:
        server = self
        page = resources.files(__package__).joinpath("audit.html").read_text(encoding="utf-8")

        class Handler(BaseHTTPRequestHandler):
            def log_message(self, *args: Any) -> None:  # keep stderr quiet
                pass

            def _authorised(self) -> dict[str, list[str]] | None:
                hosts = {f"127.0.0.1:{server.port}", f"localhost:{server.port}"}
                query = parse_qs(urlparse(self.path).query)
                if self.headers.get("Host") not in hosts:
                    self._send(421, "text/plain", b"Fel host")
                    return None
                if not secrets.compare_digest(query.get("k", [""])[0], server.key):
                    self._send(403, "text/plain", b"Fel eller saknad nyckel")
                    return None
                return query

            def _send(self, status: int, ctype: str, body: bytes) -> None:
                self.send_response(status)
                self.send_header("Content-Type", f"{ctype}; charset=utf-8")
                self.send_header("Content-Length", str(len(body)))
                self.send_header("Cache-Control", "no-store")
                self.send_header("Referrer-Policy", "no-referrer")
                self.send_header("X-Frame-Options", "DENY")
                self.send_header("X-Content-Type-Options", "nosniff")
                self.send_header(
                    "Content-Security-Policy",
                    "default-src 'none'; script-src 'unsafe-inline'; style-src 'unsafe-inline'; "
                    "connect-src 'self'; base-uri 'none'; form-action 'none'",
                )
                self.end_headers()
                self.wfile.write(body)

            def do_GET(self) -> None:
                query = self._authorised()
                if query is None:
                    return
                route = urlparse(self.path).path
                if route == "/":
                    self._send(200, "text/html", page.encode())
                elif route == "/events":
                    after = int(query.get("after", ["0"])[0] or 0)
                    body = json.dumps(server.log.since(after), ensure_ascii=False).encode()
                    self._send(200, "application/json", body)
                else:
                    self._send(404, "text/plain", b"Finns inte")

            def do_POST(self) -> None:
                if self._authorised() is None:
                    return
                if urlparse(self.path).path == "/clear":
                    server.log.clear()
                    self._send(204, "text/plain", b"")
                else:
                    self._send(404, "text/plain", b"Finns inte")

        return Handler


def enabled() -> bool:
    return os.environ.get("M365_MCP_AUDIT") == "1"
