"""Which parts of Microsoft 365 the server exposes.

Set with M365_MCP_FEATURES, a comma-separated list (or "all"). Disabled
parts are removed from the MCP server *and* left out of the OAuth scopes,
so the access token itself cannot reach them.
"""

from __future__ import annotations

import os

# feature -> delegated Microsoft Graph scopes it needs
SCOPES: dict[str, list[str]] = {
    "mail": ["Mail.ReadWrite"],
    "calendar": ["Calendars.ReadWrite"],
    "teams": ["Chat.Read", "Team.ReadBasic.All", "Channel.ReadBasic.All", "ChannelMessage.Read.All"],
    "sharepoint": ["Sites.ReadWrite.All"],
}
BASE_SCOPES = ["User.Read"]
DEFAULT = ("sharepoint",)


def enabled() -> tuple[str, ...]:
    raw = os.environ.get("M365_MCP_FEATURES", "").strip().lower()
    if not raw:
        return DEFAULT
    if raw == "all":
        return tuple(SCOPES)
    names = tuple(dict.fromkeys(n.strip() for n in raw.split(",") if n.strip()))
    unknown = [n for n in names if n not in SCOPES]
    if unknown:
        raise ValueError(
            f"Okänd del i M365_MCP_FEATURES: {', '.join(unknown)}. Giltiga: {', '.join(SCOPES)}, all"
        )
    return names


def scopes(features: tuple[str, ...]) -> list[str]:
    out = [s for f in features for s in SCOPES[f]]
    return out + BASE_SCOPES


def request_scopes() -> list[str]:
    """Scopes to ask for at sign-in: enabled features, plus Copilot's if it redacts."""
    from . import copilot

    out = scopes(enabled())
    if copilot.enabled():
        out += copilot.SCOPES
    return list(dict.fromkeys(out))
