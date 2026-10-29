"""Reversible pseudonymisation of personal data (GDPR art. 4.5).

Everything that leaves this server towards the LLM passes through
`Redactor.redact`. Personal data is swapped for stable tokens such as
``[[PERSON_1]]`` or ``[[EMAIL_2]]``. Personal identity numbers are instead
swapped for made-up numbers from the 1800s in the same national format
(see `idnumbers`). The mapping token -> original value
lives only in this process's memory (the `Vault`) and is never returned to
the model. When the model writes a draft, `Redactor.restore` puts the real
values back before the draft is sent to Microsoft Graph.
"""

from __future__ import annotations

import re
import threading
from dataclasses import dataclass, field
from typing import Callable, Iterable

from . import idnumbers

TOKEN_RE = re.compile(r"\[\[([A-Z]+)_(\d+)(?:_(FIRST|LAST))?\]\]")


class UnknownTokenError(ValueError):
    """Raised when text contains a token the vault does not know."""


def _luhn_ok(digits: str) -> bool:
    total = 0
    for i, ch in enumerate(reversed(digits)):
        n = int(ch)
        if i % 2 == 1:
            n *= 2
            if n > 9:
                n -= 9
        total += n
    return total % 10 == 0


def _valid_card(match: str) -> bool:
    digits = re.sub(r"\D", "", match)
    return 13 <= len(digits) <= 19 and _luhn_ok(digits)


@dataclass(frozen=True)
class Detector:
    kind: str
    pattern: re.Pattern[str]
    validate: Callable[[str], bool] | None = None


_STREET_SUFFIX = (
    r"(?:gatan|gata|vägen|väg|gränd|torget|torg|plan|allén|backen|stigen|"
    r"leden|platsen|stråket|kajen|esplanaden)"
)

DETECTORS: list[Detector] = [
    Detector("EMAIL", re.compile(r"[\w.+-]+@[\w-]+(?:\.[\w-]+)+", re.UNICODE)),
    Detector(
        "IBAN",
        re.compile(r"\b[A-Z]{2}\d{2}(?: ?[A-Z0-9]{4}){2,7}(?: ?[A-Z0-9]{1,4})?\b"),
    ),
    Detector(
        "CARD",
        re.compile(r"(?<!\d)(?:\d[ -]?){12,18}\d(?!\d)"),
        _valid_card,
    ),
    Detector(
        "PHONE",
        re.compile(
            r"(?<![\w+])(?:(?:\+|00)\d{1,3}[\s-]?(?:\(0\)[\s-]?)?\d{1,4}|0\d{1,3})"
            r"(?:[\s-]?\d{2,4}){2,4}(?!\d)"
        ),
    ),
    Detector(
        "IP",
        re.compile(r"\b(?:(?:25[0-5]|2[0-4]\d|1?\d?\d)\.){3}(?:25[0-5]|2[0-4]\d|1?\d?\d)\b"),
    ),
    Detector(
        "ADDRESS",
        re.compile(
            r"\b[A-ZÅÄÖ][a-zåäöé]+" + _STREET_SUFFIX + r"\s+\d+\s?[A-Za-z]?\b"
        ),
    ),
    Detector(
        "ADDRESS",
        re.compile(r"\b\d{3}\s?\d{2}\s+[A-ZÅÄÖ][A-ZÅÄÖa-zåäö]+\b"),
    ),
]

_NAME_PART_RE = re.compile(r"^[A-ZÅÄÖÆØÉ][\w'-]+$", re.UNICODE)


@dataclass
class Vault:
    """In-memory, process-local mapping between tokens and real values."""

    _by_value: dict[tuple[str, str], str] = field(default_factory=dict)
    _by_token: dict[str, str] = field(default_factory=dict)
    _counters: dict[str, int] = field(default_factory=dict)
    _lock: threading.Lock = field(default_factory=threading.Lock)
    # surrogate -> kind, for stand-ins that don't look like [[TOKEN]]
    surrogates: dict[str, str] = field(default_factory=dict)

    def token_for(self, kind: str, value: str) -> str:
        key = (kind, value.casefold() if kind in ("EMAIL", "PERSON") else value)
        with self._lock:
            if key in self._by_value:
                return self._by_value[key]
            self._counters[kind] = self._counters.get(kind, 0) + 1
            token = f"[[{kind}_{self._counters[kind]}]]"
            self._by_value[key] = token
            self._by_token[token] = value
            return token

    def surrogate_for(self, kind: str, value: str, make: Callable[[], str]) -> str:
        """A made-up stand-in for value, unique within this vault."""
        key = (kind, re.sub(r"\D", "", value))
        with self._lock:
            if key in self._by_value:
                return self._by_value[key]
            surrogate = make()
            while surrogate in self._by_token:
                surrogate = make()
            self._by_value[key] = surrogate
            self._by_token[surrogate] = value
            self.surrogates[surrogate] = kind
            return surrogate

    def alias(self, token: str, value: str) -> None:
        """Register an extra token (e.g. first name of a person)."""
        with self._lock:
            self._by_token.setdefault(token, value)

    def lookup(self, token: str) -> str | None:
        return self._by_token.get(token)

    def clear(self) -> None:
        with self._lock:
            self._by_value.clear()
            self._by_token.clear()
            self._counters.clear()
            self.surrogates.clear()


class Redactor:
    def __init__(
        self,
        vault: Vault | None = None,
        allowlist: Iterable[str] = (),
        extra_names: Iterable[str] = (),
        use_ner: bool = False,
    ) -> None:
        self.vault = vault or Vault()
        self.allowlist = {a.casefold() for a in allowlist if a.strip()}
        # name part (e.g. "Anna") -> token for that part
        self._name_parts: dict[str, str] = {}
        self._lock = threading.Lock()
        self._nlp = _load_ner() if use_ner else None
        for name in extra_names:
            self.register_person(name)

    def reset(self) -> None:
        """Forget every token mapping and learned name."""
        with self._lock:
            self._name_parts.clear()
        self.vault.clear()

    # ---- registration -------------------------------------------------

    def register_person(self, name: str | None, email: str | None = None) -> None:
        """Teach the redactor a person's name, typically from mail headers."""
        name = (name or "").strip().strip('"')
        if (not name or "@" in name) and email:
            name = _name_from_email(email)
        if not name or name.casefold() in self.allowlist:
            return
        parts = [p for p in re.split(r"[\s,]+", name) if p]
        if len(parts) < 1 or not all(_NAME_PART_RE.match(p) for p in parts):
            return
        token = self.vault.token_for("PERSON", name)
        base = token[:-2]
        with self._lock:
            self._name_parts.setdefault(name, token)
            if len(parts) >= 2:
                for suffix, part in (("FIRST", parts[0]), ("LAST", parts[-1])):
                    if len(part) < 3 or part.casefold() in self.allowlist:
                        continue
                    if part in self._name_parts:
                        continue
                    part_token = f"{base}_{suffix}]]"
                    self.vault.alias(part_token, part)
                    self._name_parts[part] = part_token

    # ---- redaction ----------------------------------------------------

    def redact(self, text: str | None) -> str:
        if not text:
            return text or ""
        spans: list[tuple[int, int, str]] = []

        for start, end, fmt in idnumbers.find(text):
            value = text[start:end]
            spans.append((start, end, self.vault.surrogate_for("PERSONNUMMER", value, fmt.surrogate)))

        for det in DETECTORS:
            for m in det.pattern.finditer(text):
                value = m.group(0)
                if value.casefold() in self.allowlist:
                    continue
                if det.validate and not det.validate(value):
                    continue
                if det.kind == "EMAIL":
                    self.register_person(None, value)
                spans.append((m.start(), m.end(), self.vault.token_for(det.kind, value)))

        if self._nlp is not None:
            for ent in self._nlp(text).ents:
                if ent.label_ in ("PER", "PERSON") and ent.text.casefold() not in self.allowlist:
                    self.register_person(ent.text)

        with self._lock:
            names = sorted(self._name_parts.items(), key=lambda kv: -len(kv[0]))
        for name, token in names:
            for m in re.finditer(rf"(?<!\w){re.escape(name)}(?!\w)", text):
                spans.append((m.start(), m.end(), token))

        return _apply(text, spans)

    def restore(self, text: str | None, strict: bool = True) -> str:
        """Replace tokens with real values. Unknown tokens raise if strict."""
        if not text:
            return text or ""

        def sub(m: re.Match[str]) -> str:
            value = self.vault.lookup(m.group(0))
            if value is None:
                if strict:
                    raise UnknownTokenError(
                        f"Okänd token {m.group(0)} – använd bara tokens som "
                        "förekommer i hämtade mejl."
                    )
                return m.group(0)
            return value

        text = TOKEN_RE.sub(sub, text)
        for surrogate in sorted(self.vault.surrogates, key=len, reverse=True):
            if surrogate in text:
                text = text.replace(surrogate, self.vault.lookup(surrogate) or surrogate)
        return text

    def placeholders(self, text: str) -> list[tuple[int, int, str, str, str | None]]:
        """Tokens and surrogates in text: (start, end, placeholder, kind, original)."""
        found = [
            (m.start(), m.end(), m.group(0), m.group(1), self.vault.lookup(m.group(0)))
            for m in TOKEN_RE.finditer(text)
        ]
        for surrogate, kind in list(self.vault.surrogates.items()):
            for m in re.finditer(re.escape(surrogate), text):
                found.append((m.start(), m.end(), surrogate, kind, self.vault.lookup(surrogate)))
        return sorted(found)


def _apply(text: str, spans: list[tuple[int, int, str]]) -> str:
    # Longest span wins on overlap; ties go to the earliest detector.
    spans.sort(key=lambda s: (s[0], -(s[1] - s[0])))
    out: list[str] = []
    pos = 0
    for start, end, token in spans:
        if start < pos:
            continue
        out.append(text[pos:start])
        out.append(token)
        pos = end
    out.append(text[pos:])
    return "".join(out)


def _name_from_email(email: str) -> str:
    local = email.split("@", 1)[0]
    parts = re.split(r"[._-]+", local)
    if len(parts) < 2 or not all(p.isalpha() and len(p) > 1 for p in parts):
        return ""
    return " ".join(p.capitalize() for p in parts)


def _load_ner():
    try:
        import spacy  # type: ignore[import-not-found]
    except ImportError as exc:  # pragma: no cover - optional dependency
        raise RuntimeError(
            "REDACT_USE_NER=1 kräver 'pip install m365-mcp-gdpr[ner]' och "
            "'python -m spacy download sv_core_news_sm'"
        ) from exc
    return spacy.load("sv_core_news_sm")
