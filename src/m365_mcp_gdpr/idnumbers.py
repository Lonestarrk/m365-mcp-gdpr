"""Nordic personal identity numbers: detection and 1800s surrogates.

Detected numbers are replaced with made-up numbers in the same national
format but with a birth year in the 1800s, so they cannot belong to a living
person. Where the format has a check digit the surrogate gets a valid one.

Supported: Sweden (personnummer, samordningsnummer), Denmark (CPR), Norway
(fødselsnummer, D-nummer) and Finland (henkilötunnus).
"""

from __future__ import annotations

import random
import re
from dataclasses import dataclass
from datetime import date
from typing import Callable

_rng = random.SystemRandom()


def _luhn_ok(digits: str) -> bool:
    total = 0
    for i, ch in enumerate(reversed(digits)):
        n = int(ch) * (2 if i % 2 else 1)
        total += n - 9 if n > 9 else n
    return total % 10 == 0


def _luhn_digit(body: str) -> str:
    for d in "0123456789":
        if _luhn_ok(body + d):
            return d
    raise AssertionError("unreachable")


def _valid_date(year: int, month: int, day: int) -> bool:
    try:
        date(year, month, day)
    except ValueError:
        return False
    return True


def _random_date(first_year: int, last_year: int) -> date:
    return date(_rng.randint(first_year, last_year), _rng.randint(1, 12), _rng.randint(1, 28))


# ---- Sweden ---------------------------------------------------------------

_SE_12 = re.compile(r"(?<![\d-])(1[89]|20)(\d{2})(\d{2})(\d{2})([-+ ]?)(\d{4})(?![\d-])")


def _se_date_ok(yy: str, mm: str, dd: str, century: int = 1900) -> bool:
    day = int(dd)
    if day > 60:  # samordningsnummer
        day -= 60
    return _valid_date(century + int(yy), int(mm), day) or _valid_date(2000 + int(yy), int(mm), day)


def _se12_match(m: re.Match[str]) -> bool:
    cc, yy, mm, dd, sep, nnnn = m.groups()
    if not _se_date_ok(yy, mm, dd, int(cc) * 100):
        return False
    # Without separator a 12-digit run could be anything; require the check digit.
    return bool(sep) or _luhn_ok(yy + mm + dd + nnnn)


def _se12_surrogate() -> str:
    d = _random_date(1800, 1899)
    body = f"{d:%y%m%d}{_rng.randint(0, 999):03d}"
    return f"18{body[:6]}-{body[6:]}{_luhn_digit(body)}"


# ---- Sweden 10 digits / Denmark CPR (same DDDDDD-NNNN shape) ------------------

_SIX_FOUR = re.compile(r"(?<![\d-])(\d{2})(\d{2})(\d{2})([-+]?)(\d{4})(?![\d-])")


def _cpr_date_ok(dd: str, mm: str, yy: str) -> bool:
    return _valid_date(2000 + int(yy), int(mm), int(dd)) or _valid_date(1900 + int(yy), int(mm), int(dd))


def _six_four_match(m: re.Match[str]) -> bool:
    a, b, c, sep, nnnn = m.groups()
    se_ok = _se_date_ok(a, b, c)
    dk_ok = _cpr_date_ok(a, b, c)
    if sep:
        return se_ok or dk_ok
    # No separator: ten bare digits are often phone or order numbers.
    return se_ok and _luhn_ok(a + b + c + nnnn)


def _six_four_surrogate() -> str:
    """A Danish CPR number for someone born 1858-1899, that is not a valid Swedish number."""
    while True:
        d = _random_date(1858, 1899)
        seventh = _rng.randint(5, 8)  # CPR century digit 5-8 with year 58-99 means 1800s
        serial = f"{seventh}{_rng.randint(0, 999):03d}"
        digits = f"{d:%d%m%y}{serial}"
        if not _luhn_ok(digits):
            return f"{digits[:6]}-{digits[6:]}"


# ---- Norway -----------------------------------------------------------------

_NO = re.compile(r"(?<!\d)(\d{6}) ?(\d{5})(?!\d)")
_NO_W1 = (3, 7, 6, 1, 8, 9, 4, 5, 2)
_NO_W2 = (5, 4, 3, 2, 7, 6, 5, 4, 3, 2)


def _mod11(digits: str, weights: tuple[int, ...]) -> int | None:
    r = 11 - sum(int(d) * w for d, w in zip(digits, weights)) % 11
    if r == 11:
        return 0
    return None if r == 10 else r


def _no_valid(eleven: str) -> bool:
    k1 = _mod11(eleven[:9], _NO_W1)
    k2 = _mod11(eleven[:10], _NO_W2)
    if k1 is None or k2 is None or eleven[9:] != f"{k1}{k2}":
        return False
    day = int(eleven[:2])
    if day > 40:  # D-nummer
        day -= 40
    return _cpr_date_ok(f"{day:02d}", eleven[2:4], eleven[4:6])


def _no_match(m: re.Match[str]) -> bool:
    return _no_valid(m.group(1) + m.group(2))


def _no_surrogate() -> str:
    while True:
        d = _random_date(1855, 1899)
        body = f"{d:%d%m%y}{_rng.randint(500, 749):03d}"  # individnummer 500-749 + year 54-99 = 1800s
        k1 = _mod11(body, _NO_W1)
        if k1 is None:
            continue
        k2 = _mod11(body + str(k1), _NO_W2)
        if k2 is None:
            continue
        return f"{body}{k1}{k2}"


# ---- Finland ----------------------------------------------------------------

_FI = re.compile(r"(?<![\w-])(\d{2})(\d{2})(\d{2})([-+ABCDEFUVWXY])(\d{3})([0-9A-Y])(?![\w-])")
_FI_CHECK = "0123456789ABCDEFHJKLMNPRSTUVWXY"


def _fi_match(m: re.Match[str]) -> bool:
    dd, mm, yy, _sign, nnn, check = m.groups()
    if not _cpr_date_ok(dd, mm, yy):
        return False
    return _FI_CHECK[int(dd + mm + yy + nnn) % 31] == check


def _fi_surrogate() -> str:
    d = _random_date(1800, 1899)
    nnn = f"{_rng.randint(2, 899):03d}"
    return f"{d:%d%m%y}+{nnn}{_FI_CHECK[int(f'{d:%d%m%y}' + nnn) % 31]}"


@dataclass(frozen=True)
class IdFormat:
    name: str
    pattern: re.Pattern[str]
    matches: Callable[[re.Match[str]], bool]
    surrogate: Callable[[], str]


# Order matters: longer formats first so 12-digit numbers aren't split.
FORMATS: list[IdFormat] = [
    IdFormat("SE", _SE_12, _se12_match, _se12_surrogate),
    IdFormat("NO", _NO, _no_match, _no_surrogate),
    IdFormat("FI", _FI, _fi_match, _fi_surrogate),
    IdFormat("SE/DK", _SIX_FOUR, _six_four_match, _six_four_surrogate),
]


def find(text: str) -> list[tuple[int, int, IdFormat]]:
    """Spans of personal identity numbers in text, longest first on overlap."""
    spans: list[tuple[int, int, IdFormat]] = []
    for fmt in FORMATS:
        for m in fmt.pattern.finditer(text):
            if fmt.matches(m):
                spans.append((m.start(), m.end(), fmt))
    return spans
