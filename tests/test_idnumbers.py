import re

import pytest

from m365_mcp_gdpr import idnumbers as idn
from m365_mcp_gdpr.redact import Redactor

REAL = {
    "se12": "19811218-9876",
    "se10": "811218-9876",
    "se_samordning": "811278-9873",
    "dk": "070761-4285",
    "no": "15058910063",
    "fi": "131052-308T",
}


@pytest.mark.parametrize("name, number", REAL.items())
def test_detected_and_replaced(name, number):
    r = Redactor()
    out = r.redact(f"Nummer: {number}.")
    assert number not in out
    assert r.restore(out) == f"Nummer: {number}."


def test_swedish_wrong_check_digit_with_dash_still_caught():
    r = Redactor()
    assert "20150312-1238" not in r.redact("personnr 20150312-1238")
    assert "150312-1238" not in r.redact("personnr 150312-1238")


def test_bare_digits_need_check_digit():
    r = Redactor()
    # Ten bare digits with a wrong check digit are usually order or phone numbers.
    assert r.redact("order 8112189875") == "order 8112189875"
    assert "8112189876" not in r.redact("pnr 8112189876")


def test_impossible_dates_ignored():
    r = Redactor()
    assert r.redact("ref 991399-1234 och 20241332-0001") == "ref 991399-1234 och 20241332-0001"


def _born_1800s_se(s: str) -> bool:
    return s.startswith("18") and idn._luhn_ok(s[2:].replace("-", ""))


@pytest.mark.parametrize("_", range(50))
def test_surrogates_are_1800s_and_valid(_):
    se = idn._se12_surrogate()
    assert re.fullmatch(r"18\d{6}-\d{4}", se) and _born_1800s_se(se)

    dk = idn._six_four_surrogate()
    assert re.fullmatch(r"\d{6}-\d{4}", dk)
    yy, seventh = int(dk[4:6]), int(dk[7])
    assert 58 <= yy <= 99 and 5 <= seventh <= 8  # CPR rule for 1858-1899
    assert not idn._luhn_ok(dk.replace("-", ""))  # never a valid Swedish number

    no = idn._no_surrogate()
    assert idn._no_valid(no) and 500 <= int(no[6:9]) <= 749 and int(no[4:6]) >= 55

    fi = idn._fi_surrogate()
    m = idn._FI.fullmatch(fi)
    assert m and m.group(4) == "+" and idn._fi_match(m)


def test_same_number_same_surrogate():
    r = Redactor()
    a = r.redact("19811218-9876")
    assert r.redact("19811218-9876") == a
    assert r.redact("198112189876") == a  # same digits, other formatting


def test_surrogate_restored_in_drafts():
    r = Redactor()
    seen = r.redact("Kundens nummer är 19811218-9876")
    fake = seen.split()[-1]
    assert fake.startswith("18")
    assert r.restore(f"Jag har noterat {fake}.") == "Jag har noterat 19811218-9876."
