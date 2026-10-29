import pytest

from m365_mcp_gdpr.redact import Redactor, UnknownTokenError

MAIL = """Hej Anna!

Tack för samtalet. Mitt personnummer är 19811218-9876 och du når mig på
070-123 45 67 eller +46 8 123 456 78. Skicka fakturan till
Storgatan 12B, 114 51 Stockholm. IBAN: SE45 5000 0000 0583 9825 7466.

Vänliga hälsningar
Erik Lund
erik.lund@example.se
"""


def make():
    r = Redactor(allowlist=["Stockholm"])
    r.register_person("Anna Svensson", "anna.svensson@kund.se")
    return r


def test_redacts_swedish_pii():
    r = make()
    out = r.redact(MAIL)
    for secret in [
        "Anna", "19811218-9876", "070-123 45 67", "+46 8 123 456 78",
        "Storgatan 12B", "114 51", "SE45 5000", "Erik Lund", "erik.lund@example.se",
    ]:
        assert secret not in out, secret
    assert "[[PERSON_1_FIRST]]" in out
    assert "personnummer är 18" in out  # replaced with a made-up 1800s number
    assert "[[EMAIL_" in out


def test_name_learned_from_email_signature():
    r = make()
    r.redact("Mvh\nerik.lund@example.se")
    assert "Erik" not in r.redact("Ring Erik imorgon")


def test_roundtrip():
    r = make()
    out = r.redact(MAIL)
    assert r.restore(out) == MAIL


def test_stable_tokens():
    r = make()
    assert r.redact("anna.svensson@kund.se") == r.redact("ANNA.SVENSSON@kund.se")


def test_unknown_token_rejected():
    r = make()
    with pytest.raises(UnknownTokenError):
        r.restore("Hej [[PERSON_99]]")


def test_invalid_numbers_not_redacted():
    r = Redactor()
    # Ogiltigt datum respektive ogiltigt kortnummer ska lämnas orörda.
    text = "Order 20241301-0001 och artikel 1234567890123"
    assert r.redact(text) == text


def test_allowlist():
    r = Redactor(allowlist=["Matilda Foodtech"])
    r.register_person("Matilda Foodtech")
    assert r.redact("Från Matilda Foodtech") == "Från Matilda Foodtech"


def test_reset_forgets_names():
    r = make()
    r.redact("Hej Anna")
    r.reset()
    r.register_person("Anna Svensson")
    assert r.restore(r.redact("Hej Anna")) == "Hej Anna"
