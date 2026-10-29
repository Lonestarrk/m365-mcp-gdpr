# outlook-mcp-gdpr

MCP-server som ger Claude åtkomst till en Outlook-/Microsoft 365-brevlåda via
Microsoft Graph – **med pseudonymisering av personuppgifter innan något lämnar
er maskin**.

Det är samma brevlåda som Microsoft 365 Copilot arbetar mot. Copilot har inget
eget API för att läsa/skriva mejl, så servern pratar direkt med Microsoft Graph.

## Vad den kan

| Verktyg | Beskrivning |
|---|---|
| `list_folders` | Mappar med antal (olästa) |
| `list_messages` | Lista mejl i en mapp, filtrera på olästa |
| `search_messages` | Fritextsök i hela brevlådan |
| `read_message` | Läs ett mejl (bilagors innehåll hämtas inte) |
| `mark_read`, `flag_message`, `set_categories`, `move_message` | Administrera |
| `create_draft`, `create_reply_draft`, `update_draft`, `delete_draft` | Utkast |

Servern **kan inte skicka mejl**. Den begär bara `Mail.ReadWrite` + `User.Read`,
så den tekniskt saknar behörighet att skicka. Alla utkast granskas och skickas av
en människa i Outlook. `delete_draft` vägrar radera något som inte är ett utkast.

## Hur GDPR-tvätten fungerar

```
Outlook ──Graph──▶ [ MCP-server: redact() ] ──tokens──▶ Claude
Outlook ◀─Graph─── [ MCP-server: restore() ] ◀─tokens── Claude
                        │
                   Vault (endast RAM)
```

1. När ett mejl hämtas lär sig servern namn och adresser ur avsändare/mottagare.
2. Ämne, förhandsvisning, brödtext, bilagenamn och adresser tvättas. Ersätts med
   stabila tokens:

   ```
   Hej [[PERSON_1_FIRST]]! Mitt personnummer är [[PERSONNUMMER_1]], ring
   [[PHONE_1]]. Adress: [[ADDRESS_1]]. /[[PERSON_2]] [[EMAIL_3]]
   ```
3. Kopplingen token → riktigt värde finns **bara i serverprocessens minne** och
   skickas aldrig till modellen. Den försvinner när servern stängs.
4. När Claude skriver ett utkast med tokens byts de tillbaka till riktiga
   värden innan utkastet sparas i Outlook. Okända tokens avvisas.

**Detekteras:** e-postadresser, personnummer/samordningsnummer (Luhn-validerade),
telefonnummer (svenska + internationella), IBAN, kortnummer (Luhn), IP-adresser,
svenska gatuadresser och postnummer+ort, samt personnamn från mejlhuvuden,
e-postadresser (`fornamn.efternamn@`) och en valfri namnlista. Med
`REDACT_USE_NER=1` används även spaCys svenska NER-modell för namn i fritext.

### Begränsningar – läs detta

- Regelbaserad tvätt är aldrig 100 %. Namn i fritext på personer som inte
  förekommer i mejlhuvuden fångas bara med NER eller namnlista. Hälsodata,
  fackmedlemskap och andra känsliga uppgifter (art. 9) i fritext fångas inte.
- Pseudonymiserade uppgifter är fortfarande personuppgifter enligt GDPR
  (art. 4.5, skäl 26). Tvätten är en skyddsåtgärd (dataminimering, art. 25 och 32)
  – den ersätter inte rättslig grund, personuppgiftsbiträdesavtal (DPA) med
  Anthropic, registerförteckning eller eventuell DPIA.
- Mejl-ID:n och datum skickas otvättade (de är inte personuppgifter i sig).

## Installation

### 1. App-registrering i Microsoft Entra ID

1. Entra admin center → *App registrations* → *New registration*.
2. *Supported account types*: Endast den här organisationen.
3. *Authentication* → *Allow public client flows* = **Yes** (för device code).
4. *API permissions* → Microsoft Graph → **Delegated**: `Mail.ReadWrite`, `User.Read`.
   Ge admin consent om er tenant kräver det.
5. Notera *Application (client) ID* och *Directory (tenant) ID*.

### 2. Installera och logga in

```bash
pip install git+https://github.com/Lonestarrk/outlook-mcp-gdpr
export OUTLOOK_MCP_CLIENT_ID=<client-id>
export OUTLOOK_MCP_TENANT_ID=<tenant-id>
outlook-mcp-gdpr login      # device code-inloggning, en gång
```

Token-cachen sparas i `~/.outlook-mcp-gdpr/token_cache.json` (filrättigheter 600).

### 3. Koppla till Claude

**Claude Code:**

```bash
claude mcp add outlook \
  -e OUTLOOK_MCP_CLIENT_ID=<client-id> \
  -e OUTLOOK_MCP_TENANT_ID=<tenant-id> \
  -e REDACT_ALLOWLIST="Matilda Foodtech,Stockholm" \
  -- outlook-mcp-gdpr
```

**Claude Desktop** (`claude_desktop_config.json`):

```json
{
  "mcpServers": {
    "outlook": {
      "command": "outlook-mcp-gdpr",
      "env": {
        "OUTLOOK_MCP_CLIENT_ID": "<client-id>",
        "OUTLOOK_MCP_TENANT_ID": "<tenant-id>",
        "REDACT_ALLOWLIST": "Matilda Foodtech"
      }
    }
  }
}
```

## Konfiguration

| Variabel | Beskrivning |
|---|---|
| `OUTLOOK_MCP_CLIENT_ID` | Krävs. Client ID från app-registreringen |
| `OUTLOOK_MCP_TENANT_ID` | Tenant-ID (standard `organizations`) |
| `OUTLOOK_MCP_CACHE_DIR` | Var token-cachen sparas |
| `REDACT_ALLOWLIST` / `REDACT_ALLOWLIST_FILE` | Termer som *inte* ska tvättas (företagsnamn, orter, produkter) |
| `REDACT_NAMES` / `REDACT_NAMES_FILE` | Extra namn som alltid ska tvättas (t.ex. kundregister-export), ett per rad |
| `REDACT_USE_NER` | `1` = aktivera spaCy-NER. Kräver `pip install 'outlook-mcp-gdpr[ner]'` och `python -m spacy download sv_core_news_sm` |

## Utveckling

```bash
pip install -e '.[dev]'
pytest
```
