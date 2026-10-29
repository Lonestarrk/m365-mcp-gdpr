# m365-mcp-gdpr

MCP-server som ger Claude åtkomst till Microsoft 365 – Outlook-brevlådan och en
SharePoint-mapp för releaser – via Microsoft Graph, **med pseudonymisering av
personuppgifter innan något lämnar er maskin**.

Det är samma data som Microsoft 365 Copilot arbetar mot. Copilot har inget eget API
för att läsa/skriva mejl, så servern pratar direkt med Microsoft Graph.

## Vad den kan

| Verktyg | Beskrivning |
|---|---|
| `list_folders` | Mappar med antal (olästa) |
| `list_messages` | Lista mejl i en mapp, filtrera på olästa |
| `search_messages` | Fritextsök i hela brevlådan |
| `read_message` | Läs ett mejl (bilagors innehåll hämtas inte) |
| `mark_read`, `flag_message`, `set_categories`, `move_message` | Administrera |
| `create_draft`, `create_reply_draft`, `update_draft`, `delete_draft` | Utkast |
| `sharepoint_upload_release` | Ladda upp en lokal fil (t.ex. release-zip) till en förbestämd SharePoint-mapp |
| `sharepoint_list_releases` | Lista filer i release-mappen, med länkar |

Servern **kan inte skicka mejl**. Den begär `Mail.ReadWrite`, `Sites.ReadWrite.All`
och `User.Read` – men inte `Mail.Send`, så den saknar tekniskt behörighet att skicka. Alla utkast granskas och skickas av
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
3. *Authentication* → *Add a platform* → **Mobile and desktop applications**, lägg till:
   - `http://localhost` (inloggning via webbläsare)
   - `ms-appx-web://Microsoft.AAD.BrokerPlugin/<client-id>` (Windows inbyggda kontoruta)
   - `msauth.com.msauth.unsignedapp://auth` (Mac med Company Portal)
4. *Authentication* → *Allow public client flows* = **Yes** (behövs bara för `--device-code`).
5. *API permissions* → Microsoft Graph → **Delegated**: `Mail.ReadWrite`,
   `Sites.ReadWrite.All`, `User.Read`. `Sites.ReadWrite.All` kräver admin consent.
   Behörigheten är delegerad: servern kommer bara åt det du själv har behörighet till.
6. Notera *Application (client) ID* och *Directory (tenant) ID*.

### 2. Installera och logga in

```bash
pip install git+https://github.com/Lonestarrk/m365-mcp-gdpr
export M365_MCP_CLIENT_ID=<client-id>
export M365_MCP_TENANT_ID=<tenant-id>
m365-mcp-gdpr login
```

Inloggningen sker med ditt vanliga jobbkonto:

- **Windows:** systemets egen kontoruta visas, med ditt jobbkonto förvalt. Ofta behövs
  inget lösenord alls om du redan är inloggad i Windows med jobbkontot.
- **Mac:** via Company Portal om den är installerad (`pip install '...[broker]'`),
  annars öppnas Microsofts inloggningssida i webbläsaren.
- **Linux:** Microsofts inloggningssida öppnas i webbläsaren.

Tvåfaktorsinloggning (t.ex. Microsoft Authenticator) sköts som vanligt i den rutan.

Du behöver bara logga in en gång. Inloggningen sparas krypterat i operativsystemets
nyckelring (Windows DPAPI, macOS Keychain, Linux libsecret) och förnyas tyst i
bakgrunden. Den håller så länge du använder den regelbundet – som standard upphör den
efter 90 dagar utan användning. Er IT kan ha regler (Conditional Access, "sign-in
frequency") som kräver ny inloggning oftare; det styrs av organisationen, inte av servern.

Om du inte har loggat in i förväg öppnas inloggningen automatiskt första gången Claude
använder servern. `m365-mcp-gdpr logout` tar bort den sparade inloggningen.
På en dator utan webbläsare: `m365-mcp-gdpr login --device-code`.

### 3. Koppla till Claude

**Claude Code:**

```bash
claude mcp add m365 \
  -e M365_MCP_CLIENT_ID=<client-id> \
  -e M365_MCP_TENANT_ID=<tenant-id> \
  -e REDACT_ALLOWLIST="Matilda Foodtech,Stockholm" \
  -- m365-mcp-gdpr
```

**Claude Desktop** (`claude_desktop_config.json`):

```json
{
  "mcpServers": {
    "m365": {
      "command": "m365-mcp-gdpr",
      "env": {
        "M365_MCP_CLIENT_ID": "<client-id>",
        "M365_MCP_TENANT_ID": "<tenant-id>",
        "REDACT_ALLOWLIST": "Matilda Foodtech"
      }
    }
  }
}
```

## SharePoint-releaser

Servern kan lägga releaser som Claude Code byggt i en fast SharePoint-mapp, så att
kollegor alltid hittar senaste versionen. Claude kan bara ladda upp till den
konfigurerade mappen (och undermappar till den), aldrig någon annanstans.

```bash
claude mcp add m365 \
  -e M365_MCP_CLIENT_ID=<client-id> \
  -e M365_MCP_TENANT_ID=<tenant-id> \
  -e SHAREPOINT_SITE="matildafoodtech.sharepoint.com:/sites/Utveckling" \
  -e SHAREPOINT_FOLDER="Releases/MinApp" \
  -e SHAREPOINT_UPLOAD_ROOTS="$HOME/projekt/minapp/dist" \
  -- m365-mcp-gdpr
```

Exempel i Claude Code: *"Bygg en release, zippa dist-mappen och ladda upp den till
SharePoint i undermappen v1.4.0."* Befintliga filer skrivs inte över om du inte ber om det.
Filer över 4 MB laddas upp i delar, så stora releaser fungerar.

## Konfiguration

| Variabel | Beskrivning |
|---|---|
| `M365_MCP_CLIENT_ID` | Krävs. Client ID från app-registreringen |
| `M365_MCP_TENANT_ID` | Tenant-ID (standard `organizations`) |
| `M365_MCP_CACHE_DIR` | Var token-cachen sparas |
| `M365_MCP_LOGIN_HINT` | Ditt jobbmejl, förifyllt i inloggningsrutan |
| `M365_MCP_AUTO_LOGIN` | `0` = öppna aldrig inloggning automatiskt, kräv `login` i förväg |
| `M365_MCP_BROKER` | `0` = använd webbläsaren även där Windows/Mac-kontorutan finns |
| `REDACT_ALLOWLIST` / `REDACT_ALLOWLIST_FILE` | Termer som *inte* ska tvättas (företagsnamn, orter, produkter) |
| `REDACT_NAMES` / `REDACT_NAMES_FILE` | Extra namn som alltid ska tvättas (t.ex. kundregister-export), ett per rad |
| `SHAREPOINT_SITE` | Webbplats som `värd:/sites/Namn`, t.ex. `matildafoodtech.sharepoint.com:/sites/Utveckling` |
| `SHAREPOINT_DRIVE` | Dokumentbibliotekets namn (standard: webbplatsens standardbibliotek) |
| `SHAREPOINT_FOLDER` | Mappen releaser laddas upp till, t.ex. `Releases/MinApp` |
| `SHAREPOINT_UPLOAD_ROOTS` | Valfritt. Kommaseparerade lokala mappar som får laddas upp från – rekommenderas, så att Claude inte kan ladda upp andra filer |
| `REDACT_USE_NER` | `1` = aktivera spaCy-NER. Kräver `pip install 'm365-mcp-gdpr[ner]'` och `python -m spacy download sv_core_news_sm` |

## Utveckling

```bash
pip install -e '.[dev]'
pytest
```
