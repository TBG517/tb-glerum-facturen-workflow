# TB Glerum facturenworkflow

Mistral Workflow die de facturenadministratie van Technisch Bureau Glerum automatiseert:

1. **Outlook**: zoekt nieuwe e-mails met PDF-bijlagen in de map `Facturen TB` (postvak `info@tb-glerum.nl`).
2. **Uitlezen**: downloadt elke PDF en leest hem uit met Mistral OCR + een taalmodel (leverancier, factuurnummer, factuurdatum, bedragen ex/btw/incl., categorie).
3. **Opslaan**: zet de bijlage in OneDrive als `JJMMDD <Naam> fact. <factuurnummer>.pdf` in `Documenten/Facturen Mistral/Bijlagen`.
4. **Excel**: voegt een rij toe onderaan het tabblad `Facturen in` van `TB Glerum Overzicht Mistral.xlsx` (kolommen: Datum, Naam, Categorie, Factuur nummer, Bedrag ex, BTW, Bedrag in, Datum betaald) — in place via de Microsoft Graph workbook-API.

Reeds verwerkte facturen worden overgeslagen (idempotent): bestaat het doelbestand al, dan gebeurt er niets.

## Workflow

- Naam: `facturen_tb_process`
- Invoer: `{"dry_run": true, "days_back": 30, "max_mails": 25}`
  - `dry_run: true` = planning tonen zonder iets te schrijven; `false` = echt verwerken.
  - `days_back` = hoe ver teruggekeken wordt in de mailmap.
  - `max_mails` = maximum aantal mails per run (veiligheidslimiet).

## Benodigde secrets (workspace secrets in Mistral Studio)

| Naam | Waarde |
| --- | --- |
| `MS_TENANT_ID` | Tenant-ID van de Microsoft 365-tenant (Entra ID → Overview) |
| `MS_CLIENT_ID` | Application (client) ID van de app registration |
| `MS_CLIENT_SECRET` | Secret-waarde van de app registration |

### App registration aanmaken (eenmalig, tenantbeheerder)

1. Ga naar **portal.azure.com** → **Microsoft Entra ID** → **App registrations** → **New registration**
2. Naam: `Facturenverwerking Mistral`; Supported account types: single tenant → **Register**
3. **Certificates & secrets** → **New client secret** → kopieer de **Value** (wordt `MS_CLIENT_SECRET`)
4. Kopieer **Application (client) ID** (`MS_CLIENT_ID`) en **Directory (tenant) ID** (`MS_TENANT_ID`)
5. **API permissions** → **Add a permission** → **Microsoft Graph** → **Application permissions**:
   - `Mail.Read` (mailbox van info@tb-glerum.nl lezen, incl. bijlagen)
   - `Files.ReadWrite.All` (OneDrive: bestanden opslaan en Excel bijwerken)
6. Klik bij beide op **Grant admin consent** (voor de hele tenant).

## Configuratie (omgevingsvariabelen, optioneel)

Standaarden staan in `apps/worker/src/worker/workflows/facturen.py` en kunnen via de omgeving worden overschreven: `TB_MAILBOX`, `TB_MAIL_FOLDER`, `TB_DRIVE_ID`, `TB_EXCEL_ITEM_ID`, `TB_EXCEL_SHEET`, `TB_DEST_FOLDER`.

## Lokaal

```bash
mistral apps install   # dependencies
mistral apps dev --filter worker
```

## Tests

```bash
python3 apps/worker/tests/test_facturen_helpers.py
```

(Deze test dekt de pure logica: datumopmaak, bestandsnaamconventie en Excel-rijindeling; integratie met Graph/Mistral draait via de dry-run van de workflow zelf.)
