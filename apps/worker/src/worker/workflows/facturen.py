"""TB Glerum facturenverwerking.

Verwerkt factuurmails uit de Outlook-map "Facturen TB" (Postvak IN-submap):

1. Zoekt nieuwe e-mails met bijlagen (sinds N dagen).
2. Downloadt de PDF-bijlagen en leest ze uit met Mistral OCR
   (leverancier, factuurnummer, factuurdatum, bedragen, categorie).
3. Slaat elke bijlage op in OneDrive als "JJMMDD <Naam> fact. <nummer>.pdf".
4. Voegt een rij toe aan het Excel-tabblad "Facturen in" van het overzicht
   in OneDrive (in place, via de Microsoft Graph workbook-API).

Ongebruikte/genegeerde bijlagen (bijv. afbeeldingshandtekeningen) worden
overgeslagen. Reeds verwerkte facturen worden gedetecteerd aan de hand van
de aanwezigheid van het doelbestand in OneDrive (idempotent herstarten).

Alle Microsoft-toegang loopt via een app registration (client credentials);
tenant/client-id/secret staan in de workspace secrets en worden uitsluitend
binnen activities gelezen.
"""

import base64
import io
import os
import re
from datetime import datetime, timedelta, timezone
from typing import Any
from urllib.parse import quote

import pydantic
import mistralai.workflows as workflows
from mistralai.client.models import DocumentURLChunk
from mistralai.workflows.plugins.mistralai import (
    ChatCompletionRequest,
    OCRRequest,
    UserMessage,
    chat_parse_to_model,
    get_mistral_client,
    mistralai_ocr,
)

GRAPH = "https://graph.microsoft.com/v1.0"
LOGIN = "https://login.microsoftonline.com"

# --- vaste configuratie (via omgeving overschrijfbaar) -----------------------


def _cfg(key: str, default: str) -> str:
    return os.environ.get(key) or default


MAILBOX = _cfg("TB_MAILBOX", "info@tb-glerum.nl")
MAIL_FOLDER = _cfg("TB_MAIL_FOLDER", "Facturen TB")
DRIVE_ID = _cfg("TB_DRIVE_ID", "b!NWy4vc2FQkyqIP8XxhWZIrf9VegJN7RKuewUNFXAXM_ske4lKF3OQ4LPhU-D_PY1")
EXCEL_ITEM_ID = _cfg("TB_EXCEL_ITEM_ID", "015F4XENZMFHQLES6LV5EIC5RARQUSEADG")
EXCEL_SHEET = _cfg("TB_EXCEL_SHEET", "Facturen in")
DEST_FOLDER = _cfg("TB_DEST_FOLDER", "Facturen Mistral/Bijlagen")

CATEGORIES = [
    "Auto",
    "Telefoon/internet",
    "Computer",
    "Kleding",
    "Gereedschap",
    "Verzekering",
    "Administratie",
    "Wegenbelasting",
    "Loods",
    "Overig",
]


# --- invoer/uitvoer ----------------------------------------------------------


class FacturenRunInput(pydantic.BaseModel):
    """Invoer voor één verwerkingsrun."""

    dry_run: bool = True
    days_back: int = 30
    max_mails: int = 25


class InvoiceRow(pydantic.BaseModel):
    """Eén verwerkte factuur; ook de Excel-rij-indeling (kolommen A t/m H)."""

    datum: str
    naam: str
    categorie: str
    factuur_nummer: str
    bedrag_ex: str = ""
    btw: str = ""
    bedrag_in: str = ""
    datum_betaald: str = ""

    def excel_row(self) -> list[str]:
        return [
            self.datum,
            self.naam,
            self.categorie,
            self.factuur_nummer,
            self.bedrag_ex,
            self.btw,
            self.bedrag_in,
            self.datum_betaald,
        ]


class ExtractedInvoice(pydantic.BaseModel):
    """Gestructureerde factuurdata uit het OCR-model."""

    naam: str = ""
    factuur_nummer: str = ""
    factuurdatum: str = ""
    bedrag_ex: str = ""
    btw: str = ""
    bedrag_in: str = ""
    categorie: str = ""


class MailSummary(pydantic.BaseModel):
    message_id: str
    subject: str
    sender_name: str
    sender_address: str
    received: str  # ISO-8601
    attachment_ids: list[str]
    attachment_names: list[str]


class ProcessResult(pydantic.BaseModel):
    planned_uploads: list[str] = []
    uploaded: list[str] = []
    rows_added: list[list[str]] = []
    skipped: list[str] = []
    errors: list[str] = []
    mails_seen: int = 0


# --- Microsoft Graph helpers (binnen activities) -----------------------------


class GraphError(workflows.WorkflowError):
    def __init__(self, message: str, status: int = 0):
        super().__init__(message, non_retryable=status in (400, 401, 403, 404))
        self.status = status


def _credentials() -> tuple[str, str, str]:
    tenant = os.environ.get("MS_TENANT_ID", "")
    client_id = os.environ.get("MS_CLIENT_ID", "")
    client_secret = os.environ.get("MS_CLIENT_SECRET", "")
    if not (tenant and client_id and client_secret):
        raise workflows.WorkflowError(
            "microsoft_credentials_missing",
            non_retryable=True,
        )
    return tenant, client_id, client_secret


async def _http() -> Any:
    import httpx

    return httpx.AsyncClient(timeout=60.0)


@workflows.activity(name="facturen-get-token", retry_policy_max_attempts=3)
async def get_token() -> str:
    """Haalt een Microsoft Graph-toegangstoken op (client credentials)."""
    tenant, client_id, client_secret = _credentials()
    client = await _http()
    try:
        resp = await client.post(
            f"{LOGIN}/{tenant}/oauth2/v2.0/token",
            data={
                "grant_type": "client_credentials",
                "client_id": client_id,
                "client_secret": client_secret,
                "scope": "https://graph.microsoft.com/.default",
            },
        )
        if resp.status_code != 200:
            raise GraphError("token_request_failed", resp.status_code)
        return resp.json()["access_token"]
    finally:
        await client.aclose()


@workflows.activity(name="facturen-list-mails", retry_policy_max_attempts=3)
async def list_new_mails(token: str, days_back: int, max_mails: int) -> list[MailSummary]:
    """Geeft mails met bijlagen uit de map Facturen TB sinds days_back dagen."""
    since_iso = (datetime.now(timezone.utc) - timedelta(days=days_back)).strftime(
        "%Y-%m-%dT%H:%M:%SZ"
    )
    client = await _http()
    headers = {"Authorization": f"Bearer {token}"}
    try:
        folder_id = ""
        for base in (
            f"{GRAPH}/users/{MAILBOX}/mailFolders",
            f"{GRAPH}/users/{MAILBOX}/mailFolders/inbox/childFolders",
        ):
            folders = await client.get(
                base,
                params={"$filter": "displayName eq '" + MAIL_FOLDER + "'"},
                headers=headers,
            )
            items = folders.json().get("value", []) if folders.status_code == 200 else []
            if items:
                folder_id = items[0]["id"]
                break
        if not folder_id:
            raise GraphError("mail_folder_not_found", 404)

        mails: list[MailSummary] = []
        url: str | None = (
            f"{GRAPH}/users/{MAILBOX}/mailFolders/{folder_id}/messages"
            f"?$filter=receivedDateTime ge {since_iso}"
            f"&$top={max_mails}&$orderby=receivedDateTime asc"
        )
        while url and len(mails) < max_mails:
            resp = await client.get(url, headers=headers)
            if resp.status_code != 200:
                raise GraphError("list_messages_failed", resp.status_code)
            for m in resp.json().get("value", []):
                atts = await client.get(
                    f"{GRAPH}/users/{MAILBOX}/messages/{m['id']}/attachments",
                    headers=headers,
                )
                att_items = atts.json().get("value", []) if atts.status_code == 200 else []
                keep = [
                    a
                    for a in att_items
                    if not a.get("isInline")
                    and (a.get("contentType") == "application/pdf"
                         or str(a.get("name", "")).lower().endswith(".pdf"))
                ]
                if keep:
                    sender = (m.get("from") or {}).get("emailAddress", {}) or {}
                    mails.append(
                        MailSummary(
                            message_id=m["id"],
                            subject=m.get("subject") or "",
                            sender_name=sender.get("name", ""),
                            sender_address=sender.get("address", ""),
                            received=m["receivedDateTime"],
                            attachment_ids=[a["id"] for a in keep],
                            attachment_names=[a["name"] for a in keep],
                        )
                    )
            url = resp.json().get("@odata.nextLink")
        return mails
    finally:
        await client.aclose()


@workflows.activity(name="facturen-download-attachment", retry_policy_max_attempts=3)
async def download_attachment(token: str, message_id: str, attachment_id: str) -> tuple[str, str]:
    """Downloadt één bijlage; geeft (bestandsnaam, base64-inhoud)."""
    client = await _http()
    try:
        resp = await client.get(
            f"{GRAPH}/users/{MAILBOX}/messages/{message_id}/attachments/{attachment_id}/$value",
            headers={"Authorization": f"Bearer {token}"},
        )
        if resp.status_code != 200:
            raise GraphError("download_attachment_failed", resp.status_code)
        meta = await client.get(
            f"{GRAPH}/users/{MAILBOX}/messages/{message_id}/attachments/{attachment_id}",
            headers={"Authorization": f"Bearer {token}"},
        )
        name = meta.json().get("name", "factuur.pdf") if meta.status_code == 200 else "factuur.pdf"
        return name, base64.b64encode(resp.content).decode()
    finally:
        await client.aclose()


@workflows.activity(name="facturen-extract-invoice", retry_policy_max_attempts=2)
async def extract_invoice(
    pdf_b64: str, subject: str, sender_name: str, received: str
) -> ExtractedInvoice:
    """Leest een factuur-PDF uit met OCR en structureren."""
    client = get_mistral_client()
    upload = client.files.upload(
        file={"file": io.BytesIO(base64.b64decode(pdf_b64)), "file_name": "factuur.pdf"},
        purpose="ocr",
    )
    signed = client.files.get_signed_url(file_id=upload.id)
    ocr = await mistralai_ocr(
        OCRRequest(
            model="mistral-ocr-latest",
            document=DocumentURLChunk(document_url=signed.url),
        )
    )
    text = "\n".join(page.markdown for page in ocr.pages)

    prompt = (
        "Hieronder staat de OCR-tekst van een factuur en de metadata van de e-mail "
        "waarmee hij binnenkwam. Vul de velden; ontbrekende waarden leeg laten, "
        "nooit verzinnen. Bedragen als getal met een punt als decimaalteken.\n"
        f"- naam: korte leveranciersnaam voor een overzicht (bijv. 'MoveMove', 'IKEA', 'KPN')\n"
        "- factuur_nummer: het factuurnummer op de factuur\n"
        "- factuurdatum: als 'DD-MM-JJJJ' (zoniet bekend: leeg)\n"
        "- bedrag_ex, btw, bedrag_in: bedragen\n"
        f"- categorie: precies een van {CATEGORIES}\n\n"
        f"E-mailonderwerp: {subject}\nE-mailafzender: {sender_name}\n"
        f"Ontvangen op: {received}\n\nFactuurtekst:\n{text[:20000]}"
    )
    request = ChatCompletionRequest(
        model="mistral-small-latest",
        messages=[UserMessage(content=prompt)],
    )
    return await chat_parse_to_model(ExtractedInvoice, request)


@workflows.activity(name="facturen-file-exists", retry_policy_max_attempts=3)
async def file_exists(token: str, folder_path: str, filename: str) -> bool:
    """Controleert of het doelbestand al bestaat (verwerking is dan al gebeurd)."""
    client = await _http()
    try:
        resp = await client.get(
            f"{GRAPH}/users/{MAILBOX}/drives/{DRIVE_ID}/root:/{folder_path}/{filename}",
            headers={"Authorization": f"Bearer {token}"},
        )
        return resp.status_code == 200
    finally:
        await client.aclose()


@workflows.activity(name="facturen-upload", retry_policy_max_attempts=1)
async def upload_file(
    token: str, folder_path: str, filename: str, pdf_b64: str
) -> str:
    """Slaat een bijlage op in OneDrive onder de opgegeven naam."""
    client = await _http()
    try:
        content = base64.b64decode(pdf_b64)
        resp = await client.put(
            f"{GRAPH}/users/{MAILBOX}/drives/{DRIVE_ID}/root:/{folder_path}/{filename}:/content",
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/pdf"},
            content=content,
        )
        if resp.status_code not in (200, 201):
            raise GraphError("upload_failed", resp.status_code)
        return filename
    finally:
        await client.aclose()


@workflows.activity(name="facturen-append-excel", retry_policy_max_attempts=1)
async def append_excel_rows(token: str, rows: list[list[str]]) -> int:
    """Voegt rijen toe onderaan het tabblad 'Facturen in' (Graph workbook-API)."""
    if not rows:
        return 0
    client = await _http()
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    session_id = ""
    try:
        sess = await client.post(
            f"{GRAPH}/users/{MAILBOX}/drives/{DRIVE_ID}/items/{EXCEL_ITEM_ID}/workbook/createSession",
            headers=headers,
            json={"persistChanges": True},
        )
        if sess.status_code in (200, 201):
            session_id = sess.json().get("id", "")
            headers["workbook-session-id"] = session_id

        used = await client.get(
            f"{GRAPH}/users/{MAILBOX}/drives/{DRIVE_ID}/items/{EXCEL_ITEM_ID}"
            f"/workbook/worksheets('{quote(EXCEL_SHEET)}')/usedRange",
            headers=headers,
        )
        if used.status_code != 200:
            raise GraphError("usedrange_failed", used.status_code)
        first_row = used.json()["rowIndex"]  # 0-based; rij 1 = kopregel
        row_count = used.json()["rowCount"]
        start = first_row + row_count + 1  # 1-based Excel-rijnummer

        width = 8
        address = f"A{start}:{chr(ord('A') + width - 1)}{start + len(rows) - 1}"
        patch = await client.patch(
            f"{GRAPH}/users/{MAILBOX}/drives/{DRIVE_ID}/items/{EXCEL_ITEM_ID}"
            f"/workbook/worksheets('{quote(EXCEL_SHEET)}')/range(address='{address}')",
            headers=headers,
            json={"values": rows},
        )
        if patch.status_code not in (200, 201):
            raise GraphError("append_rows_failed", patch.status_code)
        return len(rows)
    finally:
        if session_id:
            try:
                await client.post(
                    f"{GRAPH}/users/{MAILBOX}/drives/{DRIVE_ID}/items/{EXCEL_ITEM_ID}/workbook/closeSession",
                    headers=headers,
                )
            finally:
                await client.aclose()


# --- hulpfuncties (deterministisch, in de workflow) --------------------------


def _fmt_date(iso_like: str, fallback: str) -> str:
    """Van de eerste herkenbare datum een DD-MM-JJJJ maken."""
    for candidate in (iso_like, fallback):
        if not candidate:
            continue
        match = re.search(r"(\d{4})-(\d{2})-(\d{2})", candidate)
        if match:
            return f"{match.group(3)}-{match.group(2)}-{match.group(1)}"
        match = re.search(r"(\d{1,2})-(\d{1,2})-(\d{4})", candidate)
        if match:
            return f"{int(match.group(1)):02d}-{int(match.group(2)):02d}-{match.group(3)}"
    return ""


def _filename(row: InvoiceRow, ext: str) -> str:
    """Bouwt 'JJMMDD Naam fact. nummer.ext'."""
    dd = row.datum[:2]
    mm = row.datum[3:5]
    yy = row.datum[-2:]
    number = row.factuur_nummer or "onbekend"
    safe_name = re.sub(r"[\\/:*?\"<>|]", "", row.naam).strip() or "Onbekend"
    return f"{yy}{mm}{dd} {safe_name} fact. {number}{ext}"


# --- de workflow --------------------------------------------------------------


@workflows.workflow.define(name="facturen_tb_process")
class FacturenTbProcessWorkflow:
    @workflows.workflow.entrypoint
    async def run(self, data: FacturenRunInput) -> ProcessResult:
        result = ProcessResult()

        token = await get_token()
        mails = await list_new_mails(token, data.days_back, data.max_mails)
        result.mails_seen = len(mails)

        for mail in mails:
            for att_id, att_name in zip(mail.attachment_ids, mail.attachment_names):
                ext = os.path.splitext(att_name)[1] or ".pdf"
                try:
                    _, pdf_b64 = await download_attachment(token, mail.message_id, att_id)
                    extracted = await extract_invoice(
                        pdf_b64, mail.subject, mail.sender_name, mail.received
                    )
                    row = InvoiceRow(
                        datum=_fmt_date(extracted.factuurdatum, mail.received),
                        naam=extracted.naam or mail.sender_name or "Onbekend",
                        categorie=(
                            extracted.categorie
                            if extracted.categorie in CATEGORIES
                            else "Overig"
                        ),
                        factuur_nummer=extracted.factuur_nummer,
                        bedrag_ex=extracted.bedrag_ex,
                        btw=extracted.btw,
                        bedrag_in=extracted.bedrag_in,
                    )
                    if not row.datum:
                        row.datum = _fmt_date(mail.received, "")
                    filename = _filename(row, ext)

                    if await file_exists(token, DEST_FOLDER, filename):
                        result.skipped.append(filename)
                        continue

                    if data.dry_run:
                        result.planned_uploads.append(filename)
                        result.rows_added.append(row.excel_row())
                        continue

                    await upload_file(token, DEST_FOLDER, filename, pdf_b64)
                    result.uploaded.append(filename)
                    result.rows_added.append(row.excel_row())
                except workflows.WorkflowError as exc:
                    result.errors.append(f"{att_name}: {exc}")
                except Exception as exc:  # noqa: BLE001 - per bijlage isoleren
                    result.errors.append(f"{att_name}: {exc}")

        if not data.dry_run and result.rows_added:
            await append_excel_rows(token, result.rows_added)

        return result
