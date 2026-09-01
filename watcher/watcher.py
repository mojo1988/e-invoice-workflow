import logging
import os
import re
import shutil
import sys
import time
from datetime import date, datetime
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
import subprocess

import requests
import yaml
from openpyxl import load_workbook
from pypdf import PdfReader, PdfWriter
from reportlab.lib.units import cm
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader


API_URL = os.getenv("API_URL", "http://e-invoice-eu:3000").rstrip("/")
INPUT_DIR = Path(os.getenv("INPUT_DIR", "/data/input"))
OUTPUT_DIR = Path(os.getenv("OUTPUT_DIR", "/data/output"))
ARCHIVE_DIR = Path(os.getenv("ARCHIVE_DIR", "/data/archive"))
ERROR_DIR = Path(os.getenv("ERROR_DIR", "/data/error"))
MAPPING_DIR = Path(os.getenv("MAPPING_DIR", "/data/mapping"))
WORK_DIR = Path(os.getenv("WORK_DIR", "/data/work"))
LOG_DIR = Path(os.getenv("LOG_DIR", "/data/logs"))
LOGO_PATH = Path(os.getenv("LOGO_PATH", "/data/assets/hsp-logo.png"))
LOGO_HEIGHT_CM = Decimal(os.getenv("LOGO_HEIGHT_CM", "2.5"))
LOGO_TOP_CM = Decimal(os.getenv("LOGO_TOP_CM", "0.0"))
LOGO_MARGIN_CM = Decimal(os.getenv("LOGO_MARGIN_CM", "1.0"))
FORMAT = os.getenv("FORMAT", "Factur-X-EN16931")
SHEET_NAME = os.getenv("SHEET_NAME", "Tabelle1")
POLL_SECONDS = int(os.getenv("POLL_SECONDS", "10"))
STABLE_SECONDS = int(os.getenv("STABLE_SECONDS", "10"))
REQUEST_TIMEOUT = int(os.getenv("REQUEST_TIMEOUT", "180"))

UNIT_CODES = {
    "mh": "HUR",
    "h": "HUR",
    "std": "HUR",
    "std.": "HUR",
    "stunde": "HUR",
    "stunden": "HUR",
    "hour": "HUR",
    "hours": "HUR",
    "stk": "C62",
    "stk.": "C62",
    "stück": "C62",
    "stueck": "C62",
    "st.": "C62",
    "piece": "C62",
    "pieces": "C62",
    "kg": "KGM",
    "kilogramm": "KGM",
    "g": "GRM",
    "gramm": "GRM",
    "m": "MTR",
    "meter": "MTR",
    "km": "KMT",
    "l": "LTR",
    "liter": "LTR",
    "tag": "DAY",
    "tage": "DAY",
    "monat": "MON",
    "monate": "MON",
}


def setup_logging():
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
        handlers=[
            logging.StreamHandler(sys.stdout),
            logging.FileHandler(LOG_DIR / "watcher.log", encoding="utf-8"),
        ],
    )


def text(value):
    return "" if value is None else str(value).strip()


def optional(value):
    value = text(value)
    return value if value else None


def date_iso(value):
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()

    value = text(value)
    if not value:
        return None

    for pattern in ("%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y"):
        try:
            return datetime.strptime(value, pattern).date().isoformat()
        except ValueError:
            pass

    raise ValueError(f"Ungültiges Datum: {value!r}")


def money(value):
    if value in (None, ""):
        raise ValueError("Leerer Geldbetrag")

    raw = str(value).strip().replace(" ", "")
    if "," in raw:
        raw = raw.replace(".", "").replace(",", ".")

    number = Decimal(raw).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    return format(number, ".2f")


def quantity(value):
    if value in (None, ""):
        raise ValueError("Leere Menge")

    raw = str(value).strip().replace(" ", "").replace(",", ".")
    number = Decimal(raw).quantize(Decimal("0.0001"), rounding=ROUND_HALF_UP)
    return format(number, "f").rstrip("0").rstrip(".") or "0"


def tax_percent(value):
    if value in (None, ""):
        raise ValueError("Leerer Umsatzsteuersatz")

    raw = str(value).strip().replace("%", "").replace(",", ".")
    rate = Decimal(raw)

    if rate <= 1:
        rate *= Decimal("100")

    rate = rate.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    return format(rate, "f").rstrip("0").rstrip(".")


def normalize_unit(value):
    raw = text(value)

    if not raw:
        return "C62"

    normalized = raw.lower()
    if normalized in UNIT_CODES:
        return UNIT_CODES[normalized]

    if re.fullmatch(r"[A-Z0-9]{2,3}", raw):
        return raw.upper()

    logging.warning("Unbekannte Einheit %r; verwende C62.", raw)
    return "C62"


def parse_plz_ort(value):
    source = text(value)
    match = re.search(r"(?:\bD\s*[-–]\s*)?(\d{5})\s+(.+)$", source, re.IGNORECASE)

    if not match:
        raise ValueError(
            f"PLZ/Ort aus A6 nicht lesbar: {source!r}. "
            "Erwartet z. B. 'D-90552 Röthenbach' oder '90552 Röthenbach'."
        )

    return match.group(1), match.group(2).strip()


def safe_filename(value):
    value = re.sub(r'[\\/:*?"<>|]+', "_", value)
    value = re.sub(r"\s+", " ", value).strip(" ._")
    return value or "rechnung"


def is_stable(path):
    try:
        first = path.stat()
        time.sleep(STABLE_SECONDS)
        second = path.stat()

        return (
            first.st_size > 0
            and first.st_size == second.st_size
            and first.st_mtime_ns == second.st_mtime_ns
        )
    except FileNotFoundError:
        return False


def load_stammdaten():
    path = MAPPING_DIR / "stammdaten.yaml"

    if not path.exists():
        raise FileNotFoundError(f"Stammdaten-Datei fehlt: {path}")

    with path.open("r", encoding="utf-8") as handle:
        data = yaml.safe_load(handle) or {}

    seller = data.get("seller", {})
    payment = data.get("payment", {})
    defaults = data.get("invoiceDefaults", {})

    required = {
        "seller.name": seller.get("name"),
        "seller.street": seller.get("street"),
        "seller.postalCode": seller.get("postalCode"),
        "seller.city": seller.get("city"),
        "seller.vatId": seller.get("vatId"),
        "payment.accountName": payment.get("accountName"),
        "payment.iban": payment.get("iban"),
    }

    placeholders = {
        "FIRMENNAME ERGÄNZEN",
        "STRASSE UND HAUSNUMMER ERGÄNZEN",
        "PLZ ERGÄNZEN",
        "ORT ERGÄNZEN",
        "DEUSTIDNR_ERGAENZEN",
        "DEIBAN_ERGAENZEN",
        "BIC_ERGAENZEN",
        "KONTOINHABER ERGÄNZEN",
        "ANSPRECHPARTNER ERGÄNZEN",
        "TELEFON ERGÄNZEN",
        "E-MAIL ERGÄNZEN",
    }

    missing = []
    for key, value in required.items():
        actual = text(value)
        if not actual or actual.upper() in placeholders:
            missing.append(key)

    if missing:
        raise ValueError("Stammdaten sind noch nicht vollständig: " + ", ".join(missing))

    return seller, payment, defaults


def read_and_normalize(source_path):
    wb_edit = load_workbook(source_path, data_only=False)
    wb_values = load_workbook(source_path, data_only=True)

    if SHEET_NAME not in wb_edit.sheetnames:
        raise ValueError(
            f"Arbeitsblatt {SHEET_NAME!r} fehlt. Vorhanden: {', '.join(wb_edit.sheetnames)}"
        )

    ws_edit = wb_edit[SHEET_NAME]
    ws_values = wb_values[SHEET_NAME]

    invoice_id = text(ws_values["H8"].value)
    issue_date = date_iso(ws_values["H11"].value)
    buyer_name = text(ws_values["A7"].value)
    buyer_street = text(ws_values["A9"].value)
    buyer_postal, buyer_city = parse_plz_ort(ws_values["A10"].value)

    if not invoice_id:
        raise ValueError("Rechnungsnummer in H8 fehlt.")
    if not issue_date:
        raise ValueError("Rechnungsdatum in H11 fehlt.")
    if not buyer_name:
        raise ValueError("Käuferfirma in A7 fehlt.")
    if not buyer_street:
        raise ValueError("Käuferstraße in A9 fehlt.")

    order_id = optional(ws_values["A18"].value)
    delivery_note = optional(ws_values["H9"].value)
    period_start = date_iso(ws_values["E16"].value) if ws_values["E16"].value else None
    period_end = date_iso(ws_values["E18"].value) if ws_values["E18"].value else None
    payment_terms = " ".join(
        item for item in (
            optional(ws_values["B38"].value),
            optional(ws_values["C38"].value),
            optional(ws_values["D38"].value),
        )
        if item
    )

    tax_rate = tax_percent(ws_values["C34"].value)

    positions = []
    for row in range(23, 32):
        position_number = ws_values.cell(row=row, column=1).value
        position_quantity = ws_values.cell(row=row, column=2).value
        position_unit = ws_values.cell(row=row, column=3).value
        position_description = optional(ws_values.cell(row=row, column=4).value)
        position_price = ws_values.cell(row=row, column=7).value
        position_total = ws_values.cell(row=row, column=8).value

        if (
            position_quantity in (None, "")
            and position_price in (None, "")
            and position_total in (None, "")
        ):
            continue

        if position_quantity in (None, "") or position_price in (None, ""):
            raise ValueError(f"Unvollständige Position in Excel-Zeile {row}.")

        if position_total in (None, ""):
            calculated_total = (
                Decimal(quantity(position_quantity))
                * Decimal(money(position_price))
            )
            position_total = calculated_total

        description = (
            position_description
            or order_id
            or f"Leistungen gemäß Rechnung {invoice_id}"
        )
        unit = normalize_unit(position_unit)

        ws_edit.cell(row=row, column=3).value = unit
        ws_edit.cell(row=row, column=4).value = description
        ws_edit.cell(row=row, column=8).value = float(Decimal(money(position_total)))
        ws_edit.cell(row=row, column=26).value = "InvoiceLine"

        positions.append(
            {
                "row": row,
                "id": text(position_number) or str(len(positions) + 1),
                "quantity": quantity(position_quantity),
                "unit": unit,
                "description": description,
                "price": money(position_price),
                "total": money(position_total),
            }
        )

    if not positions:
        raise ValueError("Keine Rechnungsposition in Zeile 23 bis 31 gefunden.")

    net_total = money(sum(Decimal(item["total"]) for item in positions))
    tax_total = money(Decimal(net_total) * Decimal(tax_rate) / Decimal("100"))
    gross_total = money(Decimal(net_total) + Decimal(tax_total))

    ws_edit.column_dimensions["Z"].hidden = True
    ws_edit.column_dimensions["Z"].width = 0.1
    ws_edit.sheet_view.showGridLines = False
    ws_edit.print_area = "A1:H52"
    ws_edit.page_setup.orientation = "portrait"
    ws_edit.page_setup.paperSize = ws_edit.PAPERSIZE_A4
    ws_edit.page_setup.fitToWidth = 1
    ws_edit.page_setup.fitToHeight = 1
    ws_edit.sheet_properties.pageSetUpPr.fitToPage = True
    ws_edit.page_margins.left = 0.25
    ws_edit.page_margins.right = 0.25
    ws_edit.page_margins.top = 0.35
    ws_edit.page_margins.bottom = 0.35
    ws_edit.print_options.horizontalCentered = True

    return wb_edit, {
        "id": invoice_id,
        "issue_date": issue_date,
        "buyer_name": buyer_name,
        "buyer_street": buyer_street,
        "buyer_postal": buyer_postal,
        "buyer_city": buyer_city,
        "order_id": order_id,
        "delivery_note": delivery_note,
        "period_start": period_start,
        "period_end": period_end,
        "payment_terms": payment_terms or "30 Tage netto",
        "tax_rate": tax_rate,
        "net_total": net_total,
        "tax_total": tax_total,
        "gross_total": gross_total,
        "positions": positions,
    }


def write_normalized_xlsx(workbook, invoice_id):
    WORK_DIR.mkdir(parents=True, exist_ok=True)

    target = WORK_DIR / f"{safe_filename(invoice_id)}.normalized.xlsx"
    temporary = WORK_DIR / f".{safe_filename(invoice_id)}.normalized.tmp.xlsx"

    workbook.save(temporary)
    temporary.replace(target)

    return target


def build_mapping(data, seller, payment, defaults):
    currency = text(defaults.get("currency")) or "EUR"

    mapping = {
        "meta": {
            "sectionColumn": {
                "Tabelle1": "Z",
            },
            "empty": ["[:empty:]"],
        },
        "ubl:Invoice": {
            "cbc:ID": data["id"],
            "cbc:IssueDate": data["issue_date"],
            "cbc:InvoiceTypeCode": "380",
            "cbc:DocumentCurrencyCode": currency,
            "cac:AccountingSupplierParty": {
                "cac:Party": {
                    "cac:PartyName": {
                        "cbc:Name": text(seller["name"]),
                    },
                    "cac:PostalAddress": {
                        "cbc:StreetName": text(seller["street"]),
                        "cbc:CityName": text(seller["city"]),
                        "cbc:PostalZone": text(seller["postalCode"]),
                        "cac:Country": {
                            "cbc:IdentificationCode": text(seller.get("countryCode")) or "DE",
                        },
                    },
                    "cac:PartyLegalEntity": {
                        "cbc:RegistrationName": text(seller["name"]),
                    },
                    "cac:PartyTaxScheme": {
                        "cbc:CompanyID": text(seller["vatId"]),
                        "cac:TaxScheme": {
                            "cbc:ID": "VAT",
                        },
                    },
                },
            },
            "cac:AccountingCustomerParty": {
                "cac:Party": {
                    "cac:PartyName": {
                        "cbc:Name": data["buyer_name"],
                    },
                    "cac:PostalAddress": {
                        "cbc:StreetName": data["buyer_street"],
                        "cbc:CityName": data["buyer_city"],
                        "cbc:PostalZone": data["buyer_postal"],
                        "cac:Country": {
                            "cbc:IdentificationCode": "DE",
                        },
                    },
                    "cac:PartyLegalEntity": {
                        "cbc:RegistrationName": data["buyer_name"],
                    },
                },
            },
            "cac:Delivery": {
                "cbc:ActualDeliveryDate": data["issue_date"],
            },
            "cac:PaymentMeans": {
                "cbc:PaymentMeansCode": "58",
                "cac:PayeeFinancialAccount": {
                    "cbc:ID": text(payment["iban"]),
                    "cbc:Name": text(payment["accountName"]),
                },
            },
            "cac:PaymentTerms": {
                "cbc:Note": data["payment_terms"],
            },
            "cac:TaxTotal": {
                "cbc:TaxAmount": data["tax_total"],
                "cbc:TaxAmount@currencyID": currency,
                "cac:TaxSubtotal": {
                    "cbc:TaxableAmount": data["net_total"],
                    "cbc:TaxableAmount@currencyID": currency,
                    "cbc:TaxAmount": data["tax_total"],
                    "cbc:TaxAmount@currencyID": currency,
                    "cac:TaxCategory": {
                        "cbc:ID": "S",
                        "cbc:Percent": data["tax_rate"],
                        "cac:TaxScheme": {
                            "cbc:ID": "VAT",
                        },
                    },
                },
            },
            "cac:LegalMonetaryTotal": {
                "cbc:LineExtensionAmount": data["net_total"],
                "cbc:LineExtensionAmount@currencyID": currency,
                "cbc:TaxExclusiveAmount": data["net_total"],
                "cbc:TaxExclusiveAmount@currencyID": currency,
                "cbc:TaxInclusiveAmount": data["gross_total"],
                "cbc:TaxInclusiveAmount@currencyID": currency,
                "cbc:PayableAmount": data["gross_total"],
                "cbc:PayableAmount@currencyID": currency,
            },
            "cac:InvoiceLine": {
                "section": ":InvoiceLine",
                "cbc:ID": "=:InvoiceLine.A1",
                "cbc:InvoicedQuantity": "=:InvoiceLine.B1",
                "cbc:InvoicedQuantity@unitCode": "=:InvoiceLine.C1",
                "cbc:LineExtensionAmount": "=:InvoiceLine.H1",
                "cbc:LineExtensionAmount@currencyID": currency,
                "cac:Item": {
                    "cbc:Name": "=:InvoiceLine.D1",
                    "cac:ClassifiedTaxCategory": {
                        "cbc:ID": "S",
                        "cbc:Percent": data["tax_rate"],
                        "cac:TaxScheme": {
                            "cbc:ID": "VAT",
                        },
                    },
                },
                "cac:Price": {
                    "cbc:PriceAmount": "=:InvoiceLine.G1",
                    "cbc:PriceAmount@currencyID": currency,
                },
            },
        },
    }

    if optional(payment.get("bic")):
        mapping["ubl:Invoice"]["cac:PaymentMeans"]["cac:PayeeFinancialAccount"][
            "cac:FinancialInstitutionBranch"
        ] = {
            "cbc:ID": text(payment["bic"]),
        }

    if data["order_id"]:
        mapping["ubl:Invoice"]["cac:OrderReference"] = {
            "cbc:ID": data["order_id"],
        }

    if data["delivery_note"]:
        mapping["ubl:Invoice"]["cac:DespatchDocumentReference"] = {
            "cbc:ID": data["delivery_note"],
        }

    if data["period_start"] or data["period_end"]:
        invoice_period = {}

        if data["period_start"]:
            invoice_period["cbc:StartDate"] = data["period_start"]

        if data["period_end"]:
            invoice_period["cbc:EndDate"] = data["period_end"]

        mapping["ubl:Invoice"]["cac:InvoicePeriod"] = invoice_period

    return mapping


def write_mapping(mapping, invoice_id):
    WORK_DIR.mkdir(parents=True, exist_ok=True)

    target = WORK_DIR / f"{safe_filename(invoice_id)}.mapping.yaml"
    temporary = WORK_DIR / f".{safe_filename(invoice_id)}.mapping.tmp.yaml"

    with temporary.open("w", encoding="utf-8") as handle:
        yaml.safe_dump(
            mapping,
            handle,
            allow_unicode=True,
            default_flow_style=False,
            sort_keys=False,
        )

    temporary.replace(target)
    return target


def render_xlsx_to_pdf(xlsx_path):
    pdf_path = WORK_DIR / f"{xlsx_path.stem}.pdf"
    command = [
        "libreoffice",
        "--headless",
        "--convert-to", "pdf",
        "--outdir", str(WORK_DIR),
        str(xlsx_path),
    ]
    result = subprocess.run(command, capture_output=True, text=True, timeout=REQUEST_TIMEOUT)
    if result.returncode != 0 or not pdf_path.exists():
        raise RuntimeError(
            f"LibreOffice-PDF-Export fehlgeschlagen (rc={result.returncode}): "
            f"{result.stdout} {result.stderr}"
        )
    return pdf_path


def call_einvoice(spreadsheet_path, mapping_path, invoice_id, pdf_path):
    endpoint = f"{API_URL}/api/invoice/create/{FORMAT}"

    with spreadsheet_path.open("rb") as spreadsheet, mapping_path.open("rb") as mapping, pdf_path.open("rb") as pdf:
        response = requests.post(
            endpoint,
            data={"lang": "de"},
            files={
                "spreadsheet": (
                    spreadsheet_path.name,
                    spreadsheet,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                ),
                "mapping": (
                    mapping_path.name,
                    mapping,
                    "application/x-yaml",
                ),
                "pdf": (
                    pdf_path.name,
                    pdf,
                    "application/pdf",
                ),
            },
            timeout=REQUEST_TIMEOUT,
        )

    if response.status_code != 201:
        raise RuntimeError(
            f"E-Invoice-EU antwortete mit HTTP {response.status_code}: {response.text}"
        )

    if not response.content.startswith(b"%PDF-"):
        raise RuntimeError("E-Invoice-EU lieferte keine PDF-Datei zurück.")

    output_name = f"Invoice_{safe_filename(invoice_id)}.pdf"
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    temporary = OUTPUT_DIR / f".{output_name}.tmp"
    target = OUTPUT_DIR / output_name

    temporary.write_bytes(response.content)

    if temporary.stat().st_size < 1024:
        temporary.unlink(missing_ok=True)
        raise RuntimeError("Erzeugte PDF ist ungewöhnlich klein.")

    temporary.replace(target)
    return target



def overlay_logo_on_pdf(pdf_path):
    if not LOGO_PATH.exists():
        raise FileNotFoundError(f"Logo-Datei fehlt: {LOGO_PATH}")

    temporary_overlay = WORK_DIR / f".{pdf_path.stem}.logo-overlay.pdf"
    temporary_output = WORK_DIR / f".{pdf_path.stem}.logo-applied.pdf"

    reader = PdfReader(str(pdf_path))
    if not reader.pages:
        raise RuntimeError("PDF enthält keine Seiten.")

    first_page = reader.pages[0]
    page_width = float(first_page.mediabox.width)
    page_height = float(first_page.mediabox.height)

    image = ImageReader(str(LOGO_PATH))
    image_width, image_height = image.getSize()

    logo_height = float(LOGO_HEIGHT_CM * Decimal(str(cm)))
    logo_width = logo_height * (float(image_width) / float(image_height))
    margin = float(LOGO_MARGIN_CM * Decimal(str(cm)))
    top_margin = float(LOGO_TOP_CM * Decimal(str(cm)))

    x = page_width - margin - logo_width
    y = page_height - top_margin - logo_height

    overlay_canvas = canvas.Canvas(
        str(temporary_overlay),
        pagesize=(page_width, page_height),
    )
    overlay_canvas.drawImage(
        image,
        x,
        y,
        width=logo_width,
        height=logo_height,
        mask="auto",
    )
    overlay_canvas.save()

    overlay_reader = PdfReader(str(temporary_overlay))
    first_page.merge_page(overlay_reader.pages[0])

    writer = PdfWriter()
    for page in reader.pages:
        writer.add_page(page)

    with temporary_output.open("wb") as handle:
        writer.write(handle)

    if temporary_output.stat().st_size < 1024:
        raise RuntimeError("Logo-Overlay erzeugte eine ungültige PDF.")

    shutil.copy2(temporary_output, pdf_path)

    temporary_overlay.unlink(missing_ok=True)
    temporary_output.unlink(missing_ok=True)

    logging.info(
        "Logo eingefügt: %s | Höhe: %s cm | Rand: %s cm",
        LOGO_PATH.name,
        LOGO_HEIGHT_CM,
        LOGO_MARGIN_CM,
    )


def archive_original(source_path, invoice_id):
    year = datetime.now().strftime("%Y")
    target_dir = ARCHIVE_DIR / year
    target_dir.mkdir(parents=True, exist_ok=True)

    target = target_dir / source_path.name

    if target.exists():
        target = target_dir / (
            f"{source_path.stem}_{safe_filename(invoice_id)}{source_path.suffix}"
        )

    shutil.move(str(source_path), str(target))
    return target


def move_to_error(source_path, exc):
    ERROR_DIR.mkdir(parents=True, exist_ok=True)

    target = ERROR_DIR / source_path.name
    if target.exists():
        suffix = datetime.now().strftime("%Y%m%d-%H%M%S")
        target = ERROR_DIR / f"{source_path.stem}_{suffix}{source_path.suffix}"

    if source_path.exists():
        shutil.move(str(source_path), str(target))

    error_file = target.with_suffix(target.suffix + ".error.txt")
    error_file.write_text(
        f"Zeit: {datetime.now().isoformat(timespec='seconds')}\n"
        f"Datei: {source_path.name}\n"
        f"Fehler: {type(exc).__name__}: {exc}\n",
        encoding="utf-8",
    )

    return target, error_file


def process(source_path):
    logging.info("Verarbeite: %s", source_path.name)

    seller, payment, defaults = load_stammdaten()
    workbook, data = read_and_normalize(source_path)

    normalized_xlsx = write_normalized_xlsx(workbook, data["id"])
    mapping_yaml = write_mapping(build_mapping(data, seller, payment, defaults), data["id"])

    rendered_pdf = render_xlsx_to_pdf(normalized_xlsx)
    overlay_logo_on_pdf(rendered_pdf)
    pdf = call_einvoice(normalized_xlsx, mapping_yaml, data["id"], rendered_pdf)
    archive = archive_original(source_path, data["id"])

    logging.info(
        "Erfolgreich verarbeitet: %s | PDF: %s | Archiv: %s",
        data["id"],
        pdf,
        archive,
    )


def main():
    for directory in (
        INPUT_DIR,
        OUTPUT_DIR,
        ARCHIVE_DIR,
        ERROR_DIR,
        MAPPING_DIR,
        WORK_DIR,
        LOG_DIR,
    ):
        directory.mkdir(parents=True, exist_ok=True)

    setup_logging()

    logging.info(
        "Watcher gestartet | input=%s | api=%s | format=%s",
        INPUT_DIR,
        API_URL,
        FORMAT,
    )

    while True:
        files = sorted(
            path
            for path in INPUT_DIR.glob("*.xlsx")
            if not path.name.startswith("~$")
        )

        for source_path in files:
            try:
                if not is_stable(source_path):
                    logging.info("Datei noch nicht stabil: %s", source_path.name)
                    continue

                process(source_path)

            except Exception as exc:
                logging.exception("Verarbeitung fehlgeschlagen: %s", source_path.name)

                try:
                    target, error_file = move_to_error(source_path, exc)
                    logging.error(
                        "Nach Fehler verschoben: %s | Details: %s",
                        target,
                        error_file,
                    )
                except Exception:
                    logging.exception(
                        "Fehlerdatei konnte nicht verschoben werden: %s",
                        source_path.name,
                    )

        time.sleep(POLL_SECONDS)


if __name__ == "__main__":
    main()
