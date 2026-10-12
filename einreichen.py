#!/usr/bin/env python3
"""Rechnungseinreichung bei der muki (oesterreichische Zusatzversicherung).

Reduziert die Einreichung auf einen Aufruf:

    python3 einreichen.py --case falldaten.json [--signature unterschrift.png]

Ablauf:
  1. Das ausfuellbare muki-Formular "Ansuchen um Kostenrueckerstattung" wird
     einmalig heruntergeladen und als ansuchen_formular.pdf neben dieses Skript
     gelegt.
  2. Stammdaten kommen aus profile.json (wird mit chmod 600 geschuetzt).
  3. Falldaten kommen aus der --case Datei.
  4. Das AcroForm-Formular wird programmatisch gefuellt (Feldnamen sind fest
     verdrahtet, siehe unten - ermittelt mit pypdf get_fields()).
  5. Optional wird die Unterschrift auf die Zeile "Unterschrift:" gesetzt -
     als PNG mit Transparenz oder, falls kein Bild vorliegt, als Textname.
  6. Es entsteht out/<datum>/ mit dem gefuellten Ansuchen, einheitlich benannten
     Anhaengen und einem fertigen Mail-Entwurf Einreichung.eml.

Es wird NICHTS automatisch versendet. Der .eml-Entwurf wird selbst im
Mailprogramm geoeffnet und gesendet.
"""

from __future__ import annotations

import argparse
import imaplib
import json
import mimetypes
import os
import re
import shutil
import subprocess
import sys
import time
import urllib.request
from datetime import date, datetime
from email.message import EmailMessage
from pathlib import Path

# --------------------------------------------------------------------------- #
# Pfade / Konstanten
# --------------------------------------------------------------------------- #

SCRIPT_DIR = Path(__file__).resolve().parent
FORM_URL = (
    "https://www.muki.com/fileadmin/media/pdf/Gesundheit/FamilyPlus/"
    "Ansuchen_u_Kostenrueckerstattung_10.pdf"
)
FORM_FILE = SCRIPT_DIR / "ansuchen_formular.pdf"
PROFILE_FILE = SCRIPT_DIR / "profile.json"
OUT_ROOT = SCRIPT_DIR / "out"

MAIL_TO = "leistung-kranken@muki.com"

# Bild-Anhaenge (z. B. abfotografierte Honorarnote) werden automatisch nach PDF
# konvertiert und dabei anhand der EXIF-Orientierung korrekt gedreht.
IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".tif", ".tiff", ".webp", ".bmp"}

# Checkboxen sind /Btn mit den Zustaenden /On und /Off. pypdf erwartet fuer den
# aktiven Zustand exakt "/On" (Schluessel im Appearance-Dictionary /AP /N).
CHECKED = "/On"
UNCHECKED = "/Off"

# Fest verdrahtete AcroForm-Feldnamen (Version 12/2025).
# Ermittelt mit: PdfReader(pfad).get_fields()
FIELD_POLIZZE = "Polizzennummer"
FIELD_KUNDEN = "Kundennummer"
FIELD_KRANKENVERS = "Krankenversicherung"
FIELD_WAHLARZT = "Wahlarzt"
# Achtung: Der Formular-Designer hat die fuenf Felder physisch auf die Zeilen
# Name / Strasse / PLZ-Ort / Telefon / E-Mail gelegt, sie heissen aber alle
# "Vorund Nachname des Versicherungsnehmers N".
FIELD_NAME = "Vorund Nachname des Versicherungsnehmers 1"
FIELD_STRASSE = "Vorund Nachname des Versicherungsnehmers 2"
FIELD_PLZ = "Vorund Nachname des Versicherungsnehmers 3"
FIELD_TELEFON = "Vorund Nachname des Versicherungsnehmers 4"
FIELD_EMAIL = "Vorund Nachname des Versicherungsnehmers 5"
FIELDS_DIAGNOSE = [f"Diagnose {i}" for i in range(1, 5)]
FIELDS_LEISTUNGEN = [
    f"Es wird um Abrechnung folgender Leistungen gebeten {i}" for i in range(1, 5)
]
FIELD_UEBERWEISUNG = (
    "Um Überweisung des leistungspflichtigen Betrages auf folgendes Konto wird gebeten"
)
FIELD_KONTOINHABER = "Kontoinhaber"
FIELD_BANK = "Bank"
FIELD_IBAN = "IBAN"
FIELD_BIC = "BIC"
FIELD_KRANKENANSTALT = (
    "Um Bezahlung der offenen Rechnung bei folgender Krankenanstalt wird gebeten"
)
FIELD_ANSTALT_1 = "1"
FIELD_ANSTALT_2 = "2"
FIELD_UNTERLAGEN = ["Folgende Unterlagen liegen bei", "Folgende Unterlagen liegen bei 2"]
FIELD_DATUM = "Datum"

# Position der Unterschriftszeile, als Anteil der Seitengroesse (A4 Querformat-
# unabhaengig). Gemessen an Formular-Version 12/2025: Linie von ~352 bis ~524 pt
# bei y ~66 pt (Seite 595.32 x 841.92 pt).
SIGN_X0_FRAC = 0.591
SIGN_X1_FRAC = 0.880
SIGN_BOTTOM_FRAC = 0.0722  # Unterkante der Linie
SIGN_MAX_H_FRAC = 0.036    # maximale Bildhoehe
SIGN_TEXT_SIZE = 15        # Schriftgroesse der Text-Unterschrift


# --------------------------------------------------------------------------- #
# Hilfsfunktionen
# --------------------------------------------------------------------------- #

def log(msg: str) -> None:
    print(msg, flush=True)


def warn(msg: str) -> None:
    print(f"  [WARNUNG] {msg}", flush=True)


def sanitize(name: str) -> str:
    """Dateinamen-taugliche Variante eines Strings."""
    name = name.strip()
    name = re.sub(r"[^\w.\- ]+", "_", name, flags=re.UNICODE)
    name = re.sub(r"\s+", "_", name)
    return name.strip("_") or "Unbenannt"


def load_json(path: Path, required: bool = True) -> dict:
    if not path.exists():
        if required:
            raise SystemExit(f"Fehler: {path} nicht gefunden.")
        return {}
    with path.open("r", encoding="utf-8") as fh:
        return json.load(fh)


def ensure_form() -> None:
    """Laedt das Formular einmalig herunter, falls nicht vorhanden."""
    if FORM_FILE.exists():
        return
    log(f"Formular wird einmalig geladen:\n  {FORM_URL}")
    req = urllib.request.Request(FORM_URL, headers={"User-Agent": "muki-einreichen/1.0"})
    try:
        with urllib.request.urlopen(req, timeout=60) as resp:
            data = resp.read()
    except Exception as exc:  # noqa: BLE001
        raise SystemExit(f"Fehler: Formular konnte nicht geladen werden: {exc}")
    if not data.startswith(b"%PDF"):
        raise SystemExit("Fehler: Heruntergeladene Datei ist kein PDF.")
    FORM_FILE.write_bytes(data)
    log(f"  gespeichert als {FORM_FILE.name} ({len(data)} Bytes)")


# --------------------------------------------------------------------------- #
# Fall-/Formulardaten
# --------------------------------------------------------------------------- #

def format_datum(value: str | None) -> str:
    if not value:
        return date.today().strftime("%d.%m.%Y")
    for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%d.%m.%y"):
        try:
            return datetime.strptime(value, fmt).strftime("%d.%m.%Y")
        except ValueError:
            continue
    return value


def _fill_list(values, field_names, out: dict, label: str, warnings: list[str]) -> None:
    """Verteilt eine Liste auf fortlaufend nummerierte Formularfelder."""
    if isinstance(values, str):
        values = [values]
    values = [v for v in (values or []) if v]
    if len(values) > len(field_names):
        warnings.append(
            f"{label}: {len(values)} Eintraege, aber nur {len(field_names)} Felder "
            f"vorhanden - ueberzaehlige werden ignoriert."
        )
    for field, value in zip(field_names, values):
        out[field] = str(value)


def build_field_values(profile: dict, case: dict) -> tuple[dict, list[str]]:
    """Baut das Feldname->Wert Mapping fuer das Formular."""
    values: dict[str, str] = {}
    warnings: list[str] = []

    def put(field: str, key: str, required: bool = False) -> None:
        val = profile.get(key)
        if val in (None, ""):
            if required:
                warnings.append(f"profile.json: '{key}' fehlt ({field}).")
            return
        values[field] = str(val)

    put(FIELD_POLIZZE, "polizzennummer", required=True)
    put(FIELD_KUNDEN, "kundennummer")
    put(FIELD_NAME, "name", required=True)
    put(FIELD_STRASSE, "strasse")
    put(FIELD_PLZ, "plz_ort")
    put(FIELD_TELEFON, "telefon")
    put(FIELD_EMAIL, "email")
    put(FIELD_KONTOINHABER, "kontoinhaber")
    put(FIELD_BANK, "bank")
    put(FIELD_IBAN, "iban")
    put(FIELD_BIC, "bic")

    # Art des Versicherungsschutzes (Checkbox)
    art = (case.get("art_des_versicherungsschutzes") or "").strip().lower()
    if art in ("krankenversicherung", "kranken"):
        values[FIELD_KRANKENVERS] = CHECKED
    elif art in ("wahlarzt", "wahlärztin", "wahlarzt/honorarnote"):
        values[FIELD_WAHLARZT] = CHECKED
    elif art:
        warnings.append(
            f"Unbekannte Art des Versicherungsschutzes: {art!r} "
            "(erwartet 'Krankenversicherung' oder 'Wahlarzt')."
        )

    _fill_list(case.get("diagnose"), FIELDS_DIAGNOSE, values, "Diagnose", warnings)
    _fill_list(case.get("leistungen"), FIELDS_LEISTUNGEN, values, "Leistungen", warnings)

    # Beiliegende Unterlagen: zwei Zeilen. Alles ab dem zweiten Eintrag wird in
    # Zeile 2 zusammengefasst.
    unterlagen = case.get("unterlagen") or []
    if isinstance(unterlagen, str):
        unterlagen = [unterlagen]
    unterlagen = [u for u in unterlagen if u]
    if unterlagen:
        values[FIELD_UNTERLAGEN[0]] = str(unterlagen[0])
        if len(unterlagen) > 1:
            values[FIELD_UNTERLAGEN[1]] = "; ".join(str(u) for u in unterlagen[1:])

    # Konto / offene Rechnung
    ueberweisung = case.get("ueberweisung")
    if ueberweisung is None:
        ueberweisung = bool(profile.get("iban") or profile.get("kontoinhaber"))
    values[FIELD_UEBERWEISUNG] = CHECKED if ueberweisung else UNCHECKED

    if case.get("offene_rechnung"):
        values[FIELD_KRANKENANSTALT] = CHECKED
        _fill_list(
            case.get("krankenanstalt"),
            [FIELD_ANSTALT_1, FIELD_ANSTALT_2],
            values,
            "Krankenanstalt",
            warnings,
        )
    else:
        values[FIELD_KRANKENANSTALT] = UNCHECKED

    values[FIELD_DATUM] = format_datum(case.get("datum_formular") or case.get("datum"))
    return values, warnings


# --------------------------------------------------------------------------- #
# PDF fuellen
# --------------------------------------------------------------------------- #

def make_signature_overlay(
    sig_path: Path | None, sig_text: str | None, page_w: float, page_h: float
):
    """Erzeugt eine Overlay-Seite mit der Unterschrift (reportlab).

    Entweder als Bild (PNG mit Transparenz) oder als Text (falls keine
    Bilddatei vorliegt).
    """
    from io import BytesIO

    from pypdf import PdfReader
    from reportlab.pdfgen import canvas

    box_x0 = SIGN_X0_FRAC * page_w
    box_w = (SIGN_X1_FRAC - SIGN_X0_FRAC) * page_w
    box_h = SIGN_MAX_H_FRAC * page_h

    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=(page_w, page_h))

    if sig_path is not None:
        from PIL import Image
        from reportlab.lib.utils import ImageReader

        iw, ih = Image.open(sig_path).size
        ar = iw / ih if ih else 1.0
        w = min(box_w, box_h * ar)
        h = w / ar
        c.drawImage(
            ImageReader(str(sig_path)),
            box_x0,
            SIGN_BOTTOM_FRAC * page_h,
            width=w,
            height=h,
            mask="auto",
            preserveAspectRatio=True,
            anchor="sw",
        )
    elif sig_text:
        from reportlab.pdfbase import pdfmetrics

        font_name = "Helvetica-Oblique"
        font_path = Path("/usr/share/fonts/liberation/LiberationSerif-Italic.ttf")
        try:
            from reportlab.pdfbase.ttfonts import TTFont

            if font_path.exists():
                pdfmetrics.registerFont(TTFont("SignFont", str(font_path)))
                font_name = "SignFont"
        except Exception:  # noqa: BLE001 - Schrift ist optional
            font_name = "Helvetica-Oblique"
        size = SIGN_TEXT_SIZE
        while (
            size > 8
            and pdfmetrics.stringWidth(sig_text, font_name, size) > box_w
        ):
            size -= 0.5
        c.setFont(font_name, size)
        c.setFillColorRGB(0.05, 0.05, 0.35)
        # Grundlinie knapp ueber der Unterschriftslinie
        c.drawString(box_x0, SIGN_BOTTOM_FRAC * page_h + 2, sig_text)

    c.showPage()
    c.save()
    buf.seek(0)
    return PdfReader(buf)


def warn_if_too_long(form_path: Path, values: dict, warned: list[str]) -> None:
    """Warnt, wenn ein Text fuer sein (schmales) Formularfeld zu lang ist.

    pypdf verkleinert zu langen Text automatisch, bis er passt - das kann
    unleserlich werden. Vor allem das Feld der ersten Leistung ist schmal
    (es liegt auf der Zeile der Ueberschrift).
    """
    from pypdf import PdfReader

    reader = PdfReader(str(form_path))
    widths: dict[str, float] = {}
    for annot in reader.pages[0].get("/Annots") or []:
        o = annot.get_object()
        if o.get("/FT") == "/Tx" and o.get("/T") is not None:
            try:
                rect = [float(x) for x in o["/Rect"]]
            except (TypeError, ValueError):
                continue
            widths[str(o["/T"])] = abs(rect[2] - rect[0])
    for name, value in values.items():
        width = widths.get(name)
        if width is None or not isinstance(value, str):
            continue
        capacity = width / 4.6  # grobe Naeherung (Helvetica ~0.5em bei 9pt)
        if len(value) > capacity * 1.15:
            warned.append(
                f"Langtext in Feld {name!r} ({len(value)} Zeichen, ca. "
                f"{int(capacity)} passen) - wird evtl. verkleinert dargestellt."
            )


def fill_form(
    form_path: Path,
    out_path: Path,
    values: dict,
    sig_path: Path | None,
    sig_text: str | None = None,
) -> None:
    """Fuellt das AcroForm-Formular und schreibt das Ergebnis.

    Wichtig: Der Writer wird per clone_from aus dem Reader erzeugt, damit das
    /AcroForm-Dictionary (und damit die Feldwerte) erhalten bleibt. Die
    Unterschrift wird als Overlay per merge_page eingebracht.
    """
    from pypdf import PdfReader, PdfWriter

    reader = PdfReader(str(form_path))
    writer = PdfWriter(clone_from=reader)

    writer.set_need_appearances_writer(True)
    writer.update_page_form_field_values(
        writer.pages[0], values, auto_regenerate=False
    )

    if sig_path is not None or sig_text:
        page = writer.pages[0]
        w = float(page.mediabox.width)
        h = float(page.mediabox.height)
        overlay = make_signature_overlay(sig_path, sig_text, w, h)
        page.merge_page(overlay.pages[0])

    writer.set_need_appearances_writer(True)
    with out_path.open("wb") as fh:
        writer.write(fh)


# --------------------------------------------------------------------------- #
# Anhaenge / Ordner / Mail
# --------------------------------------------------------------------------- #

def collect_files(raw_list, label: str, warnings: list[str],
                  base_dir: Path) -> list[Path]:
    result: list[Path] = []
    if isinstance(raw_list, str):
        raw_list = [raw_list]
    for raw in raw_list or []:
        if not raw:
            continue
        p = Path(os.path.expanduser(str(raw)))
        if not p.is_absolute():
            p = (base_dir / p).resolve()
        if not p.exists():
            warnings.append(f"{label}: Datei fehlt: {p}")
            continue
        result.append(p)
    return result


def _image_to_pdf(src: Path, dst: Path) -> None:
    """Konvertiert ein (ggf. abfotografiertes) Bild zu einer PDF-Seite."""
    from PIL import Image, ImageOps

    im = ImageOps.exif_transpose(Image.open(src))
    if im.mode in ("RGBA", "LA", "P"):
        base = Image.new("RGB", im.size, "white")
        mask = im.convert("RGBA").split()[-1]
        base.paste(im.convert("RGBA"), mask=mask)
        im = base
    elif im.mode != "RGB":
        im = im.convert("RGB")
    im.save(dst, "PDF", resolution=150)


def _copy_one(src: Path, dst_pdf: Path) -> Path:
    """Kopiert eine Datei; Bilder werden nach PDF konvertiert."""
    if src.suffix.lower() in IMAGE_EXTS:
        dst = dst_pdf.with_suffix(".pdf")
        _image_to_pdf(src, dst)
        return dst
    dst = dst_pdf.with_name(dst_pdf.name + src.suffix)
    shutil.copy2(src, dst)
    return dst


def copy_attachments(case: dict, out_dir: Path, warnings: list[str],
                     base_dir: Path) -> list[Path]:
    copied: list[Path] = []
    rechnungen = collect_files(case.get("rechnungen"), "Rechnung", warnings, base_dir)
    oegk = collect_files(
        case.get("oegk_bestaetigungen"), "ÖGK-Bestätigung", warnings, base_dir
    )
    for i, p in enumerate(rechnungen, 1):
        base = out_dir / f"Rechnung_{i}_{sanitize(p.stem)}"
        copied.append(_copy_one(p, base))
    for i, p in enumerate(oegk, 1):
        base = out_dir / f"OEGK_Bestaetigung_{i}_{sanitize(p.stem)}"
        copied.append(_copy_one(p, base))
    return copied


def build_mail(
    profile: dict, case: dict, attachments: list[Path], out_path: Path
) -> None:
    name = profile.get("name", "")
    polizze = profile.get("polizzennummer", "")
    art = case.get("art_des_versicherungsschutzes", "")
    personen = case.get("versicherte_personen") or []
    if isinstance(personen, str):
        personen = [personen]
    diagnose = case.get("diagnose") or []
    if isinstance(diagnose, str):
        diagnose = [diagnose]

    lines = [
        "Sehr geehrte Damen und Herren,",
        "",
        "anbei ersuche ich um Kostenrückerstattung im Rahmen meines "
        "Versicherungsschutzes.",
        "",
        f"Polizzennummer: {polizze}",
        f"Name: {name}",
    ]
    if art:
        lines.append(f"Art des Versicherungsschutzes: {art}")
    if personen:
        lines.append(f"Versicherte Person(en): {', '.join(str(p) for p in personen)}")
    if diagnose:
        lines.append(f"Diagnose: {', '.join(str(d) for d in diagnose)}")
    lines += [
        "",
        "Beiliegende Unterlagen:",
    ]
    for a in attachments:
        lines.append(f"  - {a.name}")
    lines += [
        "",
        "Die Arztrechnung wurde bereits bei der ÖGK zur Teilrückerstattung "
        "eingereicht; die ÖGK-Bestätigung über den geleisteten Kostenersatz liegt "
        "bei.",
        "",
        "Mit freundlichen Grüßen",
        name,
    ]

    msg = EmailMessage()
    msg["To"] = MAIL_TO
    sender = profile.get("absender_email") or profile.get("email")
    if sender:
        msg["From"] = sender
    msg["Subject"] = (
        f"Ansuchen um Kostenrückerstattung – Polizzennummer {polizze} – {name}"
    )
    msg.set_content("\n".join(lines))

    for path in attachments:
        data = path.read_bytes()
        ctype, _ = mimetypes.guess_type(path.name)
        maintype, subtype = (
            ctype.split("/", 1) if ctype and "/" in ctype
            else ("application", "octet-stream")
        )
        msg.add_attachment(
            data,
            maintype=maintype,
            subtype=subtype,
            filename=path.name,
        )

    out_path.write_bytes(msg.as_bytes())


# --------------------------------------------------------------------------- #
# Gmail-Entwurf (nur Entwurf - es wird NIE gesendet)
# --------------------------------------------------------------------------- #

def _gmail_config(profile: dict) -> dict:
    return profile.get("gmail") or {}


def _gmail_app_password(cfg: dict) -> str:
    pw = os.environ.get("GMAIL_APP_PASSWORD")
    if pw:
        return pw.strip()
    try:
        out = subprocess.run(
            ["secret-tool", "lookup", "service", cfg["secret_service"],
             "account", cfg["address"]],
            check=True, capture_output=True, text=True,
        )
    except (OSError, KeyError, subprocess.CalledProcessError):
        out = None
    if out and out.stdout.strip():
        return out.stdout.strip()
    raise SystemExit(
        "Kein Gmail-App-Passwort gefunden. Bitte hinterlegen:\n"
        f"  secret-tool store --label=\"Gmail App-Passwort\" service "
        f"{cfg.get('secret_service', 'muki-gmail')} account "
        f"{cfg.get('address', '')}"
    )


def _find_drafts_folder(conn: imaplib.IMAP4_SSL, fallback: str) -> str:
    try:
        _typ, folders = conn.list()
        for line in folders or []:
            text = line.decode(errors="replace") if isinstance(line, bytes) else str(line)
            if "\\Draft" in text:
                return text.rsplit(' "', 1)[-1].rstrip('"')
    except imaplib.IMAP4.error:
        pass
    return fallback


def upload_gmail_draft(eml_bytes: bytes, profile: dict) -> str:
    """Legt die Mail als ENTWURF im Gmail-Drafts-Ordner ab. Kein Versand."""
    cfg = _gmail_config(profile)
    if not cfg:
        raise SystemExit(
            "Keine Gmail-Konfiguration gefunden. Bitte in profile.json einen "
            "'gmail'-Block ergaenzen."
        )
    password = _gmail_app_password(cfg)
    conn = imaplib.IMAP4_SSL(cfg.get("imap_host", "imap.gmail.com"),
                            cfg.get("imap_port", 993))
    try:
        conn.login(cfg["address"], password)
        folder = _find_drafts_folder(conn, cfg.get("drafts_folder", "[Gmail]/Drafts"))
        typ, resp = conn.append(
            folder, r"(\Draft)", imaplib.Time2Internaldate(time.time()), eml_bytes
        )
        if typ != "OK":
            raise SystemExit(f"Gmail-Entwurf fehlgeschlagen: {resp}")
        return folder
    finally:
        try:
            conn.logout()
        except imaplib.IMAP4.error:
            pass


def render_preview(pdf_path: Path, out_png: Path) -> bool:
    exe = shutil.which("pdftoppm")
    if not exe:
        return False
    try:
        subprocess.run(
            [exe, "-r", "110", "-png", "-f", "1", "-l", "1",
             str(pdf_path), str(out_png.with_suffix(""))],
            check=True, capture_output=True,
        )
    except subprocess.CalledProcessError:
        return False
    # pdftoppm haengt "-1" an den Dateinamen an
    produced = out_png.with_name(out_png.stem + "-1.png")
    if produced.exists() and produced != out_png:
        produced.replace(out_png)
    return out_png.exists()


# --------------------------------------------------------------------------- #
# Main
# --------------------------------------------------------------------------- #

def parse_args(argv=None):
    ap = argparse.ArgumentParser(
        description="Rechnung bei der muki einreichen (Ansuchen um "
                    "Kostenrückerstattung automatisch ausfuellen)."
    )
    ap.add_argument("--case", required=True,
                    help="Falldaten-JSON (oder '-' fuer stdin).")
    ap.add_argument("--signature", default=None,
                    help="PNG mit Transparenz fuer die Unterschrift.")
    ap.add_argument("--signature-text", default=None,
                    help="Name, der als Textunterschrift eingesetzt wird "
                         "(falls kein Bild vorliegt).")
    ap.add_argument("--profile", default=str(PROFILE_FILE),
                    help="Pfad zur profile.json (Default: neben dem Skript).")
    ap.add_argument("--out", default=str(OUT_ROOT),
                    help="Zielordner fuer die Ausgabe (Default: out/).")
    ap.add_argument("--no-preview", action="store_true",
                    help="Kein Vorschaubild rendern.")
    ap.add_argument("--gmail-draft", action="store_true",
                    help="Die Mail zusaetzlich als ENTWURF in Gmail ablegen "
                         "(kein Versand).")
    return ap.parse_args(argv)


def main(argv=None) -> int:
    args = parse_args(argv)

    ensure_form()
    profile_path = Path(os.path.expanduser(args.profile))
    if not profile_path.is_absolute():
        profile_path = (Path.cwd() / profile_path).resolve()
    profile = load_json(profile_path)
    try:
        os.chmod(profile_path, 0o600)
    except OSError:
        pass

    if args.case == "-":
        case = json.load(sys.stdin)
        case_dir = Path.cwd()
    else:
        case_path = Path(os.path.expanduser(args.case))
        if not case_path.is_absolute():
            case_path = (Path.cwd() / case_path).resolve()
        case = load_json(case_path)
        case_dir = case_path.parent

    warnings: list[str] = []
    values, value_warnings = build_field_values(profile, case)
    warnings.extend(value_warnings)

    datum = sanitize(str(case.get("datum") or date.today().isoformat()))
    out_dir = Path(os.path.expanduser(args.out))
    if not out_dir.is_absolute():
        out_dir = (Path.cwd() / out_dir).resolve()
    out_dir = out_dir / datum
    out_dir.mkdir(parents=True, exist_ok=True)

    # Unterschrift
    sig_raw = args.signature or profile.get("signature")
    sig_text = args.signature_text or profile.get("signature_text")
    sig_path: Path | None = None
    if sig_raw:
        cand = Path(os.path.expanduser(str(sig_raw)))
        if not cand.is_absolute():
            cand = (profile_path.parent / cand).resolve()
        if cand.exists():
            sig_path = cand
        else:
            warnings.append(f"Unterschrift nicht gefunden: {cand}")

    name = sanitize(profile.get("name", "Unbekannt"))
    form_out = out_dir / f"Ansuchen_Kostenrueckerstattung_{name}.pdf"
    warn_if_too_long(FORM_FILE, values, warnings)
    fill_form(FORM_FILE, form_out, values, sig_path, sig_text)

    attachments = [form_out]
    attachments += copy_attachments(case, out_dir, warnings, case_dir)

    eml_out = out_dir / "Einreichung.eml"
    build_mail(profile, case, attachments, eml_out)

    if not args.no_preview:
        preview = out_dir / "Vorschau_Ansuchen.png"
        if not render_preview(form_out, preview):
            warnings.append(
                "Vorschaubild konnte nicht erstellt werden (poppler-utils/pdftoppm fehlt?)."
            )

    log("")
    log("Einreichung vorbereitet:")
    log(f"  Ordner:   {out_dir}")
    log(f"  Ansuchen: {form_out.name}")
    for a in attachments[1:]:
        log(f"            {a.name}")
    log(f"  Mail:     {eml_out.name}  (NICHT automatisch versendet)")
    if args.gmail_draft:
        folder = upload_gmail_draft(eml_out.read_bytes(), profile)
        log(f"  Gmail:    Entwurf abgelegt in {folder}  (NICHT versendet)")
    if warnings:
        log("")
        log("Hinweise:")
        for w in warnings:
            warn(w)
    log("")
    log("Naechster Schritt: Einreichung.eml im Mailprogramm oeffnen und senden.")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except ImportError as exc:  # fehlende Abhaengigkeiten freundlich melden
        print(
            f"\nFehlende Abhaengigkeit: {exc.name}\n\n"
            "Bitte die mitgelieferte virtuelle Umgebung nutzen:\n"
            f"  source {SCRIPT_DIR / '.venv' / 'bin' / 'activate'}\n"
            f"  python3 {Path(__file__).name} --case falldaten.json\n\n"
            "Oder die Pakete installieren:  pip install 'pypdf>=6,<7' reportlab pillow\n"
        )
        raise SystemExit(1)
