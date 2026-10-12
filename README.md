# Rechnungseinreichung bei der muki – in einem Aufruf

Füllt das offizielle Formular **„Ansuchen um Kostenrückerstattung"** der
österreichischen Zusatzversicherung *muki* (AcroForm-PDF) programmatisch aus,
benennt die Anhänge einheitlich und erzeugt einen fertigen Mail-Entwurf
(`Einreichung.eml`). Es wird **nichts automatisch versendet**.

Ein kleines Werkzeug, das drei Dinge zeigt, die im Alltag oft mühsam sind:
**PDF-Formulare füllen**, **Bilder in PDFs umwandeln** und **Mails als Entwurf
ablegen** – alles ohne Framework, nur mit der Python-Stdlib plus `pypdf`,
`reportlab` und `pillow`.

## Ablauf

1. Arztrechnung (Wahlarzt/Honorarnote) **zuerst** bei der ÖGK über die
   Meine-ÖGK-App einreichen (Teilrückerstattung).
2. **Erst danach** bei muki einreichen: ausgefülltes Formular + Rechnungskopie +
   ÖGK-Bestätigung über den geleisteten Kostenersatz, per E-Mail an
   `leistung-kranken@muki.com`.

Das Tool übernimmt nur Schritt 2. Die ÖGK-Bestätigung ist Pflicht, weil muki nur
die Differenz zum ÖGK-Anteil erstattet.

## Voraussetzungen

- Python 3.10+
- `pypdf` (6.x), `reportlab`, `pillow`
- optional `poppler-utils` (`pdftoppm`) für das Vorschaubild
- Internetzugang nur beim **ersten** Lauf, um das Formular zu laden

```bash
python3 -m venv .venv && source .venv/bin/activate
pip install "pypdf>=6,<7" reportlab pillow
```

## Einrichten

### `profile.json` (Stammdaten)

`profile.example.json` kopieren und die Platzhalter ersetzen. Die Datei wird
beim Lauf automatisch auf `chmod 600` gesetzt.

```bash
cp profile.example.json profile.json
```

```json
{
  "polizzennummer": "0000000",
  "kundennummer": "000000000",
  "name": "Max Mustermann",
  "strasse": "Musterstraße 12",
  "plz_ort": "1010 Wien",
  "telefon": "+43 664 0000000",
  "email": "max.mustermann@example.com",
  "kontoinhaber": "Max Mustermann",
  "bank": "Beispielbank",
  "iban": "AT00 0000 0000 0000 0000",
  "bic": "",
  "signature": "unterschrift.png",
  "signature_text": "Max Mustermann"
}
```

`signature` (optional) wird relativ zu `profile.json` aufgelöst; per
`--signature` kann ein anderer Pfad angegeben werden. Ist keine Bilddatei
vorhanden, wird mit `signature_text` (bzw. `--signature-text`) der Name als
Textunterschrift auf die Unterschriftszeile gesetzt.

### Falldatei (`falldaten.json`)

Vorlage unter `faelle/vorlage/falldaten.json`:

```json
{
  "datum": "2026-10-07",
  "art_des_versicherungsschutzes": "Wahlarzt",
  "versicherte_personen": ["Max Mustermann"],
  "diagnose": ["Akute Bronchitis (J20.9)"],
  "leistungen": ["Facharztkonsultation", "Honorarnote Nr. 2026-4711"],
  "unterlagen": ["Arztrechnung", "ÖGK-Bestätigung"],
  "datum_formular": "07.10.2026",
  "rechnungen": ["rechnung_1.pdf"],
  "oegk_bestaetigungen": ["oegk_bestaetigung_1.pdf"]
}
```

- `art_des_versicherungsschutzes`: `"Krankenversicherung"` oder `"Wahlarzt"`
  (setzt die passende Checkbox).
- `rechnungen` / `oegk_bestaetigungen`: Pfade **relativ zur Falldatei** oder
  absolut. Bilder (`.jpg`, `.png`, …) werden automatisch gedreht und nach PDF
  konvertiert.
- `leistungen`: max. 4 Einträge; `diagnose`: max. 4; `unterlagen`: zwei Zeilen
  (alles ab dem 2. Eintrag wird in Zeile 2 zusammengefasst).
- Optional `offene_rechnung: true` + `krankenanstalt: [...]`, wenn stattdessen
  direkt die Krankenanstalt bezahlt werden soll.
- Optional `ueberweisung: false`, falls das Konto nicht angekreuzt werden soll.

## Aufruf

```bash
python3 einreichen.py --case faelle/<person>/<fall>/falldaten.json
```

Optionen: `--profile PFAD`, `--out PFAD`, `--no-preview`,
`--signature-text "Vorname Nachname"`, `--gmail-draft`,
`--case -` (liest die Falldaten von stdin).

## Ergebnis: `out/<datum>/`

| Datei | Inhalt |
|---|---|
| `Ansuchen_Kostenrueckerstattung_<Name>.pdf` | ausgefülltes Formular |
| `Rechnung_1_<originalname>.pdf` | Rechnungskopie |
| `OEGK_Bestaetigung_1_<originalname>.pdf` | ÖGK-Bestätigung |
| `Einreichung.eml` | fertiger Mail-Entwurf (An, Betreff, Text, alle PDFs) |
| `Vorschau_Ansuchen.png` | gerenderte Seite 1 zur Kontrolle |

**Die `.eml` im Mailprogramm öffnen, prüfen und selbst senden.** Das Tool sendet
nie.

## Gmail-Entwurf (optional)

Mit `--gmail-draft` wird die Mail zusätzlich per IMAP als **Entwurf** in Gmail
abgelegt (Ordner mit `\Draft`-Flag, z. B. `[Gmail]/Entwürfe`). Es wird
ausschließlich `APPEND` mit `\Draft` benutzt – **kein Versand, kein SMTP**.

Dafür in `profile.json` einen `gmail`-Block ergänzen und ein Gmail-App-Passwort
hinterlegen – entweder in der Umgebungsvariable `GMAIL_APP_PASSWORD` oder im
System-Keyring:

```bash
secret-tool store --label="Gmail App-Passwort" service muki-gmail \
  account max.mustermann@example.com
```

## Wie es funktioniert

- Das Formular (**Version 12/2025**) hat feste AcroForm-Feldnamen; sie wurden mit
  `pypdf.get_fields()` ermittelt und stehen als Konstanten oben im Skript.
  Bei einer neuen Formularversion: Feldnamen prüfen und dort anpassen.
- **Adress-Mapping:** Die fünf Felder im Formular heißen irreführend alle
  `Vorund Nachname des Versicherungsnehmers 1…5`, liegen physisch aber auf den
  Zeilen Name / Straße / PLZ-Ort / Telefon / E-Mail. Das Skript füllt sie
  entsprechend.
- Das Formular wird in `clone_from`-Manier gefüllt (AcroForm bleibt erhalten),
  `NeedAppearances` wird gesetzt; die Unterschrift wird per `merge_page` als
  Overlay eingebracht.
- Zu lange Texte werden automatisch verkleinert (das Tool warnt).

## Beispieldaten

Unter `beispiel/` liegen `falldaten.json`, `unterschrift.png`, `rechnung_1.pdf`
und `oegk_bestaetigung_1.pdf` (reine Testdokumente):

```bash
python3 einreichen.py --case beispiel/falldaten.json
```

## Excel-Übersicht

`uebersicht.py` liest alle `falldaten.json` unter `faelle/` (außer `vorlage`),
schreibt `out/Uebersicht_muki.xlsx` neu und übernimmt manuelle Spalten
(muki-Erstattung, Status, Notizen) aus einer bestehenden Tabelle.

```bash
python3 uebersicht.py
```
