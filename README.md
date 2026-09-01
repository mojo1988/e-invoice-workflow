# E-Invoice Workflow

Automatischer Workflow zur Erstellung von ZUGFeRD-Rechnungen aus Excel-Rechnungsvorlagen. Der Watcher überwacht einen Eingangsordner, verarbeitet neue Rechnungen über den Dienst `e-invoice-eu`, archiviert die Excel-Quelldatei und legt die fertige PDF im Ausgabeordner ab.

## Verbindliche Vorlage

Der Watcher unterstützt ausschließlich Kopien der verbindlichen Excel-Mastervorlage `vorlage.xlsx`. Eine beliebige Excel-Datei funktioniert nicht: Die Zellzuordnung ist im Watcher fest implementiert.

Die Mastervorlage wird lokal auf der NAS unter `data/mapping/vorlage.xlsx` abgelegt. Für jede Rechnung wird diese Datei kopiert, mit einem eigenen Namen gespeichert und die Kopie in `data/input` gelegt, beispielsweise `Rechnung-2026-001.xlsx`.

**Nicht verändern:** Tabellenblatt `Tabelle1`, Zellpositionen, Spalten, Zeilen, Formeln und Formatierung der Vorlage müssen unverändert bleiben.

### Festes Mapping

| Inhalt | Excel-Zelle/Bereich |
|---|---|
| Tabellenblatt | `Tabelle1` |
| Käuferfirma | `A7` |
| Käuferstraße | `A9` |
| Käufer-PLZ und Ort | `A10`, z. B. `D-90552 Röthenbach` |
| Rechnungsnummer | `H8` |
| Lieferscheinnummer | `H9` |
| Rechnungsdatum | `H11` |
| Auftragsnummer | `A18` |
| Leistungszeitraum Beginn | `E16` |
| Leistungszeitraum Ende | `E18` |
| Positionen | Zeilen `23–31` |
| Positionsnummer | Spalte `A` |
| Menge | Spalte `B` |
| Einheit | Spalte `C` |
| Bezeichnung | Spalte `D` |
| Preis pro Einheit | Spalte `G` |
| Positionswert | Spalte `H` |
| Umsatzsteuersatz | `C34` |
| Zahlungsbedingungen | `B38:D38` |

Verkäufer-, Bank- und Standarddaten stehen in `data/mapping/stammdaten.yaml`. Das Firmenlogo muss als `data/assets/logo.png` vorliegen.

## Installation auf Synology

Voraussetzung ist Synology Container Manager. Die Compose-Datei bindet den NAS-Ordner `/volume1/docker/e-invoice-workflow/data` im Container nach `/data` ein. Lege auf der NAS daher exakt diese Ordner an:

```text
/volume1/docker/e-invoice-workflow/data/
├── input/
├── archive/
├── output/
├── error/
├── mapping/
├── assets/
├── work/
└── logs/
```

Lege anschließend die installationsspezifischen Dateien unter diesen Pfaden ab:

```text
/volume1/docker/e-invoice-workflow/data/mapping/vorlage.xlsx
/volume1/docker/e-invoice-workflow/data/mapping/stammdaten.yaml
/volume1/docker/e-invoice-workflow/data/assets/logo.png
```

`vorlage.xlsx`, `stammdaten.yaml` und `logo.png` gehören nicht in ein öffentliches Repository, wenn sie Firmen-, Kunden-, Kontakt- oder Bankdaten enthalten. Sie müssen vor dem ersten Start lokal auf der NAS angelegt oder kopiert werden.

## Compose-YAML

Die Datei `compose.yaml` im Repository ist direkt für Synology Container Manager nutzbar: Sie verwendet ausschließlich veröffentlichte Images und enthält keinen `build:`-Abschnitt. Es wird auf der NAS weder ein Git-Clone noch ein lokaler Docker-Build benötigt.

```yaml
services:
  e-invoice-eu:
    image: gflohr/e-invoice-eu:latest
    container_name: e-invoice-eu
    restart: unless-stopped
    environment:
      NODE_ENV: production

  watcher:
    image: moj008/e-invoice-watcher:latest
    container_name: e-invoice-watcher
    restart: unless-stopped
    depends_on:
      - e-invoice-eu
    environment:
      API_URL: http://e-invoice-eu:3000
      INPUT_DIR: /data/input
      OUTPUT_DIR: /data/output
      ARCHIVE_DIR: /data/archive
      ERROR_DIR: /data/error
      MAPPING_DIR: /data/mapping
      WORK_DIR: /data/work
      LOG_DIR: /data/logs
      LOGO_PATH: /data/assets/logo.png
      FORMAT: Factur-X-EN16931
      SHEET_NAME: Tabelle1
      POLL_SECONDS: "10"
      STABLE_SECONDS: "10"
      REQUEST_TIMEOUT: "180"
    volumes:
      - /volume1/docker/e-invoice-workflow/data:/data
```

Die Compose-Datei startet zwei Container:

| Container | Aufgabe |
|---|---|
| `e-invoice-eu` | API-Dienst für die technische Erstellung der E-Rechnung. |
| `e-invoice-watcher` | Überwacht `input` und steuert Verarbeitung, Archivierung und Ausgabe. |

Die Dienste kommunizieren intern über `http://e-invoice-eu:3000`; eine externe Portfreigabe ist nicht erforderlich.

## Betrieb

1. In Container Manager ein Projekt anlegen und die `compose.yaml` aus diesem Repository importieren oder ihren Inhalt einfügen.
2. Prüfen, dass `/volume1/docker/e-invoice-workflow/data` vorhanden ist und alle Unterordner enthält.
3. Vor dem Start `vorlage.xlsx`, `stammdaten.yaml` und `logo.png` an den oben genannten NAS-Pfaden ablegen.
4. Das Projekt starten und warten, bis `e-invoice-eu` sowie `e-invoice-watcher` den Status `Running` haben.
5. Eine ausgefüllte Kopie von `vorlage.xlsx` nach `data/input` legen.
6. Die fertige PDF erscheint direkt unter `data/output`.
7. Erfolgreich verarbeitete Excel-Dateien werden nach `data/archive/JJJJ` verschoben. Fehlerhafte Dateien und Details landen in `data/error`.

## Aktualisieren

Für eine neue Watcher-Version das aktuelle Image `moj008/e-invoice-watcher:latest` in Container Manager herunterladen und das Projekt neu erstellen oder neu starten. Die eingebundenen Datenordner bleiben dabei erhalten.

## Entwicklung

Der Python-Watcher liegt unter `watcher/watcher.py`. GitHub Actions baut auf `main` ein Multi-Architektur-Image für `linux/amd64` und `linux/arm64` und veröffentlicht es auf Docker Hub.