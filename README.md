# E-Invoice Workflow

## Ordnerstruktur

Die Arbeitsordner werden beim Start des Watchers automatisch unterhalb von `/data` angelegt:

```text
/data/input
/data/output
/data/archive
/data/error
/data/mapping
/data/work
/data/logs
```

Der Host muss lediglich den Ordner `data` nach `/data` mounten. Der Ordner `/data/assets` wird nicht automatisch erstellt, weil das Logo dort nur gelesen wird. Lege ihn zusammen mit dem Logo vorher an.

Diese Dateien müssen vor dem ersten Start bereitliegen:

```text
data/mapping/stammdaten.yaml
data/mapping/vorlage.xlsx
data/assets/logo.png
```

Für die lokale Vorbereitung genügt:

```bash
mkdir -p data/{mapping,assets}
```

Die übrigen Arbeitsordner erzeugt der Watcher selbst. Bei restriktiven NAS-Rechten können sie optional vorab angelegt werden:

```bash
mkdir -p data/{input,archive,output,error,mapping,assets,work,logs}
```

## Installation auf TrueNAS SCALE

Für TrueNAS SCALE steht `compose.truenas.yaml` im Repository bereit. Passe darin den Poolnamen im Host-Pfad an:

```yaml
- /mnt/tank/appdata/e-invoice-workflow/data:/data
```

Der Host-Pfad muss auf ein geeignetes Dataset oder Unterverzeichnis eines Datasets zeigen. Erstelle mindestens die Dataset-Struktur `appdata/e-invoice-workflow` und lege die drei erforderlichen Dateien unter `data` ab.

Installation über **Apps -> Discover Apps -> Install via YAML**:

1. Einen App-Namen in Kleinbuchstaben vergeben, z. B. `e-invoice-workflow`.
2. Den Inhalt von `compose.truenas.yaml` als YAML einfügen.
3. `/mnt/tank` durch den tatsächlichen Poolpfad ersetzen.
4. Die Anwendung speichern und starten.
5. Die App-Logs kontrollieren.

Die TrueNAS-Konfiguration verwendet ausschließlich veröffentlichte Images. Es gibt keinen lokalen Build und keinen Git-Clone. Der API-Port von `e-invoice-eu` wird nur innerhalb des Compose-Netzwerks verwendet.

## Synology

Für Synology Container Manager wird `compose.yaml` verwendet. Auch diese Konfiguration verwendet veröffentlichte Images und benötigt keinen Git-Clone oder lokalen Build.

## Dateien im Datenordner

```text
data/
├── input/
├── archive/
├── output/
├── error/
├── mapping/
│   ├── stammdaten.yaml
│   └── vorlage.xlsx
├── assets/
│   └── logo.png
├── work/
└── logs/
```

`vorlage.xlsx` ist die feste Mastervorlage. Für die Verarbeitung wird eine ausgefüllte Kopie mit beliebigem Namen nach `data/input` gelegt. Der Watcher verarbeitet nur `.xlsx`-Dateien direkt unter `input`.

## Dateiname der ausgegebenen PDF

### Syntax
OUTPUT_FILENAME_TEMPLATE: invoice_{order_id}_{invoice_id}.pdf

### Mögliche Variablen
{invoice_id} – Rechnungsnummer (aus Excel-Zelle H9)

{order_id} – Bestellnummer (aus Excel-Zelle A19, falls leer, wird der Platzhalter durch nichts ersetzt)

{delivery_note} – Lieferscheinnummer (aus Excel-Zelle H10)

{issue_date} – Rechnungsdatum im ISO-Format YYYY-MM-DD (aus Excel-Zelle H12)

{buyer_name} – Name/Firma des Käufers (aus Excel-Zelle A8)

