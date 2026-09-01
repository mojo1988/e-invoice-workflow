# E-Invoice Workflow

Automatischer Workflow zur Erstellung von ZUGFeRD-Rechnungen aus Excel-Rechnungsvorlagen. Der Watcher überwacht einen Eingangsordner, verarbeitet neue Rechnungen über den Dienst `e-invoice-eu`, archiviert die Excel-Quelldatei und legt die fertige PDF im Ausgabeordner ab.

## Installation auf einer Synology NAS

Diese Anleitung verwendet Synology Container Manager und Docker Compose. Der Watcher wird als Multi-Architektur-Image bereitgestellt und kann auf unterstützten Intel/AMD- (`amd64`) sowie 64-Bit-ARM-Systemen (`arm64`) verwendet werden.

### Voraussetzungen

- Synology DSM mit installiertem **Container Manager**
- Ein Benutzerkonto mit Schreibrechten auf die verwendete Freigabe
- Die persönlichen Rechnungsdateien: `stammdaten.yaml`, `rechnungsvorlage.xlsx` und das Firmenlogo

### Ordner anlegen

Lege in File Station – beispielhaft in der Freigabe `docker` auf Volume 1 – diese Struktur an:

```text
/docker/e-invoice-workflow/
├── inbox/
├── archive/
├── output/
├── mapping/
├── assets/
└── logs/
```

Die zugrunde liegenden Container-Pfade lauten bei diesem Beispiel:

```text
/volume1/docker/e-invoice-workflow/
```

Passe `/volume1/docker` in der Compose-Datei an, wenn deine Freigabe oder dein Volume anders heißt.

### Persönliche Dateien ablegen

Diese Dateien gehören nicht in GitHub und müssen lokal auf der NAS abgelegt werden:

```text
stammdaten.yaml
  → /volume1/docker/e-invoice-workflow/mapping/stammdaten.yaml

rechnungsvorlage.xlsx
  → /volume1/docker/e-invoice-workflow/mapping/rechnungsvorlage.xlsx

<firmenlogo>.png oder <firmenlogo>.jpg
  → /volume1/docker/e-invoice-workflow/assets/
```

Die Ordner `mapping` und `assets` werden im Container schreibgeschützt eingebunden, damit Stammdaten, Vorlage und Logo nicht durch den Container verändert werden können.

### Compose-YAML

Öffne **Container Manager → Projekt → Erstellen** und wähle die Erstellung über eine Compose-/YAML-Datei. Verwende diese Konfiguration:

```yaml
services:
  e-invoice-eu:
    image: ghcr.io/itplr/e-invoice-eu:latest
    container_name: e-invoice-eu
    restart: unless-stopped

  e-invoice-watcher:
    image: moj008/e-invoice-watcher:latest
    container_name: e-invoice-watcher
    restart: unless-stopped

    depends_on:
      - e-invoice-eu

    environment:
      TZ: Europe/Berlin
      PYTHONUNBUFFERED: "1"
      E_INVOICE_API_URL: http://e-invoice-eu:8080

    volumes:
      - /volume1/docker/e-invoice-workflow/inbox:/app/inbox
      - /volume1/docker/e-invoice-workflow/archive:/app/archive
      - /volume1/docker/e-invoice-workflow/output:/app/output
      - /volume1/docker/e-invoice-workflow/mapping:/app/mapping:ro
      - /volume1/docker/e-invoice-workflow/assets:/app/assets:ro
      - /volume1/docker/e-invoice-workflow/logs:/app/logs
```

Die Compose-Datei startet zwei Container:

| Container | Aufgabe |
|---|---|
| `e-invoice-eu` | Stellt die API für die technische Erstellung der E-Rechnung bereit. |
| `e-invoice-watcher` | Überwacht den Eingangsordner und steuert Verarbeitung, Archivierung und Ausgabe. |

Der Watcher erreicht die API innerhalb des Compose-Netzwerks über `http://e-invoice-eu:8080`. Deshalb ist keine Portfreigabe auf der NAS erforderlich.

### Projekt starten

1. Vergib in Container Manager einen Projektnamen, beispielsweise `e-invoice-workflow`.
2. Wähle einen Projektordner oder füge die YAML-Datei direkt ein.
3. Prüfe, dass die Pfade unter `volumes` zu deiner Synology passen.
4. Klicke auf **Erstellen** bzw. **Starten**.
5. Warte, bis beide Container den Status **Running** anzeigen.

Container Manager lädt beim ersten Start die Images herunter. Beim Watcher wird automatisch die zur NAS-CPU passende `amd64`- oder `arm64`-Image-Variante verwendet.

### Funktion testen

1. Lege eine ausgefüllte Excel-Rechnung in `inbox` ab.
2. Prüfe anschließend:
   - Die Quelldatei wurde nach `archive` verschoben.
   - Eine ZUGFeRD-PDF liegt direkt im Ordner `output`.
   - Bei Problemen enthalten die Dateien unter `logs` sowie die Container-Protokolle in Container Manager Hinweise.

### Aktualisieren

Um eine neue Watcher-Version zu beziehen, halte das Projekt in Container Manager an, ziehe das aktuelle Image `moj008/e-invoice-watcher:latest` und starte das Projekt erneut. Alternativ kann das Projekt erneut bereitgestellt werden; vorhandene Daten in den eingebundenen NAS-Ordnern bleiben erhalten.

## Entwicklung

Der Watcher-Code liegt unter `watcher/`. Die GitHub-Actions-Datei `.github/workflows/docker-publish.yml` baut bei Änderungen auf `main` ein Multi-Architektur-Image und veröffentlicht es auf Docker Hub.
