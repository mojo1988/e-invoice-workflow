# E-Invoice Workflow

Docker-basierter Workflow zur Erstellung von Factur-X/ZUGFeRD-Rechnungen aus Excel-Rechnungsvorlagen.

> Status: Quellcode- und Docker-Konfiguration werden vor der ersten Docker-Hub-Veröffentlichung final geprüft. Keine echten Rechnungen, Logos oder Stammdaten in dieses Repository legen.

## Funktionsweise

1. Ausgefüllte XLSX in `data/input` ablegen.
2. Der Watcher prüft und normalisiert die Rechnung.
3. LibreOffice erzeugt die sichtbare PDF.
4. Das Logo wird auf der ersten Seite platziert.
5. E-Invoice-EU bettet die strukturierte Factur-X-/ZUGFeRD-XML ein.
6. Die fertige Datei wird in `data/output` abgelegt, z. B. `Invoice_HSP-2026-14704.pdf`.
7. Die Original-XLSX wird archiviert.

## Synology-Installation

Voraussetzungen: Docker/Container Manager und Docker Compose.

Ordner anlegen:

```text
data/input
data/output
data/archive
data/error
data/mapping
data/assets
data/work
data/logs
```

Dateien bereitstellen:

- `data/mapping/stammdaten.yaml` aus `data/mapping/stammdaten.example.yaml` erstellen und mit echten Stammdaten füllen.
- Logo als `data/assets/hsp-logo.png` ablegen.
- Diese `compose.yaml` verwenden.

Start:

```bash
docker compose pull
docker compose up -d
```

Logs:

```bash
docker compose logs -f watcher
```

## Docker Hub

Das Watcher-Image wird über GitHub Actions gebaut und auf Docker Hub unter `moj008/e-invoice-watcher` veröffentlicht. Zugangsdaten niemals in Dateien speichern. GitHub Repository Secrets:

- `DOCKERHUB_USERNAME`
- `DOCKERHUB_TOKEN`

## Sicherheit

Echte Rechnungen, PDFs, Logos, Bankdaten, Steuerdaten und `stammdaten.yaml` gehören nicht ins Repository. Der zuvor im Chat geteilte Docker-Hub-Token muss widerrufen sein.
