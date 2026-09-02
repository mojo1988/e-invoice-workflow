# Feature Request: Separate SMB workspace and PDF archiving

## Background

The workflow is intended to run on a NAS. Users should receive access through an SMB share, but they should not need to see configuration, mapping, assets, logs, work files, or container-related directories.

## Requested changes

### 1. Archive generated PDFs

For each successfully processed invoice, archive both source files:

- the original Excel invoice
- the generated PDF invoice

Use separate subdirectories, preferably organized by year:

```text
/data/archive/excel/<YYYY>/Original.xlsx
/data/archive/pdf/<YYYY>/Invoice_<invoice-number>.pdf
```

The existing output behavior may remain for compatibility, but the generated PDF must additionally be copied or moved into the PDF archive after successful creation and logo overlay.

### 2. Separate user-facing workspace

Move the user-facing directories into one dedicated subdirectory, for example:

```text
/data/share/input
/data/share/output
/data/share/archive/excel
/data/share/archive/pdf
```

This `share` directory is intended to be exposed as an SMB share. Users should only see the folders needed for their daily work.

The following directories should remain outside the SMB-facing workspace:

```text
/data/mapping
/data/assets
/data/work
/data/logs
```

The configuration should allow the share root to be changed using an environment variable, for example:

```text
SHARE_DIR=/data/share
```

### 3. SMB-oriented visibility

The container must not require users to access or modify:

- `mapping/stammdaten.yaml`
- `mapping/vorlage.xlsx`
- `assets/logo.png`
- temporary files in `work`
- watcher logs in `logs`

The README and both Compose examples should document a host layout where only the share directory is published through SMB.

## Acceptance criteria

- A processed Excel invoice is archived below `archive/excel`.
- Its generated PDF is archived below `archive/pdf`.
- Input, output, and archive are below one configurable share directory.
- Mapping, assets, work, and logs are outside the share directory.
- Existing processing behavior and Docker image usage remain compatible.
- The feature is implemented consistently for Synology and TrueNAS.
