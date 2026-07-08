# Catalog Ingestion Automation

## Overview

**Catalog Ingestion Automation** is a two-step utility for preparing WMS catalog files across different states.

It solves the repetitive cleanup problem in a simple flow: first format a raw Lodg file into the required _lodg structure, then apply validation later to remove invalid rows. This avoids manual column cleanup, keeps output consistent, and lets validation happen when the validation sheet becomes available.

No template file is required per run. The _lodg output columns are already hardcoded in the script, so files like _Lodg_CA are not needed each time.

### Key Capabilities

- Run 1 (`format`): accepts a Lodg-style input file, keeps only required _lodg columns, and sets `extract_text=yes`.
- Run 2 (`validate`): accepts the formatted file and validation workbook, then removes rows where `validation_status=Invalid` using `url` matching.

## Hardcoded Output Columns

The formatter always writes these columns in this order:

- `tax_rule_id`
- `country`
- `state`
- `jur_name`
- `jur_id`
- `jur_type`
- `imposition_name`
- `imposition_type_name`
- `cat_id`
- `cat_name`
- `content_type`
- `rule_type`
- `url_type`
- `url`
- `skip_tool_running`
- `tag`
- `comment`
- `last_monitoring_date`
- `extract_text`

## Prerequisites

- Python 3.10+
- `openpyxl`

## Run 1: Format (No Validation Needed)

```bash
  .venv/bin/python ingest_catalog.py format --source 'UnFormatted_Excel_sheet.xlsx' ;
```

This command only needs the input Excel file.

Defaults:

- Output defaults to `<source>_formatted.xlsx` if `--output` is not provided.
- Source sheet defaults to `Library` if present, else first sheet.
- `extract_text` value defaults to `yes`.

Useful flags:

- `--source-sheet`
- `--extract-text-value`
- `--dry-run`

## Run 2: Validate (After You Receive Validation File)

```bash
.venv/bin/python ingest_catalog.py validate --formatted 'Formatted_Excel_sheet.xlsx' --validation 'validation_Formatted_Excel_sheet.xlsx' ;
```

Defaults:

- Output defaults to `<formatted>_validated.xlsx`.
- Validation rule defaults to `validation_status = Invalid`.
- Match key defaults to `url`.

Useful flags:

- `--output`
- `--source-sheet`
- `--validation-sheet`
- `--url-column`
- `--validation-status-column`
- `--invalid-value`
- `--url-case-insensitive`
- `--dry-run`

## Output Files

Each non-dry run writes:

- Output workbook (`*_formatted.xlsx` or `*_validated.xlsx`)
- Audit file next to output (`.audit.json`)
