# Catalog Ingestion Automation

This script supports your exact two-step workflow.

1. Step 1 (`format`): take Lodg_TN-style input, keep only _lodg columns, and set `extract_text=yes`.
2. Step 2 (`validate`): take formatted output and validation workbook, remove rows with `validation_status=Invalid` by matching `url`.

The _lodg columns are hardcoded in code, so you do not need to pass `_Lodg_CA` every run.

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
"/Users/dipesh.banerjee/WMS Files/URLingestion_Scripts/.venv/bin/python" ingest_catalog.py \
  format \
  --source "Lodg_TN_Tennessee.xlsx" \
  --output "Lodg_TN_Tennessee_formatted.xlsx"
```

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
"/Users/dipesh.banerjee/WMS Files/URLingestion_Scripts/.venv/bin/python" ingest_catalog.py \
  validate \
  --formatted "Lodg_TN_Tennessee_formatted.xlsx" \
  --validation "validation__Lodg_TN_Tennessee_formatted.xlsx"
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
