# SAP Account Movement Validator

A read-only batch validator that checks, for a list of account
identifier pairs, whether each account is active in SAP and whether
it has a specific movement type — flagging accounts that are active
but missing that expected movement as candidates for manual review.

> **Note:** this is a sanitized portfolio copy. Company-specific
> identifiers (file paths, company codes) have been replaced with
> environment variables or generic placeholders — see
> `.env.example`. No real business data is included anywhere in this
> repository.

## Problem

A reconciliation team needed to validate, for a large batch of
account identifiers read from a spreadsheet, whether each account had
a specific "initial load" movement type — a manual process that meant
re-entering each identifier into SAP one by one, hundreds of times
per batch.

## What this does

For every identifier in the input CSV, three flags are derived:

1. **is_active** — the account has any movements at all.
2. **has_target_type** — filtering by the target movement type,
   at least one matching row exists.
3. **has_other_only** — the account is active but has *no* movement
   of the target type, meaning it has some other movement type
   instead. This flag requires no extra SAP query — it's derived
   purely from the first two.

## Safety

This script is **read-only**. It never confirms or writes anything
back to SAP — it only navigates screens, applies a view filter, and
reads what's rendered on screen. There is no write path anywhere in
this codebase.

## Tech notes

- **SAP GUI Scripting** via COM automation (`pywin32`).
- **Recursive screen-text extraction**: rather than reading a fixed
  coordinate or control ID, a helper walks every text-bearing control
  under the current screen and searches the combined text for a
  "no data found" message — resilient to minor layout shifts between
  SAP versions or environments.
- **Incremental CSV checkpointing**: progress is saved to disk every
  20 rows, so a mid-run failure (SAP hang, network hiccup) doesn't
  lose all prior work.
- **Per-record error isolation**: each identifier check is wrapped in
  its own `try/except`, so one problematic record doesn't abort the
  whole batch.

## Requirements

```
pip install -r requirements.txt
```

Requires an already-open, logged-in SAP GUI session with GUI
Scripting enabled.

## Configuration

Copy `.env.example` to `.env` and fill in your own values (input/
output CSV paths, SAP company code, target movement type, the "no
data" message(s) your SAP GUI shows in its logon language). The script
loads `.env` automatically on startup via `python-dotenv` — no manual
exporting needed.

## Input CSV layout

Column positions (0-indexed, configurable via the `COL_*` constants
in the script) are expected as:

| # | Column | # | Column |
|---|--------|---|--------|
| 0 | ID | 5 | HasOtherOnly1 |
| 1 | Name | 6 | Identifier2 |
| 2 | Identifier1 | 7 | IsActive2 |
| 3 | IsActive1 | 8 | HasType2 |
| 4 | HasType1 | 9 | HasOtherOnly2 |

## Usage

```bash
python sap_account_validator_portfolio.py
```

