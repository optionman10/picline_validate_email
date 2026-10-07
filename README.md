# Email Validator (Excel)

A desktop tool that validates email addresses in an Excel file and reports
how many are good, bad, or suspicious — with a reason for every problem.

![flow](https://img.shields.io/badge/python-3.10%2B-blue)

## What it checks

| Check | Severity | Example |
|---|---|---|
| Missing / multiple `@`, empty value | ❌ Error | `not-an-email` |
| Illegal characters, bad domain shape, bad TLD | ❌ Error | `jo hn@exa mple.com` |
| Leading/trailing/consecutive dots in local part | ❌ Error | `john..doe@…` |
| Local part > 64 chars / total > 254 chars | ❌ Error | |
| Disposable / temporary email domain | ❌ Error | `…@mailinator.com` |
| Domain has no MX record (DNS, cached) | ❌ Error | `…@nonexistent.com` |
| Typo in a popular domain + suggested fix | ⚠️ Warning | `…@gmial.com → gmail.com` |
| Generic/role mailbox | ⚠️ Warning | `info@`, `noreply@` |
| All-digits or random-looking local part | ⚠️ Warning | `123456@…`, `xk7qz9w2bp4@…` |
| Leading/trailing whitespace | ⚠️ Warning | `"  a@b.com  "` |

Everything not flagged is **✅ Valid**.

## Install

```bash
pip install -r requirements.txt
```

## Run

```bash
python email_validator_gui.py
```

1. **Browse…** for an `.xlsx` / `.xlsm` file — the tool auto-detects the
   sheet and column containing emails (looks for a header like *Email*,
   otherwise scans for `@`-shaped values).
2. Tick/untick **DNS domain check** (needs internet; safe to turn off for
   very large lists — it's cached but still the slowest step).
3. Click **▶ Validate** — a progress bar shows live progress; the UI never
   freezes (validation runs in a background thread).
4. See the **summary** (Valid / Warnings / Invalid) and filter the results
   table. Double-click a row to copy the email + its error description.
5. Click **💾 Export results to Excel** — creates
   `<name>_validated.xlsx` next to the source with three new columns
   (*Validation Status*, *Error Description*, *Suggested Correction*);
   invalid rows are highlighted red, warnings yellow.

## Files

| File | Purpose |
|---|---|
| `email_validator.py` | Core validation engine (pure logic, unit-testable) |
| `email_validator_gui.py` | tkinter GUI + Excel read/export |
| `test_email_validator.py` | 29 unit tests (`python -m unittest`) |
| `sample_emails.xlsx` | Generated sample input (delete anytime) |

## Tests

```bash
python -m unittest test_email_validator -v
```
