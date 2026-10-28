# Invoice Status Checker

A desktop workflow that checks invoice status in an authenticated portal and updates a workbook.

**Python, Playwright, openpyxl, Tkinter**

## What it does

- Loads invoice targets from the expected workbook sheets.
- Queries status and rejection reason through a browser session.
- Writes results back and exposes progress through the desktop UI.

## Workflow

```mermaid
flowchart LR
  N0["Workbook"] --> N1
  N1["Portal consultation"] --> N2
  N2["Status and reason"] --> N3
  N3["Updated results"]
```

## Local setup and use

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python main.py
```
Python 3.11+ is recommended. Tkinter requires a desktop Python installation.
Run `python -m playwright install chromium` before the first browser run. Set `INVOICE_LOGIN_URL` and `INVOICE_LIST_URL` to your authorized test portal. The UI collects session settings; its saved JSON is ignored by Git.

## Configuration

Environment names declared by the included examples (values must be supplied locally):

`INVOICE_LIST_URL`, `INVOICE_LOGIN_URL`.

## Boundaries and limitations

Portal selectors and workbook columns are installation-specific. Saved local configuration may contain credentials and must be protected. No portal call or real invoice was used in this review.

This is a source snapshot, not a claim of production readiness. Business results, dates and unmeasured performance are intentionally not asserted. Application UI language is preserved.

## License

No new distribution license is granted for this corporate source edition. Public visibility does not grant reuse rights; obtain permission from the relevant rights holder. Existing third-party notices remain applicable.

## Verification record

See [snapshot checks and remaining limitations](docs/VERIFICATION.md).

## Offline verification

`python -m unittest discover -s tests -v` checks duplicate-invoice grouping, completed-row exclusion, result propagation, note preservation and rejection of invalid headers with synthetic temporary workbooks. It does not open a browser or contact a portal.
