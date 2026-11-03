# Audit Prepper (with HR Data Toolbox)

One local Streamlit app with three pages:

- **Home** – pick a tool; clear all uploaded data.
- **Kronos audits** – guided 3-step flow (select → upload & configure → review & run) for: 43/843 on GRTW, GRTW shift mismatch, OT work rule with 5.2 LCT, Unprocessed TORs, Casuals with LOA 1##.1, Negative accrual bank, NBA OC > 72 h, CTO/OTO/STO cascade to LOA, and PC Portfolio Enrichment Only. Optional Outlook drafts.
- **HR Data Toolbox** – enrich a working Excel/CSV file from any number of lookup files (merged from the former `HR Data Toolbox Web v1.3` app, removed in favour of this page; same rules and output).

## Setup and run

```bash
python -m pip install -r requirements.txt
streamlit run audit_app.py
```

Windows: run `install.bat` once, then `run.bat`. Outlook drafts need Windows, classic Outlook and `pywin32`.

## Inputs and outputs

- Inputs: CSV, XLSX, XLSM, XLS exports (≤200 MB each). The header row is detected within the first 50 rows; column names match case- and whitespace-insensitively. See `AUDIT_SPEC_TEMPLATE.md` to request a new audit.
- Audit output: `[SUS] Audit - <audit name> YYYY-MM-DD.xlsx` with a single **Audit Results** sheet. Run notes (portfolio fallbacks, duplicate department paths, unmatched rows, skipped blank rows) are shown on screen and in the downloadable **run report CSV**, not in the workbook.
- HR output: `HR_Enriched_YYYY-MM-DD_HHMM.xlsx` (`Enriched Data`, plus optional report sheets). Text starting with `=` is stored as text, never as a formula.

## Reliability

- Each audit runs in isolation: a failing audit is reported as failed and the others still finish. Per tool: rows read, findings, blank rows skipped; overall: succeeded / with warnings / failed.
- Run buttons are disabled while a run is in progress.

## Temporary data policy

Uploads stay in the browser session's memory and are never written to disk. Parsed reference tables are cached in memory for ≤1 hour. Outlook attachments are written to `%TEMP%\AuditPrepper`, deleted once Outlook saves the draft; anything older than 24 h there is purged on next use. **Clear all uploaded data now** (Home) empties everything immediately. Streamlit usage telemetry is disabled in `.streamlit/config.toml`.

## Layout

```
audit_app.py              entry point (page wiring only)
audit_suite/config.py     constants, URLs, coordinator map
audit_suite/common/       text helpers, upload validation/reading, Excel writers
audit_suite/audits/       rules, leave cascade, registry, runner
audit_suite/portfolio.py  PC Portfolio matching
audit_suite/outlook_drafts.py, audit_suite/hr_toolbox/engine.py
audit_suite/ui/           Streamlit pages
tests/                    pytest, synthetic fixtures only
```

## Tests

```bash
python -m pip install -r requirements-dev.txt
python -m pytest tests
```

`tests/golden/audit_legacy.json` holds findings produced by the pre-refactor app; `test_parity_audits.py` proves the refactor matches them. `test_hr_parity.py` compared the merged HR engine with the original v1.3 engine before that folder was removed; it now skips (restore the folder from git history to re-run it).

## Limitations

- Source formulas are read as cached values (outputs are new workbooks).
- Outlook draft creation is mocked in tests; verify on the target PC by creating one draft per audit type.
- A semicolon-delimited CSV with title rows above the header may be misread (unchanged from before); export as comma CSV or XLSX.
