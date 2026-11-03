# CLAUDE.md

Monorepo of independent Fraser Health staffing/HR tools. `PROMPT.MD` is the source of truth; `PLAN.md` tracks tasks, bugs and pending decisions.

## Projects

- `Audit-Processor/` – Streamlit app: Kronos/Power BI audits + HR Data Toolbox lookup enrichment. Run `streamlit run audit_app.py`. Tests: `python -m pytest Audit-Processor/tests`.
- `Outlook-Export/` – Windows Tkinter + Outlook COM exporter with SQLite history (`outlook_export.db`). Core logic in `exporter_core/` (testable on any OS). Tests: `python -m pytest Outlook-Export/tests`.
- `RFP-Utility-Application/` – PySide6 daily-use invoice/RFP app. Tests: `python -m pytest RFP-Utility-Application/tests` (Windows-only modules are stubbed).
- `Macro-Vision/` – Windows macro recorder. No automated tests.
- `HR Data Toolbox Web v1.3/` – legacy; functionality now lives in Audit-Processor. Kept for rollback.

## Rules

- Never delete, reset or overwrite `Outlook-Export/outlook_export.db`, its backups, or export folders. Test against copies in a temp dir.
- DB schema changes must be additive, transactional, preceded by a backup, and tested on a copy.
- RFP: do not change output file names, PDF field mapping, tracker columns or folder layout without owner approval. Add regression tests first.
- Do not commit databases, exports, logs, HR data, `.env`, or build output. Existing tracked artifacts stay until the owner approves untracking (PLAN.md D1).
- Consult the owner before changing an established output or workflow; continue with unrelated safe work.
- Don't rewrite history or force-push. Focused commits per stage; update PLAN.md status.

## Testing

Use a venv: `pip install -r <project>/requirements-dev.txt` where present, then run pytest per project. Outlook COM, Excel COM and Qt UI cannot run on Linux CI — test logic with mocks and validate on the Windows target (see each README).
