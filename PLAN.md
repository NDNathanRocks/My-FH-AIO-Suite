# PLAN

Source of truth: `PROMPT.MD`. Status legend: [x] done · [~] partial · [ ] open · [?] needs owner decision.

## Project boundaries

| Project | Stack | Entry point | Sensitivity |
|---|---|---|---|
| `Audit-Processor` | Streamlit web app | `streamlit run audit_app.py` | Merge target for HR Data Toolbox. No Audit Summary sheet in output. |
| `HR Data Toolbox Web v1.3` | Streamlit web app | `run.bat` | Merged into Audit-Processor; original kept for rollback. |
| `Outlook-Export` | Tkinter + Outlook COM (Windows) | `inbox_email_exporter_v1_5_2.py` / `dist\InboxEmailExporter.exe` | Existing DB (6,108 records) + export folders must be preserved. |
| `RFP-Utility-Application` | PySide6 desktop (Windows) | `python main.py` | Daily workflow; outputs/filenames/columns must not change. |
| `Macro-Vision` | customtkinter + pynput (Windows) | `vision_suite_pro_updated.py` | Standalone; low priority. |

## Active task list

- [x] Stage 0 – Inspect repo, PLAN.md, CLAUDE.md, root `.gitignore`.
- [ ] Stage 1 – Audit-Processor: modular package, tests, remove Audit Summary sheet, merge HR toolbox.
- [ ] Stage 2 – Outlook-Export: core module, manifest + reconciliation + retry, safe migration, tests.
- [ ] Stage 3 – RFP: regression tests for transformations, safe fixes only.
- [ ] Stage 4 – Macro-Vision + per-project READMEs + final validation sweep.

## Feature-parity checklist (HR Data Toolbox → Audit-Processor)

_To be filled in Stage 1._

## Bug / risk register

| # | Project · location | Problem | Sev | Impact | Fix | Compat risk | Status |
|---|---|---|---|---|---|---|---|
| 1 | Repo root | DBs, logs, audit xlsx, build/, dist/*.exe, `__pycache__` are tracked in Git | High | HR/email metadata in history | Ignore going forward; untracking needs approval (see D1) | `git pull` would delete local copies | [?] |

## Decisions requiring approval

- **D1 – Untrack committed data/artifacts.** `Outlook-Export/*.db`, `backup/*.db`, `logs/*`, `build/`, `dist/*.exe`, `__pycache__/`, `Macro-Vision/*.json|*.txt` are tracked. `git rm --cached` would make every other clone *delete those files on next pull*. Recommend: back up the files on each machine, then untrack in one commit; purge history separately if required. Rollback: `git revert` the commit. Not done yet.

## Completed work

**Stage 0.** Inspected all five projects, recorded boundaries and risks, added root `CLAUDE.md` and `.gitignore` (ignore rules only; no tracked file removed). Validated with `git status` (no tracked files changed).
