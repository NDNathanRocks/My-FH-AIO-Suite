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
- [x] Stage 1 – Audit-Processor: modular package, tests, remove Audit Summary sheet, merge HR toolbox.
- [x] Stage 2 – Outlook-Export: core module, manifest + reconciliation + retry, safe migration, tests.
- [ ] Stage 3 – RFP: regression tests for transformations, safe fixes only.
- [ ] Stage 4 – Macro-Vision + per-project READMEs + final validation sweep.

## Feature-parity checklist (HR Data Toolbox → Audit-Processor)

Evidence: `test_hr_parity.py` (54 combos vs original engine), `test_parity_audits.py` (79 golden cases vs legacy audit_app), Playwright browser run of both pages.

| Feature (source) | Location now | Parity |
|---|---|---|
| Working + N lookup files, CSV/XLSX/XLSM (HR) | `ui/hr_page.py`, `hr_toolbox/engine.py` (+XLS) | [x] |
| Per-lookup working identifier, composite keys (HR) | engine `process` | [x] |
| Match modes exact / standardized / soft contains, per-file Match Status (HR) | engine | [x] |
| Duplicate policy list/first/last/error, unique-only, delimiters (HR) | engine | [x] |
| Output naming + conflict separate/first-nonblank/combine (HR) | engine | [x] |
| Optional report sheets, mapping-profile JSON download (HR) | engine / hr_page | [x] |
| All 8 audits + Portfolio-only tool, 3-step flow (Audit) | `audits/`, `ui/audits_page.py` | [x] |
| PC Portfolio fallbacks, Employee Details, Outlook drafts (Audit) | `portfolio.py`, `outlook_drafts.py` | [x] (drafts mocked) |
| Audit Summary sheet (Audit) | **removed by request**; notes on screen + run report CSV | intentional |

Differences: shared header styling, formula-like text kept as text, per-tool isolation/counts, temp attachments deleted after use.

## Bug / risk register

| # | Project · location | Problem | Sev | Impact | Fix | Compat risk | Status |
|---|---|---|---|---|---|---|---|
| 2 | Audit · `excel_bytes` | Excel table ref/merge counted hidden `_` columns (CTO/OTO/STO audit) → table wider than header; Excel may "repair" file | High | Corrupt-looking workbook | Size from visible columns | None | [x] tested |
| 3 | Audit · Outlook drafts | Attachment xlsx with HR data left in `%TEMP%` forever | Med | HR data retention | Delete after `Save()`, purge >24 h | None | [x] |
| 4 | HR · engine output | Text cells starting `=` written as live formulas | Med | Formula injection / broken cells | Store as text | Cell shows literal text | [x] |
| 5 | Audit · runner | One audit exception could leave misleading partial state; no counts | Med | Misleading success | Per-tool `ToolRun` status + counts + run report | None | [x] |
| 6 | Audit · TOR/NBA | Unparseable dates/hours silently excluded | Low | Silent row loss | Reported in run notes | None | [x] |
| 7 | Audit · Streamlit | Usage telemetry sent to data.streamlit.io | Low | Privacy | `gatherUsageStats=false` | None | [x] |
| 8 | Outlook · snapshot | Unreadable item recorded without any identity | High | Failed email untraceable | Folder + index + subject (if readable) in attempts, reports | None | [x] |
| 9 | Outlook · History | Commit every 100 items: crash loses ≤99 records → orphan files, re-export | High | Duplicate/untracked files | Commit per item; resume adopts valid existing file | Slightly slower writes | [x] tested |
| 10 | Outlook · save | `shutil.move` straight to final name: a network drop leaves truncated `.msg` with real name | High | Corrupt file looks exported | Copy to `.partial-*` then atomic rename; header verify | None | [x] tested |
| 11 | Outlook · duplicates | "Already exported" never checked the file exists | High | Silent missing backups | Stat+header check per duplicate; Retry Missing | None | [x] tested |
| 12 | Outlook · `upgrade_v1_5_2_database_fix.py` | Re-running regenerates and overwrites the exporter from v1.5.1 | Med | Code loss | Removed (obsolete; in git history) | None | [x] |
| 13 | Outlook · backups | `shutil.copy2` of a live DB can copy mid-transaction | Med | Unusable backup | SQLite backup API + integrity check | Backups now in `backup\` | [x] |
| 14 | Outlook · empty subject naming | (found during refactor) must stay "No Subject" | — | Name drift | Parity test vs v1.5.2 functions | None | [x] |
| 15 | Outlook · `InboxEmailExporter.spec` | Hard-coded `C:/Users/ndsouza/...` path | Low | Build only on one PC | Spec is regenerated by `build_v1_5_2.py`; untrack with D1 | None | [~] |
| 1 | Repo root | DBs, logs, audit xlsx, build/, dist/*.exe, `__pycache__` are tracked in Git | High | HR/email metadata in history | Ignore going forward; untracking needs approval (see D1) | `git pull` would delete local copies | [?] |

## Decisions requiring approval

- **D1 – Untrack committed data/artifacts.** `Outlook-Export/*.db`, `backup/*.db`, `logs/*`, `build/`, `dist/*.exe`, `__pycache__/`, `Macro-Vision/*.json|*.txt` are tracked. `git rm --cached` would make every other clone *delete those files on next pull*. Recommend: back up the files on each machine, then untrack in one commit; purge history separately if required. Rollback: `git revert` the commit. Not done yet.
- **D2 – Remove `HR Data Toolbox Web v1.3/`.** Parity is demonstrated (above). Recommend deleting the folder after you've used the merged page for one real run; rollback is `git revert`. Kept until you approve.

## Completed work

**Stage 2 – Outlook-Export reconciliation.** Moved logic into `exporter_core/` (naming, database, engine, reconcile, reports, COM adapter) behind a fake-able `MailSource`. Added additive v2 schema (manifest, attempt history, runs) with pre-migration backup and `rollback-schema`; Validate / Retry Failed / Retry Missing; per-run Excel/CSV/JSON + reconciliation report; maintenance CLI. Validated by 44 tests covering every PROMPT scenario (incl. migration of a copy of the real 6,108-record DB, interrupted run + safe resume, naming parity with v1.5.2) and the GUI `execute()` path with fake COM. Real Outlook/exe validation steps are in `Outlook-Export/README.md` (not runnable here).

**Stage 1 – Audit-Processor merge.** Split the 1,255-line `audit_app.py` into `audit_suite/` (rules, portfolio, Excel, files, Outlook, runner, UI) and merged the HR Data Toolbox as a page. Removed the Audit Summary sheet per request (notes now on screen + run-report CSV). Validated by 89 pytest tests (legacy golden parity, HR engine parity, edge cases, AppTest) and a Playwright run of both pages. Original HR folder kept for rollback (D2).

**Stage 0.** Inspected all five projects, recorded boundaries and risks, added root `CLAUDE.md` and `.gitignore` (ignore rules only; no tracked file removed). Validated with `git status` (no tracked files changed).
