# PLAN

Source of truth: `PROMPT.MD`. Status legend: [x] done · [~] partial · [ ] open · [?] needs owner decision.

## Project boundaries

| Project | Stack | Entry point | Sensitivity |
|---|---|---|---|
| `Audit-Processor` | Streamlit web app | `streamlit run audit_app.py` | Merge target for HR Data Toolbox. No Audit Summary sheet in output. |
| ~~`HR Data Toolbox Web v1.3`~~ | — | — | Removed (D2); lives on as the Audit-Processor HR page. |
| `Outlook-Export` | Tkinter + Outlook COM (Windows) | `inbox_email_exporter_v1_5_2.py` / `dist\InboxEmailExporter.exe` | Existing DB (6,108 records) + export folders must be preserved. |
| `RFP-Utility-Application` | PySide6 desktop (Windows) | `python main.py` | Daily workflow; outputs/filenames/columns must not change. |
| `Macro-Vision` | customtkinter + pynput (Windows) | `vision_suite_pro_updated.py` | Standalone; low priority. |

## Active task list

- [x] Stage 0 – Inspect repo, PLAN.md, CLAUDE.md, root `.gitignore`.
- [x] Stage 1 – Audit-Processor: modular package, tests, remove Audit Summary sheet, merge HR toolbox.
- [x] Stage 2 – Outlook-Export: core module, manifest + reconciliation + retry, safe migration, tests.
- [x] Stage 3 – RFP: regression tests for transformations, safe fixes only.
- [x] Stage 4 – Macro-Vision + per-project READMEs + final validation sweep.

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
| 15 | Outlook · `InboxEmailExporter.spec` | Hard-coded `C:/Users/ndsouza/...` path | Low | Build only on one PC | Spec is regenerated by `build_v1_5_2.py`; now untracked | None | [x] |
| 16 | RFP · `app_utils.process_csv_and_fill_pdfs` | "No N rows, use blank RFP cells" fallback references `invoice_col_name` before assignment → `UnboundLocalError` | High | Generation fails with cryptic error for that tracker state | Move column-name resolution above the filter (see D3) | Fix would start generating RFPs for blank rows | [x] owner decision |
| 17 | RFP · generator "Skip tracker update" | Read-only `EmptyCell` has no `.column` → every row "not found", no preview file | High | Option silently does nothing | Use cell position | None (normal mode unchanged) | [x] tested |
| 18 | RFP · generator | Temp tracker CSV (invoice data) never deleted from `%TEMP%` | Med | Data retention | Delete in worker `finally` | None | [x] |
| 19 | RFP · merger `_merge_files` | Wrote merged PDF straight to final name; failure leaves truncated PDF | Med | Corrupt merged invoice | Temp name + `os.replace` | None | [x] tested |
| 20 | RFP · `requirements.txt` | Pillow, psutil, pywin32 imported but not listed | Med | Fresh install fails to start | Added (ranged) | None | [x] |
| 21 | RFP · end marker | With a leading blank column, the "Total" check reads the Unit column, not Site | Low? | Possible extra RFP for a Total row | See D4 | Business rule unknown | [x] owner decision |
| 22 | RFP · `check_page.build_ha_lookup` | Blank `Last Issued Check Num` becomes text `"nan"` | Med | `nan` may be written as Remittance ID | See D5 | Changes tracker values | [x] owner decision |
| 1 | Repo root | DBs, logs, audit xlsx, build/, dist/*.exe, `__pycache__` are tracked in Git | High | HR/email metadata in history | Untracked + ignored (D1) | `git pull` deletes local copies (owner confirmed backups) | [x] D1 |

## Decisions (all resolved by owner)

- **D1 – done.** DBs, backups, logs, build/dist, `.spec`, `__pycache__` and Macro-Vision save/log files untracked (`git rm --cached`, files kept on disk) and git-ignored. Other clones lose these files on `git pull` — owner confirmed backups exist.
- **D2 – done.** `HR Data Toolbox Web v1.3/` removed; recover from git history if ever needed.
- **D3 – done (report none).** No `N` rows + blank RFP cells → "No new RFPs to generate…" message; nothing generated.
- **D4 – done.** Vendor sheets put `Total` under Notes (Excel B) and the count under Invoice #. The old check read the Invoice # cell, so without the blank padding rows the footer would have produced an RFP for invoice "4" (proven by test). Now any `Total` left of Invoice # ends the sheet.
- **D5 – done.** Blank HA check numbers are empty, never `nan`.

## Remaining limitations / target-machine validation

- Outlook COM, Excel COM, Outlook drafts and the Tk/PySide6 GUIs on Windows were not run here (logic tested with fakes; RFP pages smoke-launched offscreen). Steps: `Outlook-Export/README.md` → "Validate on the target Windows PC"; Audit drafts: create one draft per audit type.
- `Outlook-Export/dist/InboxEmailExporter.exe` is the old v1.5.2 build. It stays compatible with the migrated DB (v1.5.2 only uses `exports`), but rebuild with `py build_v1_5_2.py` to get reconciliation/retry.
- Macro-Vision: no automated tests (needs a desktop session); syntax-checked only.
- No dependency vulnerability scanner was available offline; dependencies are pinned/ranged per project.

## Completed work

**Owner decisions D1–D5.** Applied as above; RFP tests now 55 passing (no xfail), including the real AdvancedHome layout with and without padding rows.

**Stage 4 – Sweep.** Added Macro-Vision README. Ran all suites (Audit 89, Outlook 44, RFP 52 + 1 documented xfail), `py_compile` on every tracked script, ruff (pyflakes) on changed code, and searches for secrets (none), machine paths and TODOs (only in already-tracked artifacts → D1).

**Stage 3 – RFP regression coverage.** Added 53 characterization tests on the pinned dependency versions: invoice-number extraction, currency/filename helpers, real-template PDF field filling and names, tracker update (only the Y/N cell changes, formulas kept), HA lookup, merger naming/merge. Also an offscreen launch of all 9 pages. Fixed only behaviour-preserving issues (17–20); three items need your decision (D3–D5). No output names, columns or mappings changed.

**Stage 2 – Outlook-Export reconciliation.** Moved logic into `exporter_core/` (naming, database, engine, reconcile, reports, COM adapter) behind a fake-able `MailSource`. Added additive v2 schema (manifest, attempt history, runs) with pre-migration backup and `rollback-schema`; Validate / Retry Failed / Retry Missing; per-run Excel/CSV/JSON + reconciliation report; maintenance CLI. Validated by 44 tests covering every PROMPT scenario (incl. migration of a copy of the real 6,108-record DB, interrupted run + safe resume, naming parity with v1.5.2) and the GUI `execute()` path with fake COM. Real Outlook/exe validation steps are in `Outlook-Export/README.md` (not runnable here).

**Stage 1 – Audit-Processor merge.** Split the 1,255-line `audit_app.py` into `audit_suite/` (rules, portfolio, Excel, files, Outlook, runner, UI) and merged the HR Data Toolbox as a page. Removed the Audit Summary sheet per request (notes now on screen + run-report CSV). Validated by 89 pytest tests (legacy golden parity, HR engine parity, edge cases, AppTest) and a Playwright run of both pages. Original HR folder kept for rollback (D2).

**Stage 0.** Inspected all five projects, recorded boundaries and risks, added root `CLAUDE.md` and `.gitignore` (ignore rules only; no tracked file removed). Validated with `git status` (no tracked files changed).
