# Inbox Email Exporter (v1.6)

Windows desktop app that saves every item in an Outlook mailbox folder (optionally with subfolders) as `.msg` files. It records each item in `outlook_export.db`, and after every run it reconciles Outlook, the database and the export folder so you can see exactly which items are missing or failed, and why.

## Setup and run (Windows, classic Outlook)

```bash
py -m pip install -r requirements.txt
py inbox_email_exporter_v1_5_2.py
```

Build the single-file exe: `py -m pip install -r requirements-dev.txt` then `py build_v1_5_2.py` → `dist\InboxEmailExporter.exe`. Keep `config.json`, `outlook_export.db`, `logs\` and `backup\` beside the exe.

## Configuration (`config.json`)

`output_root` (must be inside `allowed_output_root`), `mailbox_name`, `folder_path`, `include_subfolders`, `active_folder_max_items`, `dry_run_default`, retry/pause timings, `minimum_free_space_mb`, `max_windows_path`, `compute_checksums` (optional SHA-256 per file). `database_commit_batch` is kept for compatibility but no longer used: every item is committed as it completes.

## Workflow

| Button | What it does |
|---|---|
| Dry Run | Lists what would be exported. Writes nothing. |
| Start Export | Exports every item not yet exported. Re-running is safe: completed items are skipped (their files are re-checked), and a valid file left behind by an interrupted run is adopted rather than saved again. |
| Validate | Scans Outlook without exporting and reports items never exported and records whose file is missing. |
| Retry Failed | Re-attempts only items whose last export failed with a retryable error. |
| Retry Missing | Exports never-exported items and re-exports completed records whose file is missing, empty or corrupt. Failed items are left for Retry Failed. |
| Reports → Check Database vs Files | Offline check of the database against the export folder; no Outlook needed. |

Each run produces, in `logs\`:
- `audit_<run>.xlsx`: summary, per-item detail (stage, error category, retryable) and item types
- `failures_<run>.csv`
- `retry_candidates_<run>.csv`
- `summary_<run>.json`
- `reconciliation_<run>.{txt,csv,json}`

Overall status is **complete**, **complete with warnings** or **incomplete**.

Run counts: discovered, eligible, exported, adopted, already exported (file verified / file missing), skipped (limit or not selected), not yet exported, failed, unreadable, stopped.

Reconciliation checks:
- Outlook items with no database record
- Unreadable Outlook items, identified by folder, index and subject where available
- Database records whose file is missing, empty, corrupt (not an OLE `.msg`) or unreadable
- Files with no database record, or whose database status is incomplete (these are adopted on the next export)
- Files at a different path than recorded, or outside the output root
- Two records pointing at one file, or several files for one email
- Duplicate StoreID+EntryID in the source, and duplicate Internet Message IDs
- Conflicting statuses
- Leftover `.partial-*` files
- Runs that never finished

## Email identity

The key is **StoreID + EntryID**, the existing primary key, unchanged. The Internet Message ID is recorded too, to flag copies of the same email. Every file name ends with the first 12 hex characters of SHA-256(EntryID), which lets files be traced back to records. Subjects are never used as identifiers. File names are byte-for-byte the same as v1.5.2 (`tests/test_naming.py`).

## Database (additive migration)

On startup the app takes a backup of `outlook_export.db` into `backup\outlook_export_pre_migration_<time>.backup.db`, using SQLite's online-backup API, and checks the backup's integrity. Then, in one transaction, it:
- adds any missing v1.5.2 columns to `exports`
- creates `export_manifest` (per-item tracing: Internet Message ID, subject, sender, date, folder, expected/actual path, size, optional checksum, statuses, attempt count, last error stage/category/retryable, timestamps), `export_attempts` (full attempt history), `export_runs` and `schema_meta`
- backfills one manifest row per existing record, marked `legacy-unverified`

Existing rows and columns are never modified, and no historical export has to be rebuilt. Legacy `folder`/`size` columns keep being filled. Message bodies are never stored.

Rollback options:
- `py outlook_export_tools.py rollback-schema` drops only the new tables, after taking another backup.
- Restore a full backup: close the app, then copy the backup file over `outlook_export.db`.

## Maintenance CLI (no Outlook needed)

```bash
py outlook_export_tools.py status
py outlook_export_tools.py backup
py outlook_export_tools.py reconcile
py outlook_export_tools.py migrate --db C:\temp\copy.db
```

## Tests

```bash
py -m pip install -r requirements-dev.txt
py -m pytest tests
```

The tests use fake Outlook objects and temporary databases and folders, plus a migration test on a copy of the real `outlook_export.db`. They run on any OS.

## Validate on the target Windows PC (not possible in CI)

1. Close the app. Copy `outlook_export.db` somewhere safe.
2. `py outlook_export_tools.py migrate --db <copy>`, then `py outlook_export_tools.py reconcile --db <copy>`. Confirm the record count matches and that missing-file counts make sense for your M: drive.
3. Launch the app. Confirm `backup\outlook_export_pre_migration_*.backup.db` was created.
4. Dry Run on a small folder (e.g. `Inbox\1. Templates`). Then Start Export, and check the reconciliation is **complete**.
5. Run Start Export again: expect 0 exported and every item "already exported, file verified".
6. Delete one exported `.msg`, click Validate (expect 1 file missing), then Retry Missing (expect 1 exported).
7. Click Stop Safely mid-run, then Start Export again: no duplicate files, and any interrupted item is "adopted".
8. Rebuild the exe with `py build_v1_5_2.py` and repeat step 5 using the exe.

## Limitations

- `.msg` validity is checked by file header only; the file is not fully parsed.
- If the configured output root changes, old records show as "outside output root" or "missing file". Use the Database tab's rename (it backs up first) before Retry Missing, so files aren't re-exported under the new root.
