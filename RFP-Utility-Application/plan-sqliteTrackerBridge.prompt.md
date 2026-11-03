# Plan: SQLite Tracker Bridge — Replace Excel Direct-Write

## What I Found (The Full Problem)

### Critical bug in `_force_excel_recalc()` in `app_utils.py`

```python
for proc in psutil.process_iter(['name']):
    if proc.info['name'] == 'EXCEL.EXE':
        os.kill(proc.pid, signal.SIGTERM)  # ← kills ALL Excel on the machine
```

This runs every time `check_page` or `generator_page` becomes visible (`showEvent`). It destroys any unsaved Excel work by anyone on the PC — that's also why the app "thinks" Excel is open immediately after: it killed a stale COM process the previous `showEvent` left behind.

### `ensure_tracker_closed_interactive()` in `app_utils.py`

Checks for `~$filename.xlsx` lock files. These persist after Excel crashes → perpetual false positives with no way to clear them short of rebooting.

### Pages that write to the tracker (blocked by the lock problem)

| Page | Columns Written | Save Method |
|---|---|---|
| `check_page.py` | `Invoice Status`, `Remittance ID` per row | `wb_edit.save()` + **BUG**: also saves read-only `wb_data` |
| `generator_page.py` | `RFP Prepared (Y/N)` = "Y" per row | `wb.save()` |
| `merger_page.py` | `Merged (Y/N)` = "Y" per row | `wb.save()` |

### Pages that only read (harmed by `showEvent` kill-loop, not the write-lock)

- `audit_page.py` — reads tracker for audit comparisons
- `report_page.py` — reads tracker for date/status fields

### Existing SQLite infrastructure

`update_page.py` already has `ExcelDatabaseHandler` (full sheet import/export). The new `TrackerBridge` is a focused, write-only complement — it does **not** replace or overlap with that.

---

## Architecture: Push-Pull via SQLite Bridge

```
App Pages             SQLite Bridge DB               Excel Tracker
(check/gen/merger) →  tracker_bridge.db    ←  Power Query refresh (on demand)
                       on shared M:\ drive      (read-only view for staff)
```

---

## SQLite Schema (`tracker_bridge.db`)

All tables use `INSERT OR REPLACE` (upsert on primary key). Each stores `updated_at` (ISO timestamp) and `updated_by` (`DOMAIN\username`) for auditability.

**`rfp_prepared_updates`** — written by `generator_page`
```sql
CREATE TABLE IF NOT EXISTS rfp_prepared_updates (
    vendor_id       TEXT NOT NULL,
    invoice_number  TEXT NOT NULL,
    rfp_prepared    TEXT DEFAULT 'Y',
    updated_at      TEXT,
    updated_by      TEXT,
    PRIMARY KEY (vendor_id, invoice_number)
)
```

**`merge_status_updates`** — written by `merger_page`
```sql
CREATE TABLE IF NOT EXISTS merge_status_updates (
    invoice_number  TEXT PRIMARY KEY,
    vendor_id       TEXT,
    company_name    TEXT,
    merged          TEXT DEFAULT 'Y',
    updated_at      TEXT,
    updated_by      TEXT
)
```

**`invoice_status_updates`** — written by `check_page`
```sql
CREATE TABLE IF NOT EXISTS invoice_status_updates (
    vendor_id       TEXT NOT NULL,
    invoice_number  TEXT NOT NULL,
    invoice_status  TEXT,
    remittance_id   TEXT,
    amount_paid     REAL,
    updated_at      TEXT,
    updated_by      TEXT,
    PRIMARY KEY (vendor_id, invoice_number)
)
```

---

## Phased Implementation Plan

### Phase 1 — `utils/tracker_bridge.py` (new file)

- `TrackerBridge(db_path)` class
- `ensure_schema()` — creates the 3 tables if they don't exist
- `upsert_rfp_prepared(vendor_id, invoice_number)` — upsert with timestamp + OS user
- `upsert_merge_status(invoice_number, vendor_id, company_name)` — upsert
- `upsert_invoice_status(vendor_id, invoice_number, status, remittance_id, amount_paid)` — upsert
- `export_sync_csvs(output_dir)` — exports all 3 tables as CSVs for Power Query consumption
- **Short-lived connections only**: open → execute → commit → close per call; SQLite's write-queuing handles concurrency
- DB path default: `M:\Staffing Services Agency Intake\tracker_bridge.db` (via `paths_config.py`)

### Phase 2 — Fix `generator_page.py` *(depends on Phase 1)*

- In `update_tracker_rows_in_sheet()`: replace `ws[rfp_col][row].value = "Y"` + `wb.save(tracker_path)` with `TrackerBridge.upsert_rfp_prepared(vid, inv)`
- Remove `ensure_tracker_closed_interactive()` call from `generate()`
- Remove `_force_excel_recalc()` call from `showEvent()`

### Phase 3 — Fix `merger_page.py` *(depends on Phase 1)*

- In `_update_tracker()`: replace `ws[merged_col][row].value = "Y"` + `wb.save()` with `TrackerBridge.upsert_merge_status(...)`
- Remove `ensure_tracker_closed_interactive()` from `run_merge()`
- Remove `force_excel_recalc()` post-merge call

### Phase 4 — Fix `check_page.py` *(depends on Phase 1)*

- In `_process_invoice_row()`: replace `ws_edit[Invoice Status]` + `ws_edit[Remittance ID]` cell writes with `TrackerBridge.upsert_invoice_status(...)`
- Fix the `wb_data.save()` bug — `wb_data` is opened `data_only=True`; saving it corrupts formula metadata; remove that line entirely
- Remove `ensure_tracker_closed_interactive()` from `on_run()`
- Remove `force_excel_recalc()` from `showEvent()`

### Phase 5 — Clean `app_utils.py` *(parallel with Phases 2–4)*

- Remove the `os.kill` loop from `_force_excel_recalc()` — keep the COM recalc logic but make it optional/manual-only, never called on `showEvent`
- `ensure_tracker_closed_interactive()` stays but is now a non-blocking informational warning only, not a gate

### Phase 6 — `settings_page.py` *(depends on Phase 1)*

- Add "Bridge DB Path" config field (persisted via `load_config`/`save_config`)
- Add "Export Sync CSVs" button → calls `TrackerBridge.export_sync_csvs(output_dir)`

### Phase 7 — `POWER_QUERY_SETUP.md` *(parallel — no code)*

- Step-by-step: Data → Get Data → From File → From Folder → point at the 3 sync CSVs
- No ODBC driver required — CSV approach avoids the SQLite ODBC admin-rights requirement on shared machines

---

## Relevant Files

| File | Change |
|---|---|
| `utils/tracker_bridge.py` | **New file** — Phase 1 |
| `utils/paths_config.py` | Add `DEFAULT_BRIDGE_DB_PATH` constant |
| `utils/app_utils.py` | Remove `os.kill` loop from `_force_excel_recalc` |
| `pages/generator_page.py` | Phase 2 — `update_tracker_rows_in_sheet()` |
| `pages/merger_page.py` | Phase 3 — `_update_tracker()` |
| `pages/check_page.py` | Phase 4 — `_process_invoice_row()`, `_save_tracker_updates()`, `on_run()` |
| `pages/settings_page.py` | Phase 6 — bridge DB path + export button |
| `pages/update_page.py` | **Unchanged** — `ExcelDatabaseHandler` stays as full-sheet import/export tool |

---

## Verification Steps

1. Run generator → confirm `rfp_prepared_updates` in SQLite, no Excel errors
2. Run merger → confirm `merge_status_updates` in SQLite
3. Run check → confirm `invoice_status_updates` in SQLite, no lock errors
4. Open `tracker.xlsx` in Excel while running the app → confirm zero interference
5. Export sync CSVs → load in Excel Power Query → verify data matches app output
6. Simulate 2 users running app simultaneously → confirm both writes succeed

---

## Scope

**Included**: All 3 write pages + `tracker_bridge.py` + settings DB path + CSV export + `_force_excel_recalc` kill-loop removal

**Excluded (future phase)**: Migrating the read path — `audit_page` and `report_page` still read directly from the Excel tracker

**Excluded**: Replacing `update_page.py`'s `ExcelDatabaseHandler` — that stays as the full-sheet import/export tool; `TrackerBridge` is the narrow write-only complement
