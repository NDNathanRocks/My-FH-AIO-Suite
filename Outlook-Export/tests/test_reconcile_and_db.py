import shutil
import sqlite3
from pathlib import Path

import pytest

from exporter_core import database
from exporter_core.database import ExportDatabase, migrate, rollback_schema
from exporter_core.engine import EngineConfig, SourceRef, run_export
from exporter_core.files import OLE_MAGIC
from exporter_core.naming import entry_hash
from exporter_core.reconcile import reconcile
from exporter_core.reports import write_reconciliation, write_run_reports
from fakes import FakeMsg, FakeSource

REAL_DB = Path(__file__).resolve().parents[1] / "outlook_export.db"
LEGACY_DDL = "CREATE TABLE exports(store_id TEXT,entry_id TEXT,folder TEXT,destination TEXT,size INTEGER,status TEXT,updated TEXT,error TEXT,PRIMARY KEY(store_id,entry_id))"
CFG = EngineConfig(retry_delay_seconds=0, memory_error_delay_seconds=0, pause_seconds=0, sleep=lambda _: None)


def legacy_db(path: Path, rows):
    with sqlite3.connect(path) as c:
        c.execute(LEGACY_DDL)
        c.execute("CREATE INDEX idx_exports_status ON exports(status)")
        c.executemany("INSERT INTO exports VALUES(?,?,?,?,?,?,?,?)", rows)
    return path


def snapshot(path):
    with sqlite3.connect(path) as c:
        return sorted(c.execute("SELECT store_id, entry_id, folder, destination, size, status, updated, error FROM exports"))


def test_migration_of_v1_schema_is_additive_backed_up_and_idempotent(tmp_path):
    rows = [("S", f"E{i}", "\\\\Mb\\Inbox", f"M:\\out\\f{i}.msg", 100 + i, "completed", "2026-09-29T09:00:00", "") for i in range(5)]
    db = legacy_db(tmp_path / "outlook_export.db", rows)
    before = snapshot(db)
    report = migrate(db, tmp_path / "backup")
    assert report["backup"] and Path(report["backup"]).exists()
    assert snapshot(Path(report["backup"])) == before
    assert snapshot(db) == before  # legacy columns/rows untouched
    assert set(report["added_columns"]) == {"item_class", "item_type", "source_folder", "file_size"}
    assert report["backfilled"] == 5
    with sqlite3.connect(db) as c:
        assert c.execute("SELECT COUNT(*) FROM export_manifest WHERE final_status='legacy-unverified'").fetchone()[0] == 5
        assert c.execute("SELECT source_folder, file_size FROM exports WHERE entry_id='E1'").fetchone() == ("\\\\Mb\\Inbox", 101)
        assert c.execute("SELECT value FROM schema_meta WHERE key='schema_version'").fetchone()[0] == "2"
    again = migrate(db, tmp_path / "backup")
    assert again["backup"] is None and not again["added_columns"] and not again["created_tables"]
    assert len(list((tmp_path / "backup").glob("*.db"))) == 1


def test_migration_failure_rolls_back(tmp_path, monkeypatch):
    db = legacy_db(tmp_path / "outlook_export.db", [("S", "E1", "f", "d", 1, "completed", "t", "")])
    before = snapshot(db)
    monkeypatch.setattr(database, "_backfill_manifest", lambda c: (_ for _ in ()).throw(RuntimeError("boom")))
    with pytest.raises(RuntimeError):
        migrate(db, tmp_path / "backup")
    with sqlite3.connect(db) as c:
        assert database.tables(c) == {"exports"}
        assert "item_type" not in database.table_columns(c, "exports")
    assert snapshot(db) == before


def test_rollback_drops_only_new_tables(tmp_path):
    db = legacy_db(tmp_path / "outlook_export.db", [("S", "E1", "f", "d", 1, "completed", "t", "")])
    migrate(db, tmp_path / "backup")
    result = rollback_schema(db, tmp_path / "backup")
    with sqlite3.connect(db) as c:
        assert database.tables(c) == {"exports"}
    assert snapshot(db)[0][:6] == ("S", "E1", "f", "d", 1, "completed")
    assert Path(result["backup"]).exists()
    migrate(db, tmp_path / "backup")  # re-migration after rollback works


def test_new_writes_keep_legacy_columns_filled(tmp_path):
    db_path = legacy_db(tmp_path / "outlook_export.db", [])
    out = tmp_path / "out"; out.mkdir()
    db = ExportDatabase(db_path, tmp_path / "backup")
    run_export(FakeSource({(): [FakeMsg("N1")]}, tmp_path), db, out, cfg=CFG)
    db.close()
    with sqlite3.connect(db_path) as c:
        folder, size, sf, fs, status = c.execute("SELECT folder, size, source_folder, file_size, status FROM exports").fetchone()
    assert folder == sf and size == fs and status == "completed"


@pytest.mark.skipif(not REAL_DB.exists(), reason="real database not present")
def test_real_database_copy_migrates_without_losing_records(tmp_path):
    copy = tmp_path / "outlook_export.db"
    shutil.copy2(REAL_DB, copy)
    before = snapshot(copy)
    report = migrate(copy, tmp_path / "backup")
    assert snapshot(copy) == before and report["backfilled"] == len(before)
    with sqlite3.connect(copy) as c:
        assert c.execute("PRAGMA integrity_check").fetchone()[0] == "ok"
    rec = reconcile(copy, tmp_path / "no_export_root_here")
    assert rec.counts["db_completed"] == len(before)
    # Paths are on M:\ (absent here) so every record reports a missing file; nothing is modified.
    assert rec.by_check().get("db_record_missing_file") == len(before)
    assert snapshot(copy) == before


def _write_msg(path: Path):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(OLE_MAGIC + b"x")


def test_reconcile_detects_every_category(tmp_path):
    out = tmp_path / "out"
    good, empty, corrupt = out / f"a - MailItem - {entry_hash('G')}.msg", out / f"b - MailItem - {entry_hash('Z')}.msg", out / f"c - MailItem - {entry_hash('C')}.msg"
    _write_msg(good); empty.write_bytes(b""); corrupt.write_bytes(b"not ole")
    shared = out / "shared - MailItem - 000000000000.msg"; _write_msg(shared)
    orphan = out / "orphan.msg"; _write_msg(orphan)
    adoptable = out / f"n - MailItem - {entry_hash('NEWSRC')}.msg"; _write_msg(adoptable)
    elsewhere = out / "moved" / f"m - MailItem - {entry_hash('MOVED')}.msg"; _write_msg(elsewhere)
    dup1 = out / f"x1 - MailItem - {entry_hash('TWIN')}.msg"; dup2 = out / "sub" / dup1.name; _write_msg(dup1); _write_msg(dup2)
    partial = out / ".partial-abc.msg"; _write_msg(partial)
    rows = [("S", "G", "f", str(good), good.stat().st_size, "completed", "t", ""),
            ("S", "Z", "f", str(empty), 0, "completed", "t", ""),
            ("S", "C", "f", str(corrupt), 7, "completed", "t", ""),
            ("S", "MISS", "f", str(out / "gone.msg"), 5, "completed", "t", ""),
            ("S", "P1", "f", str(shared), 1, "completed", "t", ""), ("S", "P2", "f", str(shared), 1, "completed", "t", ""),
            ("S", "FAIL", "f", None, None, "failed", "t", "{}"),
            ("S", "MOVED", "f", str(out / "old" / "m.msg"), 9, "completed", "t", ""),
            ("S", "OUT", "f", "Z:\\elsewhere\\o.msg", 9, "completed", "t", ""),
            ("S", "TWIN", "f", str(dup1), 9, "completed", "t", "")]
    db = legacy_db(tmp_path / "outlook_export.db", rows)
    migrate(db, tmp_path / "backup")
    with sqlite3.connect(db) as c:
        c.execute("UPDATE export_manifest SET internet_message_id='<same@x>' WHERE entry_id IN ('G','Z')")
    source = [SourceRef("S", "G", "\\\\Mb\\Inbox", (), 1), SourceRef("S", "NEWSRC", "\\\\Mb\\Inbox", (), 2),
              SourceRef("S", "", "\\\\Mb\\Inbox", (), 3, snapshot_error="MAPI_E_CORRUPT_DATA", subject="Broken"),
              SourceRef("S", "G", "\\\\Mb\\Inbox", (), 4)]
    rec = reconcile(db, out, source)
    checks = rec.by_check()
    for expected in ("db_record_missing_file", "empty_file", "corrupt_file", "duplicate_output_path", "failed_export",
                     "orphan_file", "file_missing_db_record", "file_not_at_recorded_path", "db_path_outside_output_root",
                     "multiple_files_for_email", "partial_file", "duplicate_internet_message_id", "unreadable_source_item",
                     "duplicate_identifier", "db_record_not_in_source"):
        assert checks.get(expected), (expected, checks)
    assert rec.overall_status == "incomplete"
    assert ("S", "FAIL") in rec.retry_keys("retry_failed")
    assert {("S", "MISS"), ("S", "NEWSRC"), ("S", "Z"), ("S", "C")} <= rec.retry_keys("retry_missing")
    unreadable = next(i for i in rec.issues if i.check == "unreadable_source_item")
    assert unreadable.subject == "Broken" and unreadable.path == "index 3"
    paths = write_reconciliation(tmp_path / "reports", rec, "R1")
    assert all(Path(p).exists() for p in paths.values())
    assert "INCOMPLETE" in Path(paths["txt"]).read_text()


def test_run_reports_written(tmp_path):
    out = tmp_path / "out"; out.mkdir()
    db = ExportDatabase(tmp_path / "x.db", tmp_path / "backup")
    result = run_export(FakeSource({(): [FakeMsg("R1"), FakeMsg("R2", open_error=RuntimeError("nope"))]}, tmp_path), db, out, cfg=CFG)
    db.close()
    paths = write_run_reports(tmp_path / "logs", result)
    import openpyxl
    wb = openpyxl.load_workbook(paths["excel_report"])
    assert wb.sheetnames == ["Summary", "Item Detail", "Item Types"]
    assert "R2" not in Path(paths["failures_csv"]).read_text()  # entry hash only, no raw IDs in the CSV
    assert entry_hash("R2") in Path(paths["failures_csv"]).read_text()
