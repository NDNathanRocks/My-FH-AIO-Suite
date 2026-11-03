import sqlite3
import threading
import time
from datetime import datetime

import pytest

from exporter_core.database import ExportDatabase
from exporter_core.engine import EngineConfig, run_export
from exporter_core.files import verify_msg_file
from exporter_core.naming import entry_hash
from exporter_core.reconcile import reconcile
from fakes import FakeComError, FakeMsg, FakeSource

CFG = EngineConfig(save_retries=3, retry_delay_seconds=0, memory_error_delay_seconds=0, pause_every=1000, pause_seconds=0,
                   sleep=lambda _: None)


@pytest.fixture
def env(tmp_path):
    out = tmp_path / "out"
    out.mkdir()
    db_path = tmp_path / "outlook_export.db"

    def open_db():
        return ExportDatabase(db_path, tmp_path / "backup")
    return tmp_path, out, db_path, open_db


def run(source, open_db, out, **kw):
    db = open_db()
    try:
        return run_export(source, db, out, cfg=CFG, **kw)
    finally:
        db.close()


def status_of(db_path, eid):
    with sqlite3.connect(db_path) as c:
        row = c.execute("SELECT status FROM exports WHERE entry_id=?", (eid,)).fetchone()
    return row[0] if row else None


def test_successful_export_and_idempotent_rerun(env):
    tmp, out, db_path, open_db = env
    src = FakeSource({(): [FakeMsg("E1"), FakeMsg("E2", subject="Other")], ("Sub",): [FakeMsg("E3")]}, tmp)
    r1 = run(src, open_db, out)
    s = r1.summary
    assert (s["items_seen"], s["exported"], s["failed"], s["overall_status"], s["result"]) == (3, 3, 0, "complete", "success")
    files = sorted(out.rglob("*.msg"))
    assert len(files) == 3 and all(verify_msg_file(f)[0] == "ok" for f in files)
    assert (out / "Inbox" / "Sub").is_dir()
    r2 = run(src, open_db, out)
    assert (r2.summary["exported"], r2.summary["duplicates"], r2.summary["duplicates_verified"]) == (0, 3, 3)
    assert len(list(out.rglob("*.msg"))) == 3  # no duplicates
    assert sum(m.saves for m in src.by_id.values()) == 3


def test_unreadable_and_failed_items_are_identified(env):
    tmp, out, db_path, open_db = env
    msgs = [FakeMsg("OK1"), FakeMsg("", subject="Broken one", snapshot_error="MAPI_E_CORRUPT_DATA"),
            FakeMsg("BAD", subject="Cannot open", open_error=FakeComError(-2147221233, "The item could not be found")),
            FakeMsg("MEM", subject="Too big", save_errors=[MemoryError("not enough memory")] * 3)]
    r = run(FakeSource({(): msgs}, tmp), open_db, out)
    s = r.summary
    assert (s["exported"], s["unreadable"], s["failed"], s["overall_status"]) == (1, 1, 2, "incomplete")
    by_status = {row["Subject/Name"]: row for row in r.failures}
    assert by_status["Broken one"]["Stage"] == "snapshot" and by_status["Broken one"]["Item Index"] == 2
    bad = next(f for f in r.failures if f["Entry Hash"] == entry_hash("BAD"))
    assert (bad["Stage"], bad["Error Category"], bad["Retryable"]) == ("open_item", "item_not_found", "no")
    mem = next(f for f in r.failures if f["Entry Hash"] == entry_hash("MEM"))
    assert (mem["Stage"], mem["Error Category"], mem["Retryable"]) == ("save_temp", "memory", "yes")
    assert mem["Subject/Name"] == "Too big"
    with sqlite3.connect(db_path) as c:
        attempts = c.execute("SELECT outcome, stage FROM export_attempts ORDER BY id").fetchall()
        manifest = dict(c.execute("SELECT entry_id, retryable FROM export_manifest WHERE export_status='failed'").fetchall())
    assert ("unreadable", "snapshot") in attempts and manifest == {"BAD": 0, "MEM": 1}


def test_retry_failed_only_retries_retryable_and_keeps_history(env):
    tmp, out, db_path, open_db = env
    flaky = FakeMsg("FLAKY", save_errors=[OSError("being used by another process")] * 3)
    gone = FakeMsg("GONE", open_error=FakeComError(-2147221233, "could not be found"))
    src = FakeSource({(): [FakeMsg("OK"), flaky, gone]}, tmp)
    run(src, open_db, out)
    db = open_db()
    keys = db.failed_keys(retryable_only=True)
    db.close()
    assert keys == {("STORE1", "FLAKY")}
    r = run(src, open_db, out, mode="retry_failed", selection=keys)
    assert (r.summary["exported"], r.summary["not_selected"], r.summary["failed"]) == (1, 2, 0)
    assert status_of(db_path, "FLAKY") == "completed" and status_of(db_path, "GONE") == "failed"
    assert "GONE" not in src.opened[3:]
    with sqlite3.connect(db_path) as c:
        hist = [r[0] for r in c.execute("SELECT outcome FROM export_attempts WHERE entry_id='FLAKY' ORDER BY id")]
        count = c.execute("SELECT attempt_count FROM export_manifest WHERE entry_id='FLAKY'").fetchone()[0]
    assert hist == ["failed", "exported"] and count == 2


def test_missing_file_is_reported_then_retry_missing_reexports(env):
    tmp, out, db_path, open_db = env
    src = FakeSource({(): [FakeMsg("A"), FakeMsg("B")]}, tmp)
    run(src, open_db, out)
    victim = next(f for f in out.rglob("*.msg") if entry_hash("A") in f.name)
    victim.unlink()
    r = run(src, open_db, out)  # normal run: must NOT silently re-export
    assert (r.summary["exported"], r.summary["duplicates_missing_file"], r.summary["overall_status"]) == (0, 1, "incomplete")
    rec = reconcile(db_path, out)
    assert rec.retry_keys("retry_missing") == {("STORE1", "A")}
    assert [i.check for i in rec.issues if i.severity == "error"] == ["db_record_missing_file"]
    r = run(src, open_db, out, mode="retry_missing", selection=rec.retry_keys("retry_missing"))
    assert (r.summary["exported"], r.summary["duplicates_verified"], r.summary["not_selected"]) == (1, 1, 0)
    assert victim.exists() and reconcile(db_path, out).overall_status == "complete"


def test_interrupted_run_then_safe_resume_adopts_existing_files(env):
    tmp, out, db_path, open_db = env
    msgs = [FakeMsg(f"E{i}") for i in range(10)]
    src = FakeSource({(): msgs}, tmp)
    stop = threading.Event()
    db = open_db()
    original = db.record_success
    calls = {"n": 0}

    def crash_after_file(*a, **k):
        calls["n"] += 1
        if calls["n"] == 4:
            stop.set()
            raise sqlite3.OperationalError("disk I/O error")  # file placed, DB write lost
        return original(*a, **k)
    db.record_success = crash_after_file
    r = run_export(src, db, out, cfg=CFG, stop_event=stop)
    db.close()
    assert r.summary["result"] == "stopped" and r.summary["stopped_unprocessed"] == 6
    assert r.summary["exported"] == 3 and r.summary["failed"] == 1
    db = open_db(); assert db.interrupted_runs() == []; db.close()
    rec = reconcile(db_path, out)
    assert any(i.check == "file_with_incomplete_db_status" for i in rec.issues)
    r2 = run(src, open_db, out)
    s = r2.summary
    assert (s["duplicates"], s["adopted"], s["exported"], s["failed"]) == (3, 1, 6, 0)
    assert len(list(out.rglob("*.msg"))) == 10
    assert msgs[3].saves == 1  # adopted, not saved twice
    assert reconcile(db_path, out).overall_status == "complete"


def test_unfinished_run_row_is_reported(env):
    tmp, out, db_path, open_db = env
    db = open_db()
    db.start_run("RUN_X", "export", False, "src", str(out))
    interrupted = db.interrupted_runs()
    db.close()
    rec = reconcile(db_path, out, interrupted_runs=interrupted)
    assert [i.check for i in rec.issues] == ["interrupted_run"]


def test_duplicate_subjects_get_distinct_files(env):
    tmp, out, db_path, open_db = env
    same = dict(subject="Weekly report", when=datetime(2026, 1, 1, 8, 0, 0))
    run(FakeSource({(): [FakeMsg("X1", **same), FakeMsg("X2", **same)]}, tmp), open_db, out)
    names = sorted(f.name for f in out.rglob("*.msg"))
    assert len(names) == 2 and names[0] != names[1]


def test_invalid_characters_unicode_and_long_paths(env):
    tmp, out, db_path, open_db = env
    msgs = [FakeMsg("U1", subject='Zoë: 漢字 <x> "y" | z?*/\\'), FakeMsg("L1", subject="S" * 400)]
    r = run(FakeSource({(): msgs}, tmp), open_db, out)
    assert r.summary["exported"] == 2
    for f in out.rglob("*.msg"):
        assert not set('<>:"|?*').intersection(f.name)
        assert len(str(f)) <= 259
    assert any("Zoë_ 漢字 _x_ _y_ _ z____" in f.name for f in out.rglob("*.msg"))


def test_too_long_output_root_fails_clearly(env):
    tmp, out, db_path, open_db = env
    deep = out / ("d" * 200)
    deep.mkdir()
    r = run(FakeSource({(): [FakeMsg("P1")]}, tmp, root_name="R" * 80), open_db, deep)
    f = r.failures[0]
    assert (f["Stage"], f["Error Category"], f["Retryable"]) == ("build_path", "path_too_long", "no")


def test_corrupt_existing_file_is_overwritten_not_adopted(env):
    tmp, out, db_path, open_db = env
    src = FakeSource({(): [FakeMsg("C1")]}, tmp)
    run(src, open_db, out, dry_run=True)
    run(src, open_db, out)
    f = next(out.rglob("*.msg"))
    f.write_bytes(b"garbage")
    with sqlite3.connect(db_path) as c:
        c.execute("DELETE FROM exports"); c.commit()
    r = run(src, open_db, out)
    assert r.summary["exported"] == 1 and verify_msg_file(f)[0] == "ok"


def test_dry_run_writes_nothing(env):
    tmp, out, db_path, open_db = env
    r = run(FakeSource({(): [FakeMsg("D1"), FakeMsg("D2")]}, tmp), open_db, out, dry_run=True)
    assert r.summary["proposed"] == 2 and r.summary["result"] == "success"
    assert not list(out.rglob("*")) and status_of(db_path, "D1") is None


def test_validate_mode_lists_missing_without_opening(env):
    tmp, out, db_path, open_db = env
    src = FakeSource({(): [FakeMsg("V1"), FakeMsg("V2")]}, tmp)
    run(src, open_db, out, selection={("STORE1", "V1")})
    src.opened.clear()
    r = run(src, open_db, out, mode="validate")
    assert (r.summary["duplicates_verified"], r.summary["missing_from_db"], r.summary["overall_status"]) == (1, 1, "incomplete")
    assert src.opened == [] and r.retry_candidates[0]["entry_id"] == "V2"


def test_item_limit_is_intentional_skip(env):
    tmp, out, db_path, open_db = env
    r = run(FakeSource({(): [FakeMsg(f"L{i}") for i in range(5)]}, tmp), open_db, out, limit=2)
    s = r.summary
    assert (s["exported"], s["skipped_limit"], s["balanced"], s["overall_status"]) == (2, 3, True, "complete with warnings")


def test_large_batch(env):
    tmp, out, db_path, open_db = env
    msgs = [FakeMsg(f"B{i:05d}", subject=f"Msg {i % 50}") for i in range(3000)]
    start = time.perf_counter()
    r = run(FakeSource({(): msgs}, tmp), open_db, out)
    assert r.summary["exported"] == 3000 and r.summary["balanced"]
    assert len(list(out.rglob("*.msg"))) == 3000
    rec = reconcile(db_path, out)
    assert rec.overall_status == "complete" and rec.counts["files_verified"] == 3000
    assert time.perf_counter() - start < 120


def test_no_partial_files_left_after_failed_move(env, monkeypatch):
    tmp, out, db_path, open_db = env
    import exporter_core.files as files_mod
    real = files_mod.os.replace

    def boom(a, b):
        raise PermissionError("Access is denied")
    monkeypatch.setattr(files_mod.os, "replace", boom)
    r = run(FakeSource({(): [FakeMsg("M1")]}, tmp), open_db, out)
    monkeypatch.setattr(files_mod.os, "replace", real)
    assert r.failures[0]["Stage"] == "move" and r.failures[0]["Error Category"] == "permission"
    assert not list(out.rglob("*.msg")) and not list((tmp / "temp").glob("*"))


def test_retry_missing_skips_failed_and_exports_new(env):
    tmp, out, db_path, open_db = env
    bad = FakeMsg("F1", save_errors=[OSError("being used by another process")] * 3)
    src = FakeSource({(): [FakeMsg("A1"), bad]}, tmp)
    run(src, open_db, out)
    src.tree[()].append(FakeMsg("NEW"))
    src.by_id["NEW"] = src.tree[()][-1]
    r = run(src, open_db, out, mode="retry_missing", selection=set())
    assert (r.summary["exported"], r.summary["not_selected"], r.summary["duplicates"]) == (1, 1, 1)
    assert status_of(db_path, "F1") == "failed" and status_of(db_path, "NEW") == "completed"
