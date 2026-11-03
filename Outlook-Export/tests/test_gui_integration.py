"""Drive the real GUI module's execute()/startup with fake COM modules (no Outlook, no window)."""
import importlib
import json
import sqlite3
import sys
import threading
import types
from pathlib import Path

import pytest

try:
    import tkinter  # noqa: F401
except ImportError:  # headless CI: execute() never touches Tk, so a stub is enough
    tk = types.ModuleType("tkinter")
    tk.Tk = object
    for sub in ("ttk", "messagebox", "filedialog"):
        setattr(tk, sub, types.ModuleType(f"tkinter.{sub}"))
        sys.modules[f"tkinter.{sub}"] = getattr(tk, sub)
    sys.modules["tkinter"] = tk
from fakes import FakeMsg, FakeSource  # noqa: E402


@pytest.fixture
def gui(tmp_path, monkeypatch):
    pythoncom = types.SimpleNamespace(COINIT_APARTMENTTHREADED=2, CoInitializeEx=lambda *_: None, CoUninitialize=lambda: None)
    monkeypatch.setitem(sys.modules, "pythoncom", pythoncom)
    holder = {}
    fake_com = types.ModuleType("exporter_core.outlook_com")

    def make_source(config, mode, logger):
        src = holder["source"]
        src.close = lambda: None
        return src
    fake_com.OutlookSource = make_source
    monkeypatch.setitem(sys.modules, "exporter_core.outlook_com", fake_com)
    mod = importlib.import_module("inbox_email_exporter_v1_5_2")
    for name, value in {"HERE": tmp_path, "CONFIG_PATH": tmp_path / "config.json", "DB_PATH": tmp_path / "outlook_export.db",
                        "LOG_DIR": tmp_path / "logs", "BACKUP_DIR": tmp_path / "backup", "LOCK_PATH": tmp_path / "x.lock"}.items():
        monkeypatch.setattr(mod, name, value)
    out = tmp_path / "Email Backups" / "Outlook"
    out.mkdir(parents=True)
    config = {**mod.DEFAULT_CONFIG, "output_root": str(out), "allowed_output_root": str(out.parent),
              "minimum_free_space_mb": 0, "retry_delay_seconds": 0, "memory_error_delay_seconds": 0, "pause_seconds": 0}
    return mod, config, holder, tmp_path, out


def test_startup_creates_config_and_migrates(gui):
    mod, config, holder, tmp, out = gui
    mod.startup_health_check()
    assert (tmp / "config.json").exists()
    with sqlite3.connect(tmp / "outlook_export.db") as c:
        assert {"exports", "export_manifest", "export_attempts", "export_runs"} <= {r[0] for r in c.execute("SELECT name FROM sqlite_master")}


def test_export_retry_and_reports_end_to_end(gui):
    mod, config, holder, tmp, out = gui
    flaky = FakeMsg("F1", subject="Locked", save_errors=[OSError("being used by another process")] * 4)
    holder["source"] = FakeSource({(): [FakeMsg("A1"), flaky], ("Sub",): [FakeMsg("B1", subject="Ünïcode")]}, tmp)
    args = (lambda *a: None, threading.Event(), threading.Event(), None)
    r = mod.execute(config, "configured", False, *args, mode="export")
    assert (r["exported"], r["failed"], r["overall_status"], r["reconciliation_status"]) == (2, 1, "incomplete", "incomplete")
    for key in ("excel_report", "failures_csv", "reconciliation_csv", "reconciliation_json", "reconciliation_txt"):
        assert Path(r[key]).exists(), key
    assert "Locked" in Path(r["failures_csv"]).read_text(encoding="utf-8-sig")
    assert not (tmp / "x.lock").exists()
    r = mod.execute(config, "configured", False, *args, mode="retry_failed")
    assert (r["exported"], r["not_selected"], r["reconciliation_status"]) == (1, 2, "complete")
    r = mod.execute(config, "configured", False, *args, mode="validate")
    assert (r["duplicates_verified"], r["missing_from_db"], r["overall_status"]) == (3, 0, "complete")
    assert len(list(out.rglob("*.msg"))) == 3
    summary = json.loads(Path(tmp / "logs" / f"summary_{r['run_id']}.json").read_text())
    assert summary["reconciliation_status"] == "complete"
    off = mod.offline_reconcile(config)
    assert off["status"] == "complete"


def test_lock_prevents_concurrent_runs(gui):
    mod, config, holder, tmp, out = gui
    import os
    (tmp / "x.lock").write_text(json.dumps({"pid": os.getppid()}))
    holder["source"] = FakeSource({(): [FakeMsg("A1")]}, tmp)
    with pytest.raises(RuntimeError, match="Another exporter is active"):
        mod.execute(config, "configured", False, lambda *a: None, threading.Event(), threading.Event(), None)


def test_output_outside_approved_root_is_rejected(gui):
    mod, config, holder, tmp, out = gui
    holder["source"] = FakeSource({(): [FakeMsg("A1")]}, tmp)
    bad = {**config, "allowed_output_root": str(tmp / "elsewhere")}
    with pytest.raises(RuntimeError, match="outside the approved root"):
        mod.execute(bad, "configured", False, lambda *a: None, threading.Event(), threading.Event(), None)
    assert not (tmp / "x.lock").exists()


def test_cli_tools(tmp_path, capsys):
    import outlook_export_tools as tools
    db = tmp_path / "t.db"
    assert tools.main(["migrate", "--db", str(db), "--backup-dir", str(tmp_path / "b")]) == 0
    assert tools.main(["status", "--db", str(db)]) == 0
    assert tools.main(["reconcile", "--db", str(db), "--backup-dir", str(tmp_path / "b"), "--output-root", str(tmp_path), "--report-dir", str(tmp_path / "r")]) == 0
    assert "Overall status: COMPLETE" in capsys.readouterr().out
    assert tools.main(["rollback-schema", "--db", str(db), "--backup-dir", str(tmp_path / "b"), "--yes"]) == 0
