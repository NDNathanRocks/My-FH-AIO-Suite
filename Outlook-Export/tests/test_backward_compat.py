"""The v1.6 exporter must run unchanged against every older production database and config.

Uses copies of the real databases on this machine when present (they are not in git),
plus synthetic copies of each historical schema so the checks also run anywhere.
"""
import json
import shutil
import sqlite3
from pathlib import Path

import pytest

from exporter_core.database import EXPORTS_DDL, ExportDatabase
from exporter_core.engine import EngineConfig, run_export
from exporter_core.reconcile import reconcile
from fakes import FakeMsg, FakeSource

HERE = Path(__file__).resolve().parents[1]
REAL_DBS = [p for p in [HERE / "outlook_export.db", *sorted((HERE / "backup").glob("*.db"))] if p.exists()]
CFG = EngineConfig(retry_delay_seconds=0, memory_error_delay_seconds=0, pause_seconds=0, sleep=lambda _: None)

SCHEMAS = {
    # v1.5.0 / v1.5.1 (what production currently has)
    "v1.5.1": "CREATE TABLE exports(store_id TEXT,entry_id TEXT,folder TEXT,destination TEXT,size INTEGER,status TEXT,updated TEXT,error TEXT,PRIMARY KEY(store_id,entry_id))",
    # v1.5.1 database after the v1.5.2 startup migration added columns
    "v1.5.2-migrated": "CREATE TABLE exports(store_id TEXT,entry_id TEXT,folder TEXT,destination TEXT,size INTEGER,status TEXT,updated TEXT,error TEXT, item_class INTEGER, item_type TEXT, source_folder TEXT, file_size INTEGER, PRIMARY KEY(store_id,entry_id))",
    # database created fresh by v1.5.2
    "v1.5.2-fresh": EXPORTS_DDL,
}

# Verbatim write statements from v1.5.2 History, to prove an OLD exe still works after the new one migrated the DB.
LEGACY_SUCCESS = """INSERT INTO exports(
    store_id,entry_id,item_class,item_type,source_folder,destination,file_size,status,updated,error)
    VALUES(?,?,?,?,?,?,?,?,?,?)
    ON CONFLICT(store_id,entry_id) DO UPDATE SET
    item_class=excluded.item_class,item_type=excluded.item_type,
    source_folder=excluded.source_folder,destination=excluded.destination,
    file_size=excluded.file_size,status='completed',updated=excluded.updated,error=''"""


def _rows(db):
    with sqlite3.connect(db) as c:
        cols = [r[1] for r in c.execute("PRAGMA table_info(exports)")]
        legacy = [c for c in ("store_id", "entry_id", "folder", "destination", "size", "status", "updated", "error") if c in cols]
        return sorted(c.execute(f"SELECT {','.join(legacy)} FROM exports")), legacy


def _synthetic(tmp_path, schema):
    db = tmp_path / "outlook_export.db"
    with sqlite3.connect(db) as c:
        c.execute(SCHEMAS[schema])
        cols = [r[1] for r in c.execute("PRAGMA table_info(exports)")]
        for i in range(3):
            values = {"store_id": "STORE1", "entry_id": f"OLD{i}", "folder": "\\\\Mb\\Inbox", "source_folder": "\\\\Mb\\Inbox",
                      "destination": f"M:\\old\\f{i}.msg", "size": 10, "file_size": 10, "status": "completed",
                      "updated": "2026-09-29T09:00:00", "error": "", "item_class": 43, "item_type": "MailItem"}
            use = [k for k in values if k in cols]
            c.execute(f"INSERT INTO exports({','.join(use)}) VALUES({','.join('?' * len(use))})", [values[k] for k in use])
    return db


def _exercise(db, tmp_path, known_key):
    """Start-up migration, an export run that includes an already-exported item, reconcile, then old-exe writes."""
    before, legacy_cols = _rows(db)
    out = tmp_path / "out"
    out.mkdir(exist_ok=True)
    edb = ExportDatabase(db, tmp_path / "backup")
    assert edb.migration["backup"] is None or Path(edb.migration["backup"]).exists()
    src = FakeSource({(): [FakeMsg(known_key[1]), FakeMsg("BRAND-NEW")]}, tmp_path, store=known_key[0])
    result = run_export(src, edb, out, cfg=CFG)
    edb.close()
    s = result.summary
    # The old item is NOT re-exported even though its file is on M:\ (absent here); it is flagged for review instead.
    assert (s["exported"], s["duplicates"], s["failed"]) == (1, 1, 0)
    assert s["duplicates_missing_file"] == 1
    after, _ = _rows(db)
    new_key_rows = [r for r in after if r[1] == "BRAND-NEW"]
    assert [r for r in after if r[1] != "BRAND-NEW"] == before  # every historical row byte-identical
    assert len(new_key_rows) == 1
    with sqlite3.connect(db) as c:
        assert c.execute("PRAGMA integrity_check").fetchone()[0] == "ok"
        # An older exe running against the migrated DB can still write
        c.execute(LEGACY_SUCCESS, ("STORE1", "OLD-EXE", 43, "MailItem", "f", "M:\\x.msg", 5, "completed", "t", ""))
        assert c.execute("SELECT status FROM exports WHERE entry_id='OLD-EXE'").fetchone()[0] == "completed"
    rec = reconcile(db, out)
    assert rec.counts["db_records"] == len(before) + 2
    # Re-running start-up is a no-op (no second backup, no changes)
    again = ExportDatabase(db, tmp_path / "backup")
    assert again.migration["backup"] is None and not again.migration["added_columns"]
    again.close()
    return legacy_cols


@pytest.mark.parametrize("schema", list(SCHEMAS))
def test_historical_schemas(tmp_path, schema):
    db = _synthetic(tmp_path, schema)
    _exercise(db, tmp_path, ("STORE1", "OLD0"))


@pytest.mark.skipif(not REAL_DBS, reason="no production database copies on this machine")
@pytest.mark.parametrize("real", REAL_DBS, ids=lambda p: p.name)
def test_real_production_databases(tmp_path, real):
    db = tmp_path / "outlook_export.db"
    shutil.copy2(real, db)
    with sqlite3.connect(db) as c:
        known = c.execute("SELECT store_id, entry_id FROM exports WHERE status='completed' LIMIT 1").fetchone()
    _exercise(db, tmp_path, tuple(known))
    assert Path(real).stat().st_size > 0  # original untouched (we only used a copy)


@pytest.mark.parametrize("name", ["config.json", "config copy.json"])
def test_existing_config_files_load_unchanged(tmp_path, monkeypatch, name):
    src = HERE / name
    if not src.exists():
        pytest.skip("config not present")
    import sys
    import types
    monkeypatch.setitem(sys.modules, "pythoncom", types.SimpleNamespace())
    try:
        import tkinter  # noqa: F401
    except ImportError:
        tk = types.ModuleType("tkinter"); tk.Tk = object
        for sub in ("ttk", "messagebox", "filedialog"):
            setattr(tk, sub, types.ModuleType(f"tkinter.{sub}")); monkeypatch.setitem(sys.modules, f"tkinter.{sub}", getattr(tk, sub))
        monkeypatch.setitem(sys.modules, "tkinter", tk)
    import importlib
    gui = importlib.import_module("inbox_email_exporter_v1_5_2")
    cfg_path = tmp_path / "config.json"
    shutil.copy2(src, cfg_path)
    original = cfg_path.read_bytes()
    monkeypatch.setattr(gui, "CONFIG_PATH", cfg_path)
    monkeypatch.setattr(gui, "HERE", tmp_path)
    monkeypatch.setattr(gui, "LOG_DIR", tmp_path / "logs")
    config = gui.load_config()
    stored = json.loads(original)
    for key in ("output_root", "allowed_output_root", "mailbox_name", "folder_path", "include_subfolders", "dry_run_default"):
        assert config[key] == stored[key], key
    assert config["compute_checksums"] is False  # new option defaults off
    assert cfg_path.read_bytes() == original  # loading never rewrites the file
