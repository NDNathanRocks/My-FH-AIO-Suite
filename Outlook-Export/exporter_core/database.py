"""SQLite history, manifest and migration.

Compatibility contract with v1.5.x:
* ``exports`` keeps its primary key ``(store_id, entry_id)`` and every
  existing column/row. Writes still set status ``completed``/``failed``
  exactly as before (legacy ``folder``/``size`` are also filled when present).
* New information lives in new tables (``export_manifest``,
  ``export_attempts``, ``export_runs``, ``schema_meta``). Migration only
  adds; ``rollback_schema`` drops just those tables.
* A consistent backup (SQLite online-backup API) is taken before any change.
"""
from __future__ import annotations

import json
import sqlite3
from contextlib import closing
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Iterable, Optional

from . import SCHEMA_VERSION
from .errors import FailureInfo

Key = tuple[str, str]

EXPORTS_DDL = """CREATE TABLE exports(
    store_id TEXT NOT NULL,entry_id TEXT NOT NULL,item_class INTEGER,
    item_type TEXT,source_folder TEXT,destination TEXT,file_size INTEGER,
    status TEXT,updated TEXT,error TEXT,PRIMARY KEY(store_id,entry_id))"""
LEGACY_REQUIRED = {"item_class": "INTEGER", "item_type": "TEXT", "source_folder": "TEXT", "destination": "TEXT",
                   "file_size": "INTEGER", "status": "TEXT", "updated": "TEXT", "error": "TEXT"}
NEW_TABLES = {
    "schema_meta": "CREATE TABLE schema_meta(key TEXT PRIMARY KEY, value TEXT)",
    "export_manifest": """CREATE TABLE export_manifest(
        store_id TEXT NOT NULL, entry_id TEXT NOT NULL, entry_hash TEXT, internet_message_id TEXT,
        subject TEXT, sender TEXT, item_date TEXT, mailbox TEXT, source_folder TEXT,
        item_class INTEGER, item_type TEXT, message_class TEXT,
        expected_path TEXT, actual_path TEXT, file_name TEXT, file_format TEXT, file_size INTEGER, sha256 TEXT,
        export_status TEXT, db_status TEXT, fs_status TEXT, attempt_count INTEGER DEFAULT 0,
        error_category TEXT, error_message TEXT, error_stage TEXT, retryable INTEGER,
        first_attempt TEXT, last_attempt TEXT, last_run_id TEXT, final_status TEXT, verified_at TEXT,
        PRIMARY KEY(store_id, entry_id))""",
    "export_attempts": """CREATE TABLE export_attempts(
        id INTEGER PRIMARY KEY AUTOINCREMENT, run_id TEXT, attempted_at TEXT,
        store_id TEXT, entry_id TEXT, entry_hash TEXT, source_folder TEXT, item_index INTEGER,
        subject TEXT, outcome TEXT, stage TEXT, error_category TEXT, error_type TEXT,
        error_message TEXT, retryable INTEGER, destination TEXT)""",
    "export_runs": """CREATE TABLE export_runs(
        run_id TEXT PRIMARY KEY, mode TEXT, dry_run INTEGER, source TEXT, output_root TEXT,
        started TEXT, finished TEXT, overall_status TEXT, counts_json TEXT)""",
}
NEW_INDEXES = (
    "CREATE INDEX IF NOT EXISTS idx_exports_status ON exports(status)",
    "CREATE INDEX IF NOT EXISTS idx_manifest_status ON export_manifest(export_status)",
    "CREATE INDEX IF NOT EXISTS idx_manifest_hash ON export_manifest(entry_hash)",
    "CREATE INDEX IF NOT EXISTS idx_manifest_imid ON export_manifest(internet_message_id)",
    "CREATE INDEX IF NOT EXISTS idx_attempts_key ON export_attempts(store_id, entry_id)",
)


def now() -> str:
    return datetime.now().isoformat(timespec="seconds")


def connect(db_path: Path) -> sqlite3.Connection:
    connection = sqlite3.connect(str(db_path), timeout=30)
    connection.execute("PRAGMA busy_timeout=30000")
    return connection


def table_columns(connection, table: str) -> list[str]:
    return [row[1] for row in connection.execute(f'PRAGMA table_info("{table}")')]


def tables(connection) -> set[str]:
    return {row[0] for row in connection.execute("SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'")}


def integrity(connection) -> str:
    return connection.execute("PRAGMA integrity_check").fetchone()[0]


def backup_database(db_path: Path, backup_dir: Path, label: str) -> Path:
    """Consistent copy via the SQLite backup API; verified before returning."""
    backup_dir.mkdir(parents=True, exist_ok=True)
    target = backup_dir / f"outlook_export_{label}_{datetime.now():%Y%m%d_%H%M%S}.backup.db"
    n = 1
    while target.exists():
        target = target.with_name(target.stem.replace(".backup", "") + f"_{n}.backup.db")
        n += 1
    with closing(connect(db_path)) as source, closing(sqlite3.connect(str(target))) as dest:
        source.backup(dest)
        if integrity(dest) != "ok":
            raise RuntimeError(f"Backup failed integrity check: {target}")
    return target


def pending_changes(connection) -> dict:
    existing = tables(connection)
    if "exports" not in existing:
        return {"create_exports": True, "add_columns": [], "new_tables": list(NEW_TABLES)}
    columns = table_columns(connection, "exports")
    return {"create_exports": False,
            "add_columns": [(n, k) for n, k in LEGACY_REQUIRED.items() if n not in columns],
            "new_tables": [t for t in NEW_TABLES if t not in existing]}


def migrate(db_path: Path, backup_dir: Path) -> dict:
    """Bring any v1.x database to the current schema. Idempotent and transactional."""
    db_path = Path(db_path)
    is_new = not db_path.exists() or db_path.stat().st_size == 0
    with closing(connect(db_path)) as connection:
        result = integrity(connection)
        if result != "ok":
            raise RuntimeError(f"Database integrity check failed before migration: {result}")
        changes = pending_changes(connection)
    needs_change = changes["create_exports"] or changes["add_columns"] or changes["new_tables"]
    report = {"created": is_new, "added_columns": [n for n, _ in changes["add_columns"]],
              "created_tables": list(changes["new_tables"]), "backup": None, "backfilled": 0}
    if not needs_change:
        return report
    if not is_new:
        report["backup"] = str(backup_database(db_path, backup_dir, "pre_migration"))
    with closing(connect(db_path)) as connection:
        connection.isolation_level = None  # manual transaction control
        connection.execute("BEGIN IMMEDIATE")
        try:
            if changes["create_exports"]:
                connection.execute(EXPORTS_DDL)
            columns_before = table_columns(connection, "exports")
            for name, kind in changes["add_columns"]:
                connection.execute(f'ALTER TABLE exports ADD COLUMN "{name}" {kind}')
            added = {n for n, _ in changes["add_columns"]}
            if "folder" in columns_before and "source_folder" in added:
                connection.execute("UPDATE exports SET source_folder=folder WHERE source_folder IS NULL OR source_folder=''")
            if "size" in columns_before and "file_size" in added:
                connection.execute("UPDATE exports SET file_size=size WHERE file_size IS NULL")
            for name in changes["new_tables"]:
                connection.execute(NEW_TABLES[name])
            for ddl in NEW_INDEXES:
                connection.execute(ddl)
            if "export_manifest" in changes["new_tables"]:
                report["backfilled"] = _backfill_manifest(connection)
            connection.execute("INSERT OR REPLACE INTO schema_meta(key,value) VALUES('schema_version',?)", (str(SCHEMA_VERSION),))
            connection.execute("INSERT OR REPLACE INTO schema_meta(key,value) VALUES('migrated_at',?)", (now(),))
            result = integrity(connection)
            if result != "ok":
                raise RuntimeError(f"Database integrity check failed during migration: {result}")
            connection.execute("COMMIT")
        except Exception:
            connection.execute("ROLLBACK")
            raise
    return report


def _backfill_manifest(connection) -> int:
    """One manifest row per historical export; files are not touched or rebuilt."""
    from .naming import entry_hash
    rows = connection.execute(
        "SELECT store_id, entry_id, item_class, item_type, source_folder, destination, file_size, status, updated, error FROM exports").fetchall()
    count = 0
    for store_id, entry_id, item_class, item_type, folder, destination, size, status, updated, error in rows:
        export_status = "exported" if status == "completed" else (status or "unknown")
        file_name = Path(str(destination).replace("\\", "/")).name if destination else None
        connection.execute(
            """INSERT OR IGNORE INTO export_manifest(store_id, entry_id, entry_hash, source_folder, item_class, item_type,
               expected_path, actual_path, file_name, file_format, file_size, export_status, db_status, fs_status,
               attempt_count, error_message, first_attempt, last_attempt, final_status)
               VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (store_id, entry_id, entry_hash(entry_id or ""), folder, item_class, item_type, destination, destination,
             file_name, "msg" if file_name else None, size, export_status, "legacy", "unchecked", 1 if status else 0,
             (error or "")[:1000], updated, updated, "legacy-unverified" if status == "completed" else export_status))
        count += 1
    return count


def rollback_schema(db_path: Path, backup_dir: Path) -> dict:
    """Undo the v2 tables (manifest/attempt/run history). ``exports`` is untouched."""
    backup = backup_database(db_path, backup_dir, "pre_rollback")
    with closing(connect(db_path)) as connection:
        connection.isolation_level = None
        connection.execute("BEGIN IMMEDIATE")
        try:
            dropped = [t for t in NEW_TABLES if t in tables(connection)]
            for name in dropped:
                connection.execute(f'DROP TABLE "{name}"')
            connection.execute("COMMIT")
        except Exception:
            connection.execute("ROLLBACK")
            raise
    return {"backup": str(backup), "dropped": dropped}


@dataclass
class ItemMeta:
    """Identifying metadata for one Outlook item (no body content)."""
    store_id: str
    entry_id: str
    source_folder: str = ""
    item_index: int = 0
    item_class: int = -1
    item_type: str = ""
    subject: str = ""
    sender: str = ""
    item_date: str = ""
    message_class: str = ""
    internet_message_id: str = ""
    mailbox: str = ""

    @property
    def key(self) -> Key:
        return (self.store_id, self.entry_id)


class ExportDatabase:
    """Writes are committed per item so an interruption loses at most the item in progress."""

    def __init__(self, db_path: Path, backup_dir: Path):
        self.db_path = Path(db_path)
        self.migration = migrate(self.db_path, backup_dir)
        self.connection = connect(self.db_path)
        self.connection.execute("PRAGMA journal_mode=DELETE")
        self.connection.execute("PRAGMA synchronous=FULL")
        self.columns = set(table_columns(self.connection, "exports"))
        self.completed: dict[Key, str] = {}
        size_col = "COALESCE(file_size, size)" if "size" in self.columns else "file_size"
        for store_id, entry_id, destination, _size in self.connection.execute(
                f"SELECT store_id, entry_id, destination, {size_col} FROM exports WHERE status='completed'"):
            self.completed[(store_id, entry_id)] = destination or ""

    # ---- queries ----
    def is_completed(self, key: Key) -> bool:
        return key in self.completed

    def manifest(self, key: Key) -> Optional[sqlite3.Row]:
        self.connection.row_factory = sqlite3.Row
        try:
            return self.connection.execute("SELECT * FROM export_manifest WHERE store_id=? AND entry_id=?", key).fetchone()
        finally:
            self.connection.row_factory = None

    def failed_keys(self, retryable_only: bool = True) -> set[Key]:
        sql = "SELECT store_id, entry_id FROM exports WHERE status='failed'"
        keys = set(map(tuple, self.connection.execute(sql)))
        if retryable_only:
            permanent = set(map(tuple, self.connection.execute(
                "SELECT store_id, entry_id FROM export_manifest WHERE export_status='failed' AND retryable=0")))
            keys -= permanent
        return keys

    # ---- runs ----
    def start_run(self, run_id: str, mode: str, dry_run: bool, source: str, output_root: str) -> None:
        self.connection.execute("INSERT OR REPLACE INTO export_runs(run_id, mode, dry_run, source, output_root, started) VALUES(?,?,?,?,?,?)",
                                (run_id, mode, int(dry_run), source, output_root, now()))
        self.connection.commit()

    def finish_run(self, run_id: str, overall_status: str, counts: dict) -> None:
        self.connection.execute("UPDATE export_runs SET finished=?, overall_status=?, counts_json=? WHERE run_id=?",
                                (now(), overall_status, json.dumps(counts, default=str), run_id))
        self.connection.commit()

    def interrupted_runs(self) -> list[tuple]:
        return self.connection.execute("SELECT run_id, mode, started FROM export_runs WHERE finished IS NULL ORDER BY started").fetchall()

    # ---- item writes ----
    def _attempt(self, run_id, meta: ItemMeta, outcome, destination="", failure: FailureInfo | None = None):
        from .naming import entry_hash
        self.connection.execute(
            """INSERT INTO export_attempts(run_id, attempted_at, store_id, entry_id, entry_hash, source_folder, item_index, subject,
               outcome, stage, error_category, error_type, error_message, retryable, destination) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (run_id, now(), meta.store_id, meta.entry_id, entry_hash(meta.entry_id) if meta.entry_id else "", meta.source_folder,
             meta.item_index, meta.subject[:255], outcome, failure.stage if failure else None, failure.category if failure else None,
             failure.error_type if failure else None, failure.message if failure else None,
             int(failure.retryable) if failure else None, str(destination)))

    def _manifest_upsert(self, run_id, meta: ItemMeta, **fields):
        from .naming import entry_hash
        ts = now()
        base = {"entry_hash": entry_hash(meta.entry_id), "internet_message_id": meta.internet_message_id or None,
                "subject": meta.subject[:500] or None, "sender": meta.sender[:255] or None, "item_date": meta.item_date or None,
                "mailbox": meta.mailbox or None, "source_folder": meta.source_folder, "item_class": meta.item_class,
                "item_type": meta.item_type, "message_class": meta.message_class or None, "last_attempt": ts, "last_run_id": run_id}
        base.update(fields)
        cols = list(base)
        assignments = ", ".join(f"{c}=COALESCE(excluded.{c}, export_manifest.{c})" if c in
                                ("internet_message_id", "subject", "sender", "item_date", "mailbox", "message_class") else f"{c}=excluded.{c}"
                                for c in cols)
        self.connection.execute(
            f"""INSERT INTO export_manifest(store_id, entry_id, first_attempt, attempt_count, {", ".join(cols)})
                VALUES(?, ?, ?, 1, {", ".join("?" for _ in cols)})
                ON CONFLICT(store_id, entry_id) DO UPDATE SET attempt_count=COALESCE(export_manifest.attempt_count,0)+1,
                first_attempt=COALESCE(export_manifest.first_attempt, excluded.first_attempt), {assignments}""",
            (meta.store_id, meta.entry_id, ts, *base.values()))

    def record_success(self, run_id: str, meta: ItemMeta, destination: Path, size: int, sha256: str | None, adopted: bool = False) -> None:
        dest = str(destination)
        legacy_cols, legacy_vals = "", ()
        if "folder" in self.columns:
            legacy_cols += ",folder"; legacy_vals += (meta.source_folder,)
        if "size" in self.columns:
            legacy_cols += ",size"; legacy_vals += (size,)
        legacy_updates = "".join(f",{c}=excluded.{c}" for c in legacy_cols.strip(",").split(",") if c)
        try:
            self.connection.execute(
                f"""INSERT INTO exports(store_id,entry_id,item_class,item_type,source_folder,destination,file_size,status,updated,error{legacy_cols})
                    VALUES(?,?,?,?,?,?,?,?,?,?{",?" * len(legacy_vals)})
                    ON CONFLICT(store_id,entry_id) DO UPDATE SET item_class=excluded.item_class,item_type=excluded.item_type,
                    source_folder=excluded.source_folder,destination=excluded.destination,file_size=excluded.file_size,
                    status='completed',updated=excluded.updated,error=''{legacy_updates}""",
                (meta.store_id, meta.entry_id, meta.item_class, meta.item_type, meta.source_folder, dest, size, "completed", now(), "", *legacy_vals))
            self._manifest_upsert(run_id, meta, expected_path=dest, actual_path=dest, file_name=Path(dest.replace("\\", "/")).name,
                                  file_format="msg", file_size=size, sha256=sha256, export_status="exported", db_status="completed",
                                  fs_status="ok", error_category=None, error_message=None, error_stage=None, retryable=None,
                                  final_status="adopted" if adopted else "exported", verified_at=now())
            self._attempt(run_id, meta, "adopted" if adopted else "exported", dest)
            self.connection.commit()
        except Exception:
            self.connection.rollback()
            raise
        self.completed[meta.key] = dest

    def record_failure(self, run_id: str, meta: ItemMeta, failure: FailureInfo, expected_path: str = "", details: dict | None = None) -> None:
        try:
            if meta.entry_id:
                self.connection.execute(
                    """INSERT INTO exports(store_id,entry_id,item_class,item_type,source_folder,status,updated,error)
                       VALUES(?,?,?,?,?,'failed',?,?)
                       ON CONFLICT(store_id,entry_id) DO UPDATE SET item_class=excluded.item_class,item_type=excluded.item_type,
                       source_folder=excluded.source_folder,status='failed',updated=excluded.updated,error=excluded.error""",
                    (meta.store_id, meta.entry_id, meta.item_class, meta.item_type, meta.source_folder, now(),
                     json.dumps(details or {"type": failure.error_type, "message": failure.message}, default=str)))
                self._manifest_upsert(run_id, meta, expected_path=expected_path or None, export_status="failed", db_status="failed",
                                      fs_status="missing", error_category=failure.category, error_message=failure.message,
                                      error_stage=failure.stage, retryable=int(failure.retryable), final_status="failed")
            self._attempt(run_id, meta, "failed", expected_path, failure)
            self.connection.commit()
        except Exception:
            self.connection.rollback()
            raise
        self.completed.pop(meta.key, None)

    def record_unreadable(self, run_id: str, meta: ItemMeta, failure: FailureInfo) -> None:
        self._attempt(run_id, meta, "unreadable", "", failure)
        self.connection.commit()

    def mark_verified(self, key: Key, fs_status: str, size: int | None = None) -> None:
        self.connection.execute("UPDATE export_manifest SET fs_status=?, file_size=COALESCE(?, file_size), verified_at=? WHERE store_id=? AND entry_id=?",
                                (fs_status, size, now(), *key))

    def commit(self) -> None:
        self.connection.commit()

    def close(self) -> None:
        try:
            self.connection.commit()
        finally:
            self.connection.close()


def read_export_rows(db_path: Path) -> Iterable[dict]:
    """Read-only view of exports joined with manifest for reconciliation."""
    uri = f"file:{Path(db_path).as_posix()}?mode=ro"
    with closing(sqlite3.connect(uri, uri=True)) as connection:
        connection.row_factory = sqlite3.Row
        cols = set(table_columns(connection, "exports"))
        size = "COALESCE(e.file_size, e.size)" if "size" in cols else "e.file_size"
        folder = "COALESCE(e.source_folder, e.folder)" if "folder" in cols else "e.source_folder"
        has_manifest = "export_manifest" in tables(connection)
        manifest_cols = (", m.internet_message_id, m.subject, m.sender, m.item_date, m.export_status AS m_status, m.error_category,"
                         " m.error_stage, m.retryable, m.attempt_count, m.sha256") if has_manifest else ""
        join = " LEFT JOIN export_manifest m ON m.store_id=e.store_id AND m.entry_id=e.entry_id" if has_manifest else ""
        sql = (f"SELECT e.store_id, e.entry_id, e.status, e.destination, {size} AS file_size, {folder} AS source_folder, "
               f"e.item_type, e.updated, e.error{manifest_cols} FROM exports e{join}")
        for row in connection.execute(sql):
            yield dict(row)
