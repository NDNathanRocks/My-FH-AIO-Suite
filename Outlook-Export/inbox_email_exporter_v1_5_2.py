"""Inbox Email Exporter (Windows GUI).

Business logic lives in ``exporter_core`` (testable without Outlook). This file
holds the Tkinter UI, configuration and the Outlook/COM session wiring.
The file name is kept for the existing build script and shortcuts.
"""
from __future__ import annotations

import gc
import json
import logging
import os
import queue
import shutil
import sqlite3
import sys
import threading
import time
import uuid
from datetime import datetime
from logging.handlers import RotatingFileHandler
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

from exporter_core.database import ExportDatabase, backup_database, migrate
from exporter_core.engine import EngineConfig, new_run_id, run_export
from exporter_core.errors import exception_details
from exporter_core.files import atomic_json
from exporter_core.lock import acquire_lock, release_lock
from exporter_core.naming import normalize_windows_path
from exporter_core.reconcile import reconcile
from exporter_core.reports import reconciliation_text, write_reconciliation, write_run_reports

APP = "Inbox Email Exporter"
VERSION = "1.6.0"

HERE = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
CONFIG_PATH = HERE / "config.json"
DB_PATH = HERE / "outlook_export.db"
LOG_DIR = HERE / "logs"
BACKUP_DIR = HERE / "backup"
LOCK_PATH = HERE / "outlook_export.lock"

DEFAULT_CONFIG = {
    "output_root": r"M:\Staffing Services Agency Intake\Email Backups\Outlook",
    "allowed_output_root": r"M:\Staffing Services Agency Intake\Email Backups",
    "mailbox_name": "FH Agency Staffing [FH]",
    "folder_path": "Inbox",
    "include_subfolders": True,
    "active_folder_max_items": 100,
    "dry_run_default": True,
    "pause_every": 50,
    "pause_seconds": 0.20,
    "save_retries": 4,
    "retry_delay_seconds": 2.0,
    "memory_error_delay_seconds": 15.0,
    "minimum_free_space_mb": 500,
    "max_windows_path": 240,
    "database_commit_batch": 100,
    "verify_existing_files_on_start": False,
    "compute_checksums": False,
}
MODE_LABELS = {"export": "Export", "retry_failed": "Retry failed", "retry_missing": "Retry missing", "validate": "Validate"}


def ensure_runtime() -> None:
    for folder in (HERE, LOG_DIR):
        folder.mkdir(parents=True, exist_ok=True)
        if not folder.is_dir():
            raise RuntimeError(f"Not a directory: {folder}")


def load_config() -> dict:
    ensure_runtime()
    if not CONFIG_PATH.exists():
        atomic_json(CONFIG_PATH, DEFAULT_CONFIG)
    try:
        stored = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    except ValueError as exc:
        raise RuntimeError(f"config.json is not valid JSON ({exc}). Fix or delete it to restore defaults.") from exc
    data = {**DEFAULT_CONFIG, **stored}
    if "active_folder_max_emails" in data and "active_folder_max_items" not in stored:
        data["active_folder_max_items"] = data["active_folder_max_emails"]
    for key in ("output_root", "allowed_output_root"):
        data[key] = normalize_windows_path(data.get(key, ""))
    return data


def inspect_export_schema(connection):
    tables = [row[0] for row in connection.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name")]
    table = "exports" if "exports" in tables else (tables[0] if tables else None)
    if not table:
        return {"tables": tables, "table": None, "columns": [], "source": None, "destination": None, "status": None}
    columns = [row[1] for row in connection.execute("PRAGMA table_info(" + quote_identifier(table) + ")")]

    def first(*names):
        return next((name for name in names if name in columns), None)
    return {"tables": tables, "table": table, "columns": columns,
            "source": first("source_folder", "folder", "folder_path", "source"),
            "destination": first("destination", "path", "file_path", "output_path"),
            "status": first("status", "state", "result")}


def quote_identifier(name: str) -> str:
    return '"' + str(name).replace('"', '""') + '"'


def close_log_handlers() -> None:
    logger = logging.getLogger("agency-export")
    for handler in list(logger.handlers):
        try:
            handler.flush()
            handler.close()
        finally:
            logger.removeHandler(handler)
    gc.collect()


def create_logger(run_id: str, ui_callback):
    """File logs record identifiers, hashes, folders and errors - never message bodies."""
    ensure_runtime()
    close_log_handlers()
    logger = logging.getLogger("agency-export")
    logger.setLevel(logging.DEBUG)
    logger.propagate = False
    formatter = logging.Formatter("%(asctime)s | %(levelname)s | %(message)s")
    for handler in (RotatingFileHandler(LOG_DIR / "application.log", maxBytes=10_000_000, backupCount=10, encoding="utf-8"),
                    logging.FileHandler(LOG_DIR / f"run_{run_id}.log", encoding="utf-8")):
        handler.setFormatter(formatter)
        logger.addHandler(handler)

    class UIHandler(logging.Handler):
        def emit(self, record):
            message = self.format(record)
            important = ("RUN ", "FOLDER ", "EXPORTED ", "ADOPTED ", "FAILED ", "PAUSED", "RESUMED", "STOPPING", "REPORT ", "RECONCILE", "ITEM LIMIT")
            if record.levelno >= logging.WARNING or any(token in message for token in important):
                ui_callback(record.levelname, message)
    ui = UIHandler()
    ui.setFormatter(formatter)
    logger.addHandler(ui)
    return logger


def validate_output(config: dict) -> Path:
    output = Path(os.path.expandvars(config["output_root"])).resolve()
    allowed = Path(os.path.expandvars(config["allowed_output_root"])).resolve()
    if not (output == allowed or allowed in output.parents):
        raise RuntimeError(f"Output folder {output} is outside the approved root {allowed}.")
    if not output.exists():
        raise RuntimeError(f"Output folder unavailable: {output} (is the network drive connected?)")
    free_mb = shutil.disk_usage(output).free / 1048576
    if free_mb < float(config["minimum_free_space_mb"]):
        raise RuntimeError(f"Insufficient output free space: {free_mb:,.0f} MB free, {config['minimum_free_space_mb']} MB required.")
    test = output / f".write_test_{uuid.uuid4().hex}"
    test.write_text("ok")
    test.unlink()
    return output


def execute(config, source_mode, dry_run, ui_callback, stop_event, pause_event, progress_callback=None, mode="export"):
    """One run (export / retry / validate) followed by a full reconciliation and reports."""
    import pythoncom
    from exporter_core.outlook_com import OutlookSource

    ensure_runtime()
    acquire_lock(LOCK_PATH)
    run_id = new_run_id(mode, dry_run)
    logger = create_logger(run_id, ui_callback)
    db = source = None
    pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
    try:
        output = validate_output(config)
        db = ExportDatabase(DB_PATH, BACKUP_DIR)
        if db.migration.get("backup"):
            logger.info("RUN database migrated | backup=%s", db.migration["backup"])
        interrupted = db.interrupted_runs()
        source = OutlookSource(config, source_mode, logger)
        selection = None
        if mode == "retry_missing":
            selection = reconcile(DB_PATH, output).retry_keys("retry_missing")
            logger.info("RUN retry-missing selection | records with missing/bad files=%s", len(selection))
        limit = None if source_mode == "configured" else config.get("active_folder_max_items")
        result = run_export(source, db, output, mode=mode, dry_run=dry_run, limit=limit, cfg=EngineConfig.from_config(config),
                            run_id=run_id, stop_event=stop_event, pause_event=pause_event, logger=logger,
                            progress=progress_callback, selection=selection)
        db.close()
        db = None
        reports = write_run_reports(LOG_DIR, result)
        rec = reconcile(DB_PATH, output, result.source_refs, interrupted_runs=interrupted)
        reports.update({f"reconciliation_{k}": v for k, v in write_reconciliation(LOG_DIR, rec, run_id).items()})
        summary = {**result.summary, **reports, "reconciliation_status": rec.overall_status,
                   "reconciliation_text": reconciliation_text(rec)}
        atomic_json(LOG_DIR / f"summary_{run_id}.json", summary)
        logger.info("REPORT WRITTEN | %s", reports["excel_report"])
        logger.info("RECONCILE | %s | issues=%s", rec.overall_status, len(rec.issues))
        return summary
    finally:
        if db is not None:
            db.close()
        if source is not None:
            source.close()
        gc.collect()
        pythoncom.CoUninitialize()
        close_log_handlers()
        release_lock(LOCK_PATH)


def offline_reconcile(config) -> dict:
    """Database + export folder check; no Outlook needed."""
    ensure_runtime()
    output = Path(os.path.expandvars(config["output_root"])).resolve()
    migrate(DB_PATH, BACKUP_DIR)
    with sqlite3.connect(DB_PATH) as connection:
        interrupted = connection.execute("SELECT run_id, mode, started FROM export_runs WHERE finished IS NULL").fetchall()
    rec = reconcile(DB_PATH, output, None, interrupted_runs=interrupted)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S") + "_OFFLINE"
    paths = write_reconciliation(LOG_DIR, rec, stamp)
    return {"text": reconciliation_text(rec), "status": rec.overall_status, **paths}


class App(tk.Tk):
    BASIC_FIELDS = [("output_root", "Output folder", True), ("allowed_output_root", "Approved output root", True),
                    ("mailbox_name", "Mailbox display name", False), ("folder_path", "Outlook folder path", False)]
    ADV_FIELDS = [("active_folder_max_items", "Active-folder item limit", False), ("pause_every", "Pause every N items", False),
                  ("pause_seconds", "Pause seconds", False), ("save_retries", "Save retries", False),
                  ("retry_delay_seconds", "Retry delay seconds", False), ("memory_error_delay_seconds", "Memory-error delay seconds", False),
                  ("minimum_free_space_mb", "Minimum free space MB", False), ("max_windows_path", "Maximum Windows path", False)]
    INTS = {"active_folder_max_items", "pause_every", "save_retries", "minimum_free_space_mb", "max_windows_path"}
    FLOATS = {"pause_seconds", "retry_delay_seconds", "memory_error_delay_seconds"}

    def __init__(self):
        super().__init__()
        self.title(f"{APP} {VERSION}")
        self.geometry("1120x790")
        self.minsize(980, 680)
        self.config_data = load_config()
        self.events = queue.Queue()
        self.stop_event, self.pause_event = threading.Event(), threading.Event()
        self.busy, self.dry_token, self.run_started = False, None, None
        self.make_ui()
        self.protocol("WM_DELETE_WINDOW", self.close_safely)
        self.after(250, self.clock_tick)
        self.after(100, self.poll)

    # ---------- layout ----------
    def make_ui(self):
        st = ttk.Style(self)
        st.theme_use("vista" if "vista" in st.theme_names() else "clam")
        st.configure("Title.TLabel", font=("Segoe UI", 20, "bold"), foreground="#0F3B70")
        st.configure("TButton", padding=8)
        h = ttk.Frame(self, padding=18)
        h.pack(fill="x")
        ttk.Label(h, text=APP, style="Title.TLabel").pack(anchor="w")
        ttk.Label(h, text="Outlook backup with Excel auditing, reconciliation and safe retry").pack(anchor="w")
        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True, padx=18, pady=(0, 18))
        run, settings, db, reports = (ttk.Frame(self.nb, padding=18) for _ in range(4))
        for frame, name in ((run, "Export"), (settings, "Settings"), (db, "Database"), (reports, "Reports")):
            self.nb.add(frame, text=name)

        self.mode = tk.StringVar(value="configured")
        box = ttk.LabelFrame(run, text="Source", padding=12)
        box.pack(fill="x")
        ttk.Radiobutton(box, text="Configured mailbox and folder", variable=self.mode, value="configured", command=self.invalidate).pack(anchor="w")
        ttk.Radiobutton(box, text="Currently open Outlook folder", variable=self.mode, value="active", command=self.invalidate).pack(anchor="w")
        self.use_dry = tk.BooleanVar(value=bool(self.config_data["dry_run_default"]))
        ttk.Checkbutton(box, text="Run dry run first (recommended)", variable=self.use_dry, command=self.invalidate).pack(anchor="w", pady=6)

        a = ttk.Frame(run)
        a.pack(fill="x", pady=(10, 4))
        self.dry = ttk.Button(a, text="Dry Run", command=lambda: self.start(True))
        self.dry.pack(side="left")
        self.export = ttk.Button(a, text="Start Export", command=lambda: self.start(False))
        self.export.pack(side="left", padx=8)
        self.pause = ttk.Button(a, text="Pause", command=self.toggle_pause, state="disabled")
        self.pause.pack(side="left")
        self.stop = ttk.Button(a, text="Stop Safely", command=self.request_stop, state="disabled")
        self.stop.pack(side="left", padx=8)
        self.status = ttk.Label(a, text="Ready")
        self.status.pack(side="left", padx=12)
        b = ttk.Frame(run)
        b.pack(fill="x", pady=(0, 10))
        self.validate_btn = ttk.Button(b, text="Validate (Outlook + DB + files)", command=lambda: self.start(False, "validate"))
        self.validate_btn.pack(side="left")
        self.retry_failed_btn = ttk.Button(b, text="Retry Failed", command=lambda: self.start(False, "retry_failed"))
        self.retry_failed_btn.pack(side="left", padx=8)
        self.retry_missing_btn = ttk.Button(b, text="Retry Missing", command=lambda: self.start(False, "retry_missing"))
        self.retry_missing_btn.pack(side="left")
        self.run_buttons = (self.dry, self.export, self.validate_btn, self.retry_failed_btn, self.retry_missing_btn)

        self.bar = ttk.Progressbar(run, mode="determinate", maximum=100)
        self.bar.pack(fill="x")
        m = ttk.Frame(run)
        m.pack(fill="x", pady=5)
        self.count_lbl = ttk.Label(m, text="0 / 0 items")
        self.count_lbl.pack(side="left")
        self.elapsed_lbl = ttk.Label(m, text="Elapsed 00:00:00")
        self.elapsed_lbl.pack(side="left", padx=20)
        self.eta_lbl = ttk.Label(m, text="ETA --:--:--")
        self.eta_lbl.pack(side="left")
        self.time_lbl = ttk.Label(m, text="Last update: -")
        self.time_lbl.pack(side="right")
        self.console = tk.Text(run, state="disabled", font=("Segoe UI", 9), bg="#F7F9FC", relief="flat")
        self.console.pack(fill="both", expand=True)
        for tag, color in (("INFO", "#245A9A"), ("SUCCESS", "#107C10"), ("WARNING", "#986F0B"), ("ERROR", "#C50F1F")):
            self.console.tag_configure(tag, foreground=color)

        mode = ttk.Frame(settings)
        mode.pack(fill="x")
        self.setting_mode = tk.StringVar(value="basic")
        ttk.Radiobutton(mode, text="Basic", variable=self.setting_mode, value="basic", command=self.render_settings).pack(side="left")
        ttk.Radiobutton(mode, text="Advanced", variable=self.setting_mode, value="advanced", command=self.render_settings).pack(side="left", padx=12)
        self.setting_body = ttk.Frame(settings)
        self.setting_body.pack(fill="both", expand=True, pady=12)
        self.setting_vars = {}
        self.render_settings()

        ttk.Label(db, text="Database Folder Mapping", font=("Segoe UI", 12, "bold")).pack(anchor="w")
        ttk.Label(db, text="Preview and update stored source/destination folder names. A database backup is always created before changes.", wraplength=900).pack(anchor="w", pady=5)
        self.db_old, self.db_new = tk.StringVar(), tk.StringVar()
        f = ttk.Frame(db)
        f.pack(fill="x", pady=8)
        ttk.Label(f, text="Existing text").grid(row=0, column=0, sticky="w")
        ttk.Entry(f, textvariable=self.db_old, width=75).grid(row=0, column=1, sticky="ew", padx=8)
        ttk.Label(f, text="Replacement text").grid(row=1, column=0, sticky="w")
        ttk.Entry(f, textvariable=self.db_new, width=75).grid(row=1, column=1, sticky="ew", padx=8)
        f.columnconfigure(1, weight=1)
        bb = ttk.Frame(db)
        bb.pack(fill="x")
        ttk.Button(bb, text="Refresh Database Tree", command=self.load_database_tree).pack(side="left")
        ttk.Button(bb, text="Preview Rename", command=lambda: self.db_update(False)).pack(side="left", padx=8)
        ttk.Button(bb, text="Apply Update", command=lambda: self.db_update(True)).pack(side="left")
        self.db_text = tk.Text(db, height=20, state="disabled", font=("Consolas", 9))
        self.db_text.pack(fill="both", expand=True, pady=8)

        rb = ttk.Frame(reports)
        rb.pack(fill="x")
        ttk.Button(rb, text="Open Logs and Excel Reports", command=lambda: (ensure_runtime(), os.startfile(LOG_DIR))).pack(side="left")
        ttk.Button(rb, text="Check Database vs Files (no Outlook)", command=self.run_offline_check).pack(side="left", padx=8)
        ttk.Label(reports, text="Reports: " + str(LOG_DIR) + "\nDatabase: " + str(DB_PATH) + "\nBackups: " + str(BACKUP_DIR), justify="left").pack(anchor="w", pady=10)
        self.report_text = tk.Text(reports, height=22, state="disabled", font=("Consolas", 9))
        self.report_text.pack(fill="both", expand=True)
        self.nb.bind("<<NotebookTabChanged>>", self.on_tab_changed)
        self.after(350, self.load_database_tree)

    def render_settings(self):
        for w in self.setting_body.winfo_children():
            w.destroy()
        fields = self.BASIC_FIELDS + (self.ADV_FIELDS if self.setting_mode.get() == "advanced" else [])
        for r, (key, label, browse) in enumerate(fields):
            ttk.Label(self.setting_body, text=label).grid(row=r, column=0, sticky="w", pady=6)
            var = self.setting_vars.setdefault(key, tk.StringVar(value="" if self.config_data.get(key) is None else str(self.config_data.get(key, ""))))
            ttk.Entry(self.setting_body, textvariable=var, width=72).grid(row=r, column=1, sticky="ew", padx=8, pady=6)
            if browse:
                ttk.Button(self.setting_body, text="Browse...", command=lambda k=key: self.browse(k)).grid(row=r, column=2, pady=6)
        self.rec = tk.BooleanVar(value=self.config_data["include_subfolders"])
        ttk.Checkbutton(self.setting_body, text="Include subfolders", variable=self.rec).grid(row=len(fields), column=1, sticky="w")
        self.checksums = tk.BooleanVar(value=bool(self.config_data.get("compute_checksums", False)))
        ttk.Checkbutton(self.setting_body, text="Store SHA-256 checksum of each exported file (slower)", variable=self.checksums).grid(row=len(fields) + 1, column=1, sticky="w")
        ttk.Button(self.setting_body, text="Save Settings", command=self.save_settings).grid(row=len(fields) + 2, column=1, sticky="e", pady=12)
        self.setting_body.columnconfigure(1, weight=1)

    def browse(self, key):
        path = filedialog.askdirectory(initialdir=self.setting_vars[key].get() or str(HERE), title="Select folder")
        if path:
            self.setting_vars[key].set(normalize_windows_path(path))

    def save_settings(self):
        try:
            d = dict(self.config_data)
            for k, v in self.setting_vars.items():
                x = v.get().strip()
                if k == "active_folder_max_items" and not x:
                    d[k] = None
                elif k in self.INTS:
                    d[k] = int(x)
                    if d[k] < 0:
                        raise ValueError(f"{k} must be zero or more")
                elif k in self.FLOATS:
                    d[k] = float(x)
                    if d[k] < 0:
                        raise ValueError(f"{k} must be zero or more")
                else:
                    d[k] = x
            d["output_root"] = normalize_windows_path(d.get("output_root", ""))
            d["allowed_output_root"] = normalize_windows_path(d.get("allowed_output_root", ""))
            d["include_subfolders"] = self.rec.get()
            d["dry_run_default"] = self.use_dry.get()
            d["compute_checksums"] = self.checksums.get()
            atomic_json(CONFIG_PATH, d)
            self.config_data = d
            self.invalidate()
            messagebox.showinfo("Saved", f"Saved to {CONFIG_PATH}")
        except Exception as e:
            messagebox.showerror("Invalid settings", str(e))

    # ---------- database tab ----------
    def on_tab_changed(self, event=None):
        try:
            if self.nb.tab(self.nb.select(), "text") == "Database":
                self.load_database_tree()
        except Exception as exc:
            self.show_text(self.db_text, "Database preview error: " + str(exc))

    @staticmethod
    def show_text(widget, text):
        widget.config(state="normal")
        widget.delete("1.0", "end")
        widget.insert("end", text)
        widget.config(state="disabled")

    def show_db_text(self, text):
        self.show_text(self.db_text, text)

    def load_database_tree(self):
        if not DB_PATH.exists():
            return self.show_db_text("Database not found: " + str(DB_PATH))
        con = sqlite3.connect(f"file:{DB_PATH.as_posix()}?mode=ro", uri=True)
        try:
            schema = inspect_export_schema(con)
            lines = ["DATABASE OVERVIEW", "Path: " + str(DB_PATH), "Tables: " + (", ".join(schema["tables"]) or "None"),
                     "Active table: " + str(schema["table"]), "Columns: " + (", ".join(schema["columns"]) or "None"), "", "WORKING TREE"]
            if not schema["table"]:
                lines.append("No tables were found.")
                return self.show_db_text("\n".join(lines))
            table = quote_identifier(schema["table"])
            lines.append(f"Total records: {con.execute('SELECT COUNT(*) FROM ' + table).fetchone()[0]:,}")
            src, dest, status = schema["source"], schema["destination"], schema["status"]
            if src:
                q = f"SELECT {quote_identifier(src)},COUNT(*) FROM {table} GROUP BY 1 ORDER BY 1"
                lines += ["", "Source folders:"] + ["  [" + str(r[1]) + "] " + str(r[0] or "(blank)") for r in con.execute(q)]
            if status:
                q = f"SELECT {quote_identifier(status)},COUNT(*) FROM {table} GROUP BY 1 ORDER BY 2 DESC"
                lines += ["", "Statuses:"] + ["  " + str(r[0] or "(blank)") + ": " + str(r[1]) for r in con.execute(q)]
            if "export_runs" in schema["tables"]:
                lines += ["", "Recent runs:"] + [f"  {r[0]}  {r[1]}  {r[2] or 'NOT FINISHED'}" for r in
                                                 con.execute("SELECT run_id, mode, overall_status FROM export_runs ORDER BY started DESC LIMIT 10")]
            if dest:
                q = f"SELECT {quote_identifier(dest)} FROM {table} WHERE {quote_identifier(dest)} IS NOT NULL AND {quote_identifier(dest)}<>'' LIMIT 20"
                lines += ["", "Sample destinations (first 20):"] + ["  " + normalize_windows_path(r[0]) for r in con.execute(q)]
            self.show_db_text("\n".join(lines))
        except Exception as exc:
            self.show_db_text("Database preview failed:\n" + str(exc))
        finally:
            con.close()

    def db_update(self, apply):
        old, new = self.db_old.get().strip(), self.db_new.get().strip()
        if not old or not new:
            return messagebox.showwarning("Missing values", "Enter existing and replacement text.")
        if self.busy:
            return messagebox.showwarning("Run active", "Wait for the current run to finish.")
        old, new = old.replace("/", "\\"), new.replace("/", "\\")
        con = sqlite3.connect(DB_PATH)
        con.row_factory = sqlite3.Row
        try:
            schema = inspect_export_schema(con)
            table = schema["table"]
            if not table:
                raise RuntimeError("No usable database table was found.")
            fields = [field for field in (schema["source"], schema["destination"]) if field]
            if not fields:
                raise RuntimeError("No compatible source or destination path columns were found. Columns present: " + ", ".join(schema["columns"]))
            table_q = quote_identifier(table)
            where = " OR ".join(quote_identifier(f) + " LIKE ?" for f in fields)
            params = tuple("%" + old + "%" for _ in fields)
            rows = con.execute("SELECT " + ",".join(quote_identifier(f) for f in fields) + " FROM " + table_q + " WHERE " + where + " LIMIT 1000", params).fetchall()
            preview = ["DATABASE RENAME PREVIEW", "Table: " + table, "Matched rows shown: " + str(len(rows)), ""]
            preview += [" | ".join(str(row[f] or "").replace(old, new) for f in fields) for row in rows]
            self.show_db_text("\n".join(preview))
            if apply:
                if not messagebox.askyesno("Confirm update", "Update the matched database paths? This does not move files or rename Outlook folders."):
                    return
                con.close()
                backup = backup_database(DB_PATH, BACKUP_DIR, "pre_rename")
                con = sqlite3.connect(DB_PATH)
                con.execute("BEGIN")
                sets = ",".join(quote_identifier(f) + "=REPLACE(" + quote_identifier(f) + ",?,?)" for f in fields)
                set_params = [p for _ in fields for p in (old, new)]
                cursor = con.execute("UPDATE " + table_q + " SET " + sets + " WHERE " + where, tuple(set_params) + params)
                changed = cursor.rowcount
                if "export_manifest" in schema["tables"]:
                    for col in ("expected_path", "actual_path", "source_folder"):
                        con.execute(f"UPDATE export_manifest SET {col}=REPLACE({col},?,?) WHERE {col} LIKE ?", (old, new, "%" + old + "%"))
                integrity = con.execute("PRAGMA integrity_check").fetchone()[0]
                if integrity != "ok":
                    raise RuntimeError("Database integrity check failed after update: " + str(integrity))
                con.commit()
                messagebox.showinfo("Updated", "Rows updated: " + str(changed) + "\nBackup: " + str(backup))
                self.load_database_tree()
        except Exception as exc:
            try:
                con.rollback()
            except Exception:
                pass
            messagebox.showerror("Database operation failed", str(exc))
        finally:
            con.close()

    # ---------- runs ----------
    def invalidate(self):
        self.dry_token = None
        self.status.config(text="Dry run recommended" if self.use_dry.get() else "Ready")

    def log(self, level, msg):
        self.events.put(("log", level, msg))

    def progress_event(self, *args):
        self.events.put(("progress",) + args)

    def toggle_pause(self):
        self.pause_event.clear() if self.pause_event.is_set() else self.pause_event.set()
        self.pause.config(text="Resume" if self.pause_event.is_set() else "Pause")

    def request_stop(self):
        if self.busy and messagebox.askyesno("Stop safely", "Stop after the current item?"):
            self.stop_event.set()
            self.pause_event.clear()
            self.stop.config(state="disabled")
            self.status.config(text="Stopping safely...")

    def start(self, dry, mode="export"):
        if self.busy:
            return
        token = (self.mode.get(), json.dumps(self.config_data, sort_keys=True))
        if mode == "export" and not dry and self.use_dry.get() and self.dry_token != token:
            return messagebox.showwarning("Dry run required", "Run a successful dry run or turn off the option.")
        prompts = {"export": "Start the backup now?", "retry_failed": "Retry only items whose last export failed (retryable errors)?",
                   "retry_missing": "Export items never exported, and re-export items whose file is missing, empty or corrupt?\n\nFailed items are left for Retry Failed."}
        if not dry and mode in prompts and not messagebox.askyesno("Confirm " + MODE_LABELS[mode], prompts[mode]):
            return
        self.busy = True
        self.stop_event.clear()
        self.pause_event.clear()
        self.run_started = time.monotonic()
        self.bar["value"] = 0
        for button in self.run_buttons:
            button.config(state="disabled")
        self.pause.config(state="normal")
        self.stop.config(state="normal")
        threading.Thread(target=self.worker, args=(dry, mode), daemon=True).start()

    def worker(self, dry, mode):
        try:
            self.events.put(("done", dry, execute(self.config_data, self.mode.get(), dry, self.log, self.stop_event,
                                                  self.pause_event, self.progress_event, mode)))
        except Exception as e:
            self.events.put(("failed", exception_details(e)))

    def run_offline_check(self):
        if self.busy:
            return messagebox.showwarning("Run active", "Wait for the current run to finish.")
        try:
            result = offline_reconcile(self.config_data)
            self.show_text(self.report_text, result["text"] + "\n\nCSV: " + result["csv"] + "\nJSON: " + result["json"])
        except Exception as exc:
            messagebox.showerror("Check failed", str(exc))

    @staticmethod
    def duration(v):
        if v is None:
            return "--:--:--"
        v = max(0, int(v))
        return f"{v // 3600:02d}:{v % 3600 // 60:02d}:{v % 60:02d}"

    def clock_tick(self):
        if self.busy and self.run_started:
            self.elapsed_lbl.config(text="Elapsed " + self.duration(time.monotonic() - self.run_started))
        self.after(250, self.clock_tick)

    def poll(self):
        try:
            while True:
                event = self.events.get_nowait()
                kind = event[0]
                if kind == "log":
                    level, msg = event[1:]
                    tag = "ERROR" if level == "ERROR" else "WARNING" if level == "WARNING" else "SUCCESS" if "EXPORTED" in msg else "INFO"
                    self.console.config(state="normal")
                    self.console.insert("end", msg.split(" | ", 2)[-1] + "\n", tag)
                    self.console.see("end")
                    self.console.config(state="disabled")
                elif kind == "progress":
                    processed, total, elapsed, eta, status = event[1:]
                    pct = processed / total * 100 if total else 0
                    self.bar["value"] = pct
                    self.count_lbl.config(text=f"{processed:,} / {total:,} items ({pct:.1f}%)")
                    self.elapsed_lbl.config(text="Elapsed " + self.duration(elapsed))
                    self.eta_lbl.config(text="ETA " + self.duration(eta))
                    self.time_lbl.config(text="Last update: " + datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
                    self.status.config(text=status.title())
                elif kind == "done":
                    dry, r = event[1:]
                    self.finish()
                    if dry and r["result"] == "success":
                        self.dry_token = (self.mode.get(), json.dumps(self.config_data, sort_keys=True))
                    self.show_text(self.report_text, r.get("reconciliation_text", ""))
                    lines = [f"Mode: {MODE_LABELS.get(r['mode'], r['mode'])}{' (dry run)' if r['dry_run'] else ''}",
                             f"Run status: {r['overall_status']}   Reconciliation: {r.get('reconciliation_status', '-')}", "",
                             f"Discovered: {r['items_seen']:,}   Eligible: {r['eligible']:,}",
                             f"Exported: {r['exported']:,}   Adopted: {r['adopted']:,}   Would export: {r['proposed']:,}",
                             f"Already exported (file verified): {r['duplicates_verified']:,}   (file missing): {r['duplicates_missing_file']:,}",
                             f"Skipped: {r['not_selected'] + r['skipped_limit']:,}   Not yet exported: {r['missing_from_db']:,}",
                             f"Failed: {r['failed']:,}   Unreadable: {r['unreadable']:,}   Stopped: {r['stopped_unprocessed']:,}",
                             "", "Report: " + str(r["excel_report"]), "Failures: " + str(r["failures_csv"])]
                    (messagebox.showinfo if r["overall_status"] == "complete" else messagebox.showwarning)("Run complete", "\n".join(lines))
                elif kind == "failed":
                    self.finish()
                    messagebox.showerror("Run failed", event[1]["message"])
        except queue.Empty:
            pass
        self.after(100, self.poll)

    def finish(self):
        self.busy = False
        self.pause_event.clear()
        for button in self.run_buttons:
            button.config(state="normal")
        self.pause.config(state="disabled", text="Pause")
        self.stop.config(state="disabled")

    def close_safely(self):
        if self.busy:
            return messagebox.showwarning("Run active", "Use Stop Safely first.")
        close_log_handlers()
        self.destroy()


def startup_health_check():
    """Migrate (with backup) before any Outlook scan, then confirm the folder and DB are writable/healthy."""
    ensure_runtime()
    load_config()
    migrate(DB_PATH, BACKUP_DIR)
    test = HERE / f".startup_{uuid.uuid4().hex}"
    test.write_text("ok")
    test.unlink()
    with sqlite3.connect(DB_PATH) as connection:
        result = connection.execute("PRAGMA integrity_check").fetchone()[0]
    if result != "ok":
        raise RuntimeError(f"Database integrity check failed: {result}")


if __name__ == "__main__":
    try:
        startup_health_check()
        App().mainloop()
    except Exception as exc:
        ensure_runtime()
        (LOG_DIR / f"startup_{datetime.now():%Y%m%d_%H%M%S}.json").write_text(json.dumps(exception_details(exc), indent=2))
        messagebox.showerror("Startup failed", str(exc))
