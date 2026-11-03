from __future__ import annotations

import csv
import gc
import hashlib
import json
import logging
import ntpath
from logging.handlers import RotatingFileHandler
import os
import queue
import re
import shutil
import sqlite3
import sys
import tempfile
import threading
import time
import traceback
import uuid
from collections import Counter
from dataclasses import asdict, dataclass
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

import pythoncom
import win32com.client
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

APP = "Inbox Email Exporter"
VERSION = "1.5.2"
OL_MSG_UNICODE = 9

HERE = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
CONFIG_PATH = HERE / "config.json"
DB_PATH = HERE / "outlook_export.db"
LOG_DIR = HERE / "logs"
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
    "verify_existing_files_on_start": False
}

# Outlook OlObjectClass values commonly encountered in folders.
OBJECT_CLASS_NAMES = {
    26: "AppointmentItem", 40: "ContactItem", 41: "DocumentItem",
    42: "JournalItem", 43: "MailItem", 44: "NoteItem", 45: "PostItem",
    46: "ReportItem", 47: "RemoteItem", 48: "TaskItem",
    49: "TaskRequestItem", 50: "TaskRequestUpdateItem",
    51: "TaskRequestAcceptItem", 52: "TaskRequestDeclineItem",
    53: "MeetingRequest", 54: "MeetingCancellation",
    55: "MeetingResponseNegative", 56: "MeetingResponsePositive",
    57: "MeetingResponseTentative", 68: "SharingItem"
}

STATUS_COLORS = {
    "exported": "C6EFCE", "duplicate-skipped": "DDEBF7",
    "would-export": "FFF2CC", "failed": "FFC7CE",
    "unreadable": "F4B183", "stopped-unprocessed": "E7E6E6"
}


def ensure_runtime() -> None:
    HERE.mkdir(parents=True, exist_ok=True)
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    if not LOG_DIR.is_dir():
        raise RuntimeError(f"Logs path is not a directory: {LOG_DIR}")


def normalize_windows_path(value: str) -> str:
    """Return a display/storage path using Windows backslashes.

    Preserve UNC prefixes and drive letters. This is for filesystem paths only,
    not Outlook folder paths.
    """
    value=str(value or "").strip().replace("/", "\\")
    if not value:return value
    return ntpath.normpath(value)


def quote_identifier(name: str) -> str:
    return '"'+str(name).replace('"','""')+'"'


def inspect_export_schema(connection):
    tables=[row[0] for row in connection.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name")]
    table='exports' if 'exports' in tables else (tables[0] if tables else None)
    if not table:return {"tables":tables,"table":None,"columns":[],"source":None,"destination":None,"status":None}
    columns=[row[1] for row in connection.execute("PRAGMA table_info("+quote_identifier(table)+")")]
    def first(*names):return next((name for name in names if name in columns),None)
    return {"tables":tables,"table":table,"columns":columns,
            "source":first("source_folder","folder","folder_path","source"),
            "destination":first("destination","path","file_path","output_path"),
            "status":first("status","state","result")}


def atomic_json(path: Path, data: dict) -> None:
    temp = path.with_suffix(path.suffix + ".tmp")
    temp.write_text(json.dumps(data, indent=2), encoding="utf-8")
    os.replace(temp, path)


def load_config() -> dict:
    ensure_runtime()
    if not CONFIG_PATH.exists():
        atomic_json(CONFIG_PATH, DEFAULT_CONFIG)
    data={**DEFAULT_CONFIG, **json.loads(CONFIG_PATH.read_text(encoding="utf-8"))}
    if "active_folder_max_emails" in data and "active_folder_max_items" not in data: data["active_folder_max_items"]=data["active_folder_max_emails"]
    for key in ("output_root","allowed_output_root"):
        data[key]=normalize_windows_path(data.get(key,""))
    return data


def clean_name(value, limit=90) -> str:
    value = re.sub(r'[\\/:*?"<>|]', "_", str(value or "No Subject"))
    value = re.sub(r"[\x00-\x1f]", " ", value)
    value = re.sub(r"\s+", " ", value).strip().rstrip(". ")
    return value[:limit].rstrip(". ") or "No Subject"


def entry_hash(value: str) -> str:
    return hashlib.sha256(value.encode("utf-8", "replace")).hexdigest()[:12]


def safe_property(item, name, default=""):
    try:
        return getattr(item, name)
    except Exception:
        return default


def item_datetime(item) -> datetime:
    for name in ("ReceivedTime", "SentOn", "Start", "CreationTime", "LastModificationTime"):
        value = safe_property(item, name, None)
        if value:
            return value
    return datetime.now()


def exception_details(exc: Exception) -> dict:
    details = {"type": type(exc).__name__, "message": str(exc), "repr": repr(exc), "traceback": traceback.format_exc()}
    for name in ("hresult", "strerror", "excepinfo", "argerror"):
        if hasattr(exc, name):
            details[name] = repr(getattr(exc, name))
    return details


def is_memory_error(exc: Exception) -> bool:
    text = (str(exc) + repr(exc)).lower()
    return any(value in text for value in ("not enough memory", "-2147024882", "0x8007000e", "4096"))


def close_log_handlers() -> None:
    logger = logging.getLogger("agency-export")
    for handler in list(logger.handlers):
        try:
            handler.flush(); handler.close()
        finally:
            logger.removeHandler(handler)
    logging.shutdown(); gc.collect()


def create_logger(run_id: str, ui_callback):
    ensure_runtime(); close_log_handlers(); ensure_runtime()
    logger = logging.getLogger("agency-export")
    logger.setLevel(logging.DEBUG)
    formatter = logging.Formatter("%(asctime)s | %(levelname)s | %(message)s")
    for handler in (
        RotatingFileHandler(LOG_DIR / "application.log", maxBytes=10_000_000, backupCount=10, encoding="utf-8"),
        logging.FileHandler(LOG_DIR / f"run_{run_id}.log", encoding="utf-8")
    ):
        handler.setFormatter(formatter); logger.addHandler(handler)

    class UIHandler(logging.Handler):
        def emit(self, record):
            message = self.format(record)
            important = ("RUN ", "FOLDER ", "EXPORTED ", "FAILED ", "PAUSED", "RESUMED", "STOPPING", "REPORT ")
            if record.levelno >= logging.WARNING or any(token in message for token in important):
                ui_callback(record.levelname, message)

    ui = UIHandler(); ui.setFormatter(formatter); logger.addHandler(ui)
    return logger


def pid_running(pid: int) -> bool:
    if not pid: return False
    import ctypes
    handle = ctypes.windll.kernel32.OpenProcess(0x1000, False, int(pid))
    if handle:
        ctypes.windll.kernel32.CloseHandle(handle); return True
    return False


def acquire_lock() -> None:
    ensure_runtime()
    if LOCK_PATH.exists():
        try: info = json.loads(LOCK_PATH.read_text(encoding="utf-8"))
        except Exception: info = {"pid": 0}
        pid = int(info.get("pid") or 0)
        if pid_running(pid) and pid != os.getpid():
            raise RuntimeError(f"Another exporter is active (PID {pid}).")
        LOCK_PATH.unlink(missing_ok=True)
    fd = os.open(str(LOCK_PATH), os.O_CREAT | os.O_EXCL | os.O_WRONLY)
    try: os.write(fd, json.dumps({"pid": os.getpid(), "created": datetime.now().isoformat()}).encode())
    finally: os.close(fd)


def backup_and_migrate_database():
    """Upgrade any older exports table before a run can begin.

    Migration is additive, transactional, and backed up first. Older columns
    (`folder`, `size`) remain in place for compatibility and are copied into
    the newer audit columns (`source_folder`, `file_size`).
    """
    ensure_runtime()
    connection=sqlite3.connect(DB_PATH,timeout=30)
    try:
        connection.execute("PRAGMA busy_timeout=30000")
        integrity=connection.execute("PRAGMA integrity_check").fetchone()[0]
        if integrity!="ok":raise RuntimeError(f"Database integrity check failed before migration: {integrity}")
        exists=connection.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name='exports'").fetchone()
        if not exists:
            connection.execute("""CREATE TABLE exports(
                store_id TEXT NOT NULL,entry_id TEXT NOT NULL,item_class INTEGER,
                item_type TEXT,source_folder TEXT,destination TEXT,file_size INTEGER,
                status TEXT,updated TEXT,error TEXT,PRIMARY KEY(store_id,entry_id))""")
            connection.execute("CREATE INDEX IF NOT EXISTS idx_exports_status ON exports(status)")
            connection.commit();return {"created":True,"added":[]}
        columns=[row[1] for row in connection.execute("PRAGMA table_info(exports)")]
        additions=[]
        required={"item_class":"INTEGER","item_type":"TEXT","source_folder":"TEXT","destination":"TEXT","file_size":"INTEGER","status":"TEXT","updated":"TEXT","error":"TEXT"}
        missing=[(name,kind) for name,kind in required.items() if name not in columns]
        if missing:
            stamp=datetime.now().strftime("%Y%m%d_%H%M%S")
            backup=DB_PATH.with_name(f"outlook_export_pre_migration_{stamp}.backup.db")
            connection.commit();shutil.copy2(DB_PATH,backup)
            connection.execute("BEGIN IMMEDIATE")
            for name,kind in missing:
                connection.execute(f'ALTER TABLE exports ADD COLUMN "{name}" {kind}')
                additions.append(name)
            if "folder" in columns and "source_folder" in additions:
                connection.execute("UPDATE exports SET source_folder=folder WHERE source_folder IS NULL OR source_folder='' ")
            if "size" in columns and "file_size" in additions:
                connection.execute("UPDATE exports SET file_size=size WHERE file_size IS NULL")
            connection.execute("CREATE INDEX IF NOT EXISTS idx_exports_status ON exports(status)")
            integrity=connection.execute("PRAGMA integrity_check").fetchone()[0]
            if integrity!="ok":raise RuntimeError(f"Database integrity check failed during migration: {integrity}")
            connection.commit()
            return {"created":False,"added":additions,"backup":str(backup)}
        connection.execute("CREATE INDEX IF NOT EXISTS idx_exports_status ON exports(status)")
        connection.commit();return {"created":False,"added":[]}
    except Exception:
        connection.rollback();raise
    finally:connection.close()


class History:
    """Persistent history with preflight schema migration and cached duplicates."""
    def __init__(self,batch_size:int):
        backup_and_migrate_database()
        self.connection=sqlite3.connect(DB_PATH,timeout=30)
        self.connection.execute("PRAGMA journal_mode=DELETE")
        self.connection.execute("PRAGMA synchronous=FULL")
        self.connection.execute("PRAGMA busy_timeout=30000")
        self.completed=set(self.connection.execute("SELECT store_id,entry_id FROM exports WHERE status='completed'"))
        self.batch_size=max(1,int(batch_size));self.pending=0
    def is_duplicate(self,store_id,entry_id):return (store_id,entry_id) in self.completed
    def success(self,store_id,entry_id,item_class,item_type,folder,destination,size):
        self.connection.execute("""INSERT INTO exports(
            store_id,entry_id,item_class,item_type,source_folder,destination,file_size,status,updated,error)
            VALUES(?,?,?,?,?,?,?,?,?,?)
            ON CONFLICT(store_id,entry_id) DO UPDATE SET
            item_class=excluded.item_class,item_type=excluded.item_type,
            source_folder=excluded.source_folder,destination=excluded.destination,
            file_size=excluded.file_size,status='completed',updated=excluded.updated,error=''""",
            (store_id,entry_id,item_class,item_type,folder,str(destination),size,"completed",datetime.now().isoformat(),""))
        self.completed.add((store_id,entry_id));self.pending+=1;self._maybe_commit()
    def failure(self,store_id,entry_id,item_class,item_type,folder,error):
        self.connection.execute("""INSERT INTO exports(
            store_id,entry_id,item_class,item_type,source_folder,status,updated,error)
            VALUES(?,?,?,?,?,'failed',?,?)
            ON CONFLICT(store_id,entry_id) DO UPDATE SET
            item_class=excluded.item_class,item_type=excluded.item_type,
            source_folder=excluded.source_folder,status='failed',updated=excluded.updated,error=excluded.error""",
            (store_id,entry_id,item_class,item_type,folder,datetime.now().isoformat(),error))
        self.pending+=1;self._maybe_commit()
    def _maybe_commit(self):
        if self.pending>=self.batch_size:self.connection.commit();self.pending=0
    def close(self):self.connection.commit();self.connection.close()

@dataclass
class Counts:
    folders: int = 0; items_seen: int = 0; items_readable: int = 0
    exported: int = 0; proposed: int = 0; duplicates: int = 0
    failed: int = 0; unreadable: int = 0; stopped_unprocessed: int = 0


def validate_output(config: dict) -> Path:
    output = Path(os.path.expandvars(config["output_root"])).resolve()
    allowed = Path(os.path.expandvars(config["allowed_output_root"])).resolve()
    if not (output == allowed or allowed in output.parents): raise RuntimeError("Output is outside approved root")
    if not output.exists(): raise RuntimeError(f"Output unavailable: {output}")
    if shutil.disk_usage(output).free / 1048576 < float(config["minimum_free_space_mb"]): raise RuntimeError("Insufficient output free space")
    test = output / f".write_test_{uuid.uuid4().hex}"; test.write_text("ok"); test.unlink()
    return output


def configured_folder(namespace, config):
    root = None; available = []
    for index in range(1, namespace.Folders.Count + 1):
        candidate = namespace.Folders.Item(index); available.append(str(candidate.Name))
        if str(candidate.Name).casefold() == config["mailbox_name"].casefold(): root = candidate
    if root is None: raise RuntimeError(f"Mailbox not found. Available: {available}")
    for part in [p.strip() for p in config["folder_path"].split("\\") if p.strip()]: root = root.Folders.Item(part)
    return root


def folder_tree(root, recursive):
    def walk(folder, relative):
        yield folder, relative
        if recursive:
            for index in range(1, folder.Folders.Count + 1):
                child = folder.Folders.Item(index)
                yield from walk(child, relative + (clean_name(child.Name),))
    yield from walk(root, ())


def snapshot_items(folder, logger):
    """Read every item ID/class once; release COM proxies immediately."""
    rows = []; items = folder.Items; total = items.Count
    for index in range(1, total + 1):
        item = None
        try:
            item = items.Item(index)
            entry_id = str(item.EntryID)
            item_class = int(safe_property(item, "Class", -1))
            rows.append((entry_id, item_class, index, ""))
        except Exception as exc:
            rows.append(("", -1, index, str(exc)))
            logger.exception("FAILED SNAPSHOT | folder=%s | index=%s", safe_property(folder,"FolderPath",""), index)
        finally:
            item = None
            if index % 100 == 0: gc.collect()
    items = None; gc.collect()
    return rows, total


def destination_name(item, destination_dir, entry_id, max_path):
    stamp = item_datetime(item); item_type = OBJECT_CLASS_NAMES.get(int(safe_property(item,"Class",-1)), f"Class{safe_property(item,'Class',-1)}")
    subject = clean_name(safe_property(item,"Subject", safe_property(item,"FullName", item_type)), 70)
    suffix = f" - {item_type} - {entry_hash(entry_id)}.msg"; prefix = f"{stamp:%Y-%m-%d_%H%M%S} - "
    available = max(10, int(max_path) - len(str(destination_dir)) - len(prefix) - len(suffix) - 1)
    return prefix + clean_name(subject, available) + suffix


def save_outlook_item(item, destination, config, logger):
    destination.parent.mkdir(parents=True, exist_ok=True)
    temp_dir = Path(tempfile.gettempdir()) / "InboxEmailExporter"; temp_dir.mkdir(exist_ok=True)
    temp = temp_dir / f"{uuid.uuid4().hex}.msg"; last = None
    try:
        for attempt in range(1, int(config["save_retries"]) + 1):
            try:
                temp.unlink(missing_ok=True); item.SaveAs(str(temp), OL_MSG_UNICODE)
                if not temp.exists() or temp.stat().st_size <= 0: raise RuntimeError("SaveAs created a missing or empty file")
                break
            except Exception as exc:
                last = exc; wait = float(config["memory_error_delay_seconds"] if is_memory_error(exc) else config["retry_delay_seconds"])
                logger.warning("SAVE RETRY | attempt=%s | wait=%s | error=%s", attempt, wait, exc); gc.collect(); time.sleep(wait)
        else: raise last
        shutil.move(str(temp), str(destination))
        if not destination.exists() or destination.stat().st_size <= 0: raise RuntimeError("Destination is missing or empty")
        return destination.stat().st_size
    finally: temp.unlink(missing_ok=True)


def write_excel_report(path: Path, summary: dict, rows: list[dict], class_counts: Counter):
    wb = Workbook(); ws = wb.active; ws.title = "Summary"; detail = wb.create_sheet("Item Detail"); classes = wb.create_sheet("Item Types")
    navy="1F4E78"; white="FFFFFF"; teal="DDEBF7"
    ws["A1"]="Agency Outlook Export Report"; ws["A1"].font=Font(size=18,bold=True,color=white); ws["A1"].fill=PatternFill("solid",fgColor=navy); ws.merge_cells("A1:D1")
    summary_rows=[("Run ID",summary["run_id"]),("Source",summary["source"]),("Run type","Dry run" if summary["dry_run"] else "Export"),("Started",summary["started"]),("Finished",summary["finished"]),("Result",summary["result"]),("Folders",summary["folders"]),("Items seen",summary["items_seen"]),("Items readable",summary["items_readable"]),("Exported",summary["exported"]),("Would export",summary["proposed"]),("Duplicates",summary["duplicates"]),("Failed",summary["failed"]),("Unreadable",summary["unreadable"]),("Stopped/unprocessed",summary["stopped_unprocessed"]),("Balanced",summary["balanced"])]
    ws.append([])
    for label,value in summary_rows: ws.append([label,value])
    for cell in ws[2]: cell.fill=PatternFill("solid",fgColor=teal)
    ws.column_dimensions["A"].width=24; ws.column_dimensions["B"].width=90

    headers=["Run ID","Timestamp","Status","Source Folder","Folder Relative Path","Item Index","Entry Hash","Outlook Class","Item Type","Message Class","Subject/Name","Item Date","Destination","File Size Bytes","Duplicate Reason","Error Type","HRESULT","Error Message"]
    detail.append(headers)
    for row in rows: detail.append([row.get(h,"") for h in headers])
    for c in detail[1]: c.font=Font(bold=True,color=white); c.fill=PatternFill("solid",fgColor=navy); c.alignment=Alignment(wrap_text=True)
    if detail.max_row>1:
        table=Table(displayName="ItemAudit",ref=f"A1:R{detail.max_row}"); table.tableStyleInfo=TableStyleInfo(name="TableStyleMedium2",showRowStripes=True,showColumnStripes=False); detail.add_table(table)
        for r in range(2,detail.max_row+1):
            status=detail.cell(r,3).value; color=STATUS_COLORS.get(status)
            if color: detail.cell(r,3).fill=PatternFill("solid",fgColor=color)
    widths=[20,20,22,48,36,12,16,14,28,34,55,22,80,16,35,22,18,80]
    for i,w in enumerate(widths,1): detail.column_dimensions[get_column_letter(i)].width=w
    detail.freeze_panes="A2"

    classes.append(["Outlook Class","Item Type","Count"])
    for key,count in sorted(class_counts.items(), key=lambda x:(x[0][0],x[0][1])): classes.append([key[0],key[1],count])
    for c in classes[1]: c.font=Font(bold=True,color=white); c.fill=PatternFill("solid",fgColor=navy)
    classes.column_dimensions["A"].width=16; classes.column_dimensions["B"].width=35; classes.column_dimensions["C"].width=14
    wb.save(path)


def wait_control(stop_event, pause_event, logger):
    if stop_event.is_set(): return False
    announced = False
    while pause_event.is_set() and not stop_event.is_set():
        if not announced: logger.info("PAUSED"); announced = True
        time.sleep(.25)
    if announced and not stop_event.is_set(): logger.info("RESUMED")
    return not stop_event.is_set()


def execute(config, source_mode, dry_run, ui_callback, stop_event, pause_event, progress_callback=None):
    acquire_lock(); started=datetime.now(); run_id=started.strftime("%Y%m%d_%H%M%S")+"_"+uuid.uuid4().hex[:6]+("_DRY" if dry_run else "_EXPORT")
    logger=create_logger(run_id,ui_callback); output=validate_output(config); history=History(config["database_commit_batch"]); counts=Counts(); audit=[]; class_counts=Counter(); stopped=False
    pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
    outlook=namespace=explorer=root=None
    try:
        outlook=win32com.client.Dispatch("Outlook.Application"); namespace=outlook.GetNamespace("MAPI"); explorer=outlook.ActiveExplorer()
        root=configured_folder(namespace,config) if source_mode=="configured" else explorer.CurrentFolder
        limit=None if source_mode=="configured" else config["active_folder_max_items"]
        folder_list=list(folder_tree(root,bool(config["include_subfolders"])))
        total_expected=sum(int(folder.Items.Count) for folder,_ in folder_list)
        if limit is not None: total_expected=min(total_expected,int(limit))
        progress_started=time.monotonic()
        def progress(status):
            processed=counts.exported+counts.proposed+counts.duplicates+counts.failed+counts.unreadable
            elapsed=max(0.001,time.monotonic()-progress_started); rate=processed/elapsed
            eta=(max(0,total_expected-processed)/rate) if rate>0 else None
            if progress_callback: progress_callback(processed,total_expected,elapsed,eta,status)
        logger.info("RUN START | id=%s | source=%s | loaded_duplicates=%s",run_id,root.FolderPath,len(history.completed))
        for folder,relative in folder_list:
            if stopped or not wait_control(stop_event,pause_event,logger): stopped=True; break
            counts.folders+=1; folder_path=str(folder.FolderPath); store_id=str(folder.StoreID); destination_dir=output/clean_name(root.Name)/Path(*relative)
            snapshot,total=snapshot_items(folder,logger); counts.items_seen+=total; logger.info("FOLDER SCAN | %s | items=%s",folder_path,total)
            for entry_id,item_class,index,snapshot_error in snapshot:
                if limit is not None and counts.items_readable>=int(limit): stopped=True; break
                if not wait_control(stop_event,pause_event,logger): stopped=True; break
                now=datetime.now().isoformat(timespec="seconds"); item_type=OBJECT_CLASS_NAMES.get(item_class,f"Class{item_class}")
                base={"Run ID":run_id,"Timestamp":now,"Status":"","Source Folder":folder_path,"Folder Relative Path":"\\".join(relative),"Item Index":index,"Entry Hash":entry_hash(entry_id) if entry_id else "","Outlook Class":item_class,"Item Type":item_type,"Message Class":"","Subject/Name":"","Item Date":"","Destination":"","File Size Bytes":"","Duplicate Reason":"","Error Type":"","HRESULT":"","Error Message":""}
                if snapshot_error or not entry_id:
                    counts.unreadable+=1; base.update({"Status":"unreadable","Error Type":"SnapshotError","Error Message":snapshot_error}); audit.append(base); class_counts[(item_class,item_type)]+=1; progress("unreadable"); continue
                counts.items_readable+=1; class_counts[(item_class,item_type)]+=1
                # Fast path: in-memory set avoids one SQL query and one file lookup per duplicate.
                if history.is_duplicate(store_id,entry_id):
                    counts.duplicates+=1; base.update({"Status":"duplicate-skipped","Duplicate Reason":"StoreID + EntryID already completed in database"}); audit.append(base); progress("duplicate"); continue
                item=None
                try:
                    item=namespace.GetItemFromID(entry_id,store_id); subject=str(safe_property(item,"Subject",safe_property(item,"FullName",item_type))); date=item_datetime(item); message_class=str(safe_property(item,"MessageClass","")); destination=destination_dir/destination_name(item,destination_dir,entry_id,config["max_windows_path"])
                    base.update({"Message Class":message_class,"Subject/Name":subject,"Item Date":date.isoformat(),"Destination":str(destination)})
                    if dry_run: counts.proposed+=1; base["Status"]="would-export"; audit.append(base); progress("dry-run"); continue
                    size=save_outlook_item(item,destination,config,logger); history.success(store_id,entry_id,item_class,item_type,folder_path,destination,size); counts.exported+=1; base.update({"Status":"exported","File Size Bytes":size}); audit.append(base); logger.info("EXPORTED | type=%s | %s",item_type,destination); progress("exported")
                except Exception as exc:
                    details=exception_details(exc); counts.failed+=1; base.update({"Status":"failed","Error Type":details["type"],"HRESULT":details.get("hresult",""),"Error Message":details["message"]}); audit.append(base); history.failure(store_id,entry_id,item_class,item_type,folder_path,json.dumps(details,default=str)); logger.exception("FAILED ITEM | folder=%s | index=%s | type=%s",folder_path,index,item_type); progress("failed")
                finally: item=None; gc.collect()
                if counts.items_readable%int(config["pause_every"])==0: time.sleep(float(config["pause_seconds"]))
        if stopped:
            counts.stopped_unprocessed=max(0,counts.items_seen-counts.items_readable-counts.unreadable)
            logger.warning("STOPPING SAFELY | unprocessed=%s",counts.stopped_unprocessed)
        processed=counts.exported+counts.proposed+counts.duplicates+counts.failed+counts.unreadable
        balanced=(processed+counts.stopped_unprocessed)==counts.items_seen
        finished=datetime.now(); summary={**asdict(counts),"run_id":run_id,"dry_run":dry_run,"source":str(root.FolderPath),"started":started.isoformat(),"finished":finished.isoformat(),"balanced":balanced,"result":"stopped" if stopped else "success" if not counts.failed and not counts.unreadable and balanced else "review-required"}
        report=LOG_DIR/f"audit_{run_id}.xlsx"; write_excel_report(report,summary,audit,class_counts); summary["excel_report"]=str(report); logger.info("REPORT WRITTEN | %s",report); logger.info("RUN COMPLETE | %s",summary["result"]); return summary
    finally:
        history.close(); root=explorer=namespace=outlook=None; gc.collect(); pythoncom.CoUninitialize(); close_log_handlers(); LOCK_PATH.unlink(missing_ok=True); ensure_runtime()


class App(tk.Tk):
    BASIC_FIELDS=[("output_root","Output folder",True),("allowed_output_root","Approved output root",True),("mailbox_name","Mailbox display name",False),("folder_path","Outlook folder path",False)]
    ADV_FIELDS=[("active_folder_max_items","Active-folder item limit",False),("pause_every","Pause every N items",False),("pause_seconds","Pause seconds",False),("database_commit_batch","Database commit batch",False),("save_retries","Save retries",False),("retry_delay_seconds","Retry delay seconds",False),("memory_error_delay_seconds","Memory-error delay seconds",False),("minimum_free_space_mb","Minimum free space MB",False),("max_windows_path","Maximum Windows path",False)]
    def __init__(self):
        super().__init__();self.title(f"{APP} {VERSION}");self.geometry("1120x790");self.minsize(980,680)
        self.config_data=load_config();self.events=queue.Queue();self.stop_event=threading.Event();self.pause_event=threading.Event();self.busy=False;self.dry_token=None;self.run_started=None
        self.make_ui();self.protocol("WM_DELETE_WINDOW",self.close_safely);self.after(250,self.clock_tick);self.after(100,self.poll)
    def make_ui(self):
        st=ttk.Style(self);st.theme_use("vista" if "vista" in st.theme_names() else "clam");st.configure("Title.TLabel",font=("Segoe UI",20,"bold"),foreground="#0F3B70");st.configure("TButton",padding=8)
        h=ttk.Frame(self,padding=18);h.pack(fill="x");ttk.Label(h,text=APP,style="Title.TLabel").pack(anchor="w");ttk.Label(h,text="Generic Outlook backup with Excel auditing and duplicate protection").pack(anchor="w")
        self.nb=ttk.Notebook(self);self.nb.pack(fill="both",expand=True,padx=18,pady=(0,18));run=ttk.Frame(self.nb,padding=18);settings=ttk.Frame(self.nb,padding=18);db=ttk.Frame(self.nb,padding=18);reports=ttk.Frame(self.nb,padding=18)
        for frame,name in ((run,"Export"),(settings,"Settings"),(db,"Database"),(reports,"Reports")):self.nb.add(frame,text=name)
        self.mode=tk.StringVar(value="configured");box=ttk.LabelFrame(run,text="Source",padding=12);box.pack(fill="x");ttk.Radiobutton(box,text="Configured mailbox and folder",variable=self.mode,value="configured",command=self.invalidate).pack(anchor="w");ttk.Radiobutton(box,text="Currently open Outlook folder",variable=self.mode,value="active",command=self.invalidate).pack(anchor="w")
        self.use_dry=tk.BooleanVar(value=bool(self.config_data["dry_run_default"]));ttk.Checkbutton(box,text="Run dry run first (recommended)",variable=self.use_dry,command=self.invalidate).pack(anchor="w",pady=6)
        a=ttk.Frame(run);a.pack(fill="x",pady=10);self.dry=ttk.Button(a,text="Dry Run",command=lambda:self.start(True));self.dry.pack(side="left");self.export=ttk.Button(a,text="Start Export",command=lambda:self.start(False));self.export.pack(side="left",padx=8);self.pause=ttk.Button(a,text="Pause",command=self.toggle_pause,state="disabled");self.pause.pack(side="left");self.stop=ttk.Button(a,text="Stop Safely",command=self.request_stop,state="disabled");self.stop.pack(side="left",padx=8);self.status=ttk.Label(a,text="Ready");self.status.pack(side="left",padx=12)
        self.bar=ttk.Progressbar(run,mode="determinate",maximum=100);self.bar.pack(fill="x");m=ttk.Frame(run);m.pack(fill="x",pady=5);self.count_lbl=ttk.Label(m,text="0 / 0 items");self.count_lbl.pack(side="left");self.elapsed_lbl=ttk.Label(m,text="Elapsed 00:00:00");self.elapsed_lbl.pack(side="left",padx=20);self.eta_lbl=ttk.Label(m,text="ETA --:--:--");self.eta_lbl.pack(side="left");self.time_lbl=ttk.Label(m,text="Last update: -");self.time_lbl.pack(side="right")
        self.console=tk.Text(run,state="disabled",font=("Segoe UI",9),bg="#F7F9FC",relief="flat");self.console.pack(fill="both",expand=True);[(self.console.tag_configure(t,foreground=c)) for t,c in (("INFO","#245A9A"),("SUCCESS","#107C10"),("WARNING","#986F0B"),("ERROR","#C50F1F"))]
        mode=ttk.Frame(settings);mode.pack(fill="x");self.setting_mode=tk.StringVar(value="basic");ttk.Radiobutton(mode,text="Basic",variable=self.setting_mode,value="basic",command=self.render_settings).pack(side="left");ttk.Radiobutton(mode,text="Advanced",variable=self.setting_mode,value="advanced",command=self.render_settings).pack(side="left",padx=12);self.setting_body=ttk.Frame(settings);self.setting_body.pack(fill="both",expand=True,pady=12);self.setting_vars={};self.render_settings()
        ttk.Label(db,text="Database Folder Mapping",font=("Segoe UI",12,"bold")).pack(anchor="w");ttk.Label(db,text="Preview and update stored source/destination folder names. A database backup is always created before changes.",wraplength=900).pack(anchor="w",pady=5);self.db_old=tk.StringVar();self.db_new=tk.StringVar();f=ttk.Frame(db);f.pack(fill="x",pady=8);ttk.Label(f,text="Existing text").grid(row=0,column=0,sticky="w");ttk.Entry(f,textvariable=self.db_old,width=75).grid(row=0,column=1,sticky="ew",padx=8);ttk.Label(f,text="Replacement text").grid(row=1,column=0,sticky="w");ttk.Entry(f,textvariable=self.db_new,width=75).grid(row=1,column=1,sticky="ew",padx=8);f.columnconfigure(1,weight=1);b=ttk.Frame(db);b.pack(fill="x");ttk.Button(b,text="Refresh Database Tree",command=self.load_database_tree).pack(side="left");ttk.Button(b,text="Preview Rename",command=lambda:self.db_update(False)).pack(side="left",padx=8);ttk.Button(b,text="Apply Update",command=lambda:self.db_update(True)).pack(side="left");self.db_text=tk.Text(db,height=20,state="disabled",font=("Consolas",9));self.db_text.pack(fill="both",expand=True,pady=8)
        ttk.Button(reports,text="Open Logs and Excel Reports",command=lambda:(ensure_runtime(),os.startfile(LOG_DIR))).pack(anchor="w")
        ttk.Label(reports,text="Reports: "+str(LOG_DIR)+"\nDatabase: "+str(DB_PATH),justify="left").pack(anchor="w",pady=10)
        self.nb.bind("<<NotebookTabChanged>>",self.on_tab_changed);self.after(350,self.load_database_tree)
    def render_settings(self):
        [w.destroy() for w in self.setting_body.winfo_children()];fields=self.BASIC_FIELDS+(self.ADV_FIELDS if self.setting_mode.get()=="advanced" else [])
        for r,(key,label,browse) in enumerate(fields):
            ttk.Label(self.setting_body,text=label).grid(row=r,column=0,sticky="w",pady=6);var=self.setting_vars.setdefault(key,tk.StringVar(value="" if self.config_data.get(key) is None else str(self.config_data.get(key,""))));ttk.Entry(self.setting_body,textvariable=var,width=72).grid(row=r,column=1,sticky="ew",padx=8,pady=6)
            if browse:ttk.Button(self.setting_body,text="Browse...",command=lambda k=key:self.browse(k)).grid(row=r,column=2,pady=6)
        self.rec=tk.BooleanVar(value=self.config_data["include_subfolders"]);ttk.Checkbutton(self.setting_body,text="Include subfolders",variable=self.rec).grid(row=len(fields),column=1,sticky="w");ttk.Button(self.setting_body,text="Save Settings",command=self.save_settings).grid(row=len(fields)+1,column=1,sticky="e",pady=12);self.setting_body.columnconfigure(1,weight=1)
    def browse(self,key):
        start=self.setting_vars[key].get() or str(HERE)
        path=filedialog.askdirectory(initialdir=start,title="Select folder")
        if path:self.setting_vars[key].set(normalize_windows_path(path))
    def save_settings(self):
        try:
            d=dict(self.config_data);ints={"active_folder_max_items","pause_every","database_commit_batch","save_retries","minimum_free_space_mb","max_windows_path"};floats={"pause_seconds","retry_delay_seconds","memory_error_delay_seconds"}
            for k,v in self.setting_vars.items():
                x=v.get().strip();d[k]=None if k=="active_folder_max_items" and not x else int(x) if k in ints else float(x) if k in floats else x
            d["output_root"]=normalize_windows_path(d.get("output_root",""));d["allowed_output_root"]=normalize_windows_path(d.get("allowed_output_root",""));d["include_subfolders"]=self.rec.get();d["dry_run_default"]=self.use_dry.get();atomic_json(CONFIG_PATH,d);self.config_data=d;self.invalidate();messagebox.showinfo("Saved",f"Saved to {CONFIG_PATH}")
        except Exception as e:messagebox.showerror("Invalid settings",str(e))
    def on_tab_changed(self,event=None):
        try:
            if self.nb.tab(self.nb.select(),"text")=="Database":self.load_database_tree()
        except Exception as exc:self.show_db_text("Database preview error: "+str(exc))
    def show_db_text(self,text):
        self.db_text.config(state="normal");self.db_text.delete("1.0","end");self.db_text.insert("end",text);self.db_text.config(state="disabled")
    def load_database_tree(self):
        if not DB_PATH.exists():return self.show_db_text("Database not found: "+str(DB_PATH))
        con=sqlite3.connect(DB_PATH);con.row_factory=sqlite3.Row
        try:
            schema=inspect_export_schema(con)
            lines=["DATABASE OVERVIEW", "Path: "+str(DB_PATH), "Tables: "+(", ".join(schema["tables"]) or "None"), "Active table: "+str(schema["table"]), "Columns: "+(", ".join(schema["columns"]) or "None"), "", "WORKING TREE"]
            if not schema["table"]:
                lines.append("No tables were found.");return self.show_db_text("\n".join(lines))
            table=quote_identifier(schema["table"]);total=con.execute("SELECT COUNT(*) FROM "+table).fetchone()[0];lines.append("Total records: "+f"{total:,}")
            src=schema["source"];dest=schema["destination"];status=schema["status"]
            if src:
                q="SELECT "+quote_identifier(src)+",COUNT(*) AS n FROM "+table+" GROUP BY "+quote_identifier(src)+" ORDER BY "+quote_identifier(src)
                lines += ["", "Source folders:"]+["  ["+str(r[1])+"] "+str(r[0] or "(blank)") for r in con.execute(q).fetchall()]
            else:lines += ["", "Source-folder column not found. Compatible names checked: source_folder, folder, folder_path, source."]
            if status:
                q="SELECT "+quote_identifier(status)+",COUNT(*) FROM "+table+" GROUP BY "+quote_identifier(status)+" ORDER BY COUNT(*) DESC"
                lines += ["", "Statuses:"]+["  "+str(r[0] or "(blank)")+": "+str(r[1]) for r in con.execute(q).fetchall()]
            if dest:
                q="SELECT "+quote_identifier(dest)+" FROM "+table+" WHERE "+quote_identifier(dest)+" IS NOT NULL AND "+quote_identifier(dest)+"<>'' LIMIT 20"
                lines += ["", "Sample destinations (first 20):"]+["  "+normalize_windows_path(r[0]) for r in con.execute(q).fetchall()]
            else:lines += ["", "Destination column not found. Compatible names checked: destination, path, file_path, output_path."]
            self.show_db_text("\n".join(lines))
        except Exception as exc:self.show_db_text("Database preview failed:\n"+str(exc))
        finally:con.close()
    def db_update(self,apply):
        old=self.db_old.get().strip();new=self.db_new.get().strip()
        if not old or not new:return messagebox.showwarning("Missing values","Enter existing and replacement text.")
        if self.busy:return messagebox.showwarning("Run active","Wait for the current run to finish.")
        # Standardize filesystem-style replacements while retaining Outlook backslashes.
        old=old.replace("/","\\");new=new.replace("/","\\")
        con=sqlite3.connect(DB_PATH);con.row_factory=sqlite3.Row
        try:
            schema=inspect_export_schema(con);table=schema["table"]
            if not table:raise RuntimeError("No usable database table was found.")
            fields=[field for field in (schema["source"],schema["destination"]) if field]
            if not fields:raise RuntimeError("No compatible source or destination path columns were found. Columns present: "+", ".join(schema["columns"]))
            table_q=quote_identifier(table);where=" OR ".join(quote_identifier(f)+" LIKE ?" for f in fields);params=tuple("%"+old+"%" for _ in fields)
            select=",".join(quote_identifier(f) for f in fields);rows=con.execute("SELECT "+select+" FROM "+table_q+" WHERE "+where+" LIMIT 1000",params).fetchall()
            preview=["DATABASE RENAME PREVIEW","Table: "+table,"Matched rows shown: "+str(len(rows)),""]
            for row in rows:
                preview.append(" | ".join(str(row[f] or "").replace(old,new) for f in fields))
            self.show_db_text("\n".join(preview))
            if apply:
                if not messagebox.askyesno("Confirm update","Update the matched database paths? This does not move files or rename Outlook folders."):return
                backup=DB_PATH.with_name(f"outlook_export_{datetime.now():%Y%m%d_%H%M%S}.backup.db");shutil.copy2(DB_PATH,backup);con.execute("BEGIN")
                sets=",".join(quote_identifier(f)+"=REPLACE("+quote_identifier(f)+",?,?)" for f in fields);set_params=[]
                for _ in fields:set_params.extend((old,new))
                con.execute("UPDATE "+table_q+" SET "+sets+" WHERE "+where,tuple(set_params)+params);changed=con.total_changes
                integrity=con.execute("PRAGMA integrity_check").fetchone()[0]
                if integrity!="ok":raise RuntimeError("Database integrity check failed after update: "+str(integrity))
                con.commit();messagebox.showinfo("Updated","Rows updated: "+str(changed)+"\nBackup: "+str(backup));self.load_database_tree()
        except Exception as exc:
            try:con.rollback()
            except Exception:pass
            messagebox.showerror("Database operation failed",str(exc))
        finally:con.close()
    def invalidate(self):self.dry_token=None;self.status.config(text="Dry run recommended" if self.use_dry.get() else "Ready")
    def log(self,level,msg):self.events.put(("log",level,msg))
    def progress_event(self,*args):self.events.put(("progress",)+args)
    def toggle_pause(self):self.pause_event.clear() if self.pause_event.is_set() else self.pause_event.set();self.pause.config(text="Resume" if self.pause_event.is_set() else "Pause")
    def request_stop(self):
        if self.busy and messagebox.askyesno("Stop safely","Stop after the current item?"):self.stop_event.set();self.pause_event.clear();self.stop.config(state="disabled");self.status.config(text="Stopping safely...")
    def start(self,dry):
        if self.busy:return
        token=(self.mode.get(),json.dumps(self.config_data,sort_keys=True))
        if not dry and self.use_dry.get() and self.dry_token!=token:return messagebox.showwarning("Dry run required","Run a successful dry run or turn off the option.")
        if not dry and not messagebox.askyesno("Confirm export","Start the backup now?"):return
        self.busy=True;self.stop_event.clear();self.pause_event.clear();self.run_started=time.monotonic();self.bar["value"]=0;self.dry.config(state="disabled");self.export.config(state="disabled");self.pause.config(state="normal");self.stop.config(state="normal");threading.Thread(target=self.worker,args=(dry,),daemon=True).start()
    def worker(self,dry):
        try:self.events.put(("done",dry,execute(self.config_data,self.mode.get(),dry,self.log,self.stop_event,self.pause_event,self.progress_event)))
        except Exception as e:self.events.put(("failed",exception_details(e)))
    @staticmethod
    def duration(v):
        if v is None:return "--:--:--"
        v=max(0,int(v));return f"{v//3600:02d}:{v%3600//60:02d}:{v%60:02d}"
    def clock_tick(self):
        if self.busy and self.run_started:self.elapsed_lbl.config(text="Elapsed "+self.duration(time.monotonic()-self.run_started))
        self.after(250,self.clock_tick)
    def poll(self):
        try:
            while True:
                event=self.events.get_nowait();kind=event[0]
                if kind=="log":
                    level,msg=event[1:];tag="ERROR" if level=="ERROR" else "WARNING" if level=="WARNING" else "SUCCESS" if "EXPORTED" in msg else "INFO"
                    self.console.config(state="normal");self.console.insert("end",msg.split(" | ",2)[-1]+"\n",tag);self.console.see("end");self.console.config(state="disabled")
                elif kind=="progress":
                    processed,total,elapsed,eta,status=event[1:];pct=processed/total*100 if total else 0
                    self.bar["value"]=pct;self.count_lbl.config(text=f"{processed:,} / {total:,} items ({pct:.1f}%)");self.elapsed_lbl.config(text="Elapsed "+self.duration(elapsed));self.eta_lbl.config(text="ETA "+self.duration(eta));self.time_lbl.config(text="Last update: "+datetime.now().strftime("%Y-%m-%d %H:%M:%S"));self.status.config(text=status.title())
                elif kind=="done":
                    dry,result=event[1:];self.finish()
                    if dry and result["result"]=="success":self.dry_token=(self.mode.get(),json.dumps(self.config_data,sort_keys=True))
                    messagebox.showinfo("Run complete","Result: "+str(result["result"])+"\nExported: "+f'{result["exported"]:,}'+"\nDuplicates: "+f'{result["duplicates"]:,}'+"\nReview: "+f'{result["failed"]+result["unreadable"]:,}'+"\n\nReport: "+str(result["excel_report"]))
                elif kind=="failed":self.finish();messagebox.showerror("Run failed",event[1]["message"])
        except queue.Empty:pass
        self.after(100,self.poll)
    def finish(self):self.busy=False;self.pause_event.clear();self.dry.config(state="normal");self.export.config(state="normal");self.pause.config(state="disabled",text="Pause");self.stop.config(state="disabled")
    def close_safely(self):
        if self.busy:return messagebox.showwarning("Run active","Use Stop Safely first.")
        close_log_handlers();LOCK_PATH.unlink(missing_ok=True);self.destroy()

def startup_health_check():
    backup_and_migrate_database()
    ensure_runtime(); load_config(); test=HERE/f".startup_{uuid.uuid4().hex}"; test.write_text("ok"); test.unlink(); connection=sqlite3.connect(DB_PATH); result=connection.execute("PRAGMA integrity_check").fetchone()[0]; connection.close()
    if result!="ok": raise RuntimeError(f"Database integrity check failed: {result}")


if __name__=="__main__":
    try: startup_health_check(); App().mainloop()
    except Exception as exc: ensure_runtime(); (LOG_DIR/f"startup_{datetime.now():%Y%m%d_%H%M%S}.json").write_text(json.dumps(exception_details(exc),indent=2)); messagebox.showerror("Startup failed",str(exc))
