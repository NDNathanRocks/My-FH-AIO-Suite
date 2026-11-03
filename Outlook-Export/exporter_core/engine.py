"""Export engine: discovers items, exports, records, and accounts for every item.

Outlook is reached only through ``MailSource`` / ``MailItem`` so the logic
can be tested with fakes. Every discovered item ends in exactly one outcome.
"""
from __future__ import annotations

import gc
import logging
import time
import uuid
from collections import Counter
from dataclasses import asdict, dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Callable, Iterable, Optional, Protocol

from .database import ExportDatabase, ItemMeta, Key
from .errors import StageError, classify, exception_details, is_memory_error
from .files import place_file, sha256_file, verify_msg_file
from .naming import clean_name, destination_name, entry_hash, item_type_name

MODES = ("export", "retry_failed", "retry_missing", "validate")
WINDOWS_MAX_PATH = 259


@dataclass
class FolderRef:
    name: str
    path: str
    store_id: str
    relative: tuple[str, ...]
    handle: object = None


@dataclass
class SourceRef:
    store_id: str
    entry_id: str
    folder_path: str
    relative: tuple[str, ...]
    index: int
    item_class: int = -1
    snapshot_error: str = ""
    subject: str = ""

    @property
    def key(self) -> Key:
        return (self.store_id, self.entry_id)


class MailItem(Protocol):
    subject: str
    sender: str
    item_date: datetime
    message_class: str
    internet_message_id: str

    def save_msg(self, path: Path) -> None: ...
    def release(self) -> None: ...


class MailSource(Protocol):
    root_name: str
    root_path: str
    mailbox: str
    temp_dir: Path

    def folders(self) -> list[FolderRef]: ...
    def snapshot(self, folder: FolderRef) -> list[SourceRef]: ...
    def open(self, ref: SourceRef) -> MailItem: ...


@dataclass
class EngineConfig:
    save_retries: int = 4
    retry_delay_seconds: float = 2.0
    memory_error_delay_seconds: float = 15.0
    pause_every: int = 50
    pause_seconds: float = 0.2
    max_windows_path: int = 240
    compute_checksums: bool = False
    sleep: Callable[[float], None] = time.sleep

    @classmethod
    def from_config(cls, config: dict) -> "EngineConfig":
        keys = ("save_retries", "retry_delay_seconds", "memory_error_delay_seconds", "pause_every", "pause_seconds",
                "max_windows_path", "compute_checksums")
        return cls(**{k: config[k] for k in keys if k in config})


@dataclass
class Counts:
    folders: int = 0
    items_seen: int = 0          # total discovered
    items_readable: int = 0      # had a readable identity
    exported: int = 0
    adopted: int = 0             # file already present from an interrupted run; recorded, not re-saved
    proposed: int = 0            # dry run "would export"
    duplicates: int = 0          # already exported (all)
    duplicates_verified: int = 0
    duplicates_missing_file: int = 0
    missing_from_db: int = 0     # validate mode
    not_selected: int = 0        # retry modes: outside the retry selection
    skipped_limit: int = 0       # intentionally skipped: active-folder item limit reached
    failed: int = 0
    unreadable: int = 0
    stopped_unprocessed: int = 0

    @property
    def processed(self) -> int:
        return (self.exported + self.adopted + self.proposed + self.duplicates + self.missing_from_db
                + self.not_selected + self.failed + self.unreadable)


@dataclass
class RunResult:
    summary: dict
    rows: list[dict]
    class_counts: Counter
    failures: list[dict] = field(default_factory=list)
    retry_candidates: list[dict] = field(default_factory=list)
    source_refs: list = field(default_factory=list)


def new_run_id(mode: str, dry_run: bool) -> str:
    tag = "_DRY" if dry_run else {"export": "_EXPORT", "retry_failed": "_RETRYFAILED", "retry_missing": "_RETRYMISSING", "validate": "_VALIDATE"}[mode]
    return datetime.now().strftime("%Y%m%d_%H%M%S") + "_" + uuid.uuid4().hex[:6] + tag


def _wait(stop_event, pause_event, logger) -> bool:
    if stop_event is not None and stop_event.is_set():
        return False
    announced = False
    while pause_event is not None and pause_event.is_set() and not (stop_event and stop_event.is_set()):
        if not announced:
            logger.info("PAUSED")
            announced = True
        time.sleep(.25)
    if announced and not (stop_event and stop_event.is_set()):
        logger.info("RESUMED")
    return not (stop_event is not None and stop_event.is_set())


def save_with_retries(item: MailItem, temp_dir: Path, destination: Path, cfg: EngineConfig, logger) -> int:
    """SaveAs to a local temp file (retrying transient failures), then place atomically."""
    temp_dir.mkdir(parents=True, exist_ok=True)
    temp = temp_dir / f"{uuid.uuid4().hex}.msg"
    last = None
    try:
        for attempt in range(1, max(1, int(cfg.save_retries)) + 1):
            try:
                temp.unlink(missing_ok=True)
                item.save_msg(temp)
                if not temp.exists() or temp.stat().st_size <= 0:
                    raise RuntimeError("SaveAs created a missing or empty file")
                last = None
                break
            except Exception as exc:
                last = exc
                wait = float(cfg.memory_error_delay_seconds if is_memory_error(exc) else cfg.retry_delay_seconds)
                logger.warning("SAVE RETRY | attempt=%s | wait=%s | error=%s", attempt, wait, exc)
                gc.collect()
                if attempt < cfg.save_retries:
                    cfg.sleep(wait)
        if last is not None:
            raise StageError("save_temp", last)
        try:
            size = place_file(temp, destination)
        except Exception as exc:
            raise StageError("move", exc) from exc
        status, _ = verify_msg_file(destination)
        if status != "ok":
            raise StageError("verify", RuntimeError(f"Destination is missing or empty ({status})"))
        return size
    finally:
        temp.unlink(missing_ok=True)


def _safe(getter, default=""):
    try:
        value = getter()
        return default if value is None else value
    except Exception:
        return default


def run_export(source: MailSource, db: ExportDatabase, output_root: Path, *, mode: str = "export", dry_run: bool = False,
               limit: Optional[int] = None, cfg: EngineConfig | None = None, run_id: str | None = None,
               stop_event=None, pause_event=None, logger: logging.Logger | None = None,
               progress: Callable | None = None, selection: Optional[Iterable[Key]] = None) -> RunResult:
    """Run one export / retry / validation pass. ``selection`` limits work to those keys (retry modes)."""
    if mode not in MODES:
        raise ValueError(f"Unknown mode {mode!r}")
    cfg = cfg or EngineConfig()
    logger = logger or logging.getLogger("agency-export")
    run_id = run_id or new_run_id(mode, dry_run)
    # retry_failed: only items whose database status is failed (default: all retryable failures).
    # retry_missing: never-exported items plus ``selection`` (completed records whose file is missing/bad);
    #                failed items are left for retry_failed.
    selected = None
    force: set = set()
    if mode == "retry_failed":
        selected = set(selection) if selection is not None else db.failed_keys(retryable_only=True)
    elif mode == "retry_missing":
        force = set(selection or ())
    elif selection is not None:
        selected = set(selection)
    failed_now = db.failed_keys(retryable_only=False) if mode == "retry_missing" else set()
    started = datetime.now()
    counts, rows, class_counts = Counts(), [], Counter()
    failures, retry_candidates = [], []
    stopped = limit_reached = False
    folders = source.folders()
    db.start_run(run_id, mode, dry_run, source.root_path, str(output_root))
    snapshots: list[tuple[FolderRef, list[SourceRef]]] = []
    for folder in folders:
        snapshots.append((folder, source.snapshot(folder)))
    total_expected = sum(len(s) for _, s in snapshots)
    if limit is not None:
        total_expected = min(total_expected, int(limit))
    tick = time.monotonic()

    def report(status):
        if progress:
            elapsed = max(0.001, time.monotonic() - tick)
            rate = counts.processed / elapsed
            eta = (max(0, total_expected - counts.processed) / rate) if rate > 0 else None
            progress(counts.processed, total_expected, elapsed, eta, status)

    logger.info("RUN START | id=%s | mode=%s | dry_run=%s | source=%s | loaded_completed=%s",
                run_id, mode, dry_run, source.root_path, len(db.completed))
    for folder, snapshot in snapshots:
        if limit_reached:
            break
        if stopped or not _wait(stop_event, pause_event, logger):
            stopped = True
            break
        counts.folders += 1
        destination_dir = Path(output_root) / clean_name(source.root_name) / Path(*folder.relative)
        logger.info("FOLDER SCAN | %s | items=%s", folder.path, len(snapshot))
        for ref in snapshot:
            if limit is not None and counts.items_readable >= int(limit):
                limit_reached = True
                break
            if not _wait(stop_event, pause_event, logger):
                stopped = True
                break
            item_type = item_type_name(ref.item_class)
            class_counts[(ref.item_class, item_type)] += 1
            meta = ItemMeta(ref.store_id, ref.entry_id, folder.path, ref.index, ref.item_class, item_type,
                            subject=ref.subject, mailbox=source.mailbox)
            row = {"Run ID": run_id, "Timestamp": datetime.now().isoformat(timespec="seconds"), "Status": "",
                   "Source Folder": folder.path, "Folder Relative Path": "\\".join(folder.relative), "Item Index": ref.index,
                   "Entry Hash": entry_hash(ref.entry_id) if ref.entry_id else "", "Outlook Class": ref.item_class,
                   "Item Type": item_type, "Message Class": "", "Subject/Name": ref.subject, "Item Date": "", "Destination": "",
                   "File Size Bytes": "", "Duplicate Reason": "", "Error Type": "", "HRESULT": "", "Error Message": "",
                   "Stage": "", "Error Category": "", "Retryable": "", "Sender": "", "Internet Message ID": ""}
            if ref.snapshot_error or not ref.entry_id:
                failure = classify(StageError("snapshot", RuntimeError(ref.snapshot_error or "Item has no EntryID")))
                counts.unreadable += 1
                row.update({"Status": "unreadable", "Error Type": "SnapshotError", "Error Message": failure.message,
                            "Stage": "snapshot", "Error Category": failure.category, "Retryable": "yes"})
                rows.append(row)
                failures.append(row)
                if not dry_run and mode != "validate":
                    db.record_unreadable(run_id, meta, failure)
                report("unreadable")
                continue
            counts.items_readable += 1
            key = ref.key
            if (selected is not None and key not in selected) or key in failed_now:
                counts.not_selected += 1
                report("skipped")
                continue
            if db.is_completed(key) and key not in force:
                counts.duplicates += 1
                fs_status, size = verify_msg_file(db.completed[key]) if db.completed[key] else ("missing", None)
                if fs_status == "ok":
                    counts.duplicates_verified += 1
                    reason = "StoreID + EntryID already completed in database; file verified"
                else:
                    counts.duplicates_missing_file += 1
                    reason = f"StoreID + EntryID already completed in database; file {fs_status} - use Retry Missing"
                    retry_candidates.append({"store_id": ref.store_id, "entry_id": ref.entry_id, "reason": f"db record, file {fs_status}",
                                             "path": db.completed[key], "folder": folder.path})
                row.update({"Status": "duplicate-skipped", "Duplicate Reason": reason, "Destination": db.completed[key],
                            "File Size Bytes": size if size is not None else ""})
                rows.append(row)
                report("duplicate")
                continue
            if mode == "validate":
                counts.missing_from_db += 1
                row.update({"Status": "missing-from-db", "Duplicate Reason": "Not exported yet"})
                rows.append(row)
                retry_candidates.append({"store_id": ref.store_id, "entry_id": ref.entry_id, "reason": "source item not exported",
                                         "path": "", "folder": folder.path})
                report("missing")
                continue

            item, destination = None, None
            try:
                try:
                    item = source.open(ref)
                except Exception as exc:
                    raise StageError("open_item", exc) from exc
                try:
                    subject = str(_safe(lambda: item.subject, "") or "")
                    stamp = _safe(lambda: item.item_date, None) or datetime.now()
                    meta.subject = subject
                    meta.sender = str(_safe(lambda: item.sender, ""))
                    meta.item_date = stamp.isoformat() if hasattr(stamp, "isoformat") else str(stamp)
                    meta.message_class = str(_safe(lambda: item.message_class, ""))
                    meta.internet_message_id = str(_safe(lambda: item.internet_message_id, ""))
                except Exception as exc:
                    raise StageError("read_properties", exc) from exc
                try:
                    destination = destination_dir / destination_name(stamp, item_type, subject, ref.entry_id,
                                                                      destination_dir, cfg.max_windows_path)
                    if len(str(destination)) > WINDOWS_MAX_PATH:
                        raise OSError(f"Path too long ({len(str(destination))} > {WINDOWS_MAX_PATH} characters): shorten the output folder")
                except Exception as exc:
                    raise StageError("build_path", exc) from exc
                row.update({"Message Class": meta.message_class, "Subject/Name": subject, "Item Date": meta.item_date,
                            "Destination": str(destination), "Sender": meta.sender, "Internet Message ID": meta.internet_message_id})
                if dry_run:
                    counts.proposed += 1
                    row["Status"] = "would-export"
                    rows.append(row)
                    report("dry-run")
                    continue
                existing, existing_size = verify_msg_file(destination)
                if existing == "ok":
                    # Interrupted earlier run: the file landed but the database was not updated. Adopt, don't duplicate.
                    try:
                        digest = sha256_file(destination) if cfg.compute_checksums else None
                        db.record_success(run_id, meta, destination, existing_size, digest, adopted=True)
                    except Exception as exc:
                        raise StageError("db_write", exc) from exc
                    counts.adopted += 1
                    row.update({"Status": "adopted", "File Size Bytes": existing_size,
                                "Duplicate Reason": "Valid file already existed; recorded without re-saving"})
                    rows.append(row)
                    logger.info("ADOPTED | type=%s | %s", item_type, destination)
                    report("adopted")
                    continue
                size = save_with_retries(item, source.temp_dir, destination, cfg, logger)
                try:
                    digest = sha256_file(destination) if cfg.compute_checksums else None
                    db.record_success(run_id, meta, destination, size, digest)
                except Exception as exc:
                    raise StageError("db_write", exc) from exc
                counts.exported += 1
                row.update({"Status": "exported", "File Size Bytes": size})
                rows.append(row)
                logger.info("EXPORTED | type=%s | %s", item_type, destination)
                report("exported")
            except Exception as exc:
                failure = classify(exc)
                original = exc.original if isinstance(exc, StageError) else exc
                details = exception_details(original)
                counts.failed += 1
                row.update({"Status": "failed", "Error Type": failure.error_type, "HRESULT": failure.hresult,
                            "Error Message": failure.message, "Stage": failure.stage, "Error Category": failure.category,
                            "Retryable": "yes" if failure.retryable else "no"})
                rows.append(row)
                failures.append(row)
                try:
                    db.record_failure(run_id, meta, failure, str(destination or ""), details)
                except Exception:
                    logger.exception("Could not record failure in database | entry_hash=%s", entry_hash(ref.entry_id))
                logger.error("FAILED ITEM | folder=%s | index=%s | hash=%s | stage=%s | category=%s | retryable=%s | %s",
                             folder.path, ref.index, entry_hash(ref.entry_id), failure.stage, failure.category,
                             failure.retryable, failure.message)
                report("failed")
            finally:
                if item is not None:
                    try:
                        item.release()
                    except Exception:
                        pass
                item = None
                if counts.items_readable % max(1, int(cfg.pause_every)) == 0:
                    cfg.sleep(float(cfg.pause_seconds))
    counts.items_seen = total_items(snapshots)
    remaining = max(0, counts.items_seen - counts.processed)
    if stopped:
        counts.stopped_unprocessed = remaining
        logger.warning("STOPPING SAFELY | unprocessed=%s", remaining)
    elif limit_reached:
        counts.skipped_limit = remaining
        logger.info("ITEM LIMIT REACHED | skipped=%s", remaining)
    balanced = counts.processed + counts.stopped_unprocessed + counts.skipped_limit == counts.items_seen
    overall = overall_status(counts, stopped, balanced)
    legacy_result = "stopped" if stopped or limit_reached else "success" if not counts.failed and not counts.unreadable and balanced else "review-required"
    finished = datetime.now()
    summary = {**asdict(counts), "processed": counts.processed, "run_id": run_id, "mode": mode, "dry_run": dry_run,
               "source": source.root_path, "output_root": str(output_root), "started": started.isoformat(),
               "finished": finished.isoformat(), "balanced": balanced, "result": legacy_result, "overall_status": overall,
               "eligible": counts.items_readable - counts.not_selected, "retry_candidates": len(retry_candidates)}
    db.finish_run(run_id, overall, summary)
    logger.info("RUN COMPLETE | %s | %s", legacy_result, overall)
    return RunResult(summary, rows, class_counts, failures, retry_candidates, [r for _, snap in snapshots for r in snap])


def total_items(snapshots) -> int:
    return sum(len(s) for _, s in snapshots)


def overall_status(counts: Counts, stopped: bool, balanced: bool) -> str:
    if stopped or not balanced or counts.failed or counts.unreadable or counts.missing_from_db or counts.duplicates_missing_file:
        return "incomplete"
    if counts.adopted or counts.skipped_limit:
        return "complete with warnings"
    return "complete"
