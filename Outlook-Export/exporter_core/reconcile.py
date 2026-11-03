"""Reconcile the Outlook source, the database and the export folder.

Works offline (database + files) or with a source snapshot from Outlook.
Every issue names the affected item by StoreID/EntryID (+ hash, subject,
folder, path where known) and says whether a retry is safe.
"""
from __future__ import annotations

import os
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Iterable, Optional

from .database import read_export_rows
from .files import PARTIAL_PREFIX, verify_msg_file
from .naming import entry_hash, hash_from_filename, path_key

ERROR, WARNING, INFO = "error", "warning", "info"


@dataclass
class Issue:
    check: str
    severity: str
    detail: str
    store_id: str = ""
    entry_id: str = ""
    entry_hash: str = ""
    subject: str = ""
    source_folder: str = ""
    path: str = ""
    retryable: bool = False
    retry_mode: str = ""   # retry_failed | retry_missing | ""

    def as_dict(self) -> dict:
        return self.__dict__.copy()


@dataclass
class Reconciliation:
    generated: str
    output_root: str
    counts: dict = field(default_factory=dict)
    issues: list[Issue] = field(default_factory=list)

    @property
    def overall_status(self) -> str:
        if any(i.severity == ERROR for i in self.issues):
            return "incomplete"
        if any(i.severity == WARNING for i in self.issues):
            return "complete with warnings"
        return "complete"

    def by_check(self) -> dict[str, int]:
        out: dict[str, int] = defaultdict(int)
        for issue in self.issues:
            out[issue.check] += 1
        return dict(sorted(out.items()))

    def retry_keys(self, mode: str) -> set[tuple[str, str]]:
        return {(i.store_id, i.entry_id) for i in self.issues if i.retryable and i.retry_mode == mode and i.entry_id}


def _under(path: str, root: str) -> bool:
    p, r = path_key(path), path_key(root).rstrip("\\")
    return p == r or p.startswith(r + "\\")


def scan_files(output_root: Path) -> tuple[list[Path], list[Path]]:
    """Return (msg files, partial files) under output_root."""
    msgs, partials = [], []
    if not output_root.exists():
        return msgs, partials
    for dirpath, _dirs, filenames in os.walk(output_root):
        for name in filenames:
            if not name.lower().endswith(".msg"):
                continue
            (partials if name.startswith(PARTIAL_PREFIX) else msgs).append(Path(dirpath) / name)
    return msgs, partials


def reconcile(db_path: Path, output_root: Path, source_refs: Optional[Iterable] = None, *, verify_files: bool = True,
              interrupted_runs: Iterable = ()) -> Reconciliation:
    output_root = Path(output_root)
    rec = Reconciliation(datetime.now().isoformat(timespec="seconds"), str(output_root))
    add = rec.issues.append
    rows = list(read_export_rows(db_path)) if Path(db_path).exists() else []
    by_key = {(r["store_id"], r["entry_id"]): r for r in rows}
    completed = [r for r in rows if r["status"] == "completed"]
    failed = [r for r in rows if r["status"] == "failed"]

    # --- Database records vs files ---
    dest_owner: dict[str, list[dict]] = defaultdict(list)
    referenced: set[str] = set()
    verified = 0
    for r in completed:
        dest = r["destination"] or ""
        ident = dict(store_id=r["store_id"], entry_id=r["entry_id"], entry_hash=entry_hash(r["entry_id"] or ""),
                     subject=r.get("subject") or "", source_folder=r["source_folder"] or "", path=dest)
        if not dest:
            add(Issue("db_record_without_path", ERROR, "Completed record has no destination path", retryable=True, retry_mode="retry_missing", **ident))
            continue
        dest_owner[path_key(dest)].append(r)
        referenced.add(path_key(dest))
        if not _under(dest, str(output_root)):
            add(Issue("db_path_outside_output_root", WARNING, "Destination is not under the configured output root (folder renamed or root changed?)", **ident))
        if verify_files:
            status, size = verify_msg_file(dest)
            if status == "ok":
                verified += 1
                if r["file_size"] not in (None, "") and int(r["file_size"]) != size:
                    add(Issue("file_size_changed", WARNING, f"Recorded size {r['file_size']} but file is {size} bytes", **ident))
            elif status == "missing":
                add(Issue("db_record_missing_file", ERROR, "Database says exported but the file does not exist", retryable=True, retry_mode="retry_missing", **ident))
            elif status == "empty":
                add(Issue("empty_file", ERROR, "Exported file is zero bytes", retryable=True, retry_mode="retry_missing", **ident))
            elif status == "corrupt":
                add(Issue("corrupt_file", ERROR, "File is not a valid Outlook .msg (bad header)", retryable=True, retry_mode="retry_missing", **ident))
            else:
                add(Issue("unreadable_file", WARNING, "File exists but could not be opened (locked or permission)", **ident))
        m_status = r.get("m_status")
        if m_status and m_status not in ("exported",):
            add(Issue("conflicting_status", WARNING, f"exports.status=completed but manifest status={m_status}", **ident))
    for key, owners in dest_owner.items():
        if len(owners) > 1:
            for r in owners:
                add(Issue("duplicate_output_path", ERROR, f"{len(owners)} database records point to the same file",
                          store_id=r["store_id"], entry_id=r["entry_id"], entry_hash=entry_hash(r["entry_id"]), path=r["destination"]))

    for r in failed:
        retryable = r.get("retryable")
        can_retry = retryable is None or int(retryable) == 1
        add(Issue("failed_export", ERROR, f"Export failed at stage {r.get('error_stage') or 'unknown'} ({r.get('error_category') or 'unclassified'})",
                  store_id=r["store_id"], entry_id=r["entry_id"], entry_hash=entry_hash(r["entry_id"] or ""), subject=r.get("subject") or "",
                  source_folder=r["source_folder"] or "", retryable=can_retry, retry_mode="retry_failed" if can_retry else ""))

    # --- Duplicate identities ---
    by_imid: dict[str, list[dict]] = defaultdict(list)
    for r in rows:
        if r.get("internet_message_id"):
            by_imid[r["internet_message_id"]].append(r)
    for imid, group in by_imid.items():
        if len(group) > 1:
            for r in group:
                add(Issue("duplicate_internet_message_id", WARNING, f"Same Internet Message ID on {len(group)} Outlook items (copies in several folders?)",
                          store_id=r["store_id"], entry_id=r["entry_id"], entry_hash=entry_hash(r["entry_id"]), path=r["destination"] or ""))

    # --- Files vs database ---
    files, partials = scan_files(output_root)
    hash_to_keys: dict[str, list[tuple]] = defaultdict(list)
    for key in by_key:
        hash_to_keys[entry_hash(key[1] or "")].append(key)
    source_by_hash: dict[str, list] = defaultdict(list)
    source_list = list(source_refs) if source_refs is not None else None
    for ref in source_list or ():
        if ref.entry_id:
            source_by_hash[entry_hash(ref.entry_id)].append(ref)
    files_per_hash: dict[str, list[Path]] = defaultdict(list)
    for f in files:
        h = hash_from_filename(f.name)
        if h:
            files_per_hash[h].append(f)
        if path_key(str(f)) in referenced:
            continue
        status, size = verify_msg_file(f) if verify_files else ("unchecked", None)
        if h and h in hash_to_keys:
            k = hash_to_keys[h][0]
            rec_row = by_key[k]
            if rec_row["status"] == "completed":
                add(Issue("file_not_at_recorded_path", WARNING, f"File for this item exists at a different path than recorded ({status})",
                          store_id=k[0], entry_id=k[1], entry_hash=h, path=str(f)))
            else:
                add(Issue("file_with_incomplete_db_status", WARNING, f"File exists but database status is {rec_row['status']}; next export will adopt it",
                          store_id=k[0], entry_id=k[1], entry_hash=h, path=str(f), retryable=True,
                          retry_mode="retry_failed" if rec_row["status"] == "failed" else "retry_missing"))
        elif h and h in source_by_hash:
            ref = source_by_hash[h][0]
            add(Issue("file_missing_db_record", WARNING, f"File exists without a database record (interrupted run?); next export will adopt it ({status})",
                      store_id=ref.store_id, entry_id=ref.entry_id, entry_hash=h, source_folder=ref.folder_path, path=str(f),
                      retryable=True, retry_mode="retry_missing"))
        else:
            add(Issue("orphan_file", WARNING, f"File has no matching database record ({status})", entry_hash=h or "", path=str(f)))
        if status in ("empty", "corrupt"):
            add(Issue("empty_file" if status == "empty" else "corrupt_file", ERROR, f"Unreferenced file is {status}", entry_hash=h or "", path=str(f)))
    for h, group in files_per_hash.items():
        if len(group) > 1:
            for f in group:
                add(Issue("multiple_files_for_email", WARNING, f"{len(group)} files share entry hash {h}", entry_hash=h, path=str(f)))
    for p in partials:
        add(Issue("partial_file", WARNING, "Incomplete copy left by an interrupted export; safe to delete after review", path=str(p)))
    for run in interrupted_runs:
        add(Issue("interrupted_run", WARNING, f"Run {run[0]} ({run[1]}) started {run[2]} never finished"))

    # --- Outlook source vs database ---
    source_counts = {}
    if source_list is not None:
        seen_keys, unreadable, missing = set(), 0, 0
        for ref in source_list:
            if ref.snapshot_error or not ref.entry_id:
                unreadable += 1
                add(Issue("unreadable_source_item", ERROR, f"Outlook item could not be read: {ref.snapshot_error or 'no EntryID'}",
                          source_folder=ref.folder_path, subject=getattr(ref, "subject", ""), path=f"index {ref.index}"))
                continue
            key = (ref.store_id, ref.entry_id)
            if key in seen_keys:
                add(Issue("duplicate_identifier", WARNING, "StoreID+EntryID listed twice in the source", store_id=key[0], entry_id=key[1],
                          entry_hash=entry_hash(key[1]), source_folder=ref.folder_path))
            seen_keys.add(key)
            r = by_key.get(key)
            if r is None:
                missing += 1
                add(Issue("source_missing_db_record", ERROR, "Outlook item has never been exported", store_id=key[0], entry_id=key[1],
                          entry_hash=entry_hash(key[1]), source_folder=ref.folder_path, subject=getattr(ref, "subject", ""),
                          retryable=True, retry_mode="retry_missing"))
        not_in_source = [k for k in by_key if k not in seen_keys]
        for k in not_in_source:
            add(Issue("db_record_not_in_source", INFO, "Exported earlier; no longer in the scanned Outlook folders (moved or deleted)",
                      store_id=k[0], entry_id=k[1], entry_hash=entry_hash(k[1]), path=by_key[k]["destination"] or ""))
        source_counts = {"source_items": len(source_list), "source_unreadable": unreadable, "source_missing_db": missing,
                         "db_not_in_source": len(not_in_source)}

    rec.counts = {"db_records": len(rows), "db_completed": len(completed), "db_failed": len(failed),
                  "files_found": len(files), "files_verified": verified, "partial_files": len(partials),
                  **source_counts, "issues": len(rec.issues), "retry_failed": len(rec.retry_keys("retry_failed")),
                  "retry_missing": len(rec.retry_keys("retry_missing"))}
    return rec
