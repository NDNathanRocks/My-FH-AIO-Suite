"""Outlook COM implementation of ``MailSource`` (Windows + classic Outlook + pywin32).

Not importable on other platforms; tests use fakes instead.
"""
from __future__ import annotations

import gc
import tempfile
from pathlib import Path

import pythoncom  # noqa: F401  (imported for callers that initialise COM)
import win32com.client

from .engine import FolderRef, SourceRef
from .naming import OL_MSG_UNICODE, clean_name

PR_INTERNET_MESSAGE_ID = "http://schemas.microsoft.com/mapi/proptag/0x1035001F"


def _safe(obj, name, default=""):
    try:
        return getattr(obj, name)
    except Exception:
        return default


class ComItem:
    def __init__(self, item, item_type: str):
        self._item = item
        self.subject = str(_safe(item, "Subject", _safe(item, "FullName", item_type)) or "")
        for name in ("ReceivedTime", "SentOn", "Start", "CreationTime", "LastModificationTime"):
            value = _safe(item, name, None)
            if value:
                self.item_date = value
                break
        else:
            self.item_date = None
        self.message_class = str(_safe(item, "MessageClass", ""))
        sender = _safe(item, "SenderEmailAddress", "") or _safe(item, "SenderName", "")
        self.sender = str(sender or "")
        try:
            self.internet_message_id = str(item.PropertyAccessor.GetProperty(PR_INTERNET_MESSAGE_ID) or "")
        except Exception:
            self.internet_message_id = ""

    def save_msg(self, path: Path) -> None:
        self._item.SaveAs(str(path), OL_MSG_UNICODE)

    def release(self) -> None:
        self._item = None
        gc.collect()


class OutlookSource:
    """Configured mailbox/folder (optionally recursive) or the active Explorer folder."""

    def __init__(self, config: dict, source_mode: str, logger):
        self.logger = logger
        self.outlook = win32com.client.Dispatch("Outlook.Application")
        self.namespace = self.outlook.GetNamespace("MAPI")
        if source_mode == "configured":
            self.root = self._configured_folder(config)
        else:
            self.root = self.outlook.ActiveExplorer().CurrentFolder
        self.recursive = bool(config.get("include_subfolders", True))
        self.root_name = str(self.root.Name)
        self.root_path = str(self.root.FolderPath)
        self.mailbox = self.root_path.strip("\\").split("\\", 1)[0]
        self.temp_dir = Path(tempfile.gettempdir()) / "InboxEmailExporter"

    def _configured_folder(self, config):
        root, available = None, []
        for index in range(1, self.namespace.Folders.Count + 1):
            candidate = self.namespace.Folders.Item(index)
            available.append(str(candidate.Name))
            if str(candidate.Name).casefold() == config["mailbox_name"].casefold():
                root = candidate
        if root is None:
            raise RuntimeError(f"Mailbox not found. Available: {available}")
        for part in [p.strip() for p in config["folder_path"].split("\\") if p.strip()]:
            root = root.Folders.Item(part)
        return root

    def folders(self) -> list[FolderRef]:
        out = []

        def walk(folder, relative):
            out.append(FolderRef(str(folder.Name), str(folder.FolderPath), str(folder.StoreID), relative, folder))
            if self.recursive:
                for index in range(1, folder.Folders.Count + 1):
                    child = folder.Folders.Item(index)
                    walk(child, relative + (clean_name(child.Name),))
        walk(self.root, ())
        return out

    def snapshot(self, folder: FolderRef) -> list[SourceRef]:
        """Read each item's EntryID/Class once; capture the subject for unreadable items where possible."""
        rows = []
        items = folder.handle.Items
        total = items.Count
        for index in range(1, total + 1):
            item = None
            try:
                item = items.Item(index)
                rows.append(SourceRef(folder.store_id, str(item.EntryID), folder.path, folder.relative, index,
                                      int(_safe(item, "Class", -1))))
            except Exception as exc:
                subject = str(_safe(item, "Subject", "")) if item is not None else ""
                rows.append(SourceRef(folder.store_id, "", folder.path, folder.relative, index, -1, str(exc) or type(exc).__name__, subject))
                self.logger.exception("FAILED SNAPSHOT | folder=%s | index=%s", folder.path, index)
            finally:
                item = None
                if index % 100 == 0:
                    gc.collect()
        items = None
        gc.collect()
        return rows

    def open(self, ref: SourceRef) -> ComItem:
        from .naming import item_type_name
        return ComItem(self.namespace.GetItemFromID(ref.entry_id, ref.store_id), item_type_name(ref.item_class))

    def close(self) -> None:
        self.root = self.namespace = self.outlook = None
        gc.collect()
