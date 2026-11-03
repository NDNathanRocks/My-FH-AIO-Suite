"""Test doubles for the Outlook COM layer."""
from __future__ import annotations

import logging
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

from exporter_core.engine import FolderRef, SourceRef
from exporter_core.files import OLE_MAGIC

logging.getLogger("agency-export").addHandler(logging.NullHandler())


@dataclass
class FakeMsg:
    entry_id: str
    subject: str = "Hello"
    when: datetime = datetime(2026, 9, 29, 9, 32, 30)
    item_class: int = 43
    sender: str = "sender@example.org"
    imid: str = ""
    snapshot_error: str = ""
    open_error: Exception | None = None
    save_errors: list = field(default_factory=list)  # raised in order, one per SaveAs call
    payload: bytes = b"body"
    saves: int = 0


class FakeItem:
    def __init__(self, msg: FakeMsg):
        self.msg = msg
        self.subject, self.sender, self.item_date = msg.subject, msg.sender, msg.when
        self.message_class, self.internet_message_id = "IPM.Note", msg.imid

    def save_msg(self, path: Path) -> None:
        self.msg.saves += 1
        if self.msg.save_errors:
            raise self.msg.save_errors.pop(0)
        Path(path).write_bytes(OLE_MAGIC + self.msg.payload)

    def release(self):
        pass


class FakeSource:
    def __init__(self, tree: dict[tuple, list[FakeMsg]], tmp: Path, root_name="Inbox", store="STORE1"):
        self.tree, self.store = tree, store
        self.root_name, self.root_path, self.mailbox = root_name, f"\\\\Mailbox\\{root_name}", "Mailbox"
        self.temp_dir = tmp / "temp"
        self.by_id = {m.entry_id: m for msgs in tree.values() for m in msgs}
        self.opened = []

    def folders(self):
        return [FolderRef(rel[-1] if rel else self.root_name, self.root_path + "".join("\\" + r for r in rel), self.store, rel)
                for rel in self.tree]

    def snapshot(self, folder):
        return [SourceRef(self.store, "" if m.snapshot_error else m.entry_id, folder.path, folder.relative, i, m.item_class,
                          m.snapshot_error, m.subject if m.snapshot_error else "")
                for i, m in enumerate(self.tree[folder.relative], 1)]

    def open(self, ref):
        self.opened.append(ref.entry_id)
        msg = self.by_id[ref.entry_id]
        if msg.open_error:
            raise msg.open_error
        return FakeItem(msg)


class FakeComError(Exception):
    def __init__(self, hresult, text):
        super().__init__(hresult, text)
        self.hresult = hresult
