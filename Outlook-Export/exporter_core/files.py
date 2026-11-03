"""Exported-file checks and safe file placement."""
from __future__ import annotations

import hashlib
import json
import os
import shutil
import uuid
from pathlib import Path

OLE_MAGIC = bytes.fromhex("D0CF11E0A1B11AE1")  # .msg files are OLE compound documents
PARTIAL_PREFIX = ".partial-"


def verify_msg_file(path) -> tuple[str, int | None]:
    """Return (status, size). status: ok | missing | empty | corrupt | unreadable."""
    p = Path(path)
    try:
        size = p.stat().st_size
    except FileNotFoundError:
        return "missing", None
    except OSError:
        return "unreadable", None
    if size == 0:
        return "empty", 0
    try:
        with open(p, "rb") as handle:
            head = handle.read(8)
    except OSError:
        return "unreadable", size
    return ("ok" if head == OLE_MAGIC else "corrupt"), size


def sha256_file(path, chunk: int = 1024 * 1024) -> str:
    digest = hashlib.sha256()
    with open(path, "rb") as handle:
        for block in iter(lambda: handle.read(chunk), b""):
            digest.update(block)
    return digest.hexdigest()


def place_file(temp: Path, destination: Path) -> int:
    """Copy ``temp`` beside ``destination`` under a partial name, then atomically rename.

    A crash can leave only a ``.partial-*`` file, never a truncated file with
    the final name. An existing destination is replaced only once the copy is complete.
    """
    destination.parent.mkdir(parents=True, exist_ok=True)
    partial = destination.with_name(f"{PARTIAL_PREFIX}{uuid.uuid4().hex}.msg")
    try:
        shutil.copyfile(temp, partial)
        os.replace(partial, destination)
    finally:
        try:
            partial.unlink()
        except FileNotFoundError:
            pass
    return destination.stat().st_size


def atomic_json(path: Path, data) -> None:
    temp = path.with_suffix(path.suffix + ".tmp")
    temp.write_text(json.dumps(data, indent=2, default=str), encoding="utf-8")
    os.replace(temp, path)
