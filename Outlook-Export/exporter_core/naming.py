"""File and path naming. Must stay byte-for-byte compatible with v1.5.2 so
existing exports keep their names and resumed runs land on the same paths."""
from __future__ import annotations

import hashlib
import ntpath
import re
from datetime import datetime

OL_MSG_UNICODE = 9

# Outlook OlObjectClass values commonly encountered in folders.
OBJECT_CLASS_NAMES = {
    26: "AppointmentItem", 40: "ContactItem", 41: "DocumentItem",
    42: "JournalItem", 43: "MailItem", 44: "NoteItem", 45: "PostItem",
    46: "ReportItem", 47: "RemoteItem", 48: "TaskItem",
    49: "TaskRequestItem", 50: "TaskRequestUpdateItem",
    51: "TaskRequestAcceptItem", 52: "TaskRequestDeclineItem",
    53: "MeetingRequest", 54: "MeetingCancellation",
    55: "MeetingResponseNegative", 56: "MeetingResponsePositive",
    57: "MeetingResponseTentative", 68: "SharingItem",
}

HASH_IN_NAME = re.compile(r" - ([0-9a-f]{12})\.msg$", re.IGNORECASE)


def item_type_name(item_class) -> str:
    try:
        item_class = int(item_class)
    except (TypeError, ValueError):
        item_class = -1
    return OBJECT_CLASS_NAMES.get(item_class, f"Class{item_class}")


def normalize_windows_path(value: str) -> str:
    """Display/storage form of a filesystem path using backslashes (UNC/drive preserved)."""
    value = str(value or "").strip().replace("/", "\\")
    return ntpath.normpath(value) if value else value


def path_key(value: str) -> str:
    """Case-insensitive comparison key for Windows paths."""
    return normalize_windows_path(value).casefold()


def clean_name(value, limit=90) -> str:
    value = re.sub(r'[\\/:*?"<>|]', "_", str(value or "No Subject"))
    value = re.sub(r"[\x00-\x1f]", " ", value)
    value = re.sub(r"\s+", " ", value).strip().rstrip(". ")
    return value[:limit].rstrip(". ") or "No Subject"


def entry_hash(value: str) -> str:
    return hashlib.sha256(value.encode("utf-8", "replace")).hexdigest()[:12]


def hash_from_filename(name: str) -> str | None:
    match = HASH_IN_NAME.search(str(name))
    return match.group(1).lower() if match else None


def destination_name(stamp: datetime, item_type: str, subject: str, entry_id: str, destination_dir, max_path: int) -> str:
    """``YYYY-MM-DD_HHMMSS - <subject> - <ItemType> - <hash12>.msg`` limited to ``max_path``."""
    subject = clean_name(subject, 70)
    suffix = f" - {item_type} - {entry_hash(entry_id)}.msg"
    prefix = f"{stamp:%Y-%m-%d_%H%M%S} - "
    available = max(10, int(max_path) - len(str(destination_dir)) - len(prefix) - len(suffix) - 1)
    return prefix + clean_name(subject, available) + suffix
