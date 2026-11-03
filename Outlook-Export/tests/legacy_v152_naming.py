"""Verbatim naming functions from inbox_email_exporter_v1_5_2.py (v1.5.2), kept as a test oracle."""

import hashlib

import ntpath

import re

from datetime import datetime



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



def normalize_windows_path(value: str) -> str:
    """Return a display/storage path using Windows backslashes.

    Preserve UNC prefixes and drive letters. This is for filesystem paths only,
    not Outlook folder paths.
    """
    value=str(value or "").strip().replace("/", "\\")
    if not value:return value
    return ntpath.normpath(value)



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



def destination_name(item, destination_dir, entry_id, max_path):
    stamp = item_datetime(item); item_type = OBJECT_CLASS_NAMES.get(int(safe_property(item,"Class",-1)), f"Class{safe_property(item,'Class',-1)}")
    subject = clean_name(safe_property(item,"Subject", safe_property(item,"FullName", item_type)), 70)
    suffix = f" - {item_type} - {entry_hash(entry_id)}.msg"; prefix = f"{stamp:%Y-%m-%d_%H%M%S} - "
    available = max(10, int(max_path) - len(str(destination_dir)) - len(prefix) - len(suffix) - 1)
    return prefix + clean_name(subject, available) + suffix


