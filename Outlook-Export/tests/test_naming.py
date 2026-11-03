"""Naming must be identical to v1.5.2 so existing exports and resumes line up."""
from datetime import datetime

import pytest

from exporter_core import naming

def _legacy_funcs():
    import legacy_v152_naming as legacy
    return vars(legacy)


CASES = [
    ("Re: Invoice #123 / Q3", 43, "00000000ABCDEF"),
    ('Bad <chars> : * ? " | \\ / end.', 43, "x" * 140),
    ("", 43, "EID-empty-subject"),
    ("Zoë 漢字 ✓ — unicode", 46, "EID-unicode"),
    ("A" * 300, 26, "EID-long"),
    ("tab\tand\nnewline\x01ctrl", 99, "EID-ctrl"),
    ("trailing dots...   ", 43, "EID-dots"),
]


@pytest.mark.parametrize("subject,item_class,entry_id", CASES)
@pytest.mark.parametrize("dest_dir", [r"M:\HR Transformation\Email Backups\Outlook\Inbox", "M:\\" + "deep\\" * 40])
def test_destination_name_matches_legacy(subject, item_class, entry_id, dest_dir):
    legacy = _legacy_funcs()
    stamp = datetime(2026, 9, 29, 9, 32, 30)

    class Item:
        Class, Subject, ReceivedTime = item_class, subject, stamp
    expected = legacy["destination_name"](Item(), dest_dir, entry_id, 240)
    got = naming.destination_name(stamp, naming.item_type_name(item_class), subject, entry_id, dest_dir, 240)
    assert got == expected


def test_clean_name_and_hash_match_legacy():
    legacy = _legacy_funcs()
    for value in ["", None, "a/b\\c", " x. ", "é" * 200]:
        assert naming.clean_name(value) == legacy["clean_name"](value)
    assert naming.entry_hash("abc") == legacy["entry_hash"]("abc")
    assert naming.normalize_windows_path("M:/a//b/") == legacy["normalize_windows_path"]("M:/a//b/")


def test_hash_from_filename():
    assert naming.hash_from_filename("2026-09-29_093230 - Hi - MailItem - ab12cd34ef56.msg") == "ab12cd34ef56"
    assert naming.hash_from_filename("random.msg") is None
