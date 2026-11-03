"""Characterization tests for utils/app_utils.py.

These pin the CURRENT behaviour used by the daily workflow (file names,
invoice-number extraction, currency formatting). Change them only with
owner approval.
"""
import os
from pathlib import Path

import openpyxl
import pytest

from utils import app_utils as au


@pytest.mark.parametrize("text,expected", [
    ("Inv/12:3*?", "Inv123"), ('  a"b<c>d|e  ', "abcde"), ("Normal Name", "Normal Name"), ("x\\y", "xy"),
])
def test_clean_filename(text, expected):
    assert au.clean_filename(text) == expected


@pytest.mark.parametrize("raw,clean,number,currency", [
    ("$1,234.50", "1234.50", "1234.50", "$1,234.50"),
    ("", "0", "0.00", "$0.00"),
    ("abc", "abc", "0", "$0.00"),
    ("  12 ", "12", "12.00", "$12.00"),
    ("-5", "-5", "-5.00", "$-5.00"),
])
def test_currency_helpers(raw, clean, number, currency):
    assert au.clean_currency(raw) == clean
    assert au.format_number(au.clean_currency(raw)) == number
    assert au.format_currency(raw) == currency


@pytest.mark.parametrize("value,ok", [
    ("INV-001", True), ("6058", True), ("Total", False), ("invoice #", False), ("nan", False), ("0", False),
    ("--", False), ("", False), (None, False), ("N/A", False), ("A", True),
])
def test_is_usable_invoice_id(value, ok):
    assert au._is_usable_invoice_id(value) is ok


@pytest.mark.parametrize("filename,expected", [
    ("Invoice FH_25_012 - Site Unit.pdf", "FH_25_012"),
    ("invoice 12345 - x.pdf", "12345"),
    ("Goodwill RFP INV BV 233 - RMH ICU.pdf", "INV BV 233"),
    ("Goodwill RFP 6058 - RMH ICU.pdf", "6058"),
    ("Statement FHA_HOT_RE001_WITH_RECEIPT.pdf", "FHA_HOT_RE001"),
    ("scan 2026 inv 123456ADJ.pdf", "123456ADJ"),
    ("Invoice.pdf", None),
    ("photo 12.jpg", None),
    ("ABC 12-345 (copy).pdf", "12-345"),
    (os.path.join("x", "y", "Invoice 77-A - z.pdf"), "77-A"),
])
def test_extract_invoice_number(filename, expected):
    assert au.extract_invoice_number(filename) == expected


def test_normalize_key():
    assert au.normalize_key("INV-00 12_a ") == "inv0012a"


def test_tracker_lockfile_path():
    assert au.tracker_lockfile_path(os.path.join("d", "T.xlsx")) == os.path.join("d", "~$T.xlsx")
    assert au.tracker_lockfile_path("") == ""


def test_config_round_trip_and_paths():
    assert au.get_paths()["tracker"] == ""
    au.set_path("tracker", r"M:\x\Tracker.xlsx")
    au.save_config_merge({"rfp_count": 3})
    assert au.get_tracker_path() == r"M:\x\Tracker.xlsx"
    assert au.load_config()["rfp_count"] == 3
    with pytest.raises(KeyError):
        au.set_path("bogus", "x")


def test_corrupt_config_reads_as_empty(tmp_path):
    Path(au.CONFIG_PATH).write_text("{not json", encoding="utf-8")
    assert au.load_config() == {}


def test_save_workbook_safe_replaces_and_cleans_up(tmp_path):
    path = tmp_path / "Tracker.xlsx"
    wb = openpyxl.Workbook(); wb.active["A1"] = "old"; wb.save(path)
    wb = openpyxl.load_workbook(path); wb.active["A1"] = "new"; wb.active["B1"] = "=1+1"
    au.save_workbook_safe(wb, str(path))
    again = openpyxl.load_workbook(path)
    assert again.active["A1"].value == "new" and again.active["B1"].value == "=1+1"
    assert [p.name for p in tmp_path.iterdir()] == ["Tracker.xlsx"]


def test_save_workbook_safe_failure_keeps_original(tmp_path):
    path = tmp_path / "Tracker.xlsx"
    wb = openpyxl.Workbook(); wb.active["A1"] = "keep"; wb.save(path)

    class Broken:
        def save(self, p):
            Path(p).write_bytes(b"partial")
            raise OSError("disk full")

        def close(self):
            pass
    with pytest.raises(OSError):
        au.save_workbook_safe(Broken(), str(path))
    assert openpyxl.load_workbook(path).active["A1"].value == "keep"
    assert [p.name for p in tmp_path.iterdir()] == ["Tracker.xlsx"]
