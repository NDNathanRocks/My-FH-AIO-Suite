"""Characterization tests for pure helpers inside the PySide6 pages."""
import csv
from datetime import datetime

import openpyxl
import pandas as pd
import pytest

pytest.importorskip("PySide6")
from pages import audit_page, check_page, generator_page as gen  # noqa: E402
from pages.merger_page import MergerPage, _is_junk_file  # noqa: E402


def make_tracker(path, sheets):
    """Tracker layout: two title rows, headers on row 3, data from row 4."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for name, (headers, rows) in sheets.items():
        ws = wb.create_sheet(name)
        ws["A1"] = f"{name} tracker"
        ws.append([])
        ws.append(headers)
        for r in rows:
            ws.append(r)
    wb.save(path)
    return str(path)


HEADERS = [None, "Vendor ID", "Site", "Unit", "Invoice #", "Total", "RFP Prepared (Y/N)", "Merged (Y/N)"]


def test_merger_invoice_parsing():
    assert MergerPage._invoice_from_filename("`INV BV 233` - Hello 21-21-21.pdf") == "INV BV 233"
    assert MergerPage._invoice_from_filename("`INV BV 233`.pdf") == "INV BV 233"
    assert MergerPage._invoice_from_filename("Invoice 6058 - x.pdf") == "6058"
    assert MergerPage._extract_rfp_identifier("Solutions RFP 82260-584 - LMH Emergency.pdf") == "82260-584"
    assert MergerPage._extract_rfp_identifier("Sympatico RFP FHA_HOT_005.pdf") == "FHA_HOT_005"
    assert MergerPage._extract_rfp_identifier("Acme RFP 123 45.pdf") == "123 45"
    assert MergerPage._extract_rfp_identifier("no id.pdf") is None
    assert MergerPage._norm_id_keep_delims("ab c.1_2-3!") == "ABC1_2-3"
    assert MergerPage._strict_invoice_key(6058.0) == "6058" and MergerPage._strict_invoice_key(" Ab ") == "Ab"
    assert MergerPage._strict_invoice_key(None) == ""
    assert _is_junk_file("x/Thumbs.db") and _is_junk_file("~$Tracker.xlsx") and not _is_junk_file("a.pdf")


def test_generator_helpers(tmp_path):
    assert gen._strict_cell_key(12.0) == "12" and gen._strict_cell_key(" A-1 ") == "A-1" and gen._strict_cell_key(None) == ""
    (tmp_path / "sub").mkdir()
    for name in ("Goodwill RFP INV 1 - RMH ICU.pdf", "sub/77_B - x.pdf", "notes.txt", "Goodwill RFP X.pdf"):
        (tmp_path / name).write_text("x")
    assert gen._scan_generated_invoices_from_folder(str(tmp_path)) == {"INV 1", "77_B"}
    assert gen._norm_hdr_key("RFP Prepared (Y/N)") == "rfppreparedyn" and gen._norm_hdr_key("Invoice #") == "invoice#"
    assert gen._pick_col_letter_case_insensitive({"Invoice #": "E"}, "invoice #") == "E"


def test_read_payee_info(tmp_path):
    p = tmp_path / "c.csv"
    p.write_text("﻿CompanyName , Payable To\nGoodwill , Goodwill Staffing\n", encoding="utf-8")
    assert gen._read_payee_info(str(p)) == {"Goodwill": {"CompanyName": "Goodwill", "Payable To": "Goodwill Staffing"}}


def test_find_vendor_sheet_and_dump_csv(tmp_path):
    path = make_tracker(tmp_path / "T.xlsx", {
        "Summary": (HEADERS, [[None, "999"]]),
        "Goodwill": (HEADERS, [[None, "727012", "RMH", "ICU", "INV-1", 10.5, "N", None], [None, "727012", "SMH", "ER", 6058, 20, None, None]]),
        "Select Med": (HEADERS, [[None, "716335", "RMH", "ICU", "S-1", 1, "N", None]]),
    })
    assert gen._find_vendor_sheet(path, "727012") == "Goodwill"
    assert gen._find_vendor_sheet(path, "716335", preferred_sheet="select") == "Select Med"
    assert gen._find_vendor_sheet(path, "999") is None  # Summary is excluded
    csv_path = gen._dump_sheet_to_temp_csv(path, "Goodwill")
    try:
        with open(csv_path, newline="", encoding="utf-8") as f:
            rows = list(csv.reader(f))
    finally:
        import os
        os.unlink(csv_path)
    assert rows[0] == [] and rows[1] == []
    assert rows[2] == ["Unnamed: 0", "Vendor ID", "Site", "Unit", "Invoice #", "Total", "RFP Prepared (Y/N)", "Merged (Y/N)"]
    assert rows[3] == ["", "727012", "RMH", "ICU", "INV-1", "10.5", "N", ""]
    assert rows[4][4] == "6058"


def test_update_tracker_rfp_prepared_only_touches_rfp_column(tmp_path):
    path = make_tracker(tmp_path / "T.xlsx", {
        "Goodwill": (HEADERS + ["Calc"], [[None, "1", "RMH", "ICU", "INV-1", 10, "N", None, "=F4*2"],
                                          [None, "1", "SMH", "ER", 6058, 20, "N", None, "=F5*2"]])})
    before = openpyxl.load_workbook(path)["Goodwill"]
    snapshot = {(c.row, c.column): c.value for row in before.iter_rows() for c in row}
    updated, total, missing, df = gen._update_tracker_rfp_prepared(path, "Goodwill", {"INV-1", "6058", "NOPE"})
    assert (updated, total, missing, df) == (2, 3, ["NOPE"], None)
    after = openpyxl.load_workbook(path)["Goodwill"]
    changed = {(c.row, c.column): c.value for row in after.iter_rows() for c in row if snapshot.get((c.row, c.column)) != c.value}
    assert changed == {(4, 7): "Y", (5, 7): "Y"}
    assert after["I4"].value == "=F4*2"  # formulas preserved
    assert [p.name for p in tmp_path.iterdir()] == ["T.xlsx"]


def test_update_tracker_skip_mode_returns_rows_without_writing(tmp_path):
    """Regression for PLAN.md #17: blank cells in read-only mode made every row 'not found'."""
    path = make_tracker(tmp_path / "T.xlsx", {"Goodwill": (HEADERS, [[None, "1", "RMH", "ICU", "INV-1", 10, "N", None]])})
    updated, total, missing, df = gen._update_tracker_rfp_prepared(path, "Goodwill", {"INV-1"}, skip_update=True)
    assert updated == 1 and df.loc[0, "Invoice #"] == "INV-1"
    assert openpyxl.load_workbook(path)["Goodwill"]["G4"].value == "N"


def test_check_page_helpers(tmp_path):
    assert check_page.clean_email_list("a@x.org, b@x.org,, ") == ["a@x.org", "b@x.org"]
    assert check_page.clean_email_list("nan") == []
    assert check_page.safe_float_conversion("$1,234.567") == 1234.57
    assert check_page.safe_float_conversion("bad", 9) == 9
    assert check_page.extract_name_info("jane.doe@x.org") == ("Jane", "JD")
    assert check_page.extract_name_info("solo@x.org") == ("Solo", "S")
    assert check_page.sanitize_sheet_name("a/b:c@fraserhealth.ca") == "abc"
    ha = tmp_path / "ha.xlsx"
    pd.DataFrame({"Vendor ID": ["727012", "727012", "727012"], "Invoice Number": ["INV-1", "INV 1", "X9"],
                  "Payment Line Amount": ["$100.00", "50", "bad"], "Last Issued Check Num": ["C1", "C2", ""]}).to_excel(ha, index=False)
    lk = check_page.build_ha_lookup(str(ha))
    assert lk.raw_lookup[("727012", "INV-1")] == (100.0, "C1")
    assert lk.norm_lookup[("727012", "inv1")] == (150.0, "C1,C2")
    assert lk.norm_collisions == {("727012", "inv1")} and lk.source_rows == 3
    # D5: a blank check number is recorded as empty, never the text "nan".
    assert lk.raw_lookup[("727012", "X9")] == (0.0, "")
    ha2 = tmp_path / "ha2.xlsx"
    pd.DataFrame({"Vendor ID": ["1"], "Invoice Number": ["Q1"], "Payment Total Amount": ["5"], "Last Issued Check Num": [None]}).to_excel(ha2, index=False)
    assert audit_page.build_ha_lookup(str(ha2))[("1", "q1")] == (5.0, "")


def test_audit_page_helpers(tmp_path):
    assert audit_page.sanitize_sheet_name("a[b]*c/d\\e?f:g" + "x" * 40) == ("abcdefg" + "x" * 40)[:31]
    assert audit_page.best_sheet_match("Goodwil", ["Summary", "Goodwill", "Select"]) == "Goodwill"
    assert audit_page.best_sheet_match("zzz", ["Goodwill"]) is None
    assert audit_page.to_datetime_or_none(datetime(2026, 1, 1)) == datetime(2026, 1, 1)
    assert audit_page.to_datetime_or_none("2026-01-01") is None
    assert audit_page.resolve_first_column(["Invoice Number ", "X"], ["invoice #", "invoice number"]) == "Invoice Number "
    path = make_tracker(tmp_path / "T.xlsx", {"G": (HEADERS, [[None, "1", "RMH", "ICU", "INV-00 1", 1, "N", None]])})
    ws = openpyxl.load_workbook(path)["G"]
    assert audit_page.header_map(ws)["Invoice #"] == "E"
    assert audit_page.find_invoice_row(ws, "E", "inv001") == 4


def test_merged_output_name():
    build = MergerPage._build_output_name
    assert build(None, "Goodwill RFP INV-1 - RMH ICU West.pdf", "INV-1") == "RFP INV-1 - RMH ICU West.pdf"
    assert build(None, "Goodwill RFP INV-1 - RMH.pdf", "INV-1") == "RFP INV-1 - RMH.pdf"
    assert build(None, "Goodwill RFP INV-1.pdf", "INV-1") == "RFP INV-1.pdf"


class _Log:
    def __init__(self):
        self.lines = []

    def _log(self, msg):
        self.lines.append(msg)


def _pdf(path, text):
    import pymupdf
    doc = pymupdf.open(); page = doc.new_page(); page.insert_text((72, 72), text); doc.save(path); doc.close()


def test_merge_files_order_and_atomic_output(tmp_path):
    import pymupdf
    from PIL import Image
    rfp, inv, img = tmp_path / "rfp.pdf", tmp_path / "inv.pdf", tmp_path / "scan.png"
    _pdf(rfp, "RFP"); _pdf(inv, "INVOICE"); Image.new("RGB", (50, 50), "white").save(img)
    out = tmp_path / "merged.pdf"
    ok, used = MergerPage._merge_files(_Log(), str(rfp), [str(inv), str(img), str(tmp_path / "Thumbs.db")], str(out))
    assert ok and used == [str(inv), str(img)]
    doc = pymupdf.open(out)
    assert doc.page_count == 3 and "RFP" in doc[0].get_text() and "INVOICE" in doc[1].get_text()
    doc.close()
    assert sorted(p.name for p in tmp_path.iterdir()) == ["inv.pdf", "merged.pdf", "rfp.pdf", "scan.png"]


def test_merge_failure_leaves_no_partial_output(tmp_path):
    rfp = tmp_path / "rfp.pdf"; _pdf(rfp, "RFP")
    bad = tmp_path / "bad.pdf"; bad.write_bytes(b"%PDF-1.4 broken")
    log = _Log()
    ok, used = MergerPage._merge_files(log, str(rfp), [str(bad)], str(tmp_path / "merged.pdf"))
    assert not ok and used == []
    assert not (tmp_path / "merged.pdf").exists() and not (tmp_path / "merged.pdf.partial").exists()
    assert any("Merge error" in line for line in log.lines)
