import io
import zipfile

import openpyxl
import pandas as pd
import pytest

import fixtures as fx
from golden_runner import Upload, upload
from audit_suite.audits import AUDITS
from audit_suite.audits.runner import process_leave_cascade
from audit_suite.common.excel import audit_workbook_bytes, styled_sheets_bytes
from audit_suite.common.files import InputFileError, locate_table, read_raw, read_table, validate_upload


def test_validate_upload_rules():
    assert validate_upload("C:\\Users\\x\\a.csv", b"x").name == "a.csv"
    assert validate_upload("../../etc/b.xlsx", b"x").name == "b.xlsx"
    with pytest.raises(InputFileError, match="empty"):
        validate_upload("a.csv", b"")
    with pytest.raises(InputFileError, match="unsupported"):
        validate_upload("a.exe", b"x")
    with pytest.raises(InputFileError, match="limit"):
        validate_upload("a.csv", b"x" * 11, max_bytes=10)


def test_corrupt_xlsx_is_reported_not_crashing():
    with pytest.raises(InputFileError, match="not a readable Excel workbook"):
        read_raw(Upload("bad.xlsx", b"this is not a zip"))


def test_password_or_other_zip_reported():
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        z.writestr("hello.txt", "x")
    with pytest.raises(InputFileError):
        read_raw(Upload("odd.xlsx", buf.getvalue()))


def test_missing_columns_message_lists_columns():
    with pytest.raises(InputFileError, match="Required columns: Work Rule, EE ID, Pay Code Name"):
        read_table(Upload("x.csv", b"A,B\n1,2\n"), ("Work Rule", "EE ID", "Pay Code Name"), ("Work Rule", "EE ID", "Pay Code Name"))


def test_csv_encodings_and_delimiters():
    text = "EE ID;Name;Status\n1;Zoë Ł;Pending\n"
    for enc in ("utf-8-sig", "utf-8"):
        frame = locate_table(read_raw(Upload("a.csv", text.encode(enc))), ("EE ID", "Status"))
        assert frame.loc[0, "Name"] == "Zoë Ł"
    cp = "EE ID,Name\n1,Café\n".encode("cp1252")
    assert locate_table(read_raw(Upload("a.csv", cp)), ("EE ID",)).loc[0, "Name"] == "Café"


def test_blank_rows_are_counted_not_silently_lost():
    raw = [["EE ID", "Status"], ["1", "a"], [None, None], [], ["2", "b"]]
    frame = locate_table(raw, ("EE ID",))
    assert len(frame) == 2
    assert frame.attrs["blank_rows_skipped"] == 2
    assert frame.attrs["source_rows"] == 4


def test_header_case_and_whitespace_insensitive():
    frame = read_table(Upload("a.csv", b"ee  id,STATUS,creation date\n1,Pending,01/01/2026\n"),
                       ("EE ID", "Status", "Creation Date"), ("EE ID", "Status", "Creation Date"))
    assert list(frame.columns) == ["EE ID", "Status", "Creation Date"]


def test_workbook_has_no_audit_summary_sheet():
    payload = audit_workbook_bytes(pd.DataFrame({"Audit Finding": ["x"], "EE ID": ["1"]}).assign(_hidden=1), "T", "S")
    wb = openpyxl.load_workbook(io.BytesIO(payload))
    assert wb.sheetnames == ["Audit Results"]


def test_leave_cascade_table_matches_visible_columns():
    """Regression: hidden _estimated_minutes made the Excel table wider than the header row."""
    audit = AUDITS["leave_cascade"]
    result, payload, _ = process_leave_cascade(audit, upload("s", fx.SCHEDULE_BY_PAYCODE), upload("t", fx.TIME_SUMMARY))
    ws = openpyxl.load_workbook(io.BytesIO(payload))["Audit Results"]
    table = ws.tables["AuditResultsTable"]
    visible = [c for c in result.columns if not str(c).startswith("_")]
    assert table.ref == f"A5:{openpyxl.utils.get_column_letter(len(visible))}{5 + len(result)}"
    assert [c.value for c in ws[5]][: len(visible)] == visible
    assert all(c.value is not None for c in ws[5])


def test_empty_result_workbook():
    frame = pd.DataFrame(columns=["Audit Finding", "EE ID"])
    ws = openpyxl.load_workbook(io.BytesIO(audit_workbook_bytes(frame, "T", "S")))["Audit Results"]
    assert ws["A6"].value == "No findings matched the audit rules."
    assert not ws.tables


def test_styled_sheets_keep_formula_like_text_as_text():
    payload = styled_sheets_bytes({"Data": pd.DataFrame({"A": ["=1+1", "=HYPERLINK(\"x\")", "ok"]})})
    ws = openpyxl.load_workbook(io.BytesIO(payload))["Data"]
    assert ws["A2"].data_type == "s" and ws["A2"].value == "=1+1"
    assert ws["A3"].data_type == "s"
