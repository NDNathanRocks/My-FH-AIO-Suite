"""Golden tests for RFP PDF generation: file names and filled PDF fields."""
import csv
from datetime import datetime
from pathlib import Path

import pytest

pymupdf = pytest.importorskip("pymupdf")
from utils import app_utils as au  # noqa: E402

ROOT = Path(__file__).resolve().parents[1]
TEMPLATE = ROOT / "utils" / "RFP Template.pdf"
HEADERS = ["", "Site", "Unit", "Invoice #", "Date Invoice Due", "Overall Invoice Bill Totals ", "GST Subtotals",
           "Services Subtotal", "Travel/Accom Reimbursements", "Agency Personnel Name", "Cost Centre", "Manager Name", "RFP Prepared (Y/N)"]
PAYEE = {"Payable To": "Example Staffing Inc.", "Street Address": "1 Test St", "City": "Surrey", "Province": "BC",
         "Country": "Canada", "Postal Code": "V0V 0V0"}


def tracker_csv(path, rows, headers=HEADERS):
    """Same layout _dump_sheet_to_temp_csv writes: two blank rows, header, data."""
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow([]); w.writerow([]); w.writerow(headers)
        w.writerows(rows)
    return str(path)


def row(inv, rfp="N", unit="ICU", total="$1,100.00"):
    return ["", "RMH", unit, inv, "2026-10-15", total, "50.00", "1,000.00", "", "Pat Example", "12345", "Dr Manager", rfp]


def fields(pdf):
    doc = pymupdf.open(pdf)
    try:
        return {w.field_name: w.field_value for page in doc for w in (page.widgets() or [])}
    finally:
        doc.close()


@pytest.mark.skipif(not TEMPLATE.exists(), reason="template missing")
def test_generation_names_fields_and_filters(tmp_path):
    # With a leading blank column, the end-marker check reads the cell after Site (the Unit column) - see PLAN.md D4.
    csv_path = tracker_csv(tmp_path / "t.csv", [row("INV-1"), row("INV/2", rfp="Y"), row("INV 3", unit="Medical Imaging"),
                                                ["", "", "Total", "999", "", "", "", "", "", "", "", "", "N"], row("AFTER-TOTAL")])
    out = tmp_path / "out"
    ticks = list(au.process_csv_and_fill_pdfs(csv_path, str(TEMPLATE), str(out), "Goodwill", PAYEE, only_new=True))
    assert ticks == [(1, 2, "INV-1"), (2, 2, "INV 3")]
    names = sorted(p.name for p in out.iterdir())
    assert names == ["Goodwill RFP INV 3 - RMH Medical Imaging.pdf", "Goodwill RFP INV-1 - RMH ICU.pdf"]
    f = fields(out / "Goodwill RFP INV-1 - RMH ICU.pdf")
    today = datetime.today().strftime("%d/%m/%y")
    expected = {"Invoice No": "INV-1", "DDMMYY": "15/10/26", "Text2": "12345", "Text18": "50.00", "000TOTAL": "$1,100.00",
                "Name of Requestor": "Ashley Kirkham", "PositionTitle": "Coordinator", "Requestor Signature": "A. Kirkham",
                "Date DDMMYY": today, "Text17": today, "Name of Approver": "Dr Manager", "PositionTitle_2": "Clinical Manager",
                "Must match vendor listed on invoice": "Example Staffing Inc.", "Street Address": "1 Test St", "City": "Surrey",
                "Province": "BC", "Country": "Canada", "Postal Code": "V0V 0V0", "Describe Payment": "Payment for Pat Example hours",
                "Text25": "1000.00", "Text16": "3509030 UPP", "Describe paymentRow1": "Payment for Pat Example travel/accom",
                "Text5": "0.00", "Text4": "3509200 UPP", "Text3": "12345"}
    present = {k: v for k, v in expected.items() if k in f}
    assert len(present) >= 15, sorted(f)
    for key, value in present.items():
        assert f[key] == value, key
    assert fields(out / "Goodwill RFP INV 3 - RMH Medical Imaging.pdf").get("PositionTitle_2", "Manager") == "Manager"


@pytest.mark.skipif(not TEMPLATE.exists(), reason="template missing")
def test_generate_all_when_only_new_false(tmp_path):
    csv_path = tracker_csv(tmp_path / "t.csv", [row("A1", rfp="Y"), row("A2", rfp="")])
    ticks = list(au.process_csv_and_fill_pdfs(csv_path, str(TEMPLATE), str(tmp_path / "o"), "Select", PAYEE, only_new=False))
    assert [t[2] for t in ticks] == ["A1", "A2"]


@pytest.mark.skipif(not TEMPLATE.exists(), reason="template missing")
def test_without_leading_blank_column(tmp_path):
    headers = HEADERS[1:]
    csv_path = tracker_csv(tmp_path / "t.csv", [row("B1")[1:]], headers)
    ticks = list(au.process_csv_and_fill_pdfs(csv_path, str(TEMPLATE), str(tmp_path / "o"), "Select", PAYEE))
    assert ticks == [(1, 1, "B1")]


def test_csv_errors(tmp_path):
    with pytest.raises(ValueError, match="No valid rows"):
        list(au.process_csv_and_fill_pdfs(tracker_csv(tmp_path / "e.csv", []), str(TEMPLATE), str(tmp_path), "X", PAYEE))
    no_inv = [h if h != "Invoice #" else "Bill" for h in HEADERS]
    with pytest.raises(ValueError, match="Invoice column not found"):
        list(au.process_csv_and_fill_pdfs(tracker_csv(tmp_path / "n.csv", [row("Z")], no_inv), str(TEMPLATE), str(tmp_path), "X", PAYEE))
    with pytest.raises(ValueError, match="1 rows marked 'N' had no Invoice #"):
        list(au.process_csv_and_fill_pdfs(tracker_csv(tmp_path / "m.csv", [row("Total", rfp="N"), row("")]), str(TEMPLATE), str(tmp_path), "X", PAYEE))


def test_total_left_of_invoice_column_stops_generation(tmp_path, monkeypatch):
    """D4: 'Total' in any column left of Invoice # (e.g. Notes/Site) ends the sheet."""
    monkeypatch.setattr(au, "fill_pdf_from_row_fitz", lambda *a, **k: None)
    csv_path = tracker_csv(tmp_path / "t.csv", [row("A"), ["", "Total", "", "999", "", "", "", "", "", "", "", "", "N"], row("B")])
    assert [t[2] for t in au.process_csv_and_fill_pdfs(csv_path, str(TEMPLATE), str(tmp_path / "o"), "X", PAYEE, only_new=False)] == ["A"]


ADVANCEDHOME = ["", "Notes", "Invoice #", "Date Received", "Date Invoice Due", "Site", "Unit", "RFP Prepared (Y/N)"]


@pytest.mark.parametrize("padding", [True, False])
def test_owner_layout_total_row_never_becomes_an_invoice(tmp_path, monkeypatch, padding):
    """Real AdvancedHome layout: 'Total' under Notes, invoice count (4) under Invoice #."""
    monkeypatch.setattr(au, "fill_pdf_from_row_fitz", lambda *a, **k: None)
    data = [["", "", inv, "5/19/2026", "6/18/2026", "DH", "Emergency", "N"] for inv in ("7458", "7444", "7472", "7418")]
    if padding:
        data += [["", "", "", "", "1/30/1900", "", "", ""]] * 5
    data += [["", "Total", "4", "", "", "", "", ""]]
    csv_path = tracker_csv(tmp_path / "t.csv", data, ADVANCEDHOME)
    for only_new in (True, False):
        invs = [t[2] for t in au.process_csv_and_fill_pdfs(csv_path, str(TEMPLATE), str(tmp_path / "o"), "AdvancedHome", PAYEE, only_new=only_new)]
        assert invs == ["7458", "7444", "7472", "7418"]


def test_no_n_rows_but_blank_rfp_cells_reports_none(tmp_path, monkeypatch):
    """D3: no row marked N + blank RFP cells -> clear error, nothing generated."""
    monkeypatch.setattr(au, "fill_pdf_from_row_fitz", lambda *a, **k: pytest.fail("must not generate"))
    csv_path = tracker_csv(tmp_path / "b.csv", [row("C1", rfp=""), row("C2", rfp="Y")])
    with pytest.raises(ValueError, match=r"No new RFPs to generate.*1 row\(s\) have a blank"):
        list(au.process_csv_and_fill_pdfs(csv_path, str(TEMPLATE), str(tmp_path / "o"), "X", PAYEE, only_new=True))
