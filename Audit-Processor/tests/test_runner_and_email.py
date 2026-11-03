import sys
import time
import types
from datetime import date

import pandas as pd
import pytest

import fixtures as fx
from golden_runner import Upload, upload
from audit_suite import outlook_drafts
from audit_suite.audits import AUDITS
from audit_suite.audits.runner import process_audit, process_leave_cascade, run_report_frame, run_tool
from audit_suite.audits.rules import audit_nba_oc_72, audit_unprocessed_tors
from audit_suite.portfolio import prepare_schedule_audit, prepare_supervisor_portfolio


def test_one_failed_tool_does_not_stop_others():
    good = AUDITS["casuals_loa"]
    runs = [
        run_tool("bad", "Bad", lambda: process_audit(good, Upload("x.csv", b"A,B\n1,2\n"))),
        run_tool("boom", "Boom", lambda: 1 / 0),
        run_tool(good.key, good.name, lambda: process_audit(good, upload(good.source_key, fx.TIME_SUMMARY)), good),
    ]
    assert [r.status for r in runs] == ["failed", "failed", "success"]
    assert "Required columns" in runs[0].error
    assert runs[1].error.startswith("ZeroDivisionError")
    assert runs[2].findings == 1 and runs[2].rows_read == len(fx.TIME_SUMMARY[1])
    report = run_report_frame(runs)
    assert list(report["Status"]) == ["failed", "failed", "success"]


def test_unmatched_portfolio_marks_warning():
    audit = AUDITS["negative_accrual"]
    sched, sup = prepare_schedule_audit(upload("a", fx.SCHEDULE_AUDIT)), prepare_supervisor_portfolio(upload("b", fx.SUPERVISOR))
    run = run_tool(audit.key, audit.name, lambda: process_audit(audit, upload(audit.source_key, fx.ACCRUAL), sched, sup,
                                                                {}, {"id_column": "EE ID"}))
    assert run.status == "warning"
    assert any("no matched portfolio" in n for n in run.notes)


def test_leave_cascade_requires_time_summary():
    run = run_tool("lc", "LC", lambda: process_leave_cascade(AUDITS["leave_cascade"], upload("s", fx.SCHEDULE_BY_PAYCODE), None))
    assert run.status == "failed" and "Time Summary" in run.error


def test_tor_reports_unreadable_dates_and_uses_injected_today():
    frame = pd.DataFrame({"EE ID": ["1", "2"], "Status": ["Pending", "Pending"], "Creation Date": ["01/10/2026", "not a date"]})
    result = audit_unprocessed_tors(frame, cutoff_date=date(2026, 10, 5), today=date(2026, 10, 11))
    assert list(result["Days Since Creation"]) == [10]
    assert "unreadable Creation Date" in result.attrs["summary_messages"][0]


def test_nba_reports_unreadable_hours():
    frame = pd.DataFrame({"EE ID": ["1", "1"], "Actual Total Hrs (Include Corrections)": ["80:00", "bad"],
                          "Pay Code Name": ["521-NBA OC Less 72", "522-NBA OC Grtr 72"]})
    result = audit_nba_oc_72(frame)
    assert len(result) == 1 and "unreadable hours" in result.attrs["summary_messages"][0]


def test_email_recipient_rules():
    frame = pd.DataFrame({
        "Portfolio Email": ["a@x.org; mySchedulehelpline@fraserhealth.ca", "A@x.org b@x.org", "", "c@x.org"],
        "Coordinator Name": ["VCY - Patrice Fugah", "VCY - Unknown Person", "VCY - Z", "Sustainment Team"],
    })
    to, cc, att = outlook_drafts.email_details(frame)
    assert to == ["a@x.org", "b@x.org", "c@x.org"]
    assert cc == ["mySchedulehelpline@fraserhealth.ca", "patrice.fugah2@fraserhealth.ca", "Unknown Person", "Sustainment Team"] or \
        cc == ["mySchedulehelpline@fraserhealth.ca", "patrice.fugah2@fraserhealth.ca", "Unknown Person", "Z"]
    assert len(att) == 2  # blank email row and Sustainment Team row excluded


def test_draft_without_pywin32_raises_cleanly(monkeypatch, tmp_path):
    monkeypatch.setattr(outlook_drafts.tempfile, "gettempdir", lambda: str(tmp_path))
    monkeypatch.setitem(sys.modules, "pythoncom", None)
    frame = pd.DataFrame({"Portfolio Email": ["a@x.org"], "Coordinator Name": ["VCY - A"], "Audit Finding": ["f"]})
    with pytest.raises(RuntimeError, match="pywin32"):
        outlook_drafts.create_outlook_draft(AUDITS["casuals_loa"], frame)
    assert not list(tmp_path.rglob("*.xlsx"))


def test_draft_attachment_deleted_after_save(monkeypatch, tmp_path):
    monkeypatch.setattr(outlook_drafts.tempfile, "gettempdir", lambda: str(tmp_path))
    seen = {}

    class Mail:
        HTMLBody = "<html><body>sig</body></html>"
        def Display(self, *_): pass
        def Save(self): seen["saved"] = True
        class Attachments:
            @staticmethod
            def Add(path):
                seen["path"] = path
                assert __import__("os").path.exists(path)
        GetInspector = types.SimpleNamespace(Activate=lambda: None)

    client = types.SimpleNamespace(Dispatch=lambda _: types.SimpleNamespace(CreateItem=lambda _: Mail()))
    pythoncom = types.SimpleNamespace(CoInitialize=lambda: None, CoUninitialize=lambda: None)
    monkeypatch.setattr(outlook_drafts, "_outlook", lambda: (pythoncom, client))
    monkeypatch.setattr(outlook_drafts.time, "sleep", lambda _: None)
    old = tmp_path / "AuditPrepper"; old.mkdir(); stale = old / "stale.xlsx"; stale.write_bytes(b"x")
    past = time.time() - 3 * 24 * 3600
    __import__("os").utime(stale, (past, past))
    frame = pd.DataFrame({"Portfolio Email": ["a@x.org"], "Coordinator Name": ["VCY - A"], "Audit Finding": ["f"]})
    assert outlook_drafts.create_outlook_draft(AUDITS["casuals_loa"], frame) == 1
    assert seen["saved"] and not __import__("os").path.exists(seen["path"])
    assert not stale.exists()


def test_leave_employee_email_text():
    rows = pd.DataFrame({"Leave Type": ["CTO", "CTO"], "Start Date": ["01/10/2026", "02/10/2026"], "_estimated_minutes": [90, 61],
                         "First Name": ["Tia"], "Last Name": ["Theta"], "Portfolio Email": ["p@x.org", "p@x.org"]} if False else
                        {"Leave Type": ["CTO", "CTO"], "Start Date": ["01/10/2026", "02/10/2026"], "_estimated_minutes": [90, 61],
                         "First Name": ["Tia", "Tia"], "Last Name": ["Theta", "Theta"], "Portfolio Email": ["p@x.org", "p@x.org"]})
    subject, body, cc = outlook_drafts.leave_employee_email(rows)
    assert subject == "Action Required - CTO Leave Adjustment"
    assert "Hello Tia Theta" in body and "1 hour 30 minutes" in body and "1 hour 1 minute" in body
    assert cc == ["p@x.org"]
