"""Outlook draft wording and recipient rules per audit."""
import io
import types
from datetime import date, timedelta

import openpyxl
import pandas as pd
import pytest

import fixtures as fx
from golden_runner import upload
from audit_suite import outlook_drafts
from audit_suite.audits import AUDITS
from audit_suite.audits.email_templates import BODIES, days_phrase, tor_phrase
from audit_suite.audits.runner import process_audit
from audit_suite.portfolio import prepare_schedule_audit, prepare_supervisor_portfolio

EMAIL_AUDITS = ["43_843", "grtw_mismatch", "ot_swapped_lct", "unprocessed_tors", "casuals_loa", "negative_accrual", "nba_oc_72"]


@pytest.fixture
def outlook(monkeypatch, tmp_path):
    """Fake Outlook that records the draft it would create."""
    monkeypatch.setattr(outlook_drafts.tempfile, "gettempdir", lambda: str(tmp_path))
    monkeypatch.setattr(outlook_drafts.time, "sleep", lambda _: None)
    drafts = []

    class Mail:
        HTMLBody = "<html><body><p>-- signature --</p></body></html>"

        def __init__(self):
            self.attachments = []
            drafts.append(self)
            self.Attachments = types.SimpleNamespace(Add=lambda p: self.attachments.append(open(p, "rb").read()))
            self.GetInspector = types.SimpleNamespace(Activate=lambda: None)

        def Display(self, *_): pass
        def Save(self): self.saved = True

    client = types.SimpleNamespace(Dispatch=lambda _: types.SimpleNamespace(CreateItem=lambda _: Mail()))
    monkeypatch.setattr(outlook_drafts, "_outlook", lambda: (types.SimpleNamespace(CoInitialize=lambda: None, CoUninitialize=lambda: None), client))
    return drafts


def portfolio():
    return prepare_schedule_audit(upload("sa", fx.SCHEDULE_AUDIT)), prepare_supervisor_portfolio(upload("sp", fx.SUPERVISOR))


def test_every_email_audit_has_subject_and_body():
    for key in EMAIL_AUDITS:
        assert AUDITS[key].email_subject and AUDITS[key].email_body_html == BODIES[key]


def test_ot_swapped_lct_draft_follows_standard_rules(outlook):
    audit = AUDITS["ot_swapped_lct"]
    sched, sup = portfolio()
    # OT/Day Off rows with a 5.2 LCT -> 2004 and 2005; give them portfolio paths via the schedule audit.
    sa_rows = fx.SCHEDULE_AUDIT[1] + [
        ["2004", "03/10/2026", "07:00", "FH/Site A/Unit 2/Job", "", "29/09/2026", "10:00"],
        ["2005", "04/10/2026", "07:00", "FH/Site B/Unit 9/Job", "", "29/09/2026", "10:00"]]
    sched = prepare_schedule_audit(upload("sa", (fx.SCHEDULE_AUDIT[0], sa_rows)))
    result, payload, _ = process_audit(audit, upload("s", fx.SCHEDULE_SHIFT_WR), sched, sup, {}, {"id_column": "Employee ID"})
    assert list(result["Employee ID"]) == ["2004", "2005"]
    rows = outlook_drafts.create_outlook_draft(audit, result)
    mail = outlook[0]
    assert rows == 2
    assert mail.Subject == "Action Required - OT Work Rule and Swapped LCT"
    assert mail.SentOnBehalfOfName == "mySchedulehelpline@fraserhealth.ca"
    assert mail.To == "pa2@example.org; pb9@example.org"  # help-line address removed from To
    assert mail.CC == "mySchedulehelpline@fraserhealth.ca; Sam Example; Lee Example"
    body = mail.HTMLBody
    assert body.index("We identified swapped shifts that have an OT work rule applied.") < body.index("-- signature --")
    assert "Swapped shifts are not considered OT shifts" in body and "reply with your findings.</p>" in body
    ws = openpyxl.load_workbook(io.BytesIO(mail.attachments[0]))["Audit Results"]
    assert ws.max_row == 5 + 2  # separate filtered copy, 2 rows
    assert openpyxl.load_workbook(io.BytesIO(payload))["Audit Results"].max_row == 5 + 2  # full workbook unchanged


def test_ot_draft_excludes_blank_and_sustainment_rows():
    frame = pd.DataFrame({"Portfolio Email": ["a@x.org", "", "s@x.org"], "Coordinator Name": ["VCY - A B", "VCY - C D", "Sustainment Team"],
                          "Audit Finding": ["f", "f", "f"]})
    to, cc, att = outlook_drafts.email_details(frame)
    assert to == ["a@x.org", "s@x.org"] and len(att) == 1


def test_tor_body_counts_and_days(outlook):
    audit = AUDITS["unprocessed_tors"]
    today = date.today()
    frame = pd.DataFrame({"EE ID": ["1", "2", "3"], "Status": ["Pending"] * 3,
                          "Creation Date": [(today - timedelta(days=10)).strftime("%d/%m/%Y")] * 3})
    result = audit.processor(frame, cutoff_date=today - timedelta(days=3))
    result["Portfolio Email"], result["Coordinator Name"] = "a@x.org", "VCY - A B"
    outlook_drafts.create_outlook_draft(audit, result)
    body = outlook[0].HTMLBody
    assert "We identified 3 Time Off Requests (TORs) that have remained in Submitted/Pending status for more than three (3) days and have not yet been actioned." in body
    assert "Thank you for your attention to this matter." in body
    assert tor_phrase(1) == "1 Time Off Request (TOR) that has" and tor_phrase(603) == "603 Time Off Requests (TORs) that have"
    assert days_phrase(1) == "one (1) day" and days_phrase(14) == "14 days"


def test_tor_cutoff_survives_portfolio_enrichment():
    audit = AUDITS["unprocessed_tors"]
    sched, sup = portfolio()
    cutoff = date(2026, 10, 10)
    result, _, _ = process_audit(audit, upload("t", fx.TOR_DETAILS), sched, sup, {"cutoff_date": cutoff}, {"id_column": "EE ID"})
    assert result.attrs["cutoff_date"] == cutoff


def test_negative_accrual_extra_paragraph_before_closing():
    audit = AUDITS["negative_accrual"]
    body = outlook_drafts.email_body_for(audit, pd.DataFrame({"Accrual Code": ["Stat 2026"]}))
    assert body.index("non-cascading pay codes") < body.index("Please reply with your findings once the review is complete.")
    plain = outlook_drafts.email_body_for(audit, pd.DataFrame({"Accrual Code": ["Vacation"]}))
    assert "non-cascading" not in plain and "entire reporting period" in plain


@pytest.mark.parametrize("key,phrase", [
    ("43_843", "Let me know if there are any further questions?"),
    ("grtw_mismatch", "GRTW Work Rule and/or Labor Category Transfer (LCT)."),
    ("casuals_loa", "If you’ve used a cascading paycode, please update it to the correct .2 paycode."),
    ("nba_oc_72", "<strong>521-NBA OC Less 72</strong>"),
])
def test_bodies_render_as_paragraphs(key, phrase):
    frame = pd.DataFrame({"Accrual Code": []})
    body = outlook_drafts.email_body_for(AUDITS[key], frame)
    assert body.startswith("<p>Hello,</p>") and phrase in body and "\n" not in body
