"""Smoke-test the Streamlit pages with AppTest (no browser)."""
from pathlib import Path

import pytest

pytest.importorskip("streamlit")
from streamlit.testing.v1 import AppTest  # noqa: E402

import fixtures as fx  # noqa: E402
from golden_runner import upload  # noqa: E402

ROOT = Path(__file__).resolve().parents[1]
PAGE_SCRIPT = """
import sys
sys.path.insert(0, {root!r})
from audit_suite.ui import audits_page
audits_page.page()
"""


def _app():
    return AppTest.from_string(PAGE_SCRIPT.format(root=str(ROOT)), default_timeout=60)


def _mem(key, spec):
    from audit_suite.common.files import validate_upload
    up = upload(key, spec)
    return validate_upload(up.name, up.data)


def test_entry_point_renders_home():
    at = AppTest.from_file(str(ROOT / "audit_app.py"), default_timeout=60).run()
    assert not at.exception
    assert any("HR Data Toolbox" in m.value for m in at.markdown)


def test_stage_one_lists_all_audits():
    at = _app().run()
    assert not at.exception
    labels = [c.label for c in at.checkbox]
    assert "Negative Hours Accrual Bank" in labels and "PC Portfolio Enrichment Only" in labels


def test_full_run_mixed_success_and_failure():
    at = _app()
    at.session_state["workflow_stage"] = 3
    at.session_state["chosen_audits"] = ["negative_accrual", "casuals_loa", "leave_cascade"]
    at.session_state["chosen_portfolio_only"] = False
    at.session_state["upload_cache"] = {
        "accrual_negative_cpp": _mem("acc", fx.ACCRUAL),
        "time_summary_current_pay_period": _mem("ts", fx.TIME_SUMMARY),
        "schedule_by_paycode_cpp": _mem("bad", (["Nope"], [["1"]])),  # wrong file -> must fail alone
        "schedule_audit": _mem("sa", fx.SCHEDULE_AUDIT),
        "supervisor_portfolio": _mem("sp", fx.SUPERVISOR),
    }
    at.session_state["run_config"] = {
        "selected": ["negative_accrual", "casuals_loa", "leave_cascade"], "tool": False,
        "portfolio_choices": {"negative_accrual": True, "casuals_loa": False, "leave_cascade": False},
        "options": {"negative_accrual": {"portfolio_id_column": "EE ID", "portfolio_date_column": None, "portfolio_employee_fallback": True}},
        "employee_details_enabled": False,
    }
    at.run()
    run_btn = next(b for b in at.button if b.label == "Run selected items")
    run_btn.click().run()
    assert not at.exception
    runs = at.session_state["audit_results"]
    assert [r.status for r in runs] == ["warning", "success", "failed"]
    assert runs[0].findings == 2 and runs[1].findings == 1
    assert "Required columns" in runs[2].error
    assert at.session_state["audit_running"] is False
    downloads = [b for b in at.get("download_button")]
    assert downloads, "expected download buttons"


def test_ot_audit_offers_outlook_draft_button():
    at = _app()
    at.session_state["workflow_stage"] = 3
    at.session_state["chosen_audits"] = ["ot_swapped_lct"]
    at.session_state["chosen_portfolio_only"] = False
    at.session_state["upload_cache"] = {"schedule_shift_wr_cpp": _mem("s", fx.SCHEDULE_SHIFT_WR),
                                        "schedule_audit": _mem("sa", fx.SCHEDULE_AUDIT), "supervisor_portfolio": _mem("sp", fx.SUPERVISOR)}
    at.session_state["run_config"] = {"selected": ["ot_swapped_lct"], "tool": False, "portfolio_choices": {"ot_swapped_lct": True},
                                      "options": {"ot_swapped_lct": {"portfolio_id_column": "Employee ID"}}, "employee_details_enabled": False}
    at.run()
    next(b for b in at.button if b.label == "Run selected items").click().run()
    assert not at.exception
    assert any(b.label == "Create and open Outlook email draft" for b in at.button)
