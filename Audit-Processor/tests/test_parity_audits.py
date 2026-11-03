"""The refactored package must reproduce the legacy audit_app.py findings.

``golden/audit_legacy.json`` was produced from the pre-refactor single-file
app (commit 5d09770) with ``golden_runner.run_all``. The only intended
difference is that workbooks no longer contain an 'Audit Summary' sheet.
"""
import json
from pathlib import Path
from types import SimpleNamespace

import golden_runner as g

from audit_suite.audits import AUDITS
from audit_suite.audits.leave_cascade import audit_leave_cascade
from audit_suite.audits.runner import process_audit, process_portfolio_only
from audit_suite.common.files import locate_table, read_raw, require_columns
from audit_suite.common.text import employee_key, safe_filename, time_to_minutes
from audit_suite.outlook_drafts import email_details
from audit_suite.portfolio import add_location_portfolio, prepare_schedule_audit, prepare_supervisor_portfolio

GOLDEN = json.loads((Path(__file__).parent / "golden" / "audit_legacy.json").read_text())

ADAPTER = SimpleNamespace(
    AUDITS=AUDITS, process_audit=process_audit, process_portfolio_only=process_portfolio_only,
    read_raw=read_raw, locate_table=locate_table, require_columns=require_columns,
    read_schedule_audit=prepare_schedule_audit, read_supervisor_portfolio=prepare_supervisor_portfolio,
    email_details=email_details, audit_leave_cascade=audit_leave_cascade,
    add_location_portfolio=add_location_portfolio, employee_key=employee_key,
    time_to_minutes=time_to_minutes, safe_filename=safe_filename)


def _strip_new_notes(case: dict) -> dict:
    return case


def test_same_case_keys():
    assert set(g.run_all(ADAPTER)) == set(GOLDEN)


def test_findings_identical_to_legacy():
    current = json.loads(g.dumps(g.run_all(ADAPTER)))
    for key, expected in GOLDEN.items():
        got = current[key]
        if isinstance(expected, dict) and "sheet" in expected:
            assert expected["sheet"]["sheets"] == ["Audit Results", "Audit Summary"]
            assert got["sheet"]["sheets"] == ["Audit Results"], key  # intentional change
            assert got["sheet"]["rows"] == expected["sheet"]["rows"], key
            assert got["sheet"]["title"] == expected["sheet"]["title"], key
            expected = {k: v for k, v in expected.items() if k != "sheet"}
            got = {k: v for k, v in got.items() if k != "sheet"}
        assert got == expected, key
