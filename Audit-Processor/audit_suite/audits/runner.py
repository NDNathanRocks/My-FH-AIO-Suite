"""Run audits with per-tool error isolation and explicit counts."""
from __future__ import annotations

import logging
from dataclasses import dataclass, field
from datetime import date
from typing import Callable, Optional

import pandas as pd

from ..common.excel import audit_workbook_bytes
from ..common.files import InputFileError, detect_existing_table, read_raw, read_table
from ..common.text import clean_text, employee_key, safe_filename
from ..config import (DATE_FMT, EMPLOYEE_DETAILS_EXPECTED, EMPLOYEE_DETAILS_REQUIRED,
                      TIME_SUMMARY_EXPECTED, TIME_SUMMARY_REQUIRED)
from ..portfolio import add_location_portfolio
from .leave_cascade import audit_leave_cascade
from .registry import AuditDefinition

log = logging.getLogger(__name__)


def output_filename(title: str) -> str:
    return f"[SUS] Audit - {safe_filename(title)} {date.today():{DATE_FMT}}.xlsx"


@dataclass
class ToolRun:
    """Outcome of one audit/tool. ``status`` is success, warning or failed."""
    key: str
    name: str
    status: str = "success"
    rows_read: int = 0
    blank_rows_skipped: int = 0
    findings: int = 0
    notes: list[str] = field(default_factory=list)
    error: Optional[str] = None
    result: Optional[pd.DataFrame] = None
    payload: Optional[bytes] = None
    filename: Optional[str] = None
    audit: Optional[AuditDefinition] = None

    def as_report_row(self) -> dict:
        return {"Tool": self.name, "Status": self.status, "Rows read": self.rows_read,
                "Blank rows skipped": self.blank_rows_skipped, "Findings": self.findings,
                "Notes": " | ".join(self.notes), "Error": self.error or ""}


def _apply_portfolio(result, id_column, schedule_audit, supervisor, portfolio_match):
    match = portfolio_match or {}
    return add_location_portfolio(result, match.get("id_column", id_column), schedule_audit, supervisor,
                                  match.get("date_column"), match.get("allow_employee_fallback", True))


def process_audit(audit, uploaded_file, schedule_audit=None, supervisor=None, audit_options=None, portfolio_match=None):
    """Single-source audit. Returns (result, workbook bytes, filename)."""
    table = read_table(uploaded_file, audit.expected_columns, audit.required_columns)
    result = audit.processor(table, **(audit_options or {}))
    result.attrs.setdefault("rows_read", len(table))
    result.attrs.setdefault("blank_rows_skipped", table.attrs.get("blank_rows_skipped", 0))
    if schedule_audit is not None and supervisor is not None:
        attrs = dict(result.attrs)
        result = _apply_portfolio(result, audit.employee_id_column, schedule_audit, supervisor, portfolio_match)
        result.attrs = {**attrs, **result.attrs}
    return result, audit_workbook_bytes(result, audit.name, audit.source_name), output_filename(audit.name)


def process_leave_cascade(audit, schedule_upload, time_upload, schedule_audit=None, supervisor=None, portfolio_match=None):
    if time_upload is None:
        raise InputFileError("Time Summary by Pay Code is required for this audit. Return to Step 2.")
    schedule_table = read_table(schedule_upload, audit.expected_columns, audit.required_columns)
    time_table = read_table(time_upload, TIME_SUMMARY_EXPECTED, TIME_SUMMARY_REQUIRED)
    result = audit_leave_cascade(schedule_table, time_table)
    result.attrs["rows_read"] = len(schedule_table)
    result.attrs["blank_rows_skipped"] = schedule_table.attrs.get("blank_rows_skipped", 0) + time_table.attrs.get("blank_rows_skipped", 0)
    if schedule_audit is not None and supervisor is not None:
        attrs = dict(result.attrs)
        result = _apply_portfolio(result, "EE ID", schedule_audit, supervisor, portfolio_match)
        result.attrs = {**attrs, "summary_messages": result.attrs.get("summary_messages", [])}
    return result, audit_workbook_bytes(result, audit.name, audit.source_name), output_filename(audit.name)


def process_portfolio_only(uploaded_file, schedule_audit, supervisor, id_column,
                           date_column, allow_employee_fallback, output_name):
    frame = detect_existing_table(read_raw(uploaded_file))
    if id_column not in frame.columns:
        raise InputFileError(f"The selected match column '{id_column}' was not found.")
    rows_read, blank = len(frame), frame.attrs.get("blank_rows_skipped", 0)
    result = add_location_portfolio(frame, id_column, schedule_audit, supervisor, date_column, allow_employee_fallback)
    result.attrs.update({"rows_read": rows_read, "blank_rows_skipped": blank})
    title = clean_text(output_name) or "Portfolio Enriched Audit"
    return result, audit_workbook_bytes(result, title, uploaded_file.name), output_filename(title)


def read_employee_details(uploaded_file) -> pd.DataFrame:
    details = read_table(uploaded_file, EMPLOYEE_DETAILS_EXPECTED, EMPLOYEE_DETAILS_REQUIRED)
    details["_key"] = details["EmployeeNumber"].map(employee_key)
    return details


def run_tool(key: str, name: str, fn: Callable[[], tuple], audit=None) -> ToolRun:
    """Execute one tool; any exception becomes a failed ToolRun instead of stopping the queue."""
    run = ToolRun(key=key, name=name, audit=audit)
    try:
        result, payload, filename = fn()
    except InputFileError as exc:
        run.status, run.error = "failed", str(exc)
        return run
    except Exception as exc:  # isolate unexpected failures from other tools
        log.exception("Tool failed: %s", key)
        run.status, run.error = "failed", f"{type(exc).__name__}: {exc}"
        return run
    run.result, run.payload, run.filename = result, payload, filename
    run.findings = len(result)
    run.rows_read = int(result.attrs.get("rows_read", 0))
    run.blank_rows_skipped = int(result.attrs.get("blank_rows_skipped", 0))
    run.notes = list(result.attrs.get("summary_messages", []))
    if run.blank_rows_skipped:
        run.notes.append(f"{run.blank_rows_skipped} fully blank source row(s) skipped.")
    if "Portfolio Description" in result.columns and result["Portfolio Description"].eq("No Matched Portfolio").any():
        run.status = "warning"
    return run


def run_report_frame(runs: list[ToolRun]) -> pd.DataFrame:
    return pd.DataFrame([r.as_report_row() for r in runs])
