"""Run every audit path on the synthetic fixtures and return comparable data.

Works against both the legacy single-file module (via ``load_legacy``) and the
refactored package (via an adapter exposing the same function names). Used to
prove the refactor produces identical findings.
"""
from __future__ import annotations

import io
import json
from datetime import date

import openpyxl
import pandas as pd

import fixtures as fx

TOR_CUTOFF = date(2026, 10, 10)
VOLATILE = {"Days Since Creation"}


def frame_record(frame: pd.DataFrame) -> dict:
    visible = frame.drop(columns=[c for c in frame.columns if str(c).startswith("_") or c in VOLATILE], errors="ignore")
    return {
        "columns": [str(c) for c in visible.columns],
        "rows": [["" if pd.isna(v) else str(v) for v in row] for row in visible.itertuples(index=False)],
    }


def sheet_record(payload: bytes) -> dict:
    wb = openpyxl.load_workbook(io.BytesIO(payload))
    ws = wb["Audit Results"]
    rows = [["" if v is None else str(v) for v in r] for r in ws.iter_rows(min_row=5, values_only=True)]
    return {"sheets": wb.sheetnames, "title": ws["A1"].value, "rows": rows}


class Upload:
    def __init__(self, name: str, data: bytes):
        self.name, self.data = name, data

    def getvalue(self) -> bytes:
        return self.data


def upload(key: str, spec, kind: str = "csv") -> Upload:
    header, rows = spec
    data = fx.csv_bytes(header, rows) if kind == "csv" else fx.xlsx_bytes(header, rows)
    return Upload(f"{key}.{kind}", data)


def run_all(m) -> dict:
    """``m`` exposes the legacy names: AUDITS, process_audit, read_raw, locate_table, ..."""
    out: dict = {}
    sched = m.read_schedule_audit(upload("schedule_audit", fx.SCHEDULE_AUDIT))
    sup = m.read_supervisor_portfolio(upload("supervisor", fx.SUPERVISOR))
    for key, audit in m.AUDITS.items():
        if key == "leave_cascade":
            continue
        for kind in ("csv", "xlsx"):
            src = upload(audit.source_key, fx.SOURCES[audit.source_key], kind)
            opts = {"cutoff_date": TOR_CUTOFF} if key == "unprocessed_tors" else {}
            result, payload, filename = m.process_audit(audit, src, None, None, opts, None)
            out[f"{key}|{kind}|plain"] = {"frame": frame_record(result), "filename": filename.replace(str(date.today()), "DATE")}
            if kind == "csv":
                out[f"{key}|{kind}|plain"]["sheet"] = sheet_record(payload)
            for fallback in (True, False):
                match = {"id_column": audit.employee_id_column, "date_column": None, "allow_employee_fallback": fallback}
                result, payload, _ = m.process_audit(audit, src, sched, sup, opts, match)
                out[f"{key}|{kind}|portfolio|{fallback}"] = {"frame": frame_record(result), "summary": list(result.attrs.get("summary_messages", []))}
                if "Portfolio Email" in result.columns and len(result):
                    to, cc, att = m.email_details(result)
                    out[f"{key}|{kind}|email|{fallback}"] = {"to": to, "cc": cc, "attachment_rows": len(att)}
    # Leave cascade (two sources)
    audit = m.AUDITS["leave_cascade"]
    sched_tbl = m.require_columns(m.locate_table(m.read_raw(upload("s", fx.SCHEDULE_BY_PAYCODE)), audit.expected_columns, audit.required_columns), audit.required_columns)
    time_req = ("EE ID", "Actual Total Hrs (Include Corrections)", "Pay Code Name")
    time_tbl = m.require_columns(m.locate_table(m.read_raw(upload("t", fx.TIME_SUMMARY)), time_req, time_req), time_req)
    lc = m.audit_leave_cascade(sched_tbl, time_tbl)
    out["leave_cascade"] = {"frame": frame_record(lc), "summary": list(lc.attrs.get("summary_messages", [])),
                            "estimated": [float(v) for v in lc["_estimated_minutes"]]}
    lcp = m.add_location_portfolio(lc, "EE ID", sched, sup, "Start Date", True)
    out["leave_cascade|portfolio"] = {"frame": frame_record(lcp), "summary": list(lcp.attrs.get("summary_messages", []))}
    # Portfolio-only tool, with and without a date column
    existing = upload("existing", fx.EXISTING_FILTERED)
    for date_col in ("Shift Date", None):
        for fallback in (True, False):
            result, _, filename = m.process_portfolio_only(existing, sched, sup, "Employee Number", date_col, fallback, "My Tool")
            out[f"portfolio_only|{date_col}|{fallback}"] = {"frame": frame_record(result), "summary": list(result.attrs.get("summary_messages", [])),
                                                             "filename": filename.replace(str(date.today()), "DATE")}
    out["employee_key"] = [m.employee_key(v) for v in ("1001", "1001.0", " 77 ", None, 12.0)]
    out["time_to_minutes"] = [m.time_to_minutes(v) for v in ("7:30", "=-1:15", "1:02:30", "abc", "", "100:00")]
    out["safe_filename"] = [m.safe_filename(v) for v in ('A/B:C*"?<>|', "  .x. ", "", "Normal name")]
    return out


def dumps(data: dict) -> str:
    return json.dumps(data, indent=1, sort_keys=True, default=str)
