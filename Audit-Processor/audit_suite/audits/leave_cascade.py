"""CTO/OTO/STO leave cascading to LOA (two-source audit).

Matching NCO paycode hours are allocated oldest-first per leave type; LOA hours
explicitly scheduled are removed from the LOA estimate; the remaining
estimated cascaded LOA is distributed chronologically across shortages.
"""
from __future__ import annotations

import re

import pandas as pd

from ..common.text import clean_text, employee_key, minutes_to_hhmm, parse_date_value, time_to_minutes

CODE_MAP = {"CTO": "96.1-NCO CTO", "OTO": "96.2-NCO OTO", "STO": "98.1-NCO STO"}
RESULT_COLUMNS = ["Audit Finding", "EE ID", "Last Name", "First Name", "Start Date", "Duration (Hrs)",
                  "Actual Total Hrs (Include Corrections)", "Pay Code Name (Schedules)", "Pay Code Name (Timecards)",
                  "Worker Type", "Labour Agreement", "Estimated Cascaded to LOA", "Unreconciled Leave Shortage",
                  "Employee LOA Total", "Scheduled LOA Hours Removed", "Leave Type", "_estimated_minutes"]
_LEAVE_RE = re.compile(r"\b(STO|CTO|OTO)\s+Leave\b", re.I)
_LEAVE_FILTER_RE = re.compile(r"\b(?:STO|CTO|OTO)\s+Leave\b", re.I)


def _first_nonblank(ts, column):
    return next((clean_text(v) for v in ts[column] if clean_text(v)), "") if column in ts else ""


def audit_leave_cascade(schedule: pd.DataFrame, time_summary: pd.DataFrame) -> pd.DataFrame:
    schedule = schedule.copy()
    time_summary = time_summary.copy()
    schedule["_ee"] = schedule["EE ID"].map(employee_key)
    schedule["_date"] = schedule["Start Date"].map(parse_date_value)
    schedule["_minutes"] = schedule["Duration (Hrs)"].map(time_to_minutes)
    schedule["_pc"] = schedule["Pay Code Name"].map(clean_text)
    time_summary["_ee"] = time_summary["EE ID"].map(employee_key)
    time_summary["_minutes"] = time_summary["Actual Total Hrs (Include Corrections)"].map(time_to_minutes)
    time_summary["_pc"] = time_summary["Pay Code Name"].map(clean_text)
    leaves = schedule[schedule["_pc"].str.contains(_LEAVE_FILTER_RE, na=False)].copy()
    leaves["Leave Type"] = leaves["_pc"].str.extract(_LEAVE_RE, expand=False).str.upper()
    loa_employees = set(time_summary.loc[time_summary["_pc"].str.contains("LOA", case=False, na=False), "_ee"])
    leaves = leaves[leaves["_ee"].isin(loa_employees)].sort_values(["_ee", "_date"], kind="stable")
    rows, summaries = [], []
    for ee, group in leaves.groupby("_ee", sort=False):
        ts = time_summary[time_summary["_ee"] == ee]
        available = {leave_type: float(ts.loc[ts["_pc"].str.casefold().eq(code.casefold()), "_minutes"].fillna(0).sum())
                     for leave_type, code in CODE_MAP.items()}
        total_loa = float(ts.loc[ts["_pc"].str.contains("LOA", case=False, na=False), "_minutes"].fillna(0).sum())
        scheduled_loa = float(schedule.loc[(schedule["_ee"] == ee) & schedule["_pc"].str.contains("LOA", case=False, na=False), "_minutes"].fillna(0).sum())
        cascade_pool = max(0.0, total_loa - scheduled_loa)
        staged = []
        for idx, row in group.iterrows():
            leave_type = row["Leave Type"]
            scheduled = float(row["_minutes"] or 0)
            matched = min(scheduled, available[leave_type])
            available[leave_type] -= matched
            staged.append((idx, row, matched, max(0.0, scheduled - matched)))
        worker_type, agreement = _first_nonblank(ts, "Worker Type"), _first_nonblank(ts, "Labour Agreement")
        for _idx, row, matched, shortage in staged:
            estimated = min(shortage, cascade_pool)
            cascade_pool -= estimated
            rows.append({
                "Audit Finding": "Hours Match" if shortage <= 0 else "Hours Mismatched",
                "EE ID": row["EE ID"], "Last Name": row.get("Last Name", ""), "First Name": row.get("First Name", ""),
                "Start Date": row["Start Date"], "Duration (Hrs)": minutes_to_hhmm(row["_minutes"]),
                "Actual Total Hrs (Include Corrections)": minutes_to_hhmm(matched),
                "Pay Code Name (Schedules)": row["Pay Code Name"], "Pay Code Name (Timecards)": CODE_MAP[row["Leave Type"]],
                "Worker Type": worker_type, "Labour Agreement": agreement,
                "Estimated Cascaded to LOA": minutes_to_hhmm(estimated), "Unreconciled Leave Shortage": minutes_to_hhmm(max(0, shortage - estimated)),
                "Employee LOA Total": minutes_to_hhmm(total_loa), "Scheduled LOA Hours Removed": minutes_to_hhmm(scheduled_loa),
                "Leave Type": row["Leave Type"], "_estimated_minutes": estimated,
            })
        if cascade_pool > 0:
            summaries.append(f"EE {ee}: {minutes_to_hhmm(cascade_pool)} LOA hours remained after chronological shortage allocation.")
    result = pd.DataFrame(rows) if rows else pd.DataFrame(columns=RESULT_COLUMNS)
    result.attrs["summary_messages"] = summaries
    return result
