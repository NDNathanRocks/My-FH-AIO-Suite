"""PC Portfolio enrichment.

Match order: EE ID + exact date/time -> EE ID + date -> EE ID only -> unmatched.
The matched Schedule Audit location (final path segment removed) is matched
case-insensitively to ``WFM Dept Path`` in the Supervisor-Portfolio report.
"""
from __future__ import annotations

import re

import pandas as pd

from .common.files import read_table
from .common.text import clean_text, employee_key, parse_date_value, parse_time_value
from .config import SCHEDULE_AUDIT_COLUMNS, SUPERVISOR_COLUMNS

PORTFOLIO_OUTPUT_COLUMNS = ("Portfolio Description", "Portfolio Email", "Coordinator Name",
                            "Matched WFM Dept Path", "Portfolio Match Note")


def clean_transfer_path(value) -> str:
    text = clean_text(value).rstrip("/")
    return text.rsplit("/", 1)[0].strip() if "/" in text else ""


def prepare_schedule_audit(uploaded_file) -> pd.DataFrame:
    frame = read_table(uploaded_file, SCHEDULE_AUDIT_COLUMNS, SCHEDULE_AUDIT_COLUMNS)
    frame["_ee"] = frame["EE ID"].map(employee_key)
    frame["_shift_date"] = frame["Shift Date"].map(parse_date_value).dt.date
    frame["_shift_time"] = frame["Shift Start Time"].map(parse_time_value)
    edit_date = frame["Schedule Audit Edit Date"].map(parse_date_value)
    edit_time = frame["Schedule Audit Edit Time"].map(clean_text)
    frame["_edited"] = pd.to_datetime(edit_date.dt.strftime("%Y-%m-%d") + " " + edit_time, errors="coerce")
    frame["_clean_path"] = frame["Scheduled Transfer Location"].map(clean_transfer_path)
    frame["_primary_path"] = frame["Primary location (path)"].map(lambda v: clean_text(v).strip().rstrip("/"))
    return frame.set_index(["_ee", "_shift_date"], drop=False).sort_index()


def prepare_supervisor_portfolio(uploaded_file) -> pd.DataFrame:
    frame = read_table(uploaded_file, SUPERVISOR_COLUMNS, SUPERVISOR_COLUMNS)
    frame["_path"] = frame["WFM Dept Path"].map(lambda v: clean_text(v).casefold())
    return frame.set_index("_path", drop=False).sort_index()


def _latest_nonblank(candidates, column):
    usable = candidates[candidates[column].fillna("").astype(str).str.strip().ne("")]
    if usable.empty:
        return None
    return usable.sort_values("_edited", ascending=False, na_position="last").iloc[0]


def _employee_candidates(schedule_audit, ee):
    if not ee or not isinstance(schedule_audit.index, pd.MultiIndex):
        return schedule_audit.iloc[0:0].copy()
    try:
        return schedule_audit.xs(ee, level="_ee", drop_level=False).copy()
    except KeyError:
        return schedule_audit.iloc[0:0].copy()


def _dated_candidates(schedule_audit, ee, audit_dt):
    if not ee or pd.isna(audit_dt) or not isinstance(schedule_audit.index, pd.MultiIndex):
        return schedule_audit.iloc[0:0].copy()
    try:
        return schedule_audit.loc[[(ee, audit_dt.date())]].copy()
    except KeyError:
        return schedule_audit.iloc[0:0].copy()


def resolve_location(ee, raw_date, use_date, allow_employee_fallback, schedule_audit):
    """Return (path, note) for one employee/date key."""
    audit_dt = parse_date_value(raw_date) if use_date else pd.NaT
    dated = _dated_candidates(schedule_audit, ee, audit_dt) if use_date else schedule_audit.iloc[0:0].copy()

    if use_date and not pd.isna(audit_dt) and not dated.empty:
        has_time = bool(re.search(r"\d{1,2}:\d{2}", clean_text(raw_date)))
        if has_time:
            exact = dated[dated["_shift_time"] == audit_dt.time()]
            selected = _latest_nonblank(exact, "_clean_path")
            if selected is not None:
                return selected["_clean_path"], "Matched Scheduled Transfer Location by EE ID + exact date/time"
        selected = _latest_nonblank(dated, "_clean_path")
        if selected is not None:
            reason = "Date-only Scheduled Transfer Location fallback" if has_time else "Matched Scheduled Transfer Location by EE ID + date"
            return selected["_clean_path"], reason
        selected = _latest_nonblank(dated, "_primary_path")
        if selected is not None:
            return selected["_primary_path"], "Primary location (path) fallback used for matched EE ID + date"

    if allow_employee_fallback:
        employee_rows = _employee_candidates(schedule_audit, ee)
        if not employee_rows.empty:
            selected = _latest_nonblank(employee_rows, "_clean_path")
            if selected is not None:
                return selected["_clean_path"], "EE ID-only Scheduled Transfer Location fallback used"
            selected = _latest_nonblank(employee_rows, "_primary_path")
            if selected is not None:
                return selected["_primary_path"], "EE ID-only Primary location (path) fallback used"

    return "", "No Employee or Primary Location Path Found in Audit Schedule"


def _portfolio_lookup(supervisor):
    records, messages = [], []
    for path_key, matches in supervisor.reset_index(drop=True).groupby("_path", sort=False):
        if not path_key:
            continue
        if len(matches) > 1:
            messages.append(
                f"{len(matches)} Supervisor-Portfolio rows matched path "
                f"'{clean_text(matches.iloc[0]['WFM Dept Path'])}'; first nonblank values used.")

        def first(column, matches=matches):
            return next((clean_text(v) for v in matches[column] if clean_text(v)), "")
        records.append({
            "_path": path_key,
            "Portfolio Description": first("Portfolio Description") or "No Matched Portfolio",
            "Portfolio Email": first("Portfolio Email"),
            "Coordinator Name": first("Coordinator Name"),
            "_duplicate": len(matches) > 1,
        })
    return pd.DataFrame(records), messages


def add_location_portfolio(frame, id_column, schedule_audit, supervisor,
                           date_column=None, allow_employee_fallback=True):
    previous = list(frame.attrs.get("summary_messages", []))
    result = frame.drop(columns=[c for c in PORTFOLIO_OUTPUT_COLUMNS if c in frame], errors="ignore").copy()
    if id_column not in result.columns:
        raise ValueError(f"The selected employee identifier column '{id_column}' was not found.")
    if date_column and date_column not in result.columns:
        raise ValueError(f"The selected date column '{date_column}' was not found.")

    use_date = bool(date_column)
    key_rows = pd.DataFrame(index=result.index)
    key_rows["_ee"] = result[id_column].map(employee_key)
    key_rows["_raw_date"] = result[date_column].map(clean_text) if use_date else ""

    resolved = []
    for _, key in key_rows.drop_duplicates(["_ee", "_raw_date"]).iterrows():
        ee, raw_date = key["_ee"], key["_raw_date"]
        if not ee:
            resolved.append((ee, raw_date, "", "Missing employee identifier"))
        else:
            path, note = resolve_location(ee, raw_date, use_date, allow_employee_fallback, schedule_audit)
            resolved.append((ee, raw_date, path, note))
    resolution = pd.DataFrame(resolved, columns=["_ee", "_raw_date", "Matched WFM Dept Path", "Portfolio Match Note"])
    key_rows = key_rows.merge(resolution, how="left", on=["_ee", "_raw_date"], sort=False)
    key_rows.index = result.index

    lookup, duplicate_messages = _portfolio_lookup(supervisor)
    key_rows["_path"] = key_rows["Matched WFM Dept Path"].map(lambda v: clean_text(v).casefold())
    if not lookup.empty:
        key_rows = key_rows.merge(lookup, how="left", on="_path", sort=False)
        key_rows.index = result.index
    else:
        for column in ("Portfolio Description", "Portfolio Email", "Coordinator Name"):
            key_rows[column] = pd.NA
        key_rows["_duplicate"] = False

    no_path = key_rows["Matched WFM Dept Path"].fillna("").astype(str).str.strip().eq("")
    no_portfolio = key_rows["Portfolio Description"].isna()
    key_rows.loc[no_path | no_portfolio, "Portfolio Description"] = "No Matched Portfolio"
    key_rows.loc[no_portfolio & ~no_path, "Portfolio Match Note"] = (
        "No WFM Dept Path match; " + key_rows.loc[no_portfolio & ~no_path, "Portfolio Match Note"].fillna(""))
    dup = key_rows["_duplicate"].astype("boolean").fillna(False).astype(bool)
    key_rows.loc[dup, "Portfolio Match Note"] = (
        key_rows.loc[dup, "Portfolio Match Note"].fillna("") + "; Multiple portfolio rows; first nonblank values used")

    result = pd.concat([result, key_rows[list(PORTFOLIO_OUTPUT_COLUMNS)]], axis=1)
    notes = result["Portfolio Match Note"].fillna("")
    result.attrs["summary_messages"] = previous + duplicate_messages + [
        f"{int(result['Portfolio Description'].eq('No Matched Portfolio').sum())} row(s) had no matched portfolio.",
        f"{int(notes.str.contains('Date-only', case=False).sum())} row(s) used date-only fallback.",
        f"{int(notes.str.contains('EE ID-only', case=False).sum())} row(s) used EE ID-only fallback.",
        f"{int(notes.str.contains('Primary location', case=False).sum())} row(s) used Primary location fallback.",
        f"{int(notes.str.contains('No Employee or Primary Location Path Found', case=False).sum())} row(s) had no employee/location path in Schedule Audit.",
    ]
    return result
