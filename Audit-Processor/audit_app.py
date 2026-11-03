from __future__ import annotations

import csv
import io
import random
import re
import tempfile
import time
from dataclasses import dataclass
from datetime import date, timedelta
from pathlib import Path
from types import SimpleNamespace
from typing import Callable

import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter

APP_TITLE = "Audit Prepper"
DATE_FMT = "%Y-%m-%d"
HELP_EMAIL = "mySchedulehelpline@fraserhealth.ca"
MAX_UPLOAD_BYTES = 200 * 1024 * 1024
EMAIL_RE = re.compile(r"[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}", re.IGNORECASE)

# This is intentionally a soft map. If a coordinator is absent, Outlook receives the cleaned name.
COORDINATOR_EMAILS = {
    "patrice fugah": "patrice.fugah2@fraserhealth.ca",
    "rachelle tambeau": "rachelle.tambeau@fraserhealth.ca",
    "heather whitworth": "Heather.Whitworth@fraserhealth.ca",
    "sandy ardawa": "sandy.ardawa@fraserhealth.ca",
    "rashika nand": "Rashika.Nand@fraserhealth.ca",
    "praveena gopal": "Praveena.Gopal@fraserhealth.ca",
    "ekjot mahli": "ekjot.malhi@fraserhealth.ca",
    "ekjot malhi": "ekjot.malhi@fraserhealth.ca",
}


SCHEDULE_AUDIT_URL = "https://fraserhealthaut.prd.mykronos.com/ia#/dataView/1708"
SUPERVISOR_PORTFOLIO_URL = "https://app.powerbi.com/groups/me/reports/a99ac431-f485-4899-8a9c-9821fe7951a3/227eea054fdff4092823?ctid=31f660a5-192a-4db3-92ba-ca424f1b259e&experience=power-bi&clientSideAuth=0"
EMPLOYEE_DETAILS_URL = "https://app.powerbi.com/groups/me/reports/b68931ac-cc8c-4093-a0a5-04db175577b8/48704422ffb5fa4db8e2?ctid=31f660a5-192a-4db3-92ba-ca424f1b259e&experience=power-bi"
SCHEDULE_AUDIT_COLUMNS = ("EE ID", "Shift Date", "Shift Start Time", "Scheduled Transfer Location", "Primary location (path)", "Schedule Audit Edit Date", "Schedule Audit Edit Time")
SUPERVISOR_COLUMNS = ("Coordinator Name", "Portfolio Description", "Portfolio Email", "WFM Dept Path")
DATE_COLUMN_PRIORITY = ("Start Date Time", "Shift Date Time", "Start Date", "Shift Date")


@dataclass(frozen=True)
class AuditDefinition:
    key: str
    name: str
    source_key: str
    source_name: str
    source_url: str
    export_settings: str
    expected_columns: tuple[str, ...]
    required_columns: tuple[str, ...]
    employee_id_column: str
    email_subject: str | None
    email_body_html: str | None
    steps: tuple[str, ...]
    processor: Callable[..., pd.DataFrame]


@dataclass(frozen=True)
class MemoryUpload:
    """Upload snapshot that survives Streamlit widget cleanup between stages."""
    name: str
    data: bytes

    def getvalue(self) -> bytes:
        return self.data


def clean_text(value) -> str:
    return "" if pd.isna(value) else str(value).strip()


def normalize_header(value) -> str:
    return re.sub(r"\s+", " ", clean_text(value)).casefold()


DATE_COLUMN_ALIASES = (
    "Start Date Time", "Shift Date Time", "Start Date", "Shift Date",
    "Date", "Work Date", "Workdate",
)
EMPLOYEE_ID_ALIASES = ("EE ID", "Employee ID", "EmployeeNumber", "Employee Number")
NO_DATE_OPTION = "No date column - use EE ID-only fallback"


def best_column_match(columns, aliases):
    normalized = {normalize_header(column): column for column in columns}
    return next((normalized.get(normalize_header(alias)) for alias in aliases
                 if normalized.get(normalize_header(alias)) is not None), None)


def available_date_columns(columns):
    columns = list(columns)
    preferred = []
    normalized = {normalize_header(column): column for column in columns}
    for alias in DATE_COLUMN_ALIASES:
        actual = normalized.get(normalize_header(alias))
        if actual is not None and actual not in preferred:
            preferred.append(actual)
    likely = [c for c in columns if c not in preferred and "date" in normalize_header(c)]
    remaining = [c for c in columns if c not in preferred and c not in likely]
    return preferred + likely + remaining


def contains(series: pd.Series, text: str) -> pd.Series:
    return series.fillna("").astype(str).str.contains(text, case=False, regex=False)


def employee_key(value) -> str:
    return re.sub(r"\.0$", "", clean_text(value))


def cache_upload(cache_key: str, uploaded_file) -> MemoryUpload | None:
    cache = st.session_state.setdefault("upload_cache", {})
    if uploaded_file is not None:
        data = uploaded_file.getvalue()
        if len(data) > MAX_UPLOAD_BYTES:
            raise ValueError("The uploaded file exceeds the 200 MB limit.")
        cache[cache_key] = MemoryUpload(Path(uploaded_file.name).name, bytes(data))
    return cache.get(cache_key)


def cached_upload(cache_key: str) -> MemoryUpload | None:
    return st.session_state.get("upload_cache", {}).get(cache_key)


def clear_workflow_cache() -> None:
    st.session_state["upload_cache"] = {}
    st.session_state.pop("run_config", None)
    st.session_state.pop("audit_results", None)


def read_raw(uploaded_file) -> list[list[object]]:
    data = uploaded_file.getvalue()
    suffix = Path(uploaded_file.name).suffix.lower()
    if suffix == ".csv":
        last_error = None
        for encoding in ("utf-8-sig", "utf-8", "cp1252", "latin1"):
            try:
                text = data.decode(encoding)
                try:
                    delimiter = csv.Sniffer().sniff(text[:8192], delimiters=",;\t|").delimiter
                except csv.Error:
                    delimiter = ","
                rows = list(csv.reader(io.StringIO(text, newline=""), delimiter=delimiter))
                if not rows:
                    raise ValueError("The CSV is empty.")
                return rows
            except Exception as exc:
                last_error = exc
        raise ValueError(f"The CSV could not be read. Last error: {last_error}")
    if suffix in (".xlsx", ".xlsm"):
        return pd.read_excel(io.BytesIO(data), header=None, dtype=object, engine="openpyxl").values.tolist()
    if suffix == ".xls":
        return pd.read_excel(io.BytesIO(data), header=None, dtype=object, engine="xlrd").values.tolist()
    raise ValueError("Unsupported file type. Upload CSV, XLSX, XLSM, or XLS.")


def locate_table(raw_rows, expected_columns, required_columns=None) -> pd.DataFrame:
    expected_norm = {normalize_header(c) for c in expected_columns}
    required_norm = {normalize_header(c) for c in (required_columns or expected_columns)}
    candidates = []
    for index, row in enumerate(raw_rows[:50]):
        row_norm = {normalize_header(v) for v in row if clean_text(v)}
        candidates.append((required_norm.issubset(row_norm), len(expected_norm & row_norm), index))
    has_required, _, header_index = max(candidates, default=(False, 0, -1))
    if not has_required:
        raise ValueError("Could not identify the table header. Required columns: " + ", ".join(required_columns or expected_columns))
    return table_from_header(raw_rows, header_index)


def table_from_header(raw_rows, header_index: int) -> pd.DataFrame:
    headers = [clean_text(v) for v in raw_rows[header_index]]
    while headers and not headers[-1]:
        headers.pop()
    if not headers:
        raise ValueError("The uploaded file does not contain a usable header row.")
    seen, unique_headers = {}, []
    for i, header in enumerate(headers):
        base = header or f"Unnamed Column {i + 1}"
        seen[base] = seen.get(base, 0) + 1
        unique_headers.append(base if seen[base] == 1 else f"{base} ({seen[base]})")
    width = len(unique_headers)
    body = [list(row[:width]) + [None] * max(0, width - len(row)) for row in raw_rows[header_index + 1:]]
    frame = pd.DataFrame(body, columns=unique_headers).dropna(how="all")
    return frame.loc[:, ~frame.columns.astype(str).str.startswith("Unnamed Column")].reset_index(drop=True)


@st.cache_data(show_spinner=False, max_entries=8)
def _cached_existing_table(file_name: str, file_data: bytes) -> pd.DataFrame:
    uploaded_file = MemoryUpload(file_name, file_data)
    raw = read_raw(uploaded_file)
    best_index, best_score = -1, -1
    for index, row in enumerate(raw[:50]):
        values = [clean_text(v) for v in row if clean_text(v)]
        normalized = {normalize_header(v) for v in values}
        score = len(values) + (1000 if "ee id" in normalized or "employee id" in normalized else 0)
        if score > best_score:
            best_index, best_score = index, score
    if best_index < 0:
        raise ValueError("Could not find a table header in the uploaded file.")
    return table_from_header(raw, best_index)


def read_existing_table(uploaded_file) -> pd.DataFrame:
    return _cached_existing_table(uploaded_file.name, uploaded_file.getvalue())


def require_columns(frame, columns):
    lookup = {normalize_header(c): c for c in frame.columns}
    missing = [c for c in columns if normalize_header(c) not in lookup]
    if missing:
        raise ValueError("Missing required column(s): " + ", ".join(missing))
    return frame.rename(columns={lookup[normalize_header(c)]: c for c in columns})


def sort_if_present(frame, columns):
    present = [c for c in columns if c in frame.columns]
    return frame.sort_values(present, kind="stable", na_position="last") if present and not frame.empty else frame


def time_to_minutes(value):
    text = clean_text(value).replace("=", "").replace(" ", "")
    match = re.fullmatch(r"(-?)(\d+):(\d{2})(?::(\d{2}))?", text)
    if not match:
        return None
    sign = -1 if match.group(1) == "-" else 1
    return sign * (int(match.group(2)) * 60 + int(match.group(3)) + (int(match.group(4) or 0) / 60))



def parse_date_value(value):
    return pd.to_datetime(clean_text(value), errors="coerce", dayfirst=True)


def parse_time_value(value):
    text = clean_text(value)
    parsed = pd.to_datetime(text, errors="coerce")
    return None if pd.isna(parsed) else parsed.time()


def clean_transfer_path(value):
    text = clean_text(value).rstrip("/")
    return text.rsplit("/", 1)[0].strip() if "/" in text else ""


@st.cache_data(show_spinner=False, max_entries=6)
def _cached_schedule_audit(file_name: str, file_data: bytes):
    uploaded_file = MemoryUpload(file_name, file_data)
    frame = require_columns(locate_table(read_raw(uploaded_file), SCHEDULE_AUDIT_COLUMNS, SCHEDULE_AUDIT_COLUMNS), SCHEDULE_AUDIT_COLUMNS)
    frame["_ee"] = frame["EE ID"].map(employee_key)
    frame["_shift_date"] = frame["Shift Date"].map(parse_date_value).dt.date
    frame["_shift_time"] = frame["Shift Start Time"].map(parse_time_value)
    edit_date = frame["Schedule Audit Edit Date"].map(parse_date_value)
    edit_time = frame["Schedule Audit Edit Time"].map(clean_text)
    frame["_edited"] = pd.to_datetime(edit_date.dt.strftime("%Y-%m-%d") + " " + edit_time, errors="coerce")
    frame["_clean_path"] = frame["Scheduled Transfer Location"].map(clean_transfer_path)
    frame["_primary_path"] = frame["Primary location (path)"].map(lambda v: clean_text(v).strip().rstrip("/"))
    return frame.set_index(["_ee", "_shift_date"], drop=False).sort_index()


def read_schedule_audit(uploaded_file):
    return _cached_schedule_audit(uploaded_file.name, uploaded_file.getvalue())


@st.cache_data(show_spinner=False, max_entries=6)
def _cached_supervisor_portfolio(file_name: str, file_data: bytes):
    uploaded_file = MemoryUpload(file_name, file_data)
    frame = require_columns(locate_table(read_raw(uploaded_file), SUPERVISOR_COLUMNS, SUPERVISOR_COLUMNS), SUPERVISOR_COLUMNS)
    frame["_path"] = frame["WFM Dept Path"].map(lambda v: clean_text(v).casefold())
    return frame.set_index("_path", drop=False).sort_index()


def read_supervisor_portfolio(uploaded_file):
    return _cached_supervisor_portfolio(uploaded_file.name, uploaded_file.getvalue())


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


def _resolve_location_key(ee, raw_date, use_date, allow_employee_fallback, schedule_audit):
    audit_dt = parse_date_value(raw_date) if use_date else pd.NaT
    dated = _dated_candidates(schedule_audit, ee, audit_dt) if use_date else schedule_audit.iloc[0:0].copy()

    # 1. Employee + exact date/time.
    if use_date and not pd.isna(audit_dt) and not dated.empty:
        has_time = bool(re.search(r"\d{1,2}:\d{2}", clean_text(raw_date)))
        if has_time:
            exact = dated[dated["_shift_time"] == audit_dt.time()]
            selected = _latest_nonblank(exact, "_clean_path")
            if selected is not None:
                return selected["_clean_path"], "Matched Scheduled Transfer Location by EE ID + exact date/time"

        # 2. Employee + date.
        selected = _latest_nonblank(dated, "_clean_path")
        if selected is not None:
            reason = "Date-only Scheduled Transfer Location fallback" if has_time else "Matched Scheduled Transfer Location by EE ID + date"
            return selected["_clean_path"], reason
        selected = _latest_nonblank(dated, "_primary_path")
        if selected is not None:
            return selected["_primary_path"], "Primary location (path) fallback used for matched EE ID + date"

    # 3. Employee only. Search transfer first, then primary location.
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


def add_location_portfolio(frame, id_column, schedule_audit, supervisor,
                           date_column=None, allow_employee_fallback=True):
    """Fallback order: ID+datetime -> ID+date -> ID only -> unmatched."""
    previous = list(frame.attrs.get("summary_messages", []))
    result = frame.drop(columns=[c for c in (
        "Portfolio Description", "Portfolio Email", "Coordinator Name",
        "Portfolio Match Note", "Matched WFM Dept Path") if c in frame],
        errors="ignore").copy()
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
            path, note = _resolve_location_key(ee, raw_date, use_date,
                                               allow_employee_fallback, schedule_audit)
            resolved.append((ee, raw_date, path, note))
    resolution = pd.DataFrame(resolved, columns=[
        "_ee", "_raw_date", "Matched WFM Dept Path", "Portfolio Match Note"])
    key_rows = key_rows.merge(resolution, how="left", on=["_ee", "_raw_date"], sort=False)
    key_rows.index = result.index

    # Build one row per normalized portfolio path once, then batch join.
    portfolio_records, duplicate_messages = [], []
    for path_key, matches in supervisor.reset_index(drop=True).groupby("_path", sort=False):
        if not path_key:
            continue
        if len(matches) > 1:
            duplicate_messages.append(
                f"{len(matches)} Supervisor-Portfolio rows matched path "
                f"'{clean_text(matches.iloc[0]['WFM Dept Path'])}'; first nonblank values used.")
        def first(column):
            return next((clean_text(v) for v in matches[column] if clean_text(v)), "")
        portfolio_records.append({
            "_path": path_key,
            "Portfolio Description": first("Portfolio Description") or "No Matched Portfolio",
            "Portfolio Email": first("Portfolio Email"),
            "Coordinator Name": first("Coordinator Name"),
            "_duplicate": len(matches) > 1,
        })
    lookup = pd.DataFrame(portfolio_records)
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
        "No WFM Dept Path match; " +
        key_rows.loc[no_portfolio & ~no_path, "Portfolio Match Note"].fillna(""))
    dup = key_rows["_duplicate"].fillna(False)
    key_rows.loc[dup, "Portfolio Match Note"] = (
        key_rows.loc[dup, "Portfolio Match Note"].fillna("") +
        "; Multiple portfolio rows; first nonblank values used")

    result = pd.concat([result, key_rows[[
        "Portfolio Description", "Portfolio Email", "Coordinator Name",
        "Matched WFM Dept Path", "Portfolio Match Note"]]], axis=1)
    notes = result["Portfolio Match Note"].fillna("")
    result.attrs["summary_messages"] = previous + duplicate_messages + [
        f"{int(result['Portfolio Description'].eq('No Matched Portfolio').sum())} row(s) had no matched portfolio.",
        f"{int(notes.str.contains('Date-only', case=False).sum())} row(s) used date-only fallback.",
        f"{int(notes.str.contains('EE ID-only', case=False).sum())} row(s) used EE ID-only fallback.",
        f"{int(notes.str.contains('Primary location', case=False).sum())} row(s) used Primary location fallback.",
        f"{int(notes.str.contains('No Employee or Primary Location Path Found', case=False).sum())} row(s) had no employee/location path in Schedule Audit.",
    ]
    return result


def minutes_to_hhmm(minutes):
    minutes = max(0, int(round(float(minutes or 0))))
    return f"{minutes // 60}:{minutes % 60:02d}"


def natural_duration(minutes):
    minutes = max(0, int(round(float(minutes or 0))))
    hours, mins = divmod(minutes, 60)
    parts = []
    if hours: parts.append(f"{hours} hour" + ("s" if hours != 1 else ""))
    if mins: parts.append(f"{mins} minute" + ("s" if mins != 1 else ""))
    return " ".join(parts) or "0 minutes"


def audit_leave_cascade(schedule, time_summary):
    schedule = schedule.copy()
    time_summary = time_summary.copy()
    schedule["_ee"] = schedule["EE ID"].map(employee_key)
    schedule["_date"] = schedule["Start Date"].map(parse_date_value)
    schedule["_minutes"] = schedule["Duration (Hrs)"].map(time_to_minutes)
    schedule["_pc"] = schedule["Pay Code Name"].map(clean_text)
    time_summary["_ee"] = time_summary["EE ID"].map(employee_key)
    time_summary["_minutes"] = time_summary["Actual Total Hrs (Include Corrections)"].map(time_to_minutes)
    time_summary["_pc"] = time_summary["Pay Code Name"].map(clean_text)
    leave_re = re.compile(r"\b(STO|CTO|OTO)\s+Leave\b", re.I)
    leave_filter_re = re.compile(r"\b(?:STO|CTO|OTO)\s+Leave\b", re.I)
    leaves = schedule[schedule["_pc"].str.contains(leave_filter_re, na=False)].copy()
    leaves["Leave Type"] = leaves["_pc"].str.extract(leave_re, expand=False).str.upper()
    loa_employees = set(time_summary.loc[time_summary["_pc"].str.contains("LOA", case=False, na=False), "_ee"])
    leaves = leaves[leaves["_ee"].isin(loa_employees)].sort_values(["_ee", "_date"], kind="stable")
    code_map = {"CTO": "96.1-NCO CTO", "OTO": "96.2-NCO OTO", "STO": "98.1-NCO STO"}
    rows = []
    summaries = []
    for ee, group in leaves.groupby("_ee", sort=False):
        ts = time_summary[time_summary["_ee"] == ee]
        available = {leave_type: float(ts.loc[ts["_pc"].str.casefold().eq(code.casefold()), "_minutes"].fillna(0).sum()) for leave_type, code in code_map.items()}
        total_loa = float(ts.loc[ts["_pc"].str.contains("LOA", case=False, na=False), "_minutes"].fillna(0).sum())
        scheduled_loa = float(schedule.loc[(schedule["_ee"] == ee) & schedule["_pc"].str.contains("LOA", case=False, na=False), "_minutes"].fillna(0).sum())
        cascade_pool = max(0.0, total_loa - scheduled_loa)
        staged = []
        for idx, row in group.iterrows():
            leave_type = row["Leave Type"]
            scheduled = float(row["_minutes"] or 0)
            matched = min(scheduled, available[leave_type])
            available[leave_type] -= matched
            shortage = max(0.0, scheduled - matched)
            staged.append((idx, row, matched, shortage))
        for idx, row, matched, shortage in staged:
            estimated = min(shortage, cascade_pool)
            cascade_pool -= estimated
            status = "Hours Match" if shortage <= 0 else "Hours Mismatched"
            rows.append({
                "Audit Finding": status,
                "EE ID": row["EE ID"], "Last Name": row.get("Last Name", ""), "First Name": row.get("First Name", ""),
                "Start Date": row["Start Date"], "Duration (Hrs)": minutes_to_hhmm(row["_minutes"]),
                "Actual Total Hrs (Include Corrections)": minutes_to_hhmm(matched),
                "Pay Code Name (Schedules)": row["Pay Code Name"], "Pay Code Name (Timecards)": code_map[row["Leave Type"]],
                "Worker Type": next((clean_text(v) for v in ts["Worker Type"] if clean_text(v)), "") if "Worker Type" in ts else "",
                "Labour Agreement": next((clean_text(v) for v in ts["Labour Agreement"] if clean_text(v)), "") if "Labour Agreement" in ts else "",
                "Estimated Cascaded to LOA": minutes_to_hhmm(estimated), "Unreconciled Leave Shortage": minutes_to_hhmm(max(0, shortage - estimated)),
                "Employee LOA Total": minutes_to_hhmm(total_loa), "Scheduled LOA Hours Removed": minutes_to_hhmm(scheduled_loa),
                "Leave Type": row["Leave Type"], "_estimated_minutes": estimated,
            })
        if cascade_pool > 0: summaries.append(f"EE {ee}: {minutes_to_hhmm(cascade_pool)} LOA hours remained after chronological shortage allocation.")
    result = pd.DataFrame(rows)
    if result.empty:
        result = pd.DataFrame(columns=["Audit Finding", "EE ID", "Last Name", "First Name", "Start Date", "Duration (Hrs)", "Actual Total Hrs (Include Corrections)", "Pay Code Name (Schedules)", "Pay Code Name (Timecards)", "Worker Type", "Labour Agreement", "Estimated Cascaded to LOA", "Unreconciled Leave Shortage", "Employee LOA Total", "Scheduled LOA Hours Removed", "Leave Type", "_estimated_minutes"])
    result.attrs["summary_messages"] = summaries
    return result


# ----------------------------- Audit processors -----------------------------

def audit_43_843(frame, **_):
    grtw, code43 = contains(frame["Work Rule"], "GRTW"), contains(frame["Pay Code Name"], "43")
    ids = frame["EE ID"].map(employee_key)
    matched_ids = set(ids[grtw]) & set(ids[code43])
    evidence = ids.isin(matched_ids) & (grtw | code43)
    result = frame.loc[evidence].copy()
    result.insert(0, "Audit Finding", "")
    result.loc[grtw[evidence].values, "Audit Finding"] = "GRTW work rule evidence"
    for idx in result.index[code43[evidence].values]:
        result.at[idx, "Audit Finding"] = " | ".join(filter(None, [result.at[idx, "Audit Finding"], "Pay code containing 43 evidence"]))
    return sort_if_present(result, ("EE ID", "Date", "Start Date Time")).reset_index(drop=True)


def audit_grtw_mismatch(frame, **_):
    wr, lc = contains(frame["Work Rule"], "GRTW"), contains(frame["Labour Category"], "GRTW")
    mismatch = wr ^ lc
    result = frame.loc[mismatch].copy()
    finding = pd.Series(index=frame.index, dtype=object)
    finding.loc[wr & ~lc] = "GRTW Work Rule, but Labour Category is not GRTW"
    finding.loc[lc & ~wr] = "GRTW Labour Category, but Work Rule is not GRTW"
    result.insert(0, "Audit Finding", finding.loc[mismatch].values)
    return sort_if_present(result, ("Employee ID", "Start Date Time")).reset_index(drop=True)


def audit_ot_swapped_lct(frame, **_):
    wr = frame["Work Rule"].fillna("").astype(str)
    work_rule_match = wr.str.contains("OT", case=False, regex=False) | wr.str.contains("Day Off", case=False, regex=False)
    lct_match = contains(frame["Labour Category"], "5.2")
    result = frame.loc[work_rule_match & lct_match].copy()
    result.insert(0, "Audit Finding", "OT/Day Off Work Rule with Labour Category containing 5.2")
    return sort_if_present(result, ("Employee ID", "Start Date Time")).reset_index(drop=True)


def audit_unprocessed_tors(frame, cutoff_date=None, **_):
    cutoff_date = cutoff_date or date.today() - timedelta(days=3)
    created = pd.to_datetime(frame["Creation Date"].map(clean_text), errors="coerce", dayfirst=True, format="mixed")
    matched = frame["Status"].fillna("").astype(str).str.strip().str.casefold().isin({"pending", "submitted"}) & (created.dt.date < cutoff_date)
    result = frame.loc[matched].copy()
    result.insert(0, "Audit Finding", f"Unprocessed TOR created before {cutoff_date.strftime('%d %b %Y')}")
    if not result.empty:
        result.insert(1, "Days Since Creation", (pd.Timestamp(date.today()) - created.loc[matched].dt.normalize()).dt.days.values)
        result["_sort"] = created.loc[matched].values
        result = result.sort_values("_sort", kind="stable").drop(columns="_sort")
    return result.reset_index(drop=True)


def audit_casuals_loa(frame, **_):
    casual = frame["Worker Type"].fillna("").astype(str).str.contains(r"\bCasual\b", case=False, regex=True)
    paycode = frame["Pay Code Name"].fillna("").astype(str).str.contains(r"(?<!\d)1\d{2}\.1(?!\d)", regex=True)
    result = frame.loc[casual & paycode].copy()
    result.insert(0, "Audit Finding", "Casual worker with LOA 1##.1 pay code")
    return sort_if_present(result, ("EE ID", "Pay Code Name")).reset_index(drop=True)


def audit_negative_accrual(frame, **_):
    frame = frame.copy().map(lambda v: v.replace("=", "") if isinstance(v, str) else v)
    minutes = frame["Available Balance"].map(time_to_minutes)
    matched = frame["Accrual Code"].fillna("").astype(str).str.strip().str.casefold().ne("sick") & minutes.fillna(0).lt(0)
    result = frame.loc[matched].copy()
    result.insert(0, "Audit Finding", "Non-Sick accrual code with negative available balance")
    result.insert(1, "Negative Balance (Minutes)", minutes.loc[matched].round().astype(int).values)
    if not result.empty:
        result["_sort"] = minutes.loc[matched].values
        result = result.sort_values(["_sort"] + [c for c in ("EE ID", "Accrual Code") if c in result], kind="stable").drop(columns="_sort")
    return result.reset_index(drop=True)


def audit_nba_oc_72(frame, **_):
    """Evaluate the two NBA codes together by employee and include both as evidence."""
    work = frame.copy()
    work["_employee"] = work["EE ID"].map(employee_key)
    work["_minutes"] = work["Actual Total Hrs (Include Corrections)"].map(time_to_minutes)
    code = work["Pay Code Name"].fillna("").astype(str).str.strip().str.casefold()
    work["_is_521"] = code.eq("521-nba oc less 72")
    work["_is_522"] = code.eq("522-nba oc grtr 72")
    target = work["_is_521"] | work["_is_522"]
    valid = work[target & work["_minutes"].notna()].copy()
    if valid.empty:
        result = frame.iloc[0:0].copy()
        result.insert(0, "Audit Finding", pd.Series(dtype=object))
        result.insert(1, "Employee 521 Total Hours", pd.Series(dtype=float))
        result.insert(2, "Employee 522 Total Hours", pd.Series(dtype=float))
        return result

    totals_521 = valid.loc[valid["_is_521"]].groupby("_employee")["_minutes"].sum()
    totals_522 = valid.loc[valid["_is_522"]].groupby("_employee")["_minutes"].sum()
    employees = set(valid["_employee"])
    findings = {}
    for employee in employees:
        less_minutes = float(totals_521.get(employee, 0.0))
        greater_minutes = float(totals_522.get(employee, 0.0))
        issues = []
        if less_minutes > 72 * 60:
            issues.append("521-NBA OC Less 72 total exceeds 72 hours")
        if greater_minutes > 0 and less_minutes < 72 * 60:
            issues.append("522-NBA OC Grtr 72 used before 72 hours were recorded to 521")
        if issues:
            findings[employee] = " | ".join(issues)

    evidence = valid[valid["_employee"].isin(findings)].copy()
    if evidence.empty:
        result = frame.iloc[0:0].copy()
        result.insert(0, "Audit Finding", pd.Series(dtype=object))
        result.insert(1, "Employee 521 Total Hours", pd.Series(dtype=float))
        result.insert(2, "Employee 522 Total Hours", pd.Series(dtype=float))
        return result

    result = frame.loc[evidence.index].copy()
    result.insert(0, "Audit Finding", evidence["_employee"].map(findings).values)
    result.insert(1, "Employee 521 Total Hours", evidence["_employee"].map(lambda e: round(float(totals_521.get(e, 0)) / 60, 2)).values)
    result.insert(2, "Employee 522 Total Hours", evidence["_employee"].map(lambda e: round(float(totals_522.get(e, 0)) / 60, 2)).values)
    return sort_if_present(result, ("EE ID", "Pay Code Name")).reset_index(drop=True)


# ------------------------------- Audit setup -------------------------------

AUDITS = {
    "43_843": AuditDefinition("43_843", "43 - 843 coding while on GRTW", "schedule_full_cpp", "Schedule Summary - Full", "https://fraserhealthaut.prd.mykronos.com/ia#/dataView/658", "All home locations · Current pay period", ("Work Rule", "EE ID", "Pay Code Name", "Date", "Start Date Time"), ("Work Rule", "EE ID", "Pay Code Name"), "EE ID", "Action Required - GRTW Stat Holiday Audit", "<p>Hello,</p><p>We identified employees on a <strong>Graduated Return to Work (GRTW)</strong> plan who have been coded with <strong>Statutory Holiday (SH)</strong> paycodes.</p><p>Please review and correct the entries and reply with your findings.</p>", ("Open Schedule Summary - Full.", "Select all home locations and the current pay period.", "Identify EE IDs with GRTW Work Rule evidence and separate 43 pay code evidence.", "Keep evidence rows for employees appearing in both groups."), audit_43_843),
    "grtw_mismatch": AuditDefinition("grtw_mismatch", "GRTW Shifts without a valid GRTW LCT or WR", "schedule_shift_wr_cpp", "Schedule Summary - Shift with Work Rules", "https://fraserhealthaut.prd.mykronos.com/ia#/dataView/7913", "All home locations · Current pay period", ("Employee ID", "Work Rule", "Labour Category", "Start Date Time"), ("Employee ID", "Work Rule", "Labour Category"), "Employee ID", "Action Required - GRTW Shifts without a valid GRTW LCT or WR", "<p>Hello,</p><p>We identified GRTW shifts missing a valid GRTW Work Rule and/or Labor Category Transfer.</p><p>Please review, update the shifts, and reply with your findings.</p>", ("Open Schedule Summary - Shift with Work Rules.", "Select all home locations and the current pay period.", "Keep rows where only one of Work Rule or Labour Category contains GRTW."), audit_grtw_mismatch),
    "ot_swapped_lct": AuditDefinition("ot_swapped_lct", "OT Work rule and swapped in LCT", "schedule_shift_wr_cpp", "Schedule Summary - Shift with Work Rules", "https://fraserhealthaut.prd.mykronos.com/ia#/dataView/7913", "All home locations · Current pay period", ("Employee ID", "Work Rule", "Labour Category", "Start Date Time", "End Date Time"), ("Employee ID", "Work Rule", "Labour Category"), "Employee ID", None, None, ("Open Schedule Summary - Shift with Work Rules.", "Select all home locations and the current pay period.", "Keep rows where Work Rule contains OT or Day Off.", "From those rows, keep Labour Category values containing 5.2."), audit_ot_swapped_lct),
    "unprocessed_tors": AuditDefinition("unprocessed_tors", "Unprocessed TORs", "tor_details_cpp", "Scheduling - Time Off Requests - Details", "https://fraserhealthaut.prd.mykronos.com/ia#/dataView/661", "All home locations · Current pay period", ("EE ID", "Status", "Creation Date", "Time-off Request ID"), ("EE ID", "Status", "Creation Date"), "EE ID", "Action Required - Unprocessed TORs", "<p>Hello,</p><p>We identified <strong>Time Off Requests</strong> remaining in <strong>Submitted/Pending</strong> status beyond the selected cutoff.</p><p>Please review and action the requests, then reply with your findings.</p>", ("Open Scheduling - Time Off Requests - Details.", "Select all home locations and the current pay period.", "Choose the cutoff date.", "Keep Pending or Submitted requests created before the cutoff."), audit_unprocessed_tors),
    "casuals_loa": AuditDefinition("casuals_loa", "Casuals with LOA 1##.1 paycodes", "time_summary_current_pay_period", "Time Summary by Pay Code", "https://fraserhealthaut.prd.mykronos.com/ia#/dataView/3358", "All home locations · Current pay period", ("EE ID", "Pay Code Name", "Worker Type"), ("EE ID", "Pay Code Name", "Worker Type"), "EE ID", "Action Required - Casual EE's with LOA 1##.1 Paycodes", "<p>Hello,</p><p>We identified casual employee leave coded to a <strong>LOA 1##.1 paycode</strong>.</p><p>Please update these entries to the appropriate .2 paycode and reply with your findings.</p>", ("Open Time Summary by Pay Code.", "Select all home locations and the current pay period.", "Filter Worker Type to Casual.", "Keep Pay Code Name values matching 1##.1."), audit_casuals_loa),
    "negative_accrual": AuditDefinition("negative_accrual", "Negative Hours Accrual Bank", "accrual_negative_cpp", "Accrual Details - Negative Balance", "https://fraserhealthaut.prd.mykronos.com/ia#/dataView/4306", "All home locations · Current pay period", ("EE ID", "Available Balance", "Accrual Code"), ("EE ID", "Available Balance", "Accrual Code"), "EE ID", "Action Required - Negative Accrual Bank Audit", "<p>Hello,</p><p>Please find attached the audit identifying employees with a negative accrual bank balance.</p><p>Please review the flagged entries, correct or adjust entries where needed, and reply with your findings.</p><p>For negative statutory holiday balances, please review the employee’s statutory holiday coding for the year.</p><p>Please reply with your findings once the review is complete.</p>", ("Open Accrual Details - Negative Balance.", "Select all home locations and the current pay period.", "Remove equals signs from exported balances.", "Exclude Sick accrual codes.", "Keep negative balances only."), audit_negative_accrual),
    "nba_oc_72": AuditDefinition("nba_oc_72", "NBA OC > 72 hours", "time_summary_calendar_month", "Time Summary by Pay Code", "https://fraserhealthaut.prd.mykronos.com/ia#/dataView/3358", "All home locations · Calendar month", ("EE ID", "Actual Total Hrs (Include Corrections)", "Pay Code Name"), ("EE ID", "Actual Total Hrs (Include Corrections)", "Pay Code Name"), "EE ID", "Action Required - NBA On-Call Hours Coding", "<p>Hello,</p><p>We identified NBA on-call hours that may be coded to the incorrect pay code for the audited calendar month.</p><p>The first 72 eligible hours must be coded to <strong>521-NBA OC Less 72</strong>. Only hours exceeding 72 must be coded to <strong>522-NBA OC Grtr 72</strong>.</p><p>Please review and correct the flagged entries, then reply with your findings.</p>", ("Open Time Summary by Pay Code.", "Select all home locations and the complete calendar month.", "Review 521 and 522 totals together by employee.", "Flag 521 totals above 72 hours.", "Flag 522 hours when fewer than 72 hours were recorded to 521.", "Include both 521 and 522 rows for every flagged employee as reference."), audit_nba_oc_72),    "leave_cascade": AuditDefinition("leave_cascade", "CTO, OTO, and STO Leave Cascading to LOA", "schedule_by_paycode_cpp", "Schedule Summary - By Paycode", "https://fraserhealthaut.prd.mykronos.com/ia#/dataView/4808", "All home locations · Current pay period", ("EE ID", "Last Name", "First Name", "Start Date", "Duration (Hrs)", "Pay Code Name"), ("EE ID", "Start Date", "Duration (Hrs)", "Pay Code Name"), "EE ID", "Action Required - Leave Adjustment", "", ("Open Schedule Summary - By Paycode and Time Summary by Pay Code.", "Use all home locations and the current pay period for both.", "Keep employees with CTO, OTO, or STO Leave and at least one LOA timecard paycode.", "Compare combined leave hours by type against the corresponding NCO code.", "Allocate matching hours oldest-first, then distribute estimated cascaded LOA chronologically after removing scheduled LOA hours."), lambda frame, **_: frame),

}


def safe_filename(name):
    return re.sub(r"\s+", " ", re.sub(r'[\\/:*?"<>|]+', "-", clean_text(name))).strip(" .") or "Audit"


def excel_bytes(frame, audit_name, source_name):
    safe_frame = frame.drop(columns=[c for c in frame.columns if str(c).startswith("_")], errors="ignore").copy()
    summary_messages = list(frame.attrs.get("summary_messages", []))
    for column in safe_frame.columns:
        safe_frame[column] = safe_frame[column].map(lambda v: "'" + v if isinstance(v, str) and v.startswith("=") else v)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        safe_frame.to_excel(writer, sheet_name="Audit Results", index=False, startrow=4)
        ws = writer.book["Audit Results"]
        ws.sheet_view.showGridLines = False
        ws.freeze_panes = "A6"
        ws["A1"], ws["A2"], ws["A3"] = audit_name, f"Processed: {date.today():{DATE_FMT}}", f"Source: {source_name} | Findings: {len(frame):,}"
        ws["A1"].font, ws["A1"].fill = Font(size=16, bold=True, color="FFFFFF"), PatternFill("solid", fgColor="17365D")
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(1, len(frame.columns)))
        for cell in ws[5]:
            cell.fill, cell.font = PatternFill("solid", fgColor="1F4E78"), Font(color="FFFFFF", bold=True)
        if len(frame):
            ref = f"A5:{get_column_letter(len(frame.columns))}{5 + len(frame)}"
            table = Table(displayName="AuditResultsTable", ref=ref)
            table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
            ws.add_table(table)
            if "Audit Finding" in frame.columns:
                col = list(frame.columns).index("Audit Finding") + 1
                for row in range(6, 6 + len(frame)):
                    finding = clean_text(ws.cell(row, col).value)
                    if finding == "Hours Match":
                        ws.cell(row, col).fill, ws.cell(row, col).font = PatternFill("solid", fgColor="E2F0D9"), Font(color="375623")
                    else:
                        ws.cell(row, col).fill, ws.cell(row, col).font = PatternFill("solid", fgColor="FCE4D6"), Font(color="9C0006")
        elif len(frame.columns):
            ws["A6"] = "No findings matched the audit rules."
        thin = Side(style="thin", color="D9E1F2")
        for col_idx, name in enumerate(frame.columns, 1):
            values = [str(name)] + [clean_text(v) for v in frame.iloc[:, col_idx - 1].head(250)]
            width = min(max(max(map(len, values)) + 2, 12), 42)
            ws.column_dimensions[get_column_letter(col_idx)].width = width
            for row in range(6, 6 + len(frame)):
                ws.cell(row, col_idx).alignment = Alignment(vertical="top", wrap_text=width >= 30)
                ws.cell(row, col_idx).border = Border(bottom=thin)
        ws.auto_filter.ref = None
        summary = writer.book.create_sheet("Audit Summary")
        summary.sheet_view.showGridLines = False
        summary["A1"] = "Audit Summary and Decisions"
        summary["A1"].font = Font(size=15, bold=True, color="FFFFFF")
        summary["A1"].fill = PatternFill("solid", fgColor="17365D")
        summary.column_dimensions["A"].width = 120
        if not summary_messages:
            summary_messages = ["No fallbacks, ambiguities, or missing values were recorded."]
        for row_number, message in enumerate(summary_messages, 3):
            summary.cell(row_number, 1).value = message
            summary.cell(row_number, 1).alignment = Alignment(wrap_text=True, vertical="top")
    return output.getvalue()


# --------------------------- PC Portfolio / email ---------------------------

PORTFOLIO_COLUMNS = ("EmployeeNumber", "Portfolio Description", "Portfolio Email", "Coordinator Name")


def read_portfolio(uploaded_file):
    table = require_columns(locate_table(read_raw(uploaded_file), PORTFOLIO_COLUMNS, PORTFOLIO_COLUMNS), PORTFOLIO_COLUMNS)
    table["_key"] = table["EmployeeNumber"].map(employee_key)
    table = table[table["_key"] != ""]
    def first_nonblank(values):
        return next((v for v in values if clean_text(v)), "")
    return table.groupby("_key", as_index=False)[["Portfolio Description", "Portfolio Email", "Coordinator Name"]].agg(first_nonblank)


def add_portfolio(frame, id_column, portfolio):
    result = frame.drop(columns=[c for c in ("Portfolio Description", "Portfolio Email", "Coordinator Name") if c in frame], errors="ignore").copy()
    result["_key"] = result[id_column].map(employee_key)
    return result.merge(portfolio, how="left", on="_key", sort=False).drop(columns="_key")


def process_audit(audit, uploaded_file, schedule_audit=None, supervisor=None, audit_options=None, portfolio_match=None):
    table = require_columns(locate_table(read_raw(uploaded_file), audit.expected_columns, audit.required_columns), audit.required_columns)
    result = audit.processor(table, **(audit_options or {}))
    if schedule_audit is not None and supervisor is not None:
        match = portfolio_match or {}
        result = add_location_portfolio(
            result, match.get("id_column", audit.employee_id_column),
            schedule_audit, supervisor, match.get("date_column"),
            match.get("allow_employee_fallback", True))
    filename = f"[SUS] Audit - {safe_filename(audit.name)} {date.today():{DATE_FMT}}.xlsx"
    return result, excel_bytes(result, audit.name, audit.source_name), filename


def process_portfolio_only(uploaded_file, schedule_audit, supervisor, id_column,
                           date_column, allow_employee_fallback, output_name):
    frame = read_existing_table(uploaded_file)
    if id_column not in frame.columns:
        raise ValueError(f"The selected match column '{id_column}' was not found.")
    result = add_location_portfolio(frame, id_column, schedule_audit, supervisor,
                                    date_column, allow_employee_fallback)
    title = clean_text(output_name) or "Portfolio Enriched Audit"
    return result, excel_bytes(result, title, uploaded_file.name), f"[SUS] Audit - {safe_filename(title)} {date.today():{DATE_FMT}}.xlsx"


def extract_emails(value) -> list[str]:
    """Extract real email patterns, ignoring labels and handling space-separated addresses."""
    return [match.group(0) for match in EMAIL_RE.finditer(clean_text(value))]


def unique_casefold(values):
    seen, output = set(), []
    for value in values:
        text = clean_text(value)
        key = text.casefold()
        if text and key not in seen:
            seen.add(key)
            output.append(text)
    return output


def coordinator_name(value) -> str:
    text = clean_text(value)
    return text.split("-", 1)[1].strip() if "-" in text else text


def coordinator_recipient(value) -> str | None:
    name = coordinator_name(value)
    key = name.casefold()
    if not name or key in {"vcy", "sustainment team"}:
        return None
    return COORDINATOR_EMAILS.get(key, name)


def email_details(frame):
    if "Portfolio Email" not in frame or "Coordinator Name" not in frame:
        raise ValueError("The audit result is missing Portfolio Email or Coordinator Name.")

    all_emails = []
    eligible_by_row = []
    for value in frame["Portfolio Email"]:
        parsed = [e for e in extract_emails(value) if e.casefold() != HELP_EMAIL.casefold()]
        eligible_by_row.append(bool(parsed))
        all_emails.extend(parsed)
    to_values = unique_casefold(all_emails)

    cc_values = [HELP_EMAIL]
    cc_values.extend(recipient for recipient in frame["Coordinator Name"].map(coordinator_recipient) if recipient)
    cc_values = unique_casefold(cc_values)

    coordinator_clean = frame["Coordinator Name"].map(coordinator_name).str.casefold()
    coordinator_not_excluded = frame["Coordinator Name"].map(clean_text).ne("") & coordinator_clean.ne("sustainment team")
    mask = pd.Series(eligible_by_row, index=frame.index) & coordinator_not_excluded
    return to_values, cc_values, frame.loc[mask].copy().reset_index(drop=True)


def email_body_for(audit, attachment_frame):
    body = audit.email_body_html or ""
    if audit.key == "negative_accrual" and attachment_frame["Accrual Code"].fillna("").astype(str).str.contains(r"\d", regex=True).any():
        body += "<p>The audit also includes non-cascading pay codes that require review. Please review the coding and complete any required corrections.</p>"
    return body


def create_outlook_draft(audit, frame):
    to_values, cc_values, attachment_frame = email_details(frame)
    if not to_values:
        raise ValueError("No eligible Portfolio Email recipients were found.")
    if attachment_frame.empty:
        raise ValueError("No rows remain for the email attachment after exclusions.")
    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise RuntimeError("Outlook draft creation requires Windows, classic Outlook, and pywin32.") from exc
    path = Path(tempfile.gettempdir()) / f"Audit - {safe_filename(audit.name)} {date.today():{DATE_FMT}}.xlsx"
    path.write_bytes(excel_bytes(attachment_frame, audit.name, audit.source_name))
    pythoncom.CoInitialize()
    try:
        mail = win32com.client.Dispatch("Outlook.Application").CreateItem(0)
        mail.Display(False)
        time.sleep(.35)
        existing = mail.HTMLBody or "<html><body></body></html>"
        body_match = re.search(r"<body[^>]*>", existing, re.I)
        styled = '<div style="font-family:Arial,sans-serif;font-size:11pt;line-height:1.25;">' + email_body_for(audit, attachment_frame) + '</div><div style="height:11pt;line-height:11pt;">&nbsp;</div>'
        mail.HTMLBody = existing[:body_match.end()] + styled + existing[body_match.end():] if body_match else styled + existing
        mail.SentOnBehalfOfName = HELP_EMAIL
        mail.To = "; ".join(to_values)
        mail.CC = "; ".join(cc_values)
        mail.Subject = audit.email_subject
        mail.Attachments.Add(str(path.resolve()))
        mail.Save()
        mail.GetInspector.Activate()
    finally:
        pythoncom.CoUninitialize()
    return len(attachment_frame)



def create_leave_employee_draft(audit, selected_rows, employee_email):
    leave_types = unique_casefold(selected_rows["Leave Type"])
    subject_type = leave_types[0] if len(leave_types) == 1 else "Leave"
    bullets = []
    for _, row in selected_rows.iterrows():
        bullets.append(f"<li>{row['Leave Type']} Leave on {clean_text(row['Start Date'])}: approximately {natural_duration(row.get('_estimated_minutes', 0))} cascaded to LOA.</li>")
    employee_name = " ".join(filter(None, [clean_text(selected_rows.iloc[0].get("First Name")), clean_text(selected_rows.iloc[0].get("Last Name"))]))
    body = f"<p>Hello {employee_name},</p><p>During a recent audit, we noticed the following leave occurrence(s) may have cascaded to a Leave of Absence (LOA):</p><ul>{''.join(bullets)}</ul><p>If you would prefer your leave not to cascade to LOA in the future, please ensure that you have enough corresponding leave hours banked before submitting the request.</p><p>Please note that you cannot bank these hours in the same pay period in which you are using the leave, as the hours are not yet available at the time the leave is processed.</p><p>If you have any questions, please feel free to reach out.</p>"
    portfolio_emails = unique_casefold(
        email for value in selected_rows.get("Portfolio Email", pd.Series(dtype=object)) for email in extract_emails(value)
        if email.casefold() != HELP_EMAIL.casefold()
    )
    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise RuntimeError("Outlook draft creation requires Windows, classic Outlook, and pywin32.") from exc
    path = Path(tempfile.gettempdir()) / f"Audit - {safe_filename(audit.name)} selected rows {date.today():{DATE_FMT}}.xlsx"
    path.write_bytes(excel_bytes(selected_rows, audit.name, audit.source_name))
    pythoncom.CoInitialize()
    try:
        mail = win32com.client.Dispatch("Outlook.Application").CreateItem(0)
        mail.Display(False); time.sleep(.35)
        existing = mail.HTMLBody or "<html><body></body></html>"
        match = re.search(r"<body[^>]*>", existing, re.I)
        styled = '<div style="font-family:Arial,sans-serif;font-size:11pt;line-height:1.25;">' + body + '</div><div style="height:11pt;line-height:11pt;">&nbsp;</div>'
        mail.HTMLBody = existing[:match.end()] + styled + existing[match.end():] if match else styled + existing
        mail.SentOnBehalfOfName = HELP_EMAIL
        mail.To = clean_text(employee_email)
        mail.CC = "; ".join(unique_casefold([HELP_EMAIL] + portfolio_emails))
        mail.Subject = f"Action Required - {subject_type} Leave Adjustment"
        mail.Attachments.Add(str(path.resolve())); mail.Save(); mail.GetInspector.Activate()
    finally:
        pythoncom.CoUninitialize()
    return len(selected_rows)


# ----------------------------------- UI -----------------------------------

def selected_audit_keys():
    return [key for key in AUDITS if st.session_state.get(f"select_{key}", False)]


def source_groups_for(selected):
    groups = {}
    for key in selected:
        groups.setdefault(AUDITS[key].source_key, []).append(key)
        if key == "leave_cascade":
            groups.setdefault("time_summary_current_pay_period", []).append(key)
    return groups


def progress_nav(active_stage):
    cols = st.columns(3, gap=None)
    for number, (col, label) in enumerate(zip(cols, ("1. Select Audits", "2. Upload & Configure", "3. Review & Run")), 1):
        state = "active" if number == active_stage else "done" if number < active_stage else "upcoming"
        col.markdown(f'<div class="step-pill {state}">{label}</div>', unsafe_allow_html=True)


def audit_selector():
    top_left, top_right = st.columns([4, 2], vertical_alignment="center")
    top_left.markdown("### Choose audits to run")
    with top_right:
        c1, c2 = st.columns(2)
        if c1.button("Select all", use_container_width=True):
            for key in AUDITS:
                st.session_state[f"select_{key}"] = True
            st.session_state["select_portfolio_only"] = False
            st.rerun()
        if c2.button("Clear", use_container_width=True):
            for key in AUDITS:
                st.session_state[f"select_{key}"] = False
            st.session_state["select_portfolio_only"] = False
            st.rerun()
    cols = st.columns(2, gap="medium")
    for index, (key, audit) in enumerate(AUDITS.items()):
        with cols[index % 2].container(border=True, height=124, key=f"audit_card_{key}"):
            st.checkbox(audit.name, key=f"select_{key}")
            st.caption(f"{audit.source_name} · {audit.export_settings}")
    st.markdown("#### Optional tool")
    with st.container(border=True):
        st.checkbox("PC Portfolio Enrichment Only", key="select_portfolio_only")
        st.caption("Add portfolio details to an already-filtered table without applying audit rules.")
    selected = selected_audit_keys()
    tool = bool(st.session_state.get("select_portfolio_only", False))
    return selected, tool


def upload_tile(source_key, audit_keys):
    source = AUDITS[audit_keys[0]]
    if source_key == "time_summary_current_pay_period" and source.source_key != source_key:
        source = AUDITS["casuals_loa"]
    with st.container(border=True, key=f"source_card_{source_key}"):
        st.markdown(f"#### {source.source_name}")
        st.caption(source.export_settings)
        st.markdown(f"[Open export ↗]({source.source_url})")
        st.markdown("**Used by:** " + " · ".join(AUDITS[key].name for key in audit_keys))
        uploaded = st.file_uploader(f"Upload {source.source_name}", type=["csv", "xlsx", "xlsm", "xls"], key=f"upload_{source_key}", label_visibility="collapsed")
        try:
            saved = cache_upload(source_key, uploaded)
        except ValueError as exc:
            st.error(str(exc))
            saved = None
        if saved:
            st.success(f"Ready · {saved.name}")
        return saved


def render_results(results):
    if not results:
        st.info("Run the selected audits to see results here.")
        return
    for audit, result, payload, filename, error in results:
        with st.container(border=True, key=f"result_card_{audit.key}"):
            st.markdown(f"### {audit.name}")
            if error:
                st.error(error)
                continue
            if len(result):
                st.warning(f"{len(result):,} row(s) in the result.")
                st.dataframe(result.drop(columns=[c for c in result.columns if str(c).startswith("_")], errors="ignore"), use_container_width=True, hide_index=True, height=330)
                if audit.key == "leave_cascade":
                    with st.expander("Create employee-specific Outlook draft"):
                        ee = st.text_input("Enter EE ID", key="leave_email_ee")
                        matches = result[result["EE ID"].map(employee_key) == employee_key(ee)].copy() if ee else result.iloc[0:0].copy()
                        if ee and matches.empty:
                            st.warning("No rows found for that EE ID.")
                        elif not matches.empty:
                            labels = {idx: f"{clean_text(row['Start Date'])} · {row['Leave Type']} · {row['Estimated Cascaded to LOA']} estimated LOA" for idx, row in matches.iterrows()}
                            chosen = st.multiselect("Select leave occurrences", list(labels), format_func=lambda idx: labels[idx], key="leave_email_rows")
                            details = st.session_state.get("employee_details_data")
                            employee_email = ""
                            if details is not None:
                                found = details[details["_key"] == employee_key(ee)]
                                if not found.empty: employee_email = clean_text(found.iloc[0]["Email"])
                            st.caption(f"Employee email: {employee_email or 'Blank - complete manually in Outlook'}")
                            if st.button("Create employee Outlook draft", disabled=not chosen, key="leave_email_create"):
                                try: st.success(f"Draft created with {create_leave_employee_draft(audit, result.loc[chosen], employee_email)} selected row(s).")
                                except Exception as exc: st.error(str(exc))
            else:
                st.success("No findings matched the audit rules.")
            st.download_button("Download full Excel audit file", payload, filename, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=f"download_{audit.key}", use_container_width=True)
            if all(c in result.columns for c in ("Portfolio Email", "Coordinator Name")) and audit.email_subject:
                if st.button("Create and open Outlook email draft", key=f"email_{audit.key}", use_container_width=True):
                    try:
                        st.success(f"Outlook draft created with {create_outlook_draft(audit, result):,} attachment row(s).")
                    except Exception as exc:
                        st.error(str(exc))


def app():
    st.set_page_config(page_title=APP_TITLE, page_icon="⚡", layout="wide")
    st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Manrope:wght@400;500;600;700;800&display=swap');
    html,body,.stApp,button,input,label,p,div{font-family:'Manrope','Segoe UI',Arial,sans-serif}
    .stApp{background:linear-gradient(135deg,#FBF8F3 0%,#F7FAFC 45%,#F5F8FC 100%)}
    .block-container{max-width:1180px;padding-top:1.2rem;padding-bottom:3rem}
    .hero{padding:.95rem 1.15rem;margin-bottom:.7rem;border:1px solid #D5E1EA;border-radius:14px;background:linear-gradient(110deg,#FFF8EC,#EDF6FC);box-shadow:0 5px 18px rgba(25,60,86,.07)}
    .hero-title{font-size:1.72rem;font-weight:800;color:#1E3040}.hero-sub{font-size:.82rem;color:#586B79}
    .step-pill{padding:.55rem;text-align:center;font-size:.82rem;font-weight:700;border:1px solid #D5DEE6;background:#F1F4F7;color:#667683}
    .step-pill.active{background:#DCEBF6;color:#194762}.step-pill.done{background:#E8F4EE;color:#2E6650}
    .section-head{margin:1rem 0 .7rem;padding:.8rem 1rem;border:1px solid #D8E2E9;border-radius:11px}.blue{background:#EAF4FB}.mint{background:#EAF7F1}.peach{background:#FFF2E7}
    div[data-testid="stVerticalBlockBorderWrapper"]{border-color:#D7E1E8!important;border-radius:12px!important;background:rgba(255,255,255,.88)}
    div[data-testid="stFileUploaderDropzone"]{padding:.55rem .7rem!important;min-height:0!important;background:#F3F7FA}
    div[data-testid="stFileUploaderDropzone"] small{display:none}.stButton button{border-radius:9px;font-weight:700}
    .summary-bar{padding:.7rem;border:1px solid #D9E3EA;border-radius:10px;background:white;color:#516472}
    </style>
    """, unsafe_allow_html=True)
    taglines = ["Upload once, audit systematically.", "Less sorting. More reviewing.", "One guided path from export to action."]
    st.markdown(f'<div class="hero"><div class="hero-title">⚡ Audit Prepper</div><div class="hero-sub">{random.choice(taglines)}</div></div>', unsafe_allow_html=True)
    stage = st.session_state.get("workflow_stage", 1)
    progress_nav(stage)

    if stage == 1:
        st.markdown('<div class="section-head blue"><strong>STEP 1 · Select audits</strong><br><small>Choose only the audits or optional tool needed for this run.</small></div>', unsafe_allow_html=True)
        selected, tool_selected = audit_selector()
        count = len(selected) + int(tool_selected)
        left, right = st.columns([4, 1])
        left.markdown(f'<div class="summary-bar">{count} item(s) selected</div>', unsafe_allow_html=True)
        if right.button("Continue to uploads →", type="primary", disabled=not count, use_container_width=True):
            selected_signature = list(selected) + (["__portfolio_only__"] if tool_selected else [])
            if st.session_state.get("chosen_signature") != selected_signature:
                clear_workflow_cache()
            st.session_state["chosen_audits"] = list(selected)
            st.session_state["chosen_portfolio_only"] = tool_selected
            st.session_state["chosen_signature"] = selected_signature
            st.session_state["workflow_stage"] = 2
            st.rerun()
        return

    selected = [k for k in st.session_state.get("chosen_audits", []) if k in AUDITS]
    tool_selected = bool(st.session_state.get("chosen_portfolio_only", False))
    if not selected and not tool_selected:
        st.session_state["workflow_stage"] = 1
        st.rerun()
    groups = source_groups_for(selected)
    options = {key: {} for key in selected}
    portfolio_choices = {}

    if stage == 2:
        st.markdown('<div class="section-head mint"><strong>STEP 2 · Upload and configure</strong><br><small>Upload files once and review audit options.</small></div>', unsafe_allow_html=True)
        upload_cols = st.columns(2, gap="medium")
        group_items = list(groups.items())
        for i, (source_key, audit_keys) in enumerate(group_items):
            with upload_cols[i % 2]:
                upload_tile(source_key, audit_keys)
        if tool_selected:
            with upload_cols[len(group_items) % 2].container(border=True):
                st.markdown("#### Existing filtered audit")
                generic_widget = st.file_uploader("Upload existing filtered audit", type=["csv", "xlsx", "xlsm", "xls"], key="upload_portfolio_only", label_visibility="collapsed")
                try:
                    generic_saved = cache_upload("portfolio_only_source", generic_widget)
                except ValueError as exc:
                    st.error(str(exc)); generic_saved = None
                if generic_saved:
                    st.success(f"Ready · {generic_saved.name}")

        st.markdown("#### PC Portfolio and audit options")
        option_cols = st.columns(2, gap="medium")
        for i, key in enumerate(selected):
            audit = AUDITS[key]
            with option_cols[i % 2].container(border=True):
                st.markdown(f"#### {audit.name}")
                portfolio_choices[key] = st.checkbox("Add PC Portfolio details", value=True, key=f"portfolio_{key}")
                if portfolio_choices[key] and cached_upload(audit.source_key):
                    try:
                        source_table = read_existing_table(cached_upload(audit.source_key))
                        columns = list(source_table.columns)
                        detected_id = best_column_match(columns, EMPLOYEE_ID_ALIASES) or audit.employee_id_column
                        id_index = columns.index(detected_id) if detected_id in columns else None
                        options[key]["portfolio_id_column"] = st.selectbox(
                            "Employee identifier column for portfolio matching", columns,
                            index=id_index, placeholder="Choose employee identifier",
                            key=f"portfolio_id_{key}")
                        date_columns = available_date_columns(columns)
                        detected_date = best_column_match(columns, DATE_COLUMN_ALIASES)
                        date_options = [NO_DATE_OPTION] + date_columns
                        date_index = date_options.index(detected_date) if detected_date in date_options else 0
                        chosen_date = st.selectbox(
                            "Shift/date column for portfolio matching", date_options,
                            index=date_index, key=f"portfolio_date_{key}",
                            help="The best match is selected automatically. Confirm it, choose another column, or choose No date column.")
                        options[key]["portfolio_date_column"] = None if chosen_date == NO_DATE_OPTION else chosen_date
                        options[key]["portfolio_employee_fallback"] = st.checkbox(
                            "Allow EE ID-only fallback", value=True,
                            key=f"portfolio_ee_fallback_{key}")
                    except Exception as exc:
                        st.error(f"Could not inspect source columns: {exc}")
                if key == "unprocessed_tors":
                    options[key]["cutoff_date"] = st.date_input("Flag Submitted/Pending TORs created before", date.today() - timedelta(days=3), key="tor_cutoff")
                with st.expander("Steps"):
                    for n, step in enumerate(audit.steps, 1): st.write(f"{n}. {step}")
        if tool_selected:
            generic = cached_upload("portfolio_only_source")
            with option_cols[len(selected) % 2].container(border=True):
                st.markdown("#### PC Portfolio Enrichment Only")
                st.text_input("Output audit name", value="Portfolio Enriched Audit", key="portfolio_only_name")
                if generic:
                    try:
                        preview = read_existing_table(generic)
                        columns = list(preview.columns)
                        default_column = best_column_match(columns, EMPLOYEE_ID_ALIASES)
                        default_index = columns.index(default_column) if default_column in columns else None
                        selected_id_column = st.selectbox(
                            "Employee identifier column for Schedule Audit matching", columns,
                            index=default_index, placeholder="Choose employee identifier",
                            key="portfolio_only_id_column")
                        date_columns = available_date_columns(columns)
                        detected_date = best_column_match(columns, DATE_COLUMN_ALIASES)
                        date_options = [NO_DATE_OPTION] + date_columns
                        date_index = date_options.index(detected_date) if detected_date in date_options else 0
                        st.selectbox(
                            "Shift/date column for portfolio matching", date_options,
                            index=date_index, key="portfolio_only_date_column",
                            help="The best match is selected automatically. Confirm it, choose another column, or choose No date column.")
                        st.checkbox("Allow EE ID-only fallback", value=True,
                                    key="portfolio_only_employee_fallback")
                        if default_column is None and not selected_id_column:
                            st.warning("No standard employee ID column was detected. Select the correct column before continuing.")
                        st.caption(f"{len(preview):,} row(s) detected · Parsed table is cached after the first read")
                    except Exception as exc:
                        st.error(f"Could not read the table: {exc}")

        needs_portfolio = tool_selected or any(portfolio_choices.values())
        if needs_portfolio:
            portfolio_cols = st.columns(2, gap="medium")
            with portfolio_cols[0].container(border=True):
                st.markdown("#### Audit - Schedule - by Employee")
                st.caption("All home locations · Current pay period")
                st.markdown(f"[Open export ↗]({SCHEDULE_AUDIT_URL})")
                widget = st.file_uploader("Upload Schedule Audit", type=["csv", "xlsx", "xlsm", "xls"], key="schedule_audit_upload", label_visibility="collapsed")
                try: saved = cache_upload("schedule_audit", widget)
                except ValueError as exc: st.error(str(exc)); saved = None
                if saved: st.success(f"Ready · {saved.name}")
            with portfolio_cols[1].container(border=True):
                st.markdown("#### Supervisor-Portfolio-Department Report")
                st.caption("Used to match the cleaned WFM department path")
                st.markdown(f"[Open report ↗]({SUPERVISOR_PORTFOLIO_URL})")
                widget = st.file_uploader("Upload Supervisor Portfolio report", type=["csv", "xlsx", "xlsm", "xls"], key="supervisor_portfolio_upload", label_visibility="collapsed")
                try: saved = cache_upload("supervisor_portfolio", widget)
                except ValueError as exc: st.error(str(exc)); saved = None
                if saved: st.success(f"Ready · {saved.name}")

        employee_details_enabled = False
        if "leave_cascade" in selected:
            employee_details_enabled = st.checkbox("Add Employee Details for employee-facing email drafts", value=True, key="employee_details_enabled")
            if employee_details_enabled:
                with st.container(border=True):
                    st.markdown("#### Employee Details")
                    st.caption("Optional, default enabled · Matches EE ID to EmployeeNumber for employee email")
                    st.markdown(f"[Open report ↗]({EMPLOYEE_DETAILS_URL})")
                    widget = st.file_uploader("Upload Employee Details", type=["csv", "xlsx", "xlsm", "xls"], key="employee_details_upload", label_visibility="collapsed")
                    try: saved = cache_upload("employee_details", widget)
                    except ValueError as exc: st.error(str(exc)); saved = None
                    if saved: st.success(f"Ready · {saved.name}")

        sources_ready = all(cached_upload(k) for k in groups)
        tool_ready = not tool_selected or (cached_upload("portfolio_only_source") and st.session_state.get("portfolio_only_id_column"))
        portfolio_ready = not needs_portfolio or (cached_upload("schedule_audit") and cached_upload("supervisor_portfolio"))
        employee_ready = not employee_details_enabled or cached_upload("employee_details")
        ready = bool(sources_ready and tool_ready and portfolio_ready and employee_ready)
        back, summary, next_col = st.columns([1, 3, 1.4], vertical_alignment="center")
        if back.button("← Back", use_container_width=True):
            st.session_state["workflow_stage"] = 1; st.rerun()
        summary.markdown(f'<div class="summary-bar">{len(selected) + int(tool_selected)} item(s) ready for review</div>', unsafe_allow_html=True)
        if next_col.button("Review & run →", type="primary", disabled=not ready, use_container_width=True):
            st.session_state["run_config"] = {
                "selected": selected, "tool": tool_selected,
                "portfolio_choices": portfolio_choices, "options": options,
                "tool_name": st.session_state.get("portfolio_only_name", "Portfolio Enriched Audit"),
                "tool_id": st.session_state.get("portfolio_only_id_column"),
                "tool_date": None if st.session_state.get("portfolio_only_date_column") == NO_DATE_OPTION else st.session_state.get("portfolio_only_date_column"),
                "tool_employee_fallback": st.session_state.get("portfolio_only_employee_fallback", True),
                "employee_details_enabled": employee_details_enabled,
            }
            st.session_state["workflow_stage"] = 3
            st.rerun()
        return

    config = st.session_state.get("run_config", {})
    selected = config.get("selected", selected)
    tool_selected = config.get("tool", tool_selected)
    portfolio_choices = config.get("portfolio_choices", {})
    options = config.get("options", {})
    groups = source_groups_for(selected)
    st.markdown('<div class="section-head peach"><strong>STEP 3 · Review and run</strong><br><small>Confirm the queue, process audits, and download results.</small></div>', unsafe_allow_html=True)
    with st.container(border=True):
        if selected: st.write("**Audits:** " + " · ".join(AUDITS[k].name for k in selected))
        if tool_selected: st.write("**Optional tool:** PC Portfolio Enrichment Only")
    back, run_col = st.columns([1, 2])
    if back.button("← Edit uploads", use_container_width=True):
        st.session_state["workflow_stage"] = 2; st.rerun()
    if run_col.button("Run selected items", type="primary", use_container_width=True):
        status, progress = st.status("Processing queue", expanded=True), st.progress(0)
        needs_portfolio = tool_selected or any(portfolio_choices.values())
        schedule_audit, supervisor, portfolio_error = None, None, None
        if needs_portfolio:
            try:
                if not cached_upload("schedule_audit") or not cached_upload("supervisor_portfolio"):
                    raise ValueError("Schedule Audit or Supervisor-Portfolio report is missing. Return to Step 2.")
                schedule_audit = read_schedule_audit(cached_upload("schedule_audit"))
                supervisor = read_supervisor_portfolio(cached_upload("supervisor_portfolio"))
            except Exception as exc:
                portfolio_error = str(exc)
        employee_details = None
        if config.get("employee_details_enabled") and cached_upload("employee_details"):
            try:
                employee_details = require_columns(locate_table(read_raw(cached_upload("employee_details")), ("EmployeeNumber", "Email", "Employee Name"), ("EmployeeNumber", "Email")), ("EmployeeNumber", "Email"))
                employee_details["_key"] = employee_details["EmployeeNumber"].map(employee_key)
            except Exception as exc:
                status.write(f"⚠ Employee Details could not be read: {exc}")
        st.session_state["employee_details_data"] = employee_details
        results, total = [], len(selected) + int(tool_selected)
        completed = 0
        for key in selected:
            audit = AUDITS[key]
            try:
                if portfolio_choices.get(key) and portfolio_error: raise ValueError(portfolio_error)
                upload = cached_upload(audit.source_key)
                if not upload: raise ValueError("Required source file is missing. Return to Step 2.")
                if key == "leave_cascade":
                    schedule_table = require_columns(locate_table(read_raw(upload), audit.expected_columns, audit.required_columns), audit.required_columns)
                    time_upload = cached_upload("time_summary_current_pay_period")
                    time_expected = ("EE ID", "Last Name", "First Name", "Actual Total Hrs (Include Corrections)", "Pay Code Name", "Worker Type", "Labour Agreement")
                    time_required = ("EE ID", "Actual Total Hrs (Include Corrections)", "Pay Code Name")
                    time_table = require_columns(locate_table(read_raw(time_upload), time_expected, time_required), time_required)
                    result = audit_leave_cascade(schedule_table, time_table)
                    if portfolio_choices.get(key):
                        result = add_location_portfolio(
                            result, options.get(key, {}).get("portfolio_id_column", "EE ID"),
                            schedule_audit, supervisor,
                            options.get(key, {}).get("portfolio_date_column"),
                            options.get(key, {}).get("portfolio_employee_fallback", True))
                    payload = excel_bytes(result, audit.name, audit.source_name)
                    filename = f"[SUS] Audit - {safe_filename(audit.name)} {date.today():{DATE_FMT}}.xlsx"
                else:
                    audit_option = options.get(key, {})
                    portfolio_match = {
                        "id_column": audit_option.get("portfolio_id_column", audit.employee_id_column),
                        "date_column": audit_option.get("portfolio_date_column"),
                        "allow_employee_fallback": audit_option.get("portfolio_employee_fallback", True),
                    }
                    result, payload, filename = process_audit(
                        audit, upload,
                        schedule_audit if portfolio_choices.get(key) else None,
                        supervisor if portfolio_choices.get(key) else None,
                        audit_option, portfolio_match)
                results.append((audit, result, payload, filename, None)); status.write(f"✅ {audit.name}: {len(result):,} row(s)")
            except Exception as exc:
                results.append((audit, None, None, None, str(exc))); status.write(f"❌ {audit.name}: {exc}")
            completed += 1; progress.progress(completed / total)
        if tool_selected:
            placeholder = SimpleNamespace(key="portfolio_only", name=config.get("tool_name", "Portfolio Enriched Audit"), email_subject=None)
            try:
                if portfolio_error: raise ValueError(portfolio_error)
                result, payload, filename = process_portfolio_only(
                    cached_upload("portfolio_only_source"), schedule_audit, supervisor,
                    config["tool_id"], config.get("tool_date"),
                    config.get("tool_employee_fallback", True), config["tool_name"])
                results.append((placeholder, result, payload, filename, None)); status.write(f"✅ {placeholder.name}: {len(result):,} row(s)")
            except Exception as exc:
                results.append((placeholder, None, None, None, str(exc))); status.write(f"❌ {placeholder.name}: {exc}")
            completed += 1; progress.progress(completed / total)
        st.session_state["audit_results"] = results
        status.update(label="Queue complete", state="complete", expanded=False)
    render_results(st.session_state.get("audit_results", []))


if __name__ == "__main__":
    app()
