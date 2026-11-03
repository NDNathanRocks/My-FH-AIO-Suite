"""Pure text, number and date helpers (no Streamlit, no I/O)."""
from __future__ import annotations

import re
from typing import Iterable

import pandas as pd


def clean_text(value) -> str:
    return "" if pd.isna(value) else str(value).strip()


def normalize_header(value) -> str:
    return re.sub(r"\s+", " ", clean_text(value)).casefold()


def best_column_match(columns: Iterable, aliases: Iterable[str]):
    normalized = {normalize_header(column): column for column in columns}
    return next((normalized.get(normalize_header(alias)) for alias in aliases
                 if normalized.get(normalize_header(alias)) is not None), None)


def available_date_columns(columns: Iterable, aliases: Iterable[str]) -> list:
    columns = list(columns)
    preferred = []
    normalized = {normalize_header(column): column for column in columns}
    for alias in aliases:
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


def time_to_minutes(value):
    """Parse ``H:MM`` / ``-H:MM`` / ``H:MM:SS`` (Kronos style, '=' tolerated) to minutes."""
    text = clean_text(value).replace("=", "").replace(" ", "")
    match = re.fullmatch(r"(-?)(\d+):(\d{2})(?::(\d{2}))?", text)
    if not match:
        return None
    sign = -1 if match.group(1) == "-" else 1
    return sign * (int(match.group(2)) * 60 + int(match.group(3)) + (int(match.group(4) or 0) / 60))


def minutes_to_hhmm(minutes) -> str:
    minutes = max(0, int(round(float(minutes or 0))))
    return f"{minutes // 60}:{minutes % 60:02d}"


def natural_duration(minutes) -> str:
    minutes = max(0, int(round(float(minutes or 0))))
    hours, mins = divmod(minutes, 60)
    parts = []
    if hours:
        parts.append(f"{hours} hour" + ("s" if hours != 1 else ""))
    if mins:
        parts.append(f"{mins} minute" + ("s" if mins != 1 else ""))
    return " ".join(parts) or "0 minutes"


def parse_date_value(value):
    """Day-first parse, matching the Canadian Kronos export format."""
    return pd.to_datetime(clean_text(value), errors="coerce", dayfirst=True)


def parse_time_value(value):
    text = clean_text(value)
    parsed = pd.to_datetime(text, errors="coerce")
    return None if pd.isna(parsed) else parsed.time()


def safe_filename(name) -> str:
    return re.sub(r"\s+", " ", re.sub(r'[\\/:*?"<>|]+', "-", clean_text(name))).strip(" .") or "Audit"


def unique_casefold(values) -> list[str]:
    seen, output = set(), []
    for value in values:
        text = clean_text(value)
        key = text.casefold()
        if text and key not in seen:
            seen.add(key)
            output.append(text)
    return output


def sort_if_present(frame: pd.DataFrame, columns) -> pd.DataFrame:
    present = [c for c in columns if c in frame.columns]
    return frame.sort_values(present, kind="stable", na_position="last") if present and not frame.empty else frame
