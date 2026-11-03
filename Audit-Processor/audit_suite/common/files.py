"""Shared upload validation and table reading for every tool.

Uploads are held in memory only (``MemoryUpload``); nothing is written to disk.
"""
from __future__ import annotations

import csv
import io
import re
import zipfile
from dataclasses import dataclass
from pathlib import Path

import pandas as pd

from ..config import MAX_UPLOAD_BYTES, UPLOAD_TYPES
from .text import clean_text, normalize_header


class InputFileError(ValueError):
    """A user-facing problem with an uploaded file (message is safe to display)."""


@dataclass(frozen=True)
class MemoryUpload:
    """Upload snapshot that survives Streamlit widget cleanup between stages."""
    name: str
    data: bytes

    def getvalue(self) -> bytes:
        return self.data

    def stream(self) -> io.BytesIO:
        buf = io.BytesIO(self.data)
        buf.name = self.name  # type: ignore[attr-defined]
        return buf


def validate_upload(name: str, data: bytes, allowed=UPLOAD_TYPES, max_bytes: int = MAX_UPLOAD_BYTES) -> MemoryUpload:
    """Return a MemoryUpload or raise InputFileError with an actionable message."""
    safe_name = re.split(r"[\\/]", str(name or "upload"))[-1] or "upload"  # strip any client-supplied path
    suffix = Path(safe_name).suffix.lower().lstrip(".")
    if suffix not in allowed:
        raise InputFileError(f"{safe_name}: unsupported file type. Upload {', '.join(t.upper() for t in allowed)}.")
    if not data:
        raise InputFileError(f"{safe_name}: the file is empty.")
    if len(data) > max_bytes:
        raise InputFileError(f"{safe_name}: exceeds the {max_bytes // (1024 * 1024)} MB limit.")
    return MemoryUpload(safe_name, bytes(data))


def _read_excel(data: bytes, engine: str, name: str) -> list[list[object]]:
    try:
        return pd.read_excel(io.BytesIO(data), header=None, dtype=object, engine=engine).values.tolist()
    except (zipfile.BadZipFile, KeyError) as exc:
        raise InputFileError(f"{name}: not a readable Excel workbook (corrupt, password-protected, or a different format renamed to .{engine == 'xlrd' and 'xls' or 'xlsx'}).") from exc
    except ImportError as exc:
        raise InputFileError(f"{name}: reading this format needs an extra package ({exc}). Re-run the install step.") from exc
    except ValueError as exc:
        raise InputFileError(f"{name}: could not be read as Excel ({exc}).") from exc


def read_raw(uploaded_file) -> list[list[object]]:
    """Read the first worksheet / CSV into raw rows (no header detection)."""
    data = uploaded_file.getvalue()
    name = Path(uploaded_file.name).name
    suffix = Path(name).suffix.lower()
    if not data:
        raise InputFileError(f"{name}: the file is empty.")
    if suffix == ".csv":
        last_error = None
        for encoding in ("utf-8-sig", "utf-8", "cp1252", "latin1"):
            try:
                text = data.decode(encoding)
            except UnicodeDecodeError as exc:
                last_error = exc
                continue
            try:
                delimiter = csv.Sniffer().sniff(text[:8192], delimiters=",;\t|").delimiter
            except csv.Error:
                delimiter = ","
            rows = list(csv.reader(io.StringIO(text, newline=""), delimiter=delimiter))
            if not rows:
                raise InputFileError(f"{name}: the CSV is empty.")
            return rows
        raise InputFileError(f"{name}: the CSV could not be decoded. Last error: {last_error}")
    if suffix in (".xlsx", ".xlsm"):
        return _read_excel(data, "openpyxl", name)
    if suffix == ".xls":
        return _read_excel(data, "xlrd", name)
    raise InputFileError(f"{name}: unsupported file type. Upload CSV, XLSX, XLSM, or XLS.")


def locate_table(raw_rows, expected_columns, required_columns=None) -> pd.DataFrame:
    """Find the header row (within the first 50 rows) that holds the required columns."""
    expected_norm = {normalize_header(c) for c in expected_columns}
    required_norm = {normalize_header(c) for c in (required_columns or expected_columns)}
    candidates = []
    for index, row in enumerate(raw_rows[:50]):
        row_norm = {normalize_header(v) for v in row if clean_text(v)}
        candidates.append((required_norm.issubset(row_norm), len(expected_norm & row_norm), index))
    has_required, _, header_index = max(candidates, default=(False, 0, -1))
    if not has_required:
        raise InputFileError("Could not identify the table header. Required columns: " + ", ".join(required_columns or expected_columns))
    return table_from_header(raw_rows, header_index)


def table_from_header(raw_rows, header_index: int) -> pd.DataFrame:
    """Build a frame from ``header_index``. Fully blank rows are dropped and counted
    in ``frame.attrs['blank_rows_skipped']`` so callers can report them."""
    headers = [clean_text(v) for v in raw_rows[header_index]]
    while headers and not headers[-1]:
        headers.pop()
    if not headers:
        raise InputFileError("The uploaded file does not contain a usable header row.")
    seen, unique_headers = {}, []
    for i, header in enumerate(headers):
        base = header or f"Unnamed Column {i + 1}"
        seen[base] = seen.get(base, 0) + 1
        unique_headers.append(base if seen[base] == 1 else f"{base} ({seen[base]})")
    width = len(unique_headers)
    body = [list(row[:width]) + [None] * max(0, width - len(row)) for row in raw_rows[header_index + 1:]]
    full = pd.DataFrame(body, columns=unique_headers)
    frame = full.dropna(how="all")
    blank_rows = len(full) - len(frame)
    frame = frame.loc[:, ~frame.columns.astype(str).str.startswith("Unnamed Column")].reset_index(drop=True)
    frame.attrs["source_rows"] = len(full)
    frame.attrs["blank_rows_skipped"] = int(blank_rows)
    return frame


def detect_existing_table(raw) -> pd.DataFrame:
    """Header heuristic for already-filtered tables: widest early row, preferring EE ID headers."""
    best_index, best_score = -1, -1
    for index, row in enumerate(raw[:50]):
        values = [clean_text(v) for v in row if clean_text(v)]
        normalized = {normalize_header(v) for v in values}
        score = len(values) + (1000 if "ee id" in normalized or "employee id" in normalized else 0)
        if score > best_score:
            best_index, best_score = index, score
    if best_index < 0:
        raise InputFileError("Could not find a table header in the uploaded file.")
    return table_from_header(raw, best_index)


def require_columns(frame: pd.DataFrame, columns) -> pd.DataFrame:
    lookup = {normalize_header(c): c for c in frame.columns}
    missing = [c for c in columns if normalize_header(c) not in lookup]
    if missing:
        raise InputFileError("Missing required column(s): " + ", ".join(missing))
    renamed = frame.rename(columns={lookup[normalize_header(c)]: c for c in columns})
    renamed.attrs.update(frame.attrs)
    return renamed


def read_table(uploaded_file, expected_columns, required_columns) -> pd.DataFrame:
    """read_raw + locate_table + require_columns, the standard audit-source read."""
    return require_columns(locate_table(read_raw(uploaded_file), expected_columns, required_columns), required_columns)
