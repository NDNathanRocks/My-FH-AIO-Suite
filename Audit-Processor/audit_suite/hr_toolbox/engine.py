"""HR Data Toolbox: enrich a working file from one or more lookup files.

Ported from ``HR Data Toolbox Web v1.3/engine.py`` with identical matching
rules. Changes: values are computed once per distinct match (not per row),
file errors are reported as ``InputFileError``, ``.xls`` is accepted, and text
starting with '=' is written as text rather than as an Excel formula.
"""
from __future__ import annotations

import re
import zipfile
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any

import pandas as pd

from ..common.excel import styled_sheets_bytes
from ..common.files import InputFileError

MATCH_MODE_LABELS = {"exact": "Exact only", "standardized": "Exact + standardized", "soft": "Exact + standardized + soft contains"}


@dataclass
class Result:
    enriched: pd.DataFrame
    unmatched: pd.DataFrame
    duplicates: pd.DataFrame
    summary: pd.DataFrame
    output: bytes
    matched: int
    unmatched_count: int
    duplicate_keys: int
    file_results: pd.DataFrame
    soft_matches: int = 0
    blank_key_rows: int = 0


def _name(file: Any) -> str:
    return Path(str(getattr(file, "name", "uploaded_file"))).name


def _suffix(file: Any) -> str:
    return Path(_name(file)).suffix.lower()


def _excel_engine(file: Any) -> str:
    return "xlrd" if _suffix(file) == ".xls" else "openpyxl"


def sheets(file: Any) -> list[str]:
    if _suffix(file) == ".csv":
        return ["CSV Data"]
    file.seek(0)
    try:
        out = pd.ExcelFile(file, engine=_excel_engine(file)).sheet_names
    except (zipfile.BadZipFile, ValueError, KeyError) as exc:
        raise InputFileError(f"{_name(file)}: not a readable Excel workbook ({exc}).") from exc
    finally:
        file.seek(0)
    return out


def load(file: Any, sheet: str | None = None) -> pd.DataFrame:
    file.seek(0)
    try:
        if _suffix(file) == ".csv":
            try:
                df = pd.read_csv(file, dtype="string", keep_default_na=True)
            except UnicodeDecodeError:
                file.seek(0)
                df = pd.read_csv(file, dtype="string", encoding="latin-1", keep_default_na=True)
        else:
            df = pd.read_excel(file, sheet_name=sheet, engine=_excel_engine(file), dtype=object)
    except pd.errors.EmptyDataError as exc:
        raise InputFileError(f"{_name(file)}: the file is empty.") from exc
    except (zipfile.BadZipFile, KeyError) as exc:
        raise InputFileError(f"{_name(file)}: not a readable Excel workbook (corrupt or password-protected).") from exc
    finally:
        file.seek(0)
    df.columns = [str(c).strip() for c in df.columns]
    if len(set(df.columns)) != len(df.columns):
        dupes = sorted({c for c in df.columns if list(df.columns).count(c) > 1})
        raise InputFileError(f"Duplicate column headings found in {_name(file)}: {', '.join(dupes)}.")
    return df


def _clean_scalar(value: Any, aggressive: bool = False, case_sensitive: bool = False):
    if pd.isna(value):
        return pd.NA
    if isinstance(value, (pd.Timestamp, datetime, date)):
        text = pd.Timestamp(value).strftime("%Y-%m-%d")
    else:
        text = str(value).strip()
        if re.fullmatch(r"-?\d+\.0", text):
            text = text[:-2]
        # Standardize unambiguous date-like strings (e.g. employee + shift date keys).
        if re.fullmatch(r"\d{4}[-/]\d{1,2}[-/]\d{1,2}", text):
            parsed = pd.to_datetime(text, errors="coerce")
            if not pd.isna(parsed):
                text = parsed.strftime("%Y-%m-%d")
    text = re.sub(r"\s+", " ", text).strip()
    if not case_sensitive:
        text = text.casefold()
    if aggressive:
        text = re.sub(r"\s+", "", text).casefold()
    return text if text else pd.NA


def make_key(frame: pd.DataFrame, columns: list[str], case_sensitive: bool = False, aggressive: bool = False) -> pd.Series:
    if not columns:
        raise ValueError("Select at least one identifier field.")
    missing = [c for c in columns if c not in frame.columns]
    if missing:
        raise ValueError(f"Identifier column(s) not found: {', '.join(missing)}")
    if frame.empty:
        return pd.Series(pd.NA, index=frame.index, dtype="string")
    parts = [frame[c].map(lambda v: _clean_scalar(v, aggressive, case_sensitive)) for c in columns]
    combined = pd.concat(parts, axis=1)
    complete = combined.notna().all(axis=1)
    # Blank components become "" only for joining; incomplete keys are then reset to NA.
    key = combined.fillna("").astype("string").agg("\x1f".join, axis=1)
    return key.where(complete, pd.NA)


def _aggregate(values, delimiter: str, unique_only: bool):
    cleaned = [v for v in values if not pd.isna(v) and str(v).strip()]
    cleaned = list(dict.fromkeys(str(v) for v in cleaned)) if unique_only else [str(v) for v in cleaned]
    return delimiter.join(cleaned) if cleaned else pd.NA


def _safe_name(preferred: str, used: set[str]) -> str:
    name = preferred.strip() or "Lookup Value"
    if name not in used:
        return name
    base = f"{name} (Lookup)"
    name, n = base, 2
    while name in used:
        name = f"{base} {n}"
        n += 1
    return name


def _soft_match_keys(left_key, candidates: list[str]) -> list[str]:
    """Standardized lookup keys containing the working key, or vice versa."""
    if pd.isna(left_key):
        return []
    needle = str(left_key)
    if not needle:
        return []
    return [candidate for candidate in candidates if needle in candidate or candidate in needle]


def _pick(values: list, duplicate_policy: str, delimiter: str, unique_only: bool):
    if duplicate_policy == "first":
        return next((v for v in values if not pd.isna(v)), pd.NA)
    if duplicate_policy == "last":
        valid = [v for v in values if not pd.isna(v)]
        return valid[-1] if valid else pd.NA
    return _aggregate(values, delimiter, unique_only)


def validate_spec(spec: dict, source: str, right: pd.DataFrame) -> None:
    if not spec["input_keys"]:
        raise ValueError(f"{source}: select at least one working-file identifier.")
    if len(spec["keys"]) != len(spec["input_keys"]):
        raise ValueError(f"{source}: map every selected working identifier to a lookup field.")
    if not spec["fields"]:
        raise ValueError(f"{source}: select at least one field to add.")
    missing = [f["source"] for f in spec["fields"] if f["source"] not in right.columns]
    if missing:
        raise ValueError(f"{source}: lookup field(s) not found: {', '.join(missing)}")


def process(left: pd.DataFrame, lookups: list[dict[str, Any]], *, case_sensitive: bool = False,
            duplicate_policy: str = "list", delimiter: str = ", ", unique_only: bool = True,
            include_reports: bool = False, progress=None) -> Result:
    if not lookups:
        raise ValueError("Add at least one lookup file.")

    enriched = left.copy()
    overall_match = pd.Series(False, index=enriched.index)
    all_duplicates: list[pd.DataFrame] = []
    run_rows: list[dict[str, Any]] = []
    used = set(enriched.columns)
    all_working_keys: list[str] = []
    soft_total = 0
    blank_total = 0

    for number, spec in enumerate(lookups, 1):
        right = spec["data"].copy()
        source = spec.get("name", f"Lookup {number}")
        input_keys, lookup_keys = list(spec["input_keys"]), list(spec["keys"])
        fields = list(spec["fields"])
        match_mode = spec.get("match_mode", "standardized")
        status_column = _safe_name(f"Match Status - {source}", used)
        used.add(status_column)
        all_working_keys.extend(c for c in input_keys if c not in all_working_keys)
        validate_spec({"input_keys": input_keys, "keys": lookup_keys, "fields": fields}, source, right)
        if progress:
            progress(number - 1, len(lookups), f"Matching {source}")

        base_exact = make_key(enriched, input_keys, case_sensitive, False)
        base_standard = make_key(enriched, input_keys, False, True)
        right["__key"] = make_key(right, lookup_keys, case_sensitive, False)
        right["__standard"] = make_key(right, lookup_keys, False, True)
        blank_lookup_rows = int(right["__key"].isna().sum())
        right = right[right["__key"].notna()].copy()

        duplicate_mask = right.duplicated("__key", keep=False)
        duplicates = right.loc[duplicate_mask].copy()
        duplicate_count = int(duplicates["__key"].nunique())
        if duplicate_count and duplicate_policy == "error":
            raise ValueError(f"{source}: {duplicate_count} duplicate identifier combination(s) found.")
        if not duplicates.empty:
            duplicates.insert(0, "Lookup File", source)
            all_duplicates.append(duplicates.drop(columns=["__key", "__standard"], errors="ignore"))

        source_columns = list(dict.fromkeys(f["source"] for f in fields))
        exact_index = {key: idx.tolist() for key, idx in right.groupby("__key", sort=False).groups.items()}
        standard_index = {key: idx.tolist() for key, idx in right.dropna(subset=["__standard"]).groupby("__standard", sort=False).groups.items()}
        standard_candidates = list(standard_index)
        soft_cache: dict[str, list[str]] = {}

        # selection: tuple of right-row labels in concat order, per working row
        selections: list = []
        methods: list[str] = []
        for exact_key, standard_key in zip(base_exact.tolist(), base_standard.tolist()):
            rows, method = None, "Not found"
            if not pd.isna(exact_key) and exact_key in exact_index:
                rows, method = tuple(exact_index[exact_key]), "Exact"
            elif match_mode in {"standardized", "soft"} and not pd.isna(standard_key) and standard_key in standard_index:
                rows, method = tuple(standard_index[standard_key]), "Standardized"
            elif match_mode == "soft" and not pd.isna(standard_key):
                if standard_key not in soft_cache:
                    soft_cache[standard_key] = _soft_match_keys(standard_key, standard_candidates)
                matched_keys = soft_cache[standard_key]
                if matched_keys:
                    rows = tuple(i for key in matched_keys for i in standard_index[key])
                    method = "Soft match" if len(matched_keys) == 1 else f"Soft match ({len(matched_keys)} keys)"
            selections.append(rows)
            methods.append(method)
        methods_s = pd.Series(methods, index=enriched.index, dtype="string")
        file_hits = methods_s.ne("Not found")
        overall_match |= file_hits
        enriched[status_column] = methods_s
        soft_count = int(methods_s.str.startswith("Soft match", na=False).sum())
        soft_total += soft_count

        for field in fields:
            src = field["source"]
            preferred = field.get("output", src)
            action = field.get("conflict", "separate")
            column_values = right[src]
            cache: dict = {}
            incoming_values = []
            for rows in selections:
                if rows is None:
                    incoming_values.append(pd.NA)
                    continue
                if rows not in cache:
                    cache[rows] = _pick(column_values.loc[list(rows)].tolist(), duplicate_policy, delimiter, unique_only)
                incoming_values.append(cache[rows])
            incoming = pd.Series(incoming_values, index=enriched.index, dtype="object")

            if preferred in enriched.columns and action == "first_nonblank":
                current = enriched[preferred]
                enriched[preferred] = current.where(current.notna() & current.astype("string").str.strip().ne(""), incoming)
            elif preferred in enriched.columns and action == "combine":
                def combine_pair(a, b):
                    vals = [v for v in (a, b) if not pd.isna(v) and str(v).strip()]
                    if unique_only:
                        vals = list(dict.fromkeys(str(v) for v in vals))
                    return delimiter.join(map(str, vals)) if vals else pd.NA
                enriched[preferred] = [combine_pair(a, b) for a, b in zip(enriched[preferred], incoming)]
            else:
                output = _safe_name(preferred, used)
                enriched[output] = incoming
                used.add(output)

        run_rows.append({
            "Lookup File": source,
            "Working Identifier(s)": " + ".join(input_keys),
            "Lookup Identifier(s)": " + ".join(lookup_keys),
            "Match Mode": MATCH_MODE_LABELS.get(match_mode, match_mode),
            "Rows": len(right),
            "Matched Working Rows": int(file_hits.sum()),
            "Exact Matches": int(methods_s.eq("Exact").sum()),
            "Standardized Matches": int(methods_s.eq("Standardized").sum()),
            "Soft Matches": soft_count,
            "Duplicate Keys": duplicate_count,
            "Fields Added": ", ".join(source_columns),
        })
        blank_total += blank_lookup_rows

    if progress:
        progress(len(lookups), len(lookups), "Writing workbook")
    enriched["Lookup Status"] = overall_match.map({True: "Matched at least one lookup", False: "Not found in any lookup"})
    unmatched = enriched.loc[~overall_match, all_working_keys].copy()
    unmatched.insert(0, "Source Row", unmatched.index + 2)
    unmatched["Reason"] = "No matching identifier in any configured lookup"
    duplicates_df = pd.concat(all_duplicates, ignore_index=True) if all_duplicates else pd.DataFrame()
    file_results = pd.DataFrame(run_rows)
    matched = int(overall_match.sum())
    missing = int((~overall_match).sum())
    duplicate_total = int(file_results["Duplicate Keys"].sum()) if not file_results.empty else 0
    summary = pd.DataFrame({
        "Measure": ["Rows processed", "Rows matched in at least one lookup", "Rows not found in any lookup", "Match rate", "Duplicate lookup identifiers", "Lookup files"],
        "Value": [len(enriched), matched, missing, f"{matched / len(enriched):.1%}" if len(enriched) else "0.0%", duplicate_total, len(lookups)],
    })

    sheets_out = {"Enriched Data": enriched}
    if include_reports:
        sheets_out.update({"Unmatched Records": unmatched, "Duplicate Lookup Keys": duplicates_df,
                           "Run Summary": summary, "Lookup File Results": file_results})
    output = styled_sheets_bytes(sheets_out)
    if len(enriched) != len(left):  # defensive: enrichment must never add or drop rows
        raise RuntimeError("Row count changed during enrichment; output withheld.")
    return Result(enriched, unmatched, duplicates_df, summary, output, matched, missing, duplicate_total, file_results, soft_total, blank_total)
