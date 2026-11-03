from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from io import BytesIO
import re
from typing import Any

import pandas as pd
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter


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


def _name(file: Any) -> str:
    return str(getattr(file, "name", "uploaded_file"))


def sheets(file: Any) -> list[str]:
    if _name(file).lower().endswith(".csv"):
        return ["CSV Data"]
    file.seek(0)
    out = pd.ExcelFile(file, engine="openpyxl").sheet_names
    file.seek(0)
    return out


def load(file: Any, sheet: str | None = None) -> pd.DataFrame:
    file.seek(0)
    if _name(file).lower().endswith(".csv"):
        try:
            df = pd.read_csv(file, dtype="string", keep_default_na=True)
        except UnicodeDecodeError:
            file.seek(0)
            df = pd.read_csv(file, dtype="string", encoding="latin-1", keep_default_na=True)
    else:
        df = pd.read_excel(file, sheet_name=sheet, engine="openpyxl", dtype=object)
    file.seek(0)
    df.columns = [str(c).strip() for c in df.columns]
    if len(set(df.columns)) != len(df.columns):
        raise ValueError(f"Duplicate column headings found in {_name(file)}.")
    return df


def _clean_scalar(value: Any, aggressive: bool = False, case_sensitive: bool = False) -> str | pd.NA:
    if pd.isna(value):
        return pd.NA
    if isinstance(value, (pd.Timestamp, datetime, date)):
        text = pd.Timestamp(value).strftime("%Y-%m-%d")
    else:
        text = str(value).strip()
        if re.fullmatch(r"-?\d+\.0", text):
            text = text[:-2]
        # Standardize unambiguous date-like strings, useful for composite keys such as employee + shift date.
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
    parts = [frame[c].map(lambda v: _clean_scalar(v, aggressive, case_sensitive)) for c in columns]
    combined = pd.concat(parts, axis=1)
    complete = combined.notna().all(axis=1)
    # Replace missing components temporarily so string joining never receives pd.NA.
    # Incomplete composite identifiers are restored to pd.NA immediately afterward.
    key = combined.fillna("").astype("string").agg("\x1f".join, axis=1)
    return key.where(complete, pd.NA)


def _aggregate(values: pd.Series, delimiter: str, unique_only: bool) -> Any:
    cleaned = [v for v in values.tolist() if not pd.isna(v) and str(v).strip()]
    if unique_only:
        cleaned = list(dict.fromkeys(str(v) for v in cleaned))
    else:
        cleaned = [str(v) for v in cleaned]
    return delimiter.join(cleaned) if cleaned else pd.NA


def _safe_name(preferred: str, used: set[str]) -> str:
    name = preferred.strip() or "Lookup Value"
    if name not in used:
        return name
    base = f"{name} (Lookup)"
    name = base
    n = 2
    while name in used:
        name = f"{base} {n}"
        n += 1
    return name


def _soft_match_keys(left_key: Any, candidates: list[str]) -> list[str]:
    """Return standardized lookup keys containing the working key, or vice versa."""
    if pd.isna(left_key):
        return []
    needle = str(left_key)
    if not needle:
        return []
    return [candidate for candidate in candidates if needle in candidate or candidate in needle]


def process(
    left: pd.DataFrame,
    lookups: list[dict[str, Any]],
    *,
    case_sensitive: bool = False,
    duplicate_policy: str = "list",
    delimiter: str = ", ",
    unique_only: bool = True,
    include_reports: bool = False,
) -> Result:
    if not lookups:
        raise ValueError("Add at least one lookup file.")

    enriched = left.copy()
    overall_match = pd.Series(False, index=enriched.index)
    all_duplicates: list[pd.DataFrame] = []
    run_rows: list[dict[str, Any]] = []
    used = set(enriched.columns)
    all_working_keys: list[str] = []

    for number, spec in enumerate(lookups, 1):
        right = spec["data"].copy()
        source = spec.get("name", f"Lookup {number}")
        input_keys = list(spec["input_keys"])
        lookup_keys = list(spec["keys"])
        fields = list(spec["fields"])
        match_mode = spec.get("match_mode", "standardized")
        status_column = _safe_name(f"Match Status - {source}", used)
        used.add(status_column)
        all_working_keys.extend(c for c in input_keys if c not in all_working_keys)

        if not input_keys:
            raise ValueError(f"{source}: select at least one working-file identifier.")
        if len(lookup_keys) != len(input_keys):
            raise ValueError(f"{source}: map every selected working identifier to a lookup field.")
        if not fields:
            raise ValueError(f"{source}: select at least one field to add.")
        missing = [f["source"] for f in fields if f["source"] not in right.columns]
        if missing:
            raise ValueError(f"{source}: lookup field(s) not found: {', '.join(missing)}")

        base_exact = make_key(enriched, input_keys, case_sensitive, False)
        base_standard = make_key(enriched, input_keys, False, True)
        right["__key"] = make_key(right, lookup_keys, case_sensitive, False)
        right["__standard"] = make_key(right, lookup_keys, False, True)
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
        exact_groups = {key: group for key, group in right.groupby("__key", sort=False)}
        standard_groups = {key: group for key, group in right.dropna(subset=["__standard"]).groupby("__standard", sort=False)}
        standard_candidates = list(standard_groups)

        selected_groups: dict[Any, pd.DataFrame] = {}
        methods = pd.Series("Not found", index=enriched.index, dtype="string")
        for row_index in enriched.index:
            exact_key = base_exact.loc[row_index]
            standard_key = base_standard.loc[row_index]
            groups: list[pd.DataFrame] = []
            if not pd.isna(exact_key) and exact_key in exact_groups:
                groups = [exact_groups[exact_key]]
                methods.loc[row_index] = "Exact"
            elif match_mode in {"standardized", "soft"} and not pd.isna(standard_key) and standard_key in standard_groups:
                groups = [standard_groups[standard_key]]
                methods.loc[row_index] = "Standardized"
            elif match_mode == "soft" and not pd.isna(standard_key):
                matched_keys = _soft_match_keys(standard_key, standard_candidates)
                if matched_keys:
                    groups = [standard_groups[key] for key in matched_keys]
                    methods.loc[row_index] = "Soft match" if len(matched_keys) == 1 else f"Soft match ({len(matched_keys)} keys)"
            if groups:
                selected_groups[row_index] = pd.concat(groups, ignore_index=True)

        file_hits = methods.ne("Not found")
        overall_match |= file_hits
        enriched[status_column] = methods

        for field in fields:
            src = field["source"]
            preferred = field.get("output", src)
            action = field.get("conflict", "separate")
            incoming = pd.Series(pd.NA, index=enriched.index, dtype="object")
            for row_index, group in selected_groups.items():
                values = group[src]
                if duplicate_policy == "first":
                    incoming.loc[row_index] = next((v for v in values if not pd.isna(v)), pd.NA)
                elif duplicate_policy == "last":
                    valid = [v for v in values if not pd.isna(v)]
                    incoming.loc[row_index] = valid[-1] if valid else pd.NA
                else:
                    incoming.loc[row_index] = _aggregate(values, delimiter, unique_only)

            if preferred in enriched.columns and action == "first_nonblank":
                current = enriched[preferred]
                enriched[preferred] = current.where(current.notna() & current.astype("string").str.strip().ne(""), incoming)
            elif preferred in enriched.columns and action == "combine":
                def combine_pair(a: Any, b: Any) -> Any:
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
            "Match Mode": {"exact": "Exact only", "standardized": "Exact + standardized", "soft": "Exact + standardized + soft contains"}.get(match_mode, match_mode),
            "Rows": len(right),
            "Matched Working Rows": int(file_hits.sum()),
            "Exact Matches": int(methods.eq("Exact").sum()),
            "Standardized Matches": int(methods.eq("Standardized").sum()),
            "Soft Matches": int(methods.str.startswith("Soft match", na=False).sum()),
            "Duplicate Keys": duplicate_count,
            "Fields Added": ", ".join(source_columns),
        })

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

    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        enriched.to_excel(writer, sheet_name="Enriched Data", index=False)
        if include_reports:
            unmatched.to_excel(writer, sheet_name="Unmatched Records", index=False)
            duplicates_df.to_excel(writer, sheet_name="Duplicate Lookup Keys", index=False)
            summary.to_excel(writer, sheet_name="Run Summary", index=False)
            file_results.to_excel(writer, sheet_name="Lookup File Results", index=False)
        for ws in writer.book.worksheets:
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for cell in ws[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="0F766E")
            for i, col in enumerate(ws.columns, 1):
                width = max((len(str(c.value)) if c.value is not None else 0 for c in list(col)[:250]), default=0) + 2
                ws.column_dimensions[get_column_letter(i)].width = min(max(width, 12), 55)

    return Result(enriched, unmatched, duplicates_df, summary, bio.getvalue(), matched, missing, duplicate_total, file_results)
