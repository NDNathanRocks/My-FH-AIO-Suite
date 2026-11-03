"""HR Data Toolbox page (ported from HR Data Toolbox Web v1.3)."""
from __future__ import annotations

import json
from datetime import datetime

import streamlit as st

from ..common.excel import XLSX_MIME
from ..common.files import InputFileError, validate_upload
from ..hr_toolbox import load, process, sheets
from .common import apply_styles, counts_row, hero

HR_TYPES = ("xlsx", "xlsm", "xls", "csv")
MATCH_LABELS = {"Exact + standardized": "standardized", "Exact only": "exact", "Exact + standardized + soft contains": "soft"}
CONFLICT_LABELS = {"Create a separate column": "separate", "Use first nonblank value": "first_nonblank", "Combine values": "combine"}
DUPLICATE_LABELS = {"List all matching values": "list", "Use first match": "first", "Use last match": "last", "Stop and report": "error"}


def display_frame(frame):
    preview = frame.copy()
    for column in preview.columns:
        if preview[column].dtype == "object":
            preview[column] = preview[column].astype("string")
    st.dataframe(preview, width="stretch", hide_index=True)


def _stream(uploaded):
    return validate_upload(uploaded.name, uploaded.getvalue(), allowed=HR_TYPES).stream()


def _lookup_config(index, uploaded, working_columns):
    with st.expander(f"{index + 1}. {uploaded.name}", expanded=True):
        try:
            stream = _stream(uploaded)
            sheet = st.selectbox("Worksheet", sheets(stream), key=f"hr_sheet_{index}")
            frame = load(stream, sheet)
        except InputFileError as exc:
            st.error(str(exc))
            return None, None
        except Exception as exc:
            st.error(f"Could not open {uploaded.name}: {exc}")
            return None, None
        columns = list(frame.columns)
        st.caption(f"{len(frame):,} row(s) · {len(columns)} column(s)")
        input_keys = st.multiselect(
            "Working-file identifier field(s) for this lookup", working_columns,
            default=[working_columns[min(index, len(working_columns) - 1)]] if working_columns else [],
            key=f"hr_input_keys_{index}", help="Usually one field. Select multiple fields only when this lookup needs a composite key.")
        selected_keys = []
        if input_keys:
            st.caption("Map each selected working identifier to the corresponding lookup-file field.")
            key_cols = st.columns(min(len(input_keys), 3))
            for key_index, working_key in enumerate(input_keys):
                with key_cols[key_index % len(key_cols)]:
                    default_index = columns.index(working_key) if working_key in columns else min(key_index, max(len(columns) - 1, 0))
                    selected_keys.append(st.selectbox(f"{working_key} maps to", columns, index=default_index, key=f"hr_key_{index}_{key_index}"))
        match_label = st.selectbox("Match method", list(MATCH_LABELS), key=f"hr_match_mode_{index}",
                                   help="Soft contains lets 012345 match 012345 - PC RMH after lowercasing and removing whitespace.")
        match_mode = MATCH_LABELS[match_label]
        if match_mode == "soft":
            st.warning("Soft matching can produce more than one possible key. Review the per-file Match Status column and optional Lookup File Results report.")
        available = [c for c in columns if c not in selected_keys]
        chosen = st.multiselect("Fields to add", available, key=f"hr_fields_{index}")
        field_specs = []
        for field_index, field in enumerate(chosen):
            a, b, c = st.columns([2, 2, 2])
            a.text_input("Source field", field, disabled=True, key=f"hr_source_{index}_{field_index}")
            output_name = b.text_input("Output column", field, key=f"hr_output_{index}_{field_index}")
            conflict = c.selectbox("If output already exists", list(CONFLICT_LABELS), key=f"hr_conflict_{index}_{field_index}")
            field_specs.append({"source": field, "output": output_name, "conflict": CONFLICT_LABELS[conflict]})
    spec = {"name": uploaded.name, "data": frame, "input_keys": input_keys, "keys": selected_keys, "fields": field_specs, "match_mode": match_mode}
    profile = {"file": uploaded.name, "sheet": sheet, "working_keys": input_keys, "lookup_keys": selected_keys, "match_mode": match_mode, "fields": field_specs}
    return spec, profile


def _start():
    st.session_state["hr_running"] = True


def page():
    apply_styles()
    hero("🧰 HR Data Toolbox", "Local multi-file matching. Each lookup can use a different working-file identifier. Output is always Excel.")
    working_file = st.file_uploader("Working file", type=list(HR_TYPES), key="hr_working")
    lookup_files = st.file_uploader("Lookup files", type=list(HR_TYPES), accept_multiple_files=True, key="hr_lookups")
    if not working_file or not lookup_files:
        st.info("Upload one working file and at least one lookup file to continue.")
        return
    try:
        stream = _stream(working_file)
        working_sheet = st.selectbox("Working-file worksheet", sheets(stream), key="hr_working_sheet")
        working = load(stream, working_sheet)
    except Exception as exc:
        st.error(f"Could not open the working file: {exc}")
        return
    working_columns = list(working.columns)
    st.caption(f"Working file: {len(working):,} row(s) · {len(working_columns)} column(s)")

    st.markdown("### Lookup configuration")
    st.caption("Choose the working identifier independently for every lookup. For example, File 1 can use EEID and File 2 can use Location.")
    lookup_specs, profile_lookups, unreadable = [], [], 0
    for index, uploaded in enumerate(lookup_files):
        spec, profile = _lookup_config(index, uploaded, working_columns)
        if spec is None:
            unreadable += 1
            continue
        lookup_specs.append(spec)
        profile_lookups.append(profile)

    with st.expander("Duplicate and output options"):
        case_sensitive = st.checkbox("Case-sensitive exact matching", value=False, key="hr_case")
        duplicate_policy = DUPLICATE_LABELS[st.selectbox("When a lookup contains duplicate identifier matches", list(DUPLICATE_LABELS), key="hr_dup")]
        unique_only = st.checkbox("Keep unique duplicate values only", value=True, disabled=duplicate_policy != "list", key="hr_unique")
        delimiter_label = st.selectbox("List delimiter", ["Comma", "Semicolon", "Space", "New line", "Custom"], disabled=duplicate_policy != "list", key="hr_delim")
        delimiter = {"Comma": ", ", "Semicolon": "; ", "Space": " ", "New line": "\n"}.get(delimiter_label, "")
        if delimiter_label == "Custom":
            delimiter = st.text_input("Custom delimiter", value=" | ", key="hr_custom_delim")
        include_reports = st.checkbox("Include statistics and data-quality sheets/tabs", value=False, key="hr_reports")

    profile = json.dumps({"working_sheet": working_sheet, "lookups": profile_lookups, "case_sensitive": case_sensitive,
                          "duplicate_policy": duplicate_policy, "unique_only": unique_only, "delimiter": delimiter,
                          "include_reports": include_reports}, indent=2)
    st.download_button("Download mapping profile", profile, "hr_toolbox_profile.json", "application/json")

    ready = bool(lookup_specs) and not unreadable and all(
        s["input_keys"] and len(s["input_keys"]) == len(s["keys"]) and s["fields"] for s in lookup_specs)
    running = st.session_state.get("hr_running", False)
    st.button("Run enrichment", type="primary", width="stretch", disabled=not ready or running, on_click=_start)
    if unreadable:
        st.caption(f"{unreadable} lookup file(s) could not be opened. Remove or replace them to continue.")
    elif not ready:
        st.caption("For every lookup, select its working identifier, map it to a lookup identifier, and select at least one field to add.")
    if running:
        try:
            st.session_state.pop("hr_result", None)
            bar = st.progress(0.0, text="Starting…")

            def progress(done, total, text):
                bar.progress(min(done / max(total, 1), 1.0), text=text)
            st.session_state["hr_result"] = process(working, lookup_specs, case_sensitive=case_sensitive, duplicate_policy=duplicate_policy,
                                                    delimiter=delimiter, unique_only=unique_only, include_reports=include_reports, progress=progress)
            st.session_state["hr_show_reports"] = include_reports
            bar.progress(1.0, text="Done")
        except Exception as exc:
            st.error(f"Enrichment failed; no workbook was produced. {exc}")
        finally:
            st.session_state["hr_running"] = False

    result = st.session_state.get("hr_result")
    if result is None:
        return
    st.success(f"Completed: {result.matched:,} of {len(result.enriched):,} working rows matched at least one lookup.")
    counts_row([("Rows processed", len(result.enriched)), ("Matched", result.matched), ("Not found", result.unmatched_count),
                ("Soft matches", result.soft_matches), ("Duplicate keys", result.duplicate_keys)])
    if result.blank_key_rows:
        st.warning(f"{result.blank_key_rows:,} lookup row(s) had a blank identifier and could not be matched.")
    display_frame(result.enriched.head(200))
    if st.session_state.get("hr_show_reports", False):
        st.markdown("### Data quality")
        t1, t2, t3 = st.tabs(["Unmatched", "Duplicates", "Lookup results"])
        with t1:
            display_frame(result.unmatched)
        with t2:
            display_frame(result.duplicates)
        with t3:
            display_frame(result.file_results)
    st.download_button("Download enriched Excel workbook", result.output, f"HR_Enriched_{datetime.now():%Y-%m-%d_%H%M}.xlsx",
                       XLSX_MIME, type="primary", width="stretch")
