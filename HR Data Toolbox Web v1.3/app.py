from datetime import datetime
import json
import streamlit as st
from engine import load, process, sheets


def display_frame(frame):
    preview = frame.copy()
    for column in preview.columns:
        if preview[column].dtype == "object":
            preview[column] = preview[column].astype("string")
    st.dataframe(preview, width="stretch", hide_index=True)


st.set_page_config(page_title="HR Data Toolbox", page_icon="🧰", layout="wide")
st.markdown("## HR Data Toolbox")
st.caption("Local-first multi-file matching. Each lookup can use a different working-file identifier. Output is always Excel.")

working_file = st.file_uploader("Working file", type=["xlsx", "xlsm", "csv"], key="working")
lookup_files = st.file_uploader("Lookup files", type=["xlsx", "xlsm", "csv"], accept_multiple_files=True)
if not working_file or not lookup_files:
    st.info("Upload one working file and at least one lookup file to continue.")
    st.stop()

try:
    working_sheet = st.selectbox("Working-file worksheet", sheets(working_file))
    working = load(working_file, working_sheet)
except Exception as exc:
    st.error(f"Could not open the working file: {exc}")
    st.stop()
working_columns = list(working.columns)

st.markdown("### Lookup configuration")
st.caption("Choose the working identifier independently for every lookup. For example, File 1 can use EEID and File 2 can use Location.")
lookup_specs = []
profile_lookups = []
for index, uploaded in enumerate(lookup_files):
    with st.expander(f"{index + 1}. {uploaded.name}", expanded=True):
        try:
            sheet = st.selectbox("Worksheet", sheets(uploaded), key=f"sheet_{index}")
            frame = load(uploaded, sheet)
        except Exception as exc:
            st.error(f"Could not open {uploaded.name}: {exc}")
            continue
        columns = list(frame.columns)
        input_keys = st.multiselect(
            "Working-file identifier field(s) for this lookup",
            working_columns,
            default=[working_columns[min(index, len(working_columns) - 1)]] if working_columns else [],
            key=f"input_keys_{index}",
            help="Usually one field. Select multiple fields only when this lookup needs a composite key.",
        )
        selected_keys = []
        if input_keys:
            st.caption("Map each selected working identifier to the corresponding lookup-file field.")
            key_cols = st.columns(min(len(input_keys), 3))
            for key_index, working_key in enumerate(input_keys):
                with key_cols[key_index % len(key_cols)]:
                    default_index = columns.index(working_key) if working_key in columns else min(key_index, max(len(columns) - 1, 0))
                    selected_keys.append(st.selectbox(f"{working_key} maps to", columns, index=default_index, key=f"key_{index}_{key_index}"))

        match_label = st.selectbox(
            "Match method",
            ["Exact + standardized", "Exact only", "Exact + standardized + soft contains"],
            key=f"match_mode_{index}",
            help="Soft contains lets 012345 match 012345 - PC RMH after lowercasing and removing whitespace.",
        )
        match_mode = {
            "Exact only": "exact",
            "Exact + standardized": "standardized",
            "Exact + standardized + soft contains": "soft",
        }[match_label]
        if match_mode == "soft":
            st.warning("Soft matching can produce more than one possible key. Review the per-file Match Status column and optional Lookup File Results report.")

        available = [c for c in columns if c not in selected_keys]
        chosen = st.multiselect("Fields to add", available, key=f"fields_{index}")
        field_specs = []
        for field_index, field in enumerate(chosen):
            a, b, c = st.columns([2, 2, 2])
            with a:
                st.text_input("Source field", field, disabled=True, key=f"source_{index}_{field_index}")
            with b:
                output_name = st.text_input("Output column", field, key=f"output_{index}_{field_index}")
            with c:
                conflict_label = st.selectbox("If output already exists", ["Create a separate column", "Use first nonblank value", "Combine values"], key=f"conflict_{index}_{field_index}")
            action = {"Create a separate column": "separate", "Use first nonblank value": "first_nonblank", "Combine values": "combine"}[conflict_label]
            field_specs.append({"source": field, "output": output_name, "conflict": action})

        lookup_specs.append({"name": uploaded.name, "data": frame, "input_keys": input_keys, "keys": selected_keys, "fields": field_specs, "match_mode": match_mode})
        profile_lookups.append({"file": uploaded.name, "sheet": sheet, "working_keys": input_keys, "lookup_keys": selected_keys, "match_mode": match_mode, "fields": field_specs})

with st.expander("Duplicate and output options"):
    case_sensitive = st.checkbox("Case-sensitive exact matching", value=False)
    duplicate_label = st.selectbox("When a lookup contains duplicate identifier matches", ["List all matching values", "Use first match", "Use last match", "Stop and report"])
    duplicate_policy = {"List all matching values": "list", "Use first match": "first", "Use last match": "last", "Stop and report": "error"}[duplicate_label]
    unique_only = st.checkbox("Keep unique duplicate values only", value=True, disabled=duplicate_policy != "list")
    delimiter_label = st.selectbox("List delimiter", ["Comma", "Semicolon", "Space", "New line", "Custom"], disabled=duplicate_policy != "list")
    delimiter = {"Comma": ", ", "Semicolon": "; ", "Space": " ", "New line": "\n"}.get(delimiter_label, "")
    if delimiter_label == "Custom":
        delimiter = st.text_input("Custom delimiter", value=" | ")
    include_reports = st.checkbox("Include statistics and data-quality sheets/tabs", value=False)

profile = json.dumps({"working_sheet": working_sheet, "lookups": profile_lookups, "case_sensitive": case_sensitive, "duplicate_policy": duplicate_policy, "unique_only": unique_only, "delimiter": delimiter, "include_reports": include_reports}, indent=2)
st.download_button("Download mapping profile", profile, "hr_toolbox_profile.json", "application/json")

ready = bool(lookup_specs) and all(spec["input_keys"] and len(spec["input_keys"]) == len(spec["keys"]) and spec["fields"] for spec in lookup_specs)
if st.button("Run enrichment", type="primary", width="stretch", disabled=not ready):
    try:
        with st.spinner("Matching records and creating the Excel workbook..."):
            st.session_state.result = process(working, lookup_specs, case_sensitive=case_sensitive, duplicate_policy=duplicate_policy, delimiter=delimiter, unique_only=unique_only, include_reports=include_reports)
            st.session_state.show_reports = include_reports
    except Exception as exc:
        st.error(str(exc))
if not ready:
    st.caption("For every lookup, select its working identifier, map it to a lookup identifier, and select at least one field to add.")

if "result" in st.session_state:
    result = st.session_state.result
    st.success(f"Completed: {result.matched:,} of {len(result.enriched):,} working rows matched at least one lookup.")
    display_frame(result.enriched.head(200))
    if st.session_state.get("show_reports", False):
        st.markdown("### Data quality")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Rows", f"{len(result.enriched):,}")
        c2.metric("Matched", f"{result.matched:,}")
        c3.metric("Not found", f"{result.unmatched_count:,}")
        c4.metric("Duplicate keys", f"{result.duplicate_keys:,}")
        t1, t2, t3 = st.tabs(["Unmatched", "Duplicates", "Lookup results"])
        with t1: display_frame(result.unmatched)
        with t2: display_frame(result.duplicates)
        with t3: display_frame(result.file_results)
    st.download_button("Download enriched Excel workbook", result.output, f"HR_Enriched_{datetime.now():%Y-%m-%d_%H%M}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary", width="stretch")
