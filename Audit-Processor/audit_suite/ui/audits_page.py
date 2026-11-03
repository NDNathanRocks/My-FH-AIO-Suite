"""Three-stage audit workflow: select audits -> upload & configure -> review & run."""
from __future__ import annotations

import random
from datetime import date, datetime, timedelta

import streamlit as st

from ..audits import AUDITS
from ..audits.registry import SOURCE_LABELS
from ..audits.runner import (process_audit, process_leave_cascade, process_portfolio_only,
                             read_employee_details, run_report_frame, run_tool)
from ..common.excel import XLSX_MIME, visible_columns
from ..common.files import detect_existing_table, read_raw
from ..common.text import available_date_columns, best_column_match, clean_text, employee_key
from ..config import (DATE_COLUMN_ALIASES, EMPLOYEE_DETAILS_URL, EMPLOYEE_ID_ALIASES, NO_DATE_OPTION,
                      SCHEDULE_AUDIT_URL, SUPERVISOR_PORTFOLIO_URL)
from ..outlook_drafts import create_leave_employee_draft, create_outlook_draft
from .common import (apply_styles, cached_upload, clear_uploaded_data, counts_row, hero,
                     schedule_audit_table, supervisor_table, upload_widget)

STATUS_ICON = {"success": "✅", "warning": "⚠️", "failed": "❌"}


@st.cache_data(show_spinner=False, max_entries=8, ttl=3600)
def _existing_table(name: str, data: bytes):
    from ..common.files import MemoryUpload
    return detect_existing_table(read_raw(MemoryUpload(name, data)))


def read_existing_table(upload):
    return _existing_table(upload.name, upload.getvalue())


def selected_audit_keys():
    return [key for key in AUDITS if st.session_state.get(f"select_{key}", False)]


def source_groups_for(selected):
    groups = {}
    for key in selected:
        groups.setdefault(AUDITS[key].source_key, []).append(key)
        for extra in AUDITS[key].extra_sources:
            groups.setdefault(extra, []).append(key)
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
        if c1.button("Select all", width="stretch"):
            for key in AUDITS:
                st.session_state[f"select_{key}"] = True
            st.session_state["select_portfolio_only"] = False
            st.rerun()
        if c2.button("Clear", width="stretch"):
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
    return selected_audit_keys(), bool(st.session_state.get("select_portfolio_only", False))


def upload_tile(source_key, audit_keys):
    audit = AUDITS[audit_keys[0]]
    name, url, settings = (audit.source_name, audit.source_url, audit.export_settings)
    if source_key != audit.source_key:
        name, url, settings = SOURCE_LABELS.get(source_key, (source_key, "", ""))
    with st.container(border=True, key=f"source_card_{source_key}"):
        st.markdown(f"#### {name}")
        st.caption(settings)
        if url:
            st.markdown(f"[Open export ↗]({url})")
        st.markdown("**Used by:** " + " · ".join(AUDITS[key].name for key in audit_keys))
        return upload_widget(f"Upload {name}", source_key, f"upload_{source_key}")


def portfolio_column_pickers(columns, default_id, key_prefix, id_label):
    detected_id = best_column_match(columns, EMPLOYEE_ID_ALIASES) or default_id
    id_index = columns.index(detected_id) if detected_id in columns else None
    id_col = st.selectbox(id_label, columns, index=id_index, placeholder="Choose employee identifier", key=f"{key_prefix}id")
    date_options = [NO_DATE_OPTION] + available_date_columns(columns, DATE_COLUMN_ALIASES)
    detected_date = best_column_match(columns, DATE_COLUMN_ALIASES)
    date_index = date_options.index(detected_date) if detected_date in date_options else 0
    chosen_date = st.selectbox("Shift/date column for portfolio matching", date_options, index=date_index, key=f"{key_prefix}date",
                               help="The best match is selected automatically. Confirm it, choose another column, or choose No date column.")
    fallback = st.checkbox("Allow EE ID-only fallback", value=True, key=f"{key_prefix}ee_fallback")
    return id_col, (None if chosen_date == NO_DATE_OPTION else chosen_date), fallback, detected_id


def render_results(runs):
    if not runs:
        st.info("Run the selected audits to see results here.")
        return
    totals = {"success": 0, "warning": 0, "failed": 0}
    for run in runs:
        totals[run.status] += 1
    counts_row([("Tools run", len(runs)), ("Succeeded", totals["success"]), ("With warnings", totals["warning"]), ("Failed", totals["failed"])])
    if totals["failed"]:
        st.error(f"{totals['failed']} tool(s) failed. Their results are not included; other tools completed normally.")
    report = run_report_frame(runs).to_csv(index=False).encode("utf-8-sig")
    st.download_button("Download run report (CSV)", report, f"Audit run report {datetime.now():%Y-%m-%d %H%M}.csv", "text/csv",
                       key="download_run_report")
    for run in runs:
        audit = run.audit
        with st.container(border=True, key=f"result_card_{run.key}"):
            st.markdown(f"### {STATUS_ICON[run.status]} {run.name}")
            if run.error:
                st.error(run.error)
                continue
            counts_row([("Rows read", run.rows_read), ("Findings", run.findings), ("Blank rows skipped", run.blank_rows_skipped)])
            result = run.result
            if run.notes:
                with st.expander(f"Run notes ({len(run.notes)})", expanded=run.status == "warning"):
                    for note in run.notes:
                        st.write("• " + note)
            if len(result):
                st.dataframe(visible_columns(result), width="stretch", hide_index=True, height=330)
                if audit is not None and audit.key == "leave_cascade":
                    _leave_email_tools(audit, result)
            else:
                st.success("No findings matched the audit rules.")
            st.download_button("Download full Excel audit file", run.payload, run.filename, XLSX_MIME, key=f"download_{run.key}", width="stretch")
            if audit is not None and audit.email_subject and all(c in result.columns for c in ("Portfolio Email", "Coordinator Name")):
                if st.button("Create and open Outlook email draft", key=f"email_{run.key}", width="stretch"):
                    try:
                        st.success(f"Outlook draft created with {create_outlook_draft(audit, result):,} attachment row(s).")
                    except Exception as exc:
                        st.error(str(exc))


def _leave_email_tools(audit, result):
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
                if not found.empty:
                    employee_email = clean_text(found.iloc[0]["Email"])
            st.caption(f"Employee email: {employee_email or 'Blank - complete manually in Outlook'}")
            if st.button("Create employee Outlook draft", disabled=not chosen, key="leave_email_create"):
                try:
                    st.success(f"Draft created with {create_leave_employee_draft(audit, result.loc[chosen], employee_email)} selected row(s).")
                except Exception as exc:
                    st.error(str(exc))


def _stage_one():
    st.markdown('<div class="section-head blue"><strong>STEP 1 · Select audits</strong><br><small>Choose only the audits or optional tool needed for this run.</small></div>', unsafe_allow_html=True)
    selected, tool_selected = audit_selector()
    count = len(selected) + int(tool_selected)
    left, right = st.columns([4, 1])
    left.markdown(f'<div class="summary-bar">{count} item(s) selected</div>', unsafe_allow_html=True)
    if right.button("Continue to uploads →", type="primary", disabled=not count, width="stretch"):
        signature = list(selected) + (["__portfolio_only__"] if tool_selected else [])
        if st.session_state.get("chosen_signature") != signature:
            clear_uploaded_data()
        st.session_state.update({"chosen_audits": list(selected), "chosen_portfolio_only": tool_selected,
                                 "chosen_signature": signature, "workflow_stage": 2})
        st.rerun()


def _stage_two(selected, tool_selected):
    groups = source_groups_for(selected)
    options = {key: {} for key in selected}
    portfolio_choices = {}
    st.markdown('<div class="section-head mint"><strong>STEP 2 · Upload and configure</strong><br><small>Upload files once and review audit options.</small></div>', unsafe_allow_html=True)
    upload_cols = st.columns(2, gap="medium")
    group_items = list(groups.items())
    for i, (source_key, audit_keys) in enumerate(group_items):
        with upload_cols[i % 2]:
            upload_tile(source_key, audit_keys)
    if tool_selected:
        with upload_cols[len(group_items) % 2].container(border=True):
            st.markdown("#### Existing filtered audit")
            upload_widget("Upload existing filtered audit", "portfolio_only_source", "upload_portfolio_only")

    st.markdown("#### PC Portfolio and audit options")
    option_cols = st.columns(2, gap="medium")
    for i, key in enumerate(selected):
        audit = AUDITS[key]
        with option_cols[i % 2].container(border=True):
            st.markdown(f"#### {audit.name}")
            portfolio_choices[key] = st.checkbox("Add PC Portfolio details", value=True, key=f"portfolio_{key}")
            if portfolio_choices[key] and cached_upload(audit.source_key):
                try:
                    columns = list(read_existing_table(cached_upload(audit.source_key)).columns)
                    id_col, date_col, fallback, _ = portfolio_column_pickers(columns, audit.employee_id_column, f"portfolio_{key}_",
                                                                             "Employee identifier column for portfolio matching")
                    options[key].update({"portfolio_id_column": id_col, "portfolio_date_column": date_col, "portfolio_employee_fallback": fallback})
                except Exception as exc:
                    st.error(f"Could not inspect source columns: {exc}")
            if key == "unprocessed_tors":
                options[key]["cutoff_date"] = st.date_input("Flag Submitted/Pending TORs created before", date.today() - timedelta(days=3), key="tor_cutoff")
            with st.expander("Steps"):
                for n, step in enumerate(audit.steps, 1):
                    st.write(f"{n}. {step}")
    tool_config = {}
    if tool_selected:
        generic = cached_upload("portfolio_only_source")
        with option_cols[len(selected) % 2].container(border=True):
            st.markdown("#### PC Portfolio Enrichment Only")
            tool_config["name"] = st.text_input("Output audit name", value="Portfolio Enriched Audit", key="portfolio_only_name")
            if generic:
                try:
                    preview = read_existing_table(generic)
                    id_col, date_col, fallback, detected = portfolio_column_pickers(list(preview.columns), None, "portfolio_only_",
                                                                                    "Employee identifier column for Schedule Audit matching")
                    tool_config.update({"id": id_col, "date": date_col, "fallback": fallback})
                    if detected is None and not id_col:
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
            upload_widget("Upload Schedule Audit", "schedule_audit", "schedule_audit_upload")
        with portfolio_cols[1].container(border=True):
            st.markdown("#### Supervisor-Portfolio-Department Report")
            st.caption("Used to match the cleaned WFM department path")
            st.markdown(f"[Open report ↗]({SUPERVISOR_PORTFOLIO_URL})")
            upload_widget("Upload Supervisor Portfolio report", "supervisor_portfolio", "supervisor_portfolio_upload")

    employee_details_enabled = False
    if "leave_cascade" in selected:
        employee_details_enabled = st.checkbox("Add Employee Details for employee-facing email drafts", value=True, key="employee_details_enabled")
        if employee_details_enabled:
            with st.container(border=True):
                st.markdown("#### Employee Details")
                st.caption("Optional, default enabled · Matches EE ID to EmployeeNumber for employee email")
                st.markdown(f"[Open report ↗]({EMPLOYEE_DETAILS_URL})")
                upload_widget("Upload Employee Details", "employee_details", "employee_details_upload")

    missing = [k for k in groups if not cached_upload(k)]
    if tool_selected and not (cached_upload("portfolio_only_source") and tool_config.get("id")):
        missing.append("existing filtered audit + identifier column")
    if needs_portfolio and not (cached_upload("schedule_audit") and cached_upload("supervisor_portfolio")):
        missing.append("Schedule Audit and Supervisor-Portfolio report")
    if employee_details_enabled and not cached_upload("employee_details"):
        missing.append("Employee Details")
    back, summary, next_col = st.columns([1, 3, 1.4], vertical_alignment="center")
    if back.button("← Back", width="stretch"):
        st.session_state["workflow_stage"] = 1
        st.rerun()
    text = f"{len(selected) + int(tool_selected)} item(s) ready for review" if not missing else "Still needed: " + "; ".join(missing)
    summary.markdown(f'<div class="summary-bar">{text}</div>', unsafe_allow_html=True)
    if next_col.button("Review & run →", type="primary", disabled=bool(missing), width="stretch"):
        st.session_state["run_config"] = {
            "selected": selected, "tool": tool_selected, "portfolio_choices": portfolio_choices, "options": options,
            "tool_name": tool_config.get("name", "Portfolio Enriched Audit"), "tool_id": tool_config.get("id"),
            "tool_date": tool_config.get("date"), "tool_employee_fallback": tool_config.get("fallback", True),
            "employee_details_enabled": employee_details_enabled,
        }
        st.session_state["workflow_stage"] = 3
        st.rerun()


def _run_queue(config):
    selected, tool_selected = config["selected"], config["tool"]
    portfolio_choices, options = config["portfolio_choices"], config["options"]
    status, progress = st.status("Processing queue", expanded=True), st.progress(0.0)
    needs_portfolio = tool_selected or any(portfolio_choices.values())
    schedule_audit = supervisor = portfolio_error = None
    if needs_portfolio:
        try:
            sa, sp = cached_upload("schedule_audit"), cached_upload("supervisor_portfolio")
            if not sa or not sp:
                raise ValueError("Schedule Audit or Supervisor-Portfolio report is missing. Return to Step 2.")
            status.write("Reading Schedule Audit and Supervisor-Portfolio report…")
            schedule_audit, supervisor = schedule_audit_table(sa.name, sa.data), supervisor_table(sp.name, sp.data)
        except Exception as exc:
            portfolio_error = f"Portfolio reference files could not be read: {exc}"
            status.write(f"❌ {portfolio_error}")
    employee_details = None
    if config.get("employee_details_enabled") and cached_upload("employee_details"):
        try:
            employee_details = read_employee_details(cached_upload("employee_details"))
        except Exception as exc:
            status.write(f"⚠ Employee Details could not be read: {exc}")
    st.session_state["employee_details_data"] = employee_details

    runs, total = [], len(selected) + int(tool_selected)

    def portfolio_args(key):
        if not portfolio_choices.get(key):
            return None, None, None
        if portfolio_error:
            raise ValueError(portfolio_error)
        opt = options.get(key, {})
        return schedule_audit, supervisor, {"id_column": opt.get("portfolio_id_column", AUDITS[key].employee_id_column),
                                            "date_column": opt.get("portfolio_date_column"),
                                            "allow_employee_fallback": opt.get("portfolio_employee_fallback", True)}

    for done, key in enumerate(selected, 1):
        audit = AUDITS[key]
        status.write(f"Running {audit.name}…")

        def job(key=key, audit=audit):
            sa, sp, match = portfolio_args(key)
            upload = cached_upload(audit.source_key)
            if not upload:
                raise ValueError("Required source file is missing. Return to Step 2.")
            if key == "leave_cascade":
                return process_leave_cascade(audit, upload, cached_upload("time_summary_current_pay_period"), sa, sp, match)
            audit_options = {k: v for k, v in options.get(key, {}).items() if not k.startswith("portfolio_")}
            return process_audit(audit, upload, sa, sp, audit_options, match)

        run = run_tool(key, audit.name, job, audit)
        runs.append(run)
        status.write(f"{STATUS_ICON[run.status]} {audit.name}: " + (run.error or f"{run.findings:,} finding(s) from {run.rows_read:,} row(s)"))
        progress.progress(done / total)
    if tool_selected:
        name = config.get("tool_name") or "Portfolio Enriched Audit"

        def tool_job():
            if portfolio_error:
                raise ValueError(portfolio_error)
            return process_portfolio_only(cached_upload("portfolio_only_source"), schedule_audit, supervisor, config["tool_id"],
                                          config.get("tool_date"), config.get("tool_employee_fallback", True), name)

        run = run_tool("portfolio_only", name, tool_job)
        runs.append(run)
        status.write(f"{STATUS_ICON[run.status]} {name}: " + (run.error or f"{run.findings:,} row(s)"))
        progress.progress(1.0)
    failed = sum(r.status == "failed" for r in runs)
    status.update(label="Queue complete" if not failed else f"Queue finished · {failed} failed",
                  state="complete" if not failed else "error", expanded=bool(failed))
    return runs


def _start_run():
    st.session_state["audit_running"] = True


def _stage_three():
    config = st.session_state.get("run_config")
    if not config:
        st.session_state["workflow_stage"] = 2
        st.rerun()
    selected, tool_selected = config["selected"], config["tool"]
    st.markdown('<div class="section-head peach"><strong>STEP 3 · Review and run</strong><br><small>Confirm the queue, process audits, and download results.</small></div>', unsafe_allow_html=True)
    with st.container(border=True):
        if selected:
            st.write("**Audits:** " + " · ".join(AUDITS[k].name for k in selected))
        if tool_selected:
            st.write("**Optional tool:** PC Portfolio Enrichment Only")
    running = st.session_state.get("audit_running", False)
    back, run_col = st.columns([1, 2])
    if back.button("← Edit uploads", width="stretch", disabled=running):
        st.session_state["workflow_stage"] = 2
        st.rerun()
    run_col.button("Run selected items", type="primary", width="stretch", disabled=running, on_click=_start_run)
    if running:
        try:
            st.session_state.pop("audit_results", None)
            st.session_state["audit_results"] = _run_queue(config)
        finally:
            st.session_state["audit_running"] = False
    render_results(st.session_state.get("audit_results", []))


def page():
    apply_styles()
    taglines = ["Upload once, audit systematically.", "Less sorting. More reviewing.", "One guided path from export to action."]
    hero("⚡ Audit Prepper", random.choice(taglines))
    stage = st.session_state.get("workflow_stage", 1)
    progress_nav(stage)
    if stage == 1:
        _stage_one()
        return
    selected = [k for k in st.session_state.get("chosen_audits", []) if k in AUDITS]
    tool_selected = bool(st.session_state.get("chosen_portfolio_only", False))
    if not selected and not tool_selected:
        st.session_state["workflow_stage"] = 1
        st.rerun()
    if stage == 2:
        _stage_two(selected, tool_selected)
    else:
        _stage_three()
