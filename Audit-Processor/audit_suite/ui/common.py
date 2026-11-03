"""Streamlit helpers shared by every page: upload cache, styles, result widgets."""
from __future__ import annotations

import streamlit as st

from ..common.files import InputFileError, MemoryUpload, validate_upload
from ..portfolio import prepare_schedule_audit, prepare_supervisor_portfolio

CSS = """
<style>
html,body,.stApp,button,input,label,p,div{font-family:'Segoe UI',Arial,sans-serif}
.stApp{background:linear-gradient(135deg,#FBF8F3 0%,#F7FAFC 45%,#F5F8FC 100%)}
.block-container{max-width:1180px;padding-top:1.2rem;padding-bottom:3rem}
.hero{padding:.95rem 1.15rem;margin-bottom:.7rem;border:1px solid #D5E1EA;border-radius:14px;background:linear-gradient(110deg,#FFF8EC,#EDF6FC);box-shadow:0 5px 18px rgba(25,60,86,.07)}
.hero-title{font-size:1.72rem;font-weight:800;color:#1E3040}.hero-sub{font-size:.82rem;color:#586B79}
.step-pill{padding:.55rem;text-align:center;font-size:.82rem;font-weight:700;border:1px solid #D5DEE6;background:#F1F4F7;color:#667683}
.step-pill.active{background:#DCEBF6;color:#194762}.step-pill.done{background:#E8F4EE;color:#2E6650}
.section-head{margin:1rem 0 .7rem;padding:.8rem 1rem;border:1px solid #D8E2E9;border-radius:11px}.blue{background:#EAF4FB}.mint{background:#EAF7F1}.peach{background:#FFF2E7}
div[data-testid="stVerticalBlockBorderWrapper"]{border-color:#D7E1E8!important;border-radius:12px!important;background:rgba(255,255,255,.88)}
div[data-testid="stFileUploaderDropzone"]{padding:.55rem .7rem!important;min-height:0!important;background:#F3F7FA}
.stButton button{border-radius:9px;font-weight:700}
.summary-bar{padding:.7rem;border:1px solid #D9E3EA;border-radius:10px;background:white;color:#516472}
</style>
"""


def apply_styles():
    st.markdown(CSS, unsafe_allow_html=True)


def hero(title: str, subtitle: str):
    st.markdown(f'<div class="hero"><div class="hero-title">{title}</div><div class="hero-sub">{subtitle}</div></div>', unsafe_allow_html=True)


# ---- In-memory upload cache (never written to disk) ----

def cache_upload(cache_key: str, uploaded_file) -> MemoryUpload | None:
    cache = st.session_state.setdefault("upload_cache", {})
    if uploaded_file is not None:
        cache[cache_key] = validate_upload(uploaded_file.name, uploaded_file.getvalue())
    return cache.get(cache_key)


def cached_upload(cache_key: str) -> MemoryUpload | None:
    return st.session_state.get("upload_cache", {}).get(cache_key)


def clear_uploaded_data() -> None:
    """Forget every upload, parsed table and result held for this browser session."""
    for key in ("upload_cache", "run_config", "audit_results", "employee_details_data", "hr_result"):
        st.session_state.pop(key, None)
    st.cache_data.clear()


def upload_widget(label: str, cache_key: str, widget_key: str, types=("csv", "xlsx", "xlsm", "xls")):
    widget = st.file_uploader(label, type=list(types), key=widget_key, label_visibility="collapsed")
    try:
        saved = cache_upload(cache_key, widget)
    except InputFileError as exc:
        st.error(str(exc))
        return None
    if saved:
        st.success(f"Ready · {saved.name}")
    return saved


# Parsed reference tables are cached in memory for an hour at most.
@st.cache_data(show_spinner=False, max_entries=6, ttl=3600)
def schedule_audit_table(name: str, data: bytes):
    return prepare_schedule_audit(MemoryUpload(name, data))


@st.cache_data(show_spinner=False, max_entries=6, ttl=3600)
def supervisor_table(name: str, data: bytes):
    return prepare_supervisor_portfolio(MemoryUpload(name, data))


def counts_row(items: list[tuple[str, int | str]]):
    cols = st.columns(len(items))
    for col, (label, value) in zip(cols, items):
        col.metric(label, f"{value:,}" if isinstance(value, int) else value)
