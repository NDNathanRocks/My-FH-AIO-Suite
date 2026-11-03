import streamlit as st

from .. import __version__
from .common import apply_styles, clear_uploaded_data, hero


def page():
    apply_styles()
    hero("⚡ Audit Prepper", f"Excel audit and HR data-processing tools · v{__version__}")
    left, right = st.columns(2, gap="medium")
    with left.container(border=True):
        st.markdown("### Kronos audits")
        st.caption("GRTW, OT/LCT, TORs, casual LOA, negative accruals, NBA on-call, CTO/OTO/STO cascade, and PC Portfolio enrichment. Optional Outlook drafts.")
        st.page_link(st.session_state["_pages"]["audits"], label="Open audits →", icon="📋")
    with right.container(border=True):
        st.markdown("### HR Data Toolbox")
        st.caption("Enrich a working Excel/CSV file from any number of lookup files, with exact, standardized, or soft matching per lookup.")
        st.page_link(st.session_state["_pages"]["hr"], label="Open HR Data Toolbox →", icon="🧰")
    with st.expander("Data handling"):
        st.write("Uploads are kept in this browser session's memory only and are never written to disk. "
                 "Outlook draft attachments are written to a private temp folder and deleted once Outlook saves the draft. "
                 "Parsed reference tables are cached in memory for up to one hour.")
        if st.button("Clear all uploaded data now"):
            clear_uploaded_data()
            st.success("Uploaded data and results cleared.")
