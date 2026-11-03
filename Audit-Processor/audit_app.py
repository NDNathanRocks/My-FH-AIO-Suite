"""Audit Prepper entry point: ``streamlit run audit_app.py``.

All logic lives in the ``audit_suite`` package; this file only wires pages.
"""
from __future__ import annotations

import logging

import streamlit as st

from audit_suite.config import APP_TITLE
from audit_suite.ui import audits_page, home, hr_page

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(name)s: %(message)s")


def app():
    st.set_page_config(page_title=APP_TITLE, page_icon="⚡", layout="wide")
    pages = {
        "home": st.Page(home.page, title="Home", icon="🏠", url_path="home", default=True),
        "audits": st.Page(audits_page.page, title="Kronos audits", icon="📋", url_path="audits"),
        "hr": st.Page(hr_page.page, title="HR Data Toolbox", icon="🧰", url_path="hr-toolbox"),
    }
    st.session_state["_pages"] = pages
    st.navigation(list(pages.values())).run()


if __name__ == "__main__":
    app()
