"""Outlook draft creation (Windows + classic Outlook + pywin32 only).

Attachments are written to a private temp folder, attached, and deleted as
soon as Outlook has saved the draft, so no HR data is left on disk.
"""
from __future__ import annotations

import re
import tempfile
import time
from datetime import date
from pathlib import Path

import pandas as pd

from .common.excel import audit_workbook_bytes
from .common.text import clean_text, natural_duration, safe_filename, unique_casefold
from .config import COORDINATOR_EMAILS, DATE_FMT, EMAIL_RE, HELP_EMAIL, TEMP_MAX_AGE_SECONDS, TEMP_SUBDIR


def extract_emails(value) -> list[str]:
    """Extract real email patterns, ignoring labels and handling space-separated addresses."""
    return [match.group(0) for match in EMAIL_RE.finditer(clean_text(value))]


def coordinator_name(value) -> str:
    text = clean_text(value)
    return text.split("-", 1)[1].strip() if "-" in text else text


def coordinator_recipient(value):
    name = coordinator_name(value)
    key = name.casefold()
    if not name or key in {"vcy", "sustainment team"}:
        return None
    return COORDINATOR_EMAILS.get(key, name)


def email_details(frame):
    """Return (to, cc, attachment rows) following the standard recipient rules."""
    if "Portfolio Email" not in frame or "Coordinator Name" not in frame:
        raise ValueError("The audit result is missing Portfolio Email or Coordinator Name.")
    all_emails, eligible_by_row = [], []
    for value in frame["Portfolio Email"]:
        parsed = [e for e in extract_emails(value) if e.casefold() != HELP_EMAIL.casefold()]
        eligible_by_row.append(bool(parsed))
        all_emails.extend(parsed)
    to_values = unique_casefold(all_emails)
    cc_values = unique_casefold([HELP_EMAIL] + [r for r in frame["Coordinator Name"].map(coordinator_recipient) if r])
    coordinator_clean = frame["Coordinator Name"].map(coordinator_name).str.casefold()
    coordinator_not_excluded = frame["Coordinator Name"].map(clean_text).ne("") & coordinator_clean.ne("sustainment team")
    mask = pd.Series(eligible_by_row, index=frame.index) & coordinator_not_excluded
    return to_values, cc_values, frame.loc[mask].copy().reset_index(drop=True)


def email_body_for(audit, attachment_frame) -> str:
    body = audit.email_body_html or ""
    if audit.key == "negative_accrual" and attachment_frame["Accrual Code"].fillna("").astype(str).str.contains(r"\d", regex=True).any():
        body += "<p>The audit also includes non-cascading pay codes that require review. Please review the coding and complete any required corrections.</p>"
    return body


def temp_dir() -> Path:
    """Private temp folder; stale files from crashed sessions are purged on each use."""
    folder = Path(tempfile.gettempdir()) / TEMP_SUBDIR
    folder.mkdir(exist_ok=True)
    cutoff = time.time() - TEMP_MAX_AGE_SECONDS
    for item in folder.glob("*.xlsx"):
        try:
            if item.stat().st_mtime < cutoff:
                item.unlink()
        except OSError:
            pass
    return folder


def _outlook():
    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise RuntimeError("Outlook draft creation requires Windows, classic Outlook, and pywin32.") from exc
    return pythoncom, win32com.client


def _create_draft(subject, to, cc, body_html, attachment_name, attachment_bytes):
    pythoncom, client = _outlook()
    path = temp_dir() / attachment_name
    path.write_bytes(attachment_bytes)
    pythoncom.CoInitialize()
    try:
        mail = client.Dispatch("Outlook.Application").CreateItem(0)
        mail.Display(False)
        time.sleep(.35)  # let Outlook insert the user's signature
        existing = mail.HTMLBody or "<html><body></body></html>"
        body_match = re.search(r"<body[^>]*>", existing, re.I)
        styled = '<div style="font-family:Arial,sans-serif;font-size:11pt;line-height:1.25;">' + body_html + '</div><div style="height:11pt;line-height:11pt;">&nbsp;</div>'
        mail.HTMLBody = existing[:body_match.end()] + styled + existing[body_match.end():] if body_match else styled + existing
        mail.SentOnBehalfOfName = HELP_EMAIL
        mail.To = to
        mail.CC = cc
        mail.Subject = subject
        mail.Attachments.Add(str(path.resolve()))
        mail.Save()
        mail.GetInspector.Activate()
    finally:
        pythoncom.CoUninitialize()
        try:
            path.unlink()
        except OSError:
            pass


def create_outlook_draft(audit, frame) -> int:
    to_values, cc_values, attachment_frame = email_details(frame)
    if not to_values:
        raise ValueError("No eligible Portfolio Email recipients were found.")
    if attachment_frame.empty:
        raise ValueError("No rows remain for the email attachment after exclusions.")
    name = f"Audit - {safe_filename(audit.name)} {date.today():{DATE_FMT}}.xlsx"
    _create_draft(audit.email_subject, "; ".join(to_values), "; ".join(cc_values), email_body_for(audit, attachment_frame),
                  name, audit_workbook_bytes(attachment_frame, audit.name, audit.source_name))
    return len(attachment_frame)


def leave_employee_email(selected_rows) -> tuple[str, str, list[str]]:
    """Return (subject, html body, portfolio cc emails) for the employee leave draft."""
    leave_types = unique_casefold(selected_rows["Leave Type"])
    subject_type = leave_types[0] if len(leave_types) == 1 else "Leave"
    bullets = "".join(
        f"<li>{row['Leave Type']} Leave on {clean_text(row['Start Date'])}: approximately {natural_duration(row.get('_estimated_minutes', 0))} cascaded to LOA.</li>"
        for _, row in selected_rows.iterrows())
    first = selected_rows.iloc[0]
    employee_name = " ".join(filter(None, [clean_text(first.get("First Name")), clean_text(first.get("Last Name"))]))
    body = (f"<p>Hello {employee_name},</p><p>During a recent audit, we noticed the following leave occurrence(s) may have cascaded to a Leave of Absence (LOA):</p><ul>{bullets}</ul>"
            "<p>If you would prefer your leave not to cascade to LOA in the future, please ensure that you have enough corresponding leave hours banked before submitting the request.</p>"
            "<p>Please note that you cannot bank these hours in the same pay period in which you are using the leave, as the hours are not yet available at the time the leave is processed.</p>"
            "<p>If you have any questions, please feel free to reach out.</p>")
    portfolio_emails = unique_casefold(
        email for value in selected_rows.get("Portfolio Email", pd.Series(dtype=object)) for email in extract_emails(value)
        if email.casefold() != HELP_EMAIL.casefold())
    return f"Action Required - {subject_type} Leave Adjustment", body, portfolio_emails


def create_leave_employee_draft(audit, selected_rows, employee_email) -> int:
    subject, body, portfolio_emails = leave_employee_email(selected_rows)
    name = f"Audit - {safe_filename(audit.name)} selected rows {date.today():{DATE_FMT}}.xlsx"
    _create_draft(subject, clean_text(employee_email), "; ".join(unique_casefold([HELP_EMAIL] + portfolio_emails)), body,
                  name, audit_workbook_bytes(selected_rows, audit.name, audit.source_name))
    return len(selected_rows)
