"""Audit definitions: source export, required columns, email text and rule."""
from __future__ import annotations

from dataclasses import dataclass
from typing import Callable, Optional

import pandas as pd

from . import rules
from .email_templates import BODIES, SUBJECTS

KRONOS = "https://fraserhealthaut.prd.mykronos.com/ia#/dataView/"
CPP = "All home locations · Current pay period"


@dataclass(frozen=True)
class AuditDefinition:
    key: str
    name: str
    source_key: str
    source_name: str
    source_url: str
    export_settings: str
    expected_columns: tuple[str, ...]
    required_columns: tuple[str, ...]
    employee_id_column: str
    email_subject: Optional[str]
    email_body_html: Optional[str]
    steps: tuple[str, ...]
    processor: Optional[Callable[..., pd.DataFrame]]
    extra_sources: tuple[str, ...] = ()


AUDITS: dict[str, AuditDefinition] = {a.key: a for a in (
    AuditDefinition(
        "43_843", "43 - 843 coding while on GRTW", "schedule_full_cpp", "Schedule Summary - Full", KRONOS + "658", CPP,
        ("Work Rule", "EE ID", "Pay Code Name", "Date", "Start Date Time"), ("Work Rule", "EE ID", "Pay Code Name"), "EE ID",
        SUBJECTS["43_843"], BODIES["43_843"],
        ("Open Schedule Summary - Full.", "Select all home locations and the current pay period.",
         "Identify EE IDs with GRTW Work Rule evidence and separate 43 pay code evidence.", "Keep evidence rows for employees appearing in both groups."),
        rules.audit_43_843),
    AuditDefinition(
        "grtw_mismatch", "GRTW Shifts without a valid GRTW LCT or WR", "schedule_shift_wr_cpp", "Schedule Summary - Shift with Work Rules", KRONOS + "7913", CPP,
        ("Employee ID", "Work Rule", "Labour Category", "Start Date Time"), ("Employee ID", "Work Rule", "Labour Category"), "Employee ID",
        SUBJECTS["grtw_mismatch"], BODIES["grtw_mismatch"],
        ("Open Schedule Summary - Shift with Work Rules.", "Select all home locations and the current pay period.",
         "Keep rows where only one of Work Rule or Labour Category contains GRTW."),
        rules.audit_grtw_mismatch),
    AuditDefinition(
        "ot_swapped_lct", "OT Work rule and swapped in LCT", "schedule_shift_wr_cpp", "Schedule Summary - Shift with Work Rules", KRONOS + "7913", CPP,
        ("Employee ID", "Work Rule", "Labour Category", "Start Date Time", "End Date Time"), ("Employee ID", "Work Rule", "Labour Category"), "Employee ID",
        SUBJECTS["ot_swapped_lct"], BODIES["ot_swapped_lct"],
        ("Open Schedule Summary - Shift with Work Rules.", "Select all home locations and the current pay period.",
         "Keep rows where Work Rule contains OT or Day Off.", "From those rows, keep Labour Category values containing 5.2."),
        rules.audit_ot_swapped_lct),
    AuditDefinition(
        "unprocessed_tors", "Unprocessed TORs", "tor_details_cpp", "Scheduling - Time Off Requests - Details", KRONOS + "661", CPP,
        ("EE ID", "Status", "Creation Date", "Time-off Request ID"), ("EE ID", "Status", "Creation Date"), "EE ID",
        SUBJECTS["unprocessed_tors"], BODIES["unprocessed_tors"],
        ("Open Scheduling - Time Off Requests - Details.", "Select all home locations and the current pay period.",
         "Choose the cutoff date.", "Keep Pending or Submitted requests created before the cutoff."),
        rules.audit_unprocessed_tors),
    AuditDefinition(
        "casuals_loa", "Casuals with LOA 1##.1 paycodes", "time_summary_current_pay_period", "Time Summary by Pay Code", KRONOS + "3358", CPP,
        ("EE ID", "Pay Code Name", "Worker Type"), ("EE ID", "Pay Code Name", "Worker Type"), "EE ID",
        SUBJECTS["casuals_loa"], BODIES["casuals_loa"],
        ("Open Time Summary by Pay Code.", "Select all home locations and the current pay period.",
         "Filter Worker Type to Casual.", "Keep Pay Code Name values matching 1##.1."),
        rules.audit_casuals_loa),
    AuditDefinition(
        "negative_accrual", "Negative Hours Accrual Bank", "accrual_negative_cpp", "Accrual Details - Negative Balance", KRONOS + "4306", CPP,
        ("EE ID", "Available Balance", "Accrual Code"), ("EE ID", "Available Balance", "Accrual Code"), "EE ID",
        SUBJECTS["negative_accrual"], BODIES["negative_accrual"],
        ("Open Accrual Details - Negative Balance.", "Select all home locations and the current pay period.",
         "Remove equals signs from exported balances.", "Exclude Sick accrual codes.", "Keep negative balances only."),
        rules.audit_negative_accrual),
    AuditDefinition(
        "nba_oc_72", "NBA OC > 72 hours", "time_summary_calendar_month", "Time Summary by Pay Code", KRONOS + "3358", "All home locations · Calendar month",
        ("EE ID", "Actual Total Hrs (Include Corrections)", "Pay Code Name"), ("EE ID", "Actual Total Hrs (Include Corrections)", "Pay Code Name"), "EE ID",
        SUBJECTS["nba_oc_72"], BODIES["nba_oc_72"],
        ("Open Time Summary by Pay Code.", "Select all home locations and the complete calendar month.",
         "Review 521 and 522 totals together by employee.", "Flag 521 totals above 72 hours.",
         "Flag 522 hours when fewer than 72 hours were recorded to 521.", "Include both 521 and 522 rows for every flagged employee as reference."),
        rules.audit_nba_oc_72),
    AuditDefinition(
        "leave_cascade", "CTO, OTO, and STO Leave Cascading to LOA", "schedule_by_paycode_cpp", "Schedule Summary - By Paycode", KRONOS + "4808", CPP,
        ("EE ID", "Last Name", "First Name", "Start Date", "Duration (Hrs)", "Pay Code Name"), ("EE ID", "Start Date", "Duration (Hrs)", "Pay Code Name"), "EE ID",
        "Action Required - Leave Adjustment", "",
        ("Open Schedule Summary - By Paycode and Time Summary by Pay Code.", "Use all home locations and the current pay period for both.",
         "Keep employees with CTO, OTO, or STO Leave and at least one LOA timecard paycode.",
         "Compare combined leave hours by type against the corresponding NCO code.",
         "Allocate matching hours oldest-first, then distribute estimated cascaded LOA chronologically after removing scheduled LOA hours."),
        None, extra_sources=("time_summary_current_pay_period",)),
)}

# Display metadata for sources that are only ever an extra source.
SOURCE_LABELS = {
    "time_summary_current_pay_period": ("Time Summary by Pay Code", KRONOS + "3358", CPP),
}
