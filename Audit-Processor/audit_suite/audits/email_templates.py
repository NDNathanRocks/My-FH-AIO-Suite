"""Outlook draft wording per audit (text before the user's normal Outlook signature).

Edit the text here; each blank-line-separated block becomes one paragraph.
Placeholders in {braces} are filled in when the draft is created.

Standard recipient rules (all audits): sent on behalf of the help-line mailbox;
To = unique Portfolio Email values; CC = help-line mailbox + cleaned Coordinator
Name values; blank recipients, the help-line address (from To) and Sustainment
Team are excluded. The attachment is a separate filtered copy; the full audit
workbook is unchanged.
"""
from __future__ import annotations

import html

SUBJECTS = {
    "43_843": "Action Required - GRTW Stat Holiday Audit",
    "grtw_mismatch": "Action Required - GRTW Shifts without a valid GRTW LCT or WR",
    "ot_swapped_lct": "Action Required - OT Work Rule and Swapped LCT",
    "unprocessed_tors": "Action Required - Unprocessed TORs",
    "casuals_loa": "Action Required - Casual EE's with LOA 1##.1 Paycodes",
    "negative_accrual": "Action Required - Negative Accrual Bank Audit",
    "nba_oc_72": "Action Required - NBA On-Call Hours Coding",
}

BODIES = {
    "43_843": """Hello,

We identified the following employee(s) who are currently on a Graduated Return to Work (GRTW) plan and have been coded with Statutory Holiday (SH) paycodes.

Please review and correct these entries as employees, that have not worked a non GRTW (normal) shift before their SH paycode, are not eligible to receive SH pay.

Let me know if there are any further questions?""",

    "grtw_mismatch": """Hello,

We identified the following GRTW shifts that are missing a valid GRTW Work Rule and/or Labor Category Transfer (LCT).

Please review and update these shifts with the appropriate GRTW Work Rule and GRTW Labor Category Transfer to ensure accurate pay and reporting.

Please reply with your findings once the review is complete.""",

    "ot_swapped_lct": """Hello,

We identified swapped shifts that have an OT work rule applied.

Swapped shifts are not considered OT shifts and therefore should not have a swapped LCT attached. The LCT may have been left on the original shift when an employee reported an absence or began a leave. However, it would not apply to the employee who received the shift as a result of the absence or leave.

Please review the identified entries and remove or correct any OT work rules or swapped LCTs that are not applicable and reply with your findings.""",

    # {tor_phrase} -> "603 Time Off Requests (TORs) that have" / "1 Time Off Request (TOR) that has"
    # {days_phrase} -> "three (3) days"
    "unprocessed_tors": """Hello,

We identified {tor_phrase} remained in Submitted/Pending status for more than {days_phrase} and have not yet been actioned.

As a reminder, TORs in a Submitted/Pending status require review and approval/processing as soon as possible to ensure accurate scheduling, payroll processing, and leave management.

Please review and action the above TOR(s). If the request is no longer required or cannot be processed, please follow up with the employee accordingly.

Thank you for your attention to this matter.""",

    "casuals_loa": """Hello,

We identified the following casual employee leave(s) coded to a LOA 1##.1 paycode.

As a reminder, casual employees should be coded using the corresponding .2 paycode, not the .1 paycode.

Please review and update these paycodes to the appropriate casual employee paycode ending in .2. If you’ve used a cascading paycode, please update it to the correct .2 paycode.""",

    "negative_accrual": """Hello,

Please find attached the audit identifying employees with a negative accrual bank balance.

Please review the flagged entries and confirm that the appropriate pay codes were used. Where necessary, please correct or adjust the entries for the entire reporting period.

For negative statutory holiday balances, please review the employee’s statutory holiday coding for the year. If the employee is on leave, please confirm whether any statutory holiday entries need to be removed or adjusted.

Please reply with your findings once the review is complete.""",

    "nba_oc_72": """Hello,

We identified NBA on-call hours that may be coded to the incorrect pay code for the audited calendar month.

The first 72 eligible hours must be coded to <b>521-NBA OC Less 72</b>. Only hours exceeding 72 must be coded to <b>522-NBA OC Grtr 72</b>.

Please review and correct the flagged entries, then reply with your findings.""",
}

# Added to the Negative Accrual email (before the closing line) when the attachment
# contains numbered (non-cascading) accrual codes.
NEGATIVE_ACCRUAL_NON_CASCADING = ("The audit also includes non-cascading pay codes that require review. "
                                  "Please review the coding and complete any required corrections.")

_NUMBER_WORDS = {1: "one", 2: "two", 3: "three", 4: "four", 5: "five", 6: "six", 7: "seven", 8: "eight", 9: "nine", 10: "ten"}


def days_phrase(days: int) -> str:
    word = _NUMBER_WORDS.get(days)
    unit = "day" if days == 1 else "days"
    return f"{word} ({days}) {unit}" if word else f"{days} {unit}"


def tor_phrase(count: int) -> str:
    return f"{count:,} Time Off Request (TOR) that has" if count == 1 else f"{count:,} Time Off Requests (TORs) that have"


def to_html(text: str) -> str:
    """Paragraphs -> <p>; text is escaped except the <b> tags used deliberately above."""
    paragraphs = [p.strip() for p in text.strip().split("\n\n") if p.strip()]
    out = []
    for p in paragraphs:
        safe = html.escape(p, quote=False).replace("&lt;b&gt;", "<strong>").replace("&lt;/b&gt;", "</strong>")
        out.append(f"<p>{safe}</p>")
    return "".join(out)
