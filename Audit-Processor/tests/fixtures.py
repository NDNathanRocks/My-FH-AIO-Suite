"""Synthetic, anonymised fixtures for Audit-Processor tests.

No real employee data: IDs, names and paths are invented. Each builder
returns CSV/XLSX bytes shaped like the real Kronos / Power BI exports,
including the title rows that appear above the header in those exports.
"""
from __future__ import annotations

import csv
import io

import pandas as pd


def csv_bytes(header: list[str], rows: list[list[object]], preamble: int = 2) -> bytes:
    buf = io.StringIO()
    writer = csv.writer(buf)
    for i in range(preamble):
        writer.writerow([f"Report title line {i + 1}"])
    writer.writerow(header)
    writer.writerows(rows)
    return buf.getvalue().encode("utf-8")


def xlsx_bytes(header: list[str], rows: list[list[object]], preamble: int = 2) -> bytes:
    data = [[f"Report title line {i + 1}"] + [None] * (len(header) - 1) for i in range(preamble)]
    data += [header] + [list(r) for r in rows]
    buf = io.BytesIO()
    pd.DataFrame(data).to_excel(buf, header=False, index=False)
    return buf.getvalue()


SCHEDULE_FULL = (
    ["Work Rule", "EE ID", "Pay Code Name", "Date", "Start Date Time", "Extra"],
    [
        ["GRTW Plan", "1001", "Regular", "2026-10-01", "2026-10-01 07:00", "a"],
        ["Day Shift", "1001", "43-Stat Holiday", "2026-10-02", "2026-10-02 07:00", "b"],
        ["GRTW Plan", "1002", "Regular", "2026-10-01", "2026-10-01 07:00", "c"],
        ["Day Shift", "1003", "843-Stat Worked", "2026-10-03", "2026-10-03 07:00", "d"],
        ["Day Shift", "1001.0", "Regular", "2026-10-04", "2026-10-04 07:00", "e"],
        [None, None, None, None, None, None],
    ],
)

SCHEDULE_SHIFT_WR = (
    ["Employee ID", "Work Rule", "Labour Category", "Start Date Time", "End Date Time"],
    [
        ["2001", "GRTW 4h", "Unit A", "2026-10-01 07:00", "2026-10-01 11:00"],
        ["2002", "Regular", "GRTW Transfer", "2026-10-01 07:00", "2026-10-01 15:00"],
        ["2003", "GRTW 4h", "GRTW Transfer", "2026-10-02 07:00", "2026-10-02 11:00"],
        ["2004", "OT 1.5", "5.2 Swap", "2026-10-03 07:00", "2026-10-03 15:00"],
        ["2005", "Day Off Worked", "Dept 5.2", "2026-10-04 07:00", "2026-10-04 15:00"],
        ["2006", "Regular", "5.2 Swap", "2026-10-05 07:00", "2026-10-05 15:00"],
    ],
)

TOR_DETAILS = (
    ["EE ID", "Status", "Creation Date", "Time-off Request ID"],
    [
        ["3001", "Submitted", "01/10/2026", "T1"],
        ["3002", "Pending", "20/10/2026", "T2"],
        ["3003", "Approved", "01/10/2026", "T3"],
        ["3004", " pending ", "02/10/2026", "T4"],
    ],
)

TIME_SUMMARY = (
    ["EE ID", "Last Name", "First Name", "Actual Total Hrs (Include Corrections)", "Pay Code Name", "Worker Type", "Labour Agreement"],
    [
        ["4001", "Alpha", "Ann", "7:30", "101.1-LOA Unpaid", "Casual", "BCNU"],
        ["4002", "Beta", "Bob", "7:30", "101.2-LOA Unpaid", "Casual", "BCNU"],
        ["4003", "Gamma", "Cy", "7:30", "110.1-LOA Unpaid", "Regular Full Time", "HSA"],
        ["4004", "Delta", "Di", "8:00", "1101.1-Other", "Casual", "HSA"],
        ["5001", "Eps", "Ed", "40:00", "521-NBA OC Less 72", "Regular", "BCNU"],
        ["5001", "Eps", "Ed", "10:00", "522-NBA OC Grtr 72", "Regular", "BCNU"],
        ["5002", "Zeta", "Zo", "80:00", "521-NBA OC Less 72", "Regular", "BCNU"],
        ["5003", "Eta", "El", "72:00", "521-NBA OC Less 72", "Regular", "BCNU"],
        ["5003", "Eta", "El", "4:00", "522-NBA OC Grtr 72", "Regular", "BCNU"],
        ["6001", "Theta", "Tia", "7:30", "96.1-NCO CTO", "Regular", "BCNU"],
        ["6001", "Theta", "Tia", "15:00", "105.1-LOA Unpaid", "Regular", "BCNU"],
        ["6002", "Iota", "Ike", "3:00", "98.1-NCO STO", "Regular", "HSA"],
        ["6002", "Iota", "Ike", "4:30", "105.1-LOA Unpaid", "Regular", "HSA"],
    ],
)

ACCRUAL = (
    ["EE ID", "Available Balance", "Accrual Code"],
    [
        ["7001", "=-7:30", "Vacation"],
        ["7002", "-3:15", "Sick"],
        ["7003", "4:00", "Vacation"],
        ["7004", "-0:45", "Stat 2026"],
        ["7005", "", "Vacation"],
    ],
)

SCHEDULE_BY_PAYCODE = (
    ["EE ID", "Last Name", "First Name", "Start Date", "Duration (Hrs)", "Pay Code Name"],
    [
        ["6001", "Theta", "Tia", "01/10/2026", "7:30", "CTO Leave"],
        ["6001", "Theta", "Tia", "02/10/2026", "7:30", "CTO Leave"],
        ["6001", "Theta", "Tia", "03/10/2026", "7:30", "LOA Unpaid"],
        ["6002", "Iota", "Ike", "05/10/2026", "7:30", "STO Leave"],
        ["6003", "Kappa", "Kai", "05/10/2026", "7:30", "OTO Leave"],
    ],
)

SCHEDULE_AUDIT = (
    ["EE ID", "Shift Date", "Shift Start Time", "Scheduled Transfer Location", "Primary location (path)",
     "Schedule Audit Edit Date", "Schedule Audit Edit Time"],
    [
        ["1001", "01/10/2026", "07:00", "FH/Site A/Unit 1/Job", "FH/Site A/Unit 1", "28/09/2026", "09:00"],
        ["1001", "01/10/2026", "07:00", "FH/Site A/Unit 2/Job", "FH/Site A/Unit 1", "29/09/2026", "09:00"],
        ["1003", "03/10/2026", "07:00", "", "FH/Site B/Unit 9", "29/09/2026", "09:00"],
        ["2001", "01/10/2026", "07:00", "FH/Site C/Unit 3/Job", "", "29/09/2026", "10:00"],
        ["6001", "01/10/2026", "07:00", "FH/Site A/Unit 1/Job", "", "29/09/2026", "10:00"],
        ["7001", "15/09/2026", "07:00", "FH/Site Z/Unit 0/Job", "", "29/09/2026", "10:00"],
    ],
)

SUPERVISOR = (
    ["Coordinator Name", "Portfolio Description", "Portfolio Email", "WFM Dept Path"],
    [
        ["VCY - Pat Example", "Portfolio A1", "pa1@example.org", "FH/Site A/Unit 1"],
        ["VCY - Sam Example", "Portfolio A2", "pa2@example.org; mySchedulehelpline@fraserhealth.ca", "fh/site a/unit 2"],
        ["Sustainment Team", "Portfolio A2 dup", "dup@example.org", "FH/Site A/Unit 2"],
        ["VCY - Lee Example", "Portfolio B9", "pb9@example.org", "FH/Site B/Unit 9"],
    ],
)

EXISTING_FILTERED = (
    ["Employee Number", "Shift Date", "Note"],
    [["1001", "01/10/2026 07:00", "x"], ["1003", "03/10/2026", "y"], ["9999", "01/10/2026", "z"], ["", "", "blank id"]],
)

EMPLOYEE_DETAILS = (
    ["EmployeeNumber", "Email", "Employee Name"],
    [["6001", "tia@example.org", "Tia Theta"], ["6002", "", "Ike Iota"]],
)

SOURCES = {
    "schedule_full_cpp": SCHEDULE_FULL,
    "schedule_shift_wr_cpp": SCHEDULE_SHIFT_WR,
    "tor_details_cpp": TOR_DETAILS,
    "time_summary_current_pay_period": TIME_SUMMARY,
    "time_summary_calendar_month": TIME_SUMMARY,
    "accrual_negative_cpp": ACCRUAL,
    "schedule_by_paycode_cpp": SCHEDULE_BY_PAYCODE,
}
