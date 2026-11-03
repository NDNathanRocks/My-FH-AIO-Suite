"""Static configuration shared by the audit tools."""
from __future__ import annotations

import re

APP_TITLE = "Audit Prepper"
DATE_FMT = "%Y-%m-%d"
HELP_EMAIL = "mySchedulehelpline@fraserhealth.ca"
MAX_UPLOAD_BYTES = 200 * 1024 * 1024
UPLOAD_TYPES = ("csv", "xlsx", "xlsm", "xls")
EMAIL_RE = re.compile(r"[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}", re.IGNORECASE)

# Intentionally a soft map: if a coordinator is absent, Outlook receives the cleaned name.
COORDINATOR_EMAILS = {
    "patrice fugah": "patrice.fugah2@fraserhealth.ca",
    "rachelle tambeau": "rachelle.tambeau@fraserhealth.ca",
    "heather whitworth": "Heather.Whitworth@fraserhealth.ca",
    "sandy ardawa": "sandy.ardawa@fraserhealth.ca",
    "rashika nand": "Rashika.Nand@fraserhealth.ca",
    "praveena gopal": "Praveena.Gopal@fraserhealth.ca",
    "ekjot mahli": "ekjot.malhi@fraserhealth.ca",
    "ekjot malhi": "ekjot.malhi@fraserhealth.ca",
}

SCHEDULE_AUDIT_URL = "https://fraserhealthaut.prd.mykronos.com/ia#/dataView/1708"
SUPERVISOR_PORTFOLIO_URL = "https://app.powerbi.com/groups/me/reports/a99ac431-f485-4899-8a9c-9821fe7951a3/227eea054fdff4092823?ctid=31f660a5-192a-4db3-92ba-ca424f1b259e&experience=power-bi&clientSideAuth=0"
EMPLOYEE_DETAILS_URL = "https://app.powerbi.com/groups/me/reports/b68931ac-cc8c-4093-a0a5-04db175577b8/48704422ffb5fa4db8e2?ctid=31f660a5-192a-4db3-92ba-ca424f1b259e&experience=power-bi"

SCHEDULE_AUDIT_COLUMNS = ("EE ID", "Shift Date", "Shift Start Time", "Scheduled Transfer Location",
                          "Primary location (path)", "Schedule Audit Edit Date", "Schedule Audit Edit Time")
SUPERVISOR_COLUMNS = ("Coordinator Name", "Portfolio Description", "Portfolio Email", "WFM Dept Path")
EMPLOYEE_DETAILS_EXPECTED = ("EmployeeNumber", "Email", "Employee Name")
EMPLOYEE_DETAILS_REQUIRED = ("EmployeeNumber", "Email")
TIME_SUMMARY_EXPECTED = ("EE ID", "Last Name", "First Name", "Actual Total Hrs (Include Corrections)",
                         "Pay Code Name", "Worker Type", "Labour Agreement")
TIME_SUMMARY_REQUIRED = ("EE ID", "Actual Total Hrs (Include Corrections)", "Pay Code Name")

DATE_COLUMN_ALIASES = ("Start Date Time", "Shift Date Time", "Start Date", "Shift Date",
                       "Date", "Work Date", "Workdate")
EMPLOYEE_ID_ALIASES = ("EE ID", "Employee ID", "EmployeeNumber", "Employee Number")
NO_DATE_OPTION = "No date column - use EE ID-only fallback"

# Outlook draft attachments are written here and deleted as soon as Outlook has
# copied them into the draft. Anything older than this is purged on next use.
TEMP_SUBDIR = "AuditPrepper"
TEMP_MAX_AGE_SECONDS = 24 * 3600
