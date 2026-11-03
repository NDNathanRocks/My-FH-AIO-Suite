# Audit Prepper

## Run

```bash
python -m pip install -r requirements.txt
streamlit run audit_app.py
```

## Current workflow

Audit Prepper uses a guided three-stage flow and stores uploaded files independently from temporary Streamlit widgets.

Portfolio enrichment no longer uses employee-number matching. For applicable audits it:

1. Matches audit employee ID and date/time against **Audit - Schedule - by Employee**.
2. Selects the latest edited nonblank Scheduled Transfer Location.
3. Removes its final path segment.
4. Matches the cleaned path exactly and case-insensitively to WFM Dept Path in the **Supervisor-Portfolio-Department Report**.
5. Appends Portfolio Description, Portfolio Email, and Coordinator Name.

Every generated workbook includes an **Audit Summary** sheet describing unmatched portfolios, date-only fallbacks, duplicate department-path decisions, and other important reconciliation decisions.

The CTO/OTO/STO cascade audit uses Schedule Summary - By Paycode plus Time Summary by Pay Code, allocates matching paycode hours chronologically, removes explicitly scheduled LOA hours from the LOA estimate, and then distributes remaining estimated cascaded LOA chronologically across leave shortages.

## Performance and portfolio-only identifier

The optional portfolio-only workflow still requires an employee identifier because that value links each uploaded audit row to `EE ID` in **Audit - Schedule - by Employee**. It is not matched directly against the Supervisor-Portfolio report. The location found from Schedule Audit is what ultimately matches `WFM Dept Path`.

Large input files are cached by filename and file bytes. Schedule Audit is indexed by employee/date and Supervisor-Portfolio is indexed by normalized WFM path, avoiding a full scan for each audit row. The selector recognizes `EE ID`, `Employee ID`, `EmployeeNumber`, and `Employee Number`; if none exists, it remains blank rather than selecting an unrelated first column.
