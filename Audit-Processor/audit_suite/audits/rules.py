"""Audit rules. Each takes the located source table and returns finding rows,
with an ``Audit Finding`` column first. Rules never modify the input frame."""
from __future__ import annotations

from datetime import date, timedelta

import pandas as pd

from ..common.text import clean_text, contains, employee_key, sort_if_present, time_to_minutes


def audit_43_843(frame, **_):
    grtw, code43 = contains(frame["Work Rule"], "GRTW"), contains(frame["Pay Code Name"], "43")
    ids = frame["EE ID"].map(employee_key)
    matched_ids = set(ids[grtw]) & set(ids[code43])
    evidence = ids.isin(matched_ids) & (grtw | code43)
    result = frame.loc[evidence].copy()
    result.insert(0, "Audit Finding", "")
    result.loc[grtw[evidence].values, "Audit Finding"] = "GRTW work rule evidence"
    for idx in result.index[code43[evidence].values]:
        result.at[idx, "Audit Finding"] = " | ".join(filter(None, [result.at[idx, "Audit Finding"], "Pay code containing 43 evidence"]))
    return sort_if_present(result, ("EE ID", "Date", "Start Date Time")).reset_index(drop=True)


def audit_grtw_mismatch(frame, **_):
    wr, lc = contains(frame["Work Rule"], "GRTW"), contains(frame["Labour Category"], "GRTW")
    mismatch = wr ^ lc
    result = frame.loc[mismatch].copy()
    finding = pd.Series(index=frame.index, dtype=object)
    finding.loc[wr & ~lc] = "GRTW Work Rule, but Labour Category is not GRTW"
    finding.loc[lc & ~wr] = "GRTW Labour Category, but Work Rule is not GRTW"
    result.insert(0, "Audit Finding", finding.loc[mismatch].values)
    return sort_if_present(result, ("Employee ID", "Start Date Time")).reset_index(drop=True)


def audit_ot_swapped_lct(frame, **_):
    wr = frame["Work Rule"].fillna("").astype(str)
    work_rule_match = wr.str.contains("OT", case=False, regex=False) | wr.str.contains("Day Off", case=False, regex=False)
    lct_match = contains(frame["Labour Category"], "5.2")
    result = frame.loc[work_rule_match & lct_match].copy()
    result.insert(0, "Audit Finding", "OT/Day Off Work Rule with Labour Category containing 5.2")
    return sort_if_present(result, ("Employee ID", "Start Date Time")).reset_index(drop=True)


def audit_unprocessed_tors(frame, cutoff_date=None, today=None, **_):
    today = today or date.today()
    cutoff_date = cutoff_date or today - timedelta(days=3)
    created = pd.to_datetime(frame["Creation Date"].map(clean_text), errors="coerce", dayfirst=True, format="mixed")
    matched = frame["Status"].fillna("").astype(str).str.strip().str.casefold().isin({"pending", "submitted"}) & (created.dt.date < cutoff_date)
    result = frame.loc[matched].copy()
    result.insert(0, "Audit Finding", f"Unprocessed TOR created before {cutoff_date.strftime('%d %b %Y')}")
    if not result.empty:
        result.insert(1, "Days Since Creation", (pd.Timestamp(today) - created.loc[matched].dt.normalize()).dt.days.values)
        result["_sort"] = created.loc[matched].values
        result = result.sort_values("_sort", kind="stable").drop(columns="_sort")
    result = result.reset_index(drop=True)
    unparsed = int((created.isna() & frame["Creation Date"].map(clean_text).ne("")).sum())
    if unparsed:
        result.attrs["summary_messages"] = [f"{unparsed} row(s) had an unreadable Creation Date and were not evaluated."]
    return result


def audit_casuals_loa(frame, **_):
    casual = frame["Worker Type"].fillna("").astype(str).str.contains(r"\bCasual\b", case=False, regex=True)
    paycode = frame["Pay Code Name"].fillna("").astype(str).str.contains(r"(?<!\d)1\d{2}\.1(?!\d)", regex=True)
    result = frame.loc[casual & paycode].copy()
    result.insert(0, "Audit Finding", "Casual worker with LOA 1##.1 pay code")
    return sort_if_present(result, ("EE ID", "Pay Code Name")).reset_index(drop=True)


def audit_negative_accrual(frame, **_):
    frame = frame.copy().map(lambda v: v.replace("=", "") if isinstance(v, str) else v)
    minutes = frame["Available Balance"].map(time_to_minutes)
    matched = frame["Accrual Code"].fillna("").astype(str).str.strip().str.casefold().ne("sick") & minutes.fillna(0).lt(0)
    result = frame.loc[matched].copy()
    result.insert(0, "Audit Finding", "Non-Sick accrual code with negative available balance")
    result.insert(1, "Negative Balance (Minutes)", minutes.loc[matched].round().astype(int).values)
    if not result.empty:
        result["_sort"] = minutes.loc[matched].values
        result = result.sort_values(["_sort"] + [c for c in ("EE ID", "Accrual Code") if c in result], kind="stable").drop(columns="_sort")
    return result.reset_index(drop=True)


def _empty_nba(frame):
    result = frame.iloc[0:0].copy()
    result.insert(0, "Audit Finding", pd.Series(dtype=object))
    result.insert(1, "Employee 521 Total Hours", pd.Series(dtype=float))
    result.insert(2, "Employee 522 Total Hours", pd.Series(dtype=float))
    return result


def audit_nba_oc_72(frame, **_):
    """Evaluate the two NBA codes together by employee and include both as evidence."""
    work = frame.copy()
    work["_employee"] = work["EE ID"].map(employee_key)
    work["_minutes"] = work["Actual Total Hrs (Include Corrections)"].map(time_to_minutes)
    code = work["Pay Code Name"].fillna("").astype(str).str.strip().str.casefold()
    work["_is_521"] = code.eq("521-nba oc less 72")
    work["_is_522"] = code.eq("522-nba oc grtr 72")
    target = work["_is_521"] | work["_is_522"]
    valid = work[target & work["_minutes"].notna()].copy()
    if valid.empty:
        return _empty_nba(frame)

    totals_521 = valid.loc[valid["_is_521"]].groupby("_employee")["_minutes"].sum()
    totals_522 = valid.loc[valid["_is_522"]].groupby("_employee")["_minutes"].sum()
    findings = {}
    for employee in set(valid["_employee"]):
        less_minutes = float(totals_521.get(employee, 0.0))
        greater_minutes = float(totals_522.get(employee, 0.0))
        issues = []
        if less_minutes > 72 * 60:
            issues.append("521-NBA OC Less 72 total exceeds 72 hours")
        if greater_minutes > 0 and less_minutes < 72 * 60:
            issues.append("522-NBA OC Grtr 72 used before 72 hours were recorded to 521")
        if issues:
            findings[employee] = " | ".join(issues)

    evidence = valid[valid["_employee"].isin(findings)].copy()
    if evidence.empty:
        return _empty_nba(frame)
    result = frame.loc[evidence.index].copy()
    result.insert(0, "Audit Finding", evidence["_employee"].map(findings).values)
    result.insert(1, "Employee 521 Total Hours", evidence["_employee"].map(lambda e: round(float(totals_521.get(e, 0)) / 60, 2)).values)
    result.insert(2, "Employee 522 Total Hours", evidence["_employee"].map(lambda e: round(float(totals_522.get(e, 0)) / 60, 2)).values)
    result = sort_if_present(result, ("EE ID", "Pay Code Name")).reset_index(drop=True)
    skipped = int((target & work["_minutes"].isna()).sum())
    if skipped:
        result.attrs["summary_messages"] = [f"{skipped} 521/522 row(s) had unreadable hours and were excluded from totals."]
    return result
