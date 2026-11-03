"""Run and reconciliation reports: Excel (as before), CSV, JSON and a short text summary."""
from __future__ import annotations

import csv
from collections import Counter
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

from .files import atomic_json

STATUS_COLORS = {"exported": "C6EFCE", "duplicate-skipped": "DDEBF7", "would-export": "FFF2CC", "failed": "FFC7CE",
                 "unreadable": "F4B183", "stopped-unprocessed": "E7E6E6", "adopted": "E2EFDA", "missing-from-db": "FCE4D6"}
DETAIL_HEADERS = ["Run ID", "Timestamp", "Status", "Source Folder", "Folder Relative Path", "Item Index", "Entry Hash", "Outlook Class",
                  "Item Type", "Message Class", "Subject/Name", "Item Date", "Destination", "File Size Bytes", "Duplicate Reason",
                  "Error Type", "HRESULT", "Error Message", "Stage", "Error Category", "Retryable", "Sender", "Internet Message ID"]
SUMMARY_FIELDS = [("Run ID", "run_id"), ("Mode", "mode"), ("Source", "source"), ("Output root", "output_root"), ("Started", "started"),
                  ("Finished", "finished"), ("Result", "result"), ("Overall status", "overall_status"), ("Folders", "folders"),
                  ("Total discovered", "items_seen"), ("Readable", "items_readable"), ("Eligible", "eligible"), ("Exported", "exported"),
                  ("Adopted (file existed)", "adopted"), ("Would export", "proposed"), ("Already exported", "duplicates"),
                  ("Already exported + file verified", "duplicates_verified"), ("Already exported, file missing", "duplicates_missing_file"),
                  ("Missing from database", "missing_from_db"), ("Skipped (not selected)", "not_selected"),
                  ("Skipped (item limit)", "skipped_limit"), ("Failed", "failed"), ("Unreadable", "unreadable"),
                  ("Stopped/unprocessed", "stopped_unprocessed"), ("Retry candidates", "retry_candidates"), ("Balanced", "balanced")]


def _write_csv(path: Path, headers: list[str], rows: list[dict]) -> None:
    tmp = path.with_suffix(".csv.tmp")
    with open(tmp, "w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.DictWriter(handle, fieldnames=headers, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)
    tmp.replace(path)


def write_excel_report(path: Path, summary: dict, rows: list[dict], class_counts: Counter) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    detail, classes = wb.create_sheet("Item Detail"), wb.create_sheet("Item Types")
    navy, white, teal = "1F4E78", "FFFFFF", "DDEBF7"
    ws["A1"] = "Agency Outlook Export Report"
    ws["A1"].font, ws["A1"].fill = Font(size=18, bold=True, color=white), PatternFill("solid", fgColor=navy)
    ws.merge_cells("A1:D1")
    ws.append([])
    for label, key in SUMMARY_FIELDS:
        value = summary.get(key, "")
        if key == "mode" and summary.get("dry_run"):
            value = f"{value} (dry run)"
        ws.append([label, value])
    for cell in ws[2]:
        cell.fill = PatternFill("solid", fgColor=teal)
    ws.column_dimensions["A"].width, ws.column_dimensions["B"].width = 34, 90
    detail.append(DETAIL_HEADERS)
    for row in rows:
        detail.append([row.get(h, "") for h in DETAIL_HEADERS])
    for c in detail[1]:
        c.font, c.fill, c.alignment = Font(bold=True, color=white), PatternFill("solid", fgColor=navy), Alignment(wrap_text=True)
    if detail.max_row > 1:
        table = Table(displayName="ItemAudit", ref=f"A1:{get_column_letter(len(DETAIL_HEADERS))}{detail.max_row}")
        table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True, showColumnStripes=False)
        detail.add_table(table)
        for r in range(2, detail.max_row + 1):
            color = STATUS_COLORS.get(detail.cell(r, 3).value)
            if color:
                detail.cell(r, 3).fill = PatternFill("solid", fgColor=color)
    widths = [20, 20, 22, 48, 36, 12, 16, 14, 28, 34, 55, 22, 80, 16, 35, 22, 18, 80, 14, 18, 10, 30, 40]
    for i, w in enumerate(widths, 1):
        detail.column_dimensions[get_column_letter(i)].width = w
    detail.freeze_panes = "A2"
    classes.append(["Outlook Class", "Item Type", "Count"])
    for key, count in sorted(class_counts.items(), key=lambda x: (x[0][0], x[0][1])):
        classes.append([key[0], key[1], count])
    for c in classes[1]:
        c.font, c.fill = Font(bold=True, color=white), PatternFill("solid", fgColor=navy)
    tmp = path.with_suffix(".xlsx.tmp")
    wb.save(tmp)
    tmp.replace(path)


def write_run_reports(out_dir: Path, result) -> dict:
    """Excel (existing format, extended columns) + failures CSV + JSON summary."""
    out_dir.mkdir(parents=True, exist_ok=True)
    run_id = result.summary["run_id"]
    paths = {"excel_report": out_dir / f"audit_{run_id}.xlsx", "failures_csv": out_dir / f"failures_{run_id}.csv",
             "retry_csv": out_dir / f"retry_candidates_{run_id}.csv", "summary_json": out_dir / f"summary_{run_id}.json"}
    write_excel_report(paths["excel_report"], result.summary, result.rows, result.class_counts)
    _write_csv(paths["failures_csv"], DETAIL_HEADERS, result.failures)
    _write_csv(paths["retry_csv"], ["store_id", "entry_id", "reason", "path", "folder"], result.retry_candidates)
    atomic_json(paths["summary_json"], {**result.summary, "reports": {k: str(v) for k, v in paths.items()}})
    return {k: str(v) for k, v in paths.items()}


def reconciliation_text(rec) -> str:
    c = rec.counts
    lines = [f"Reconciliation {rec.generated}", f"Output root: {rec.output_root}", f"Overall status: {rec.overall_status.upper()}", "",
             f"Database records: {c.get('db_records', 0):,} (completed {c.get('db_completed', 0):,}, failed {c.get('db_failed', 0):,})",
             f"Files found: {c.get('files_found', 0):,}  verified: {c.get('files_verified', 0):,}  partial: {c.get('partial_files', 0):,}"]
    if "source_items" in c:
        lines.append(f"Outlook items: {c['source_items']:,}  unreadable: {c['source_unreadable']:,}  never exported: {c['source_missing_db']:,}"
                     f"  exported but no longer in source: {c['db_not_in_source']:,}")
    lines += [f"Safe to retry - failed: {c.get('retry_failed', 0):,}  missing: {c.get('retry_missing', 0):,}", "", "Issues by check:"]
    lines += [f"  {k}: {v:,}" for k, v in rec.by_check().items()] or ["  none"]
    return "\n".join(lines)


def write_reconciliation(out_dir: Path, rec, run_id: str) -> dict:
    out_dir.mkdir(parents=True, exist_ok=True)
    paths = {"json": out_dir / f"reconciliation_{run_id}.json", "csv": out_dir / f"reconciliation_{run_id}.csv",
             "txt": out_dir / f"reconciliation_{run_id}.txt"}
    atomic_json(paths["json"], {"generated": rec.generated, "output_root": rec.output_root, "overall_status": rec.overall_status,
                                "counts": rec.counts, "by_check": rec.by_check(), "issues": [i.as_dict() for i in rec.issues]})
    headers = list(rec.issues[0].as_dict()) if rec.issues else ["check", "severity", "detail"]
    _write_csv(paths["csv"], headers, [i.as_dict() for i in rec.issues])
    tmp = paths["txt"].with_suffix(".txt.tmp")
    tmp.write_text(reconciliation_text(rec), encoding="utf-8")
    tmp.replace(paths["txt"])
    return {k: str(v) for k, v in paths.items()}
