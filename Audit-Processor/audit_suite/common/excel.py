"""Workbook writers shared by all tools. Output is built in memory."""
from __future__ import annotations

import io
from datetime import date

import pandas as pd
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

from ..config import DATE_FMT
from .text import clean_text

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def visible_columns(frame: pd.DataFrame) -> pd.DataFrame:
    """Drop internal helper columns (prefixed with ``_``)."""
    return frame.drop(columns=[c for c in frame.columns if str(c).startswith("_")], errors="ignore")


def audit_workbook_bytes(frame: pd.DataFrame, audit_name: str, source_name: str) -> bytes:
    """Single-sheet 'Audit Results' workbook.

    Run notes (fallbacks, duplicates, unmatched counts) are intentionally NOT
    written to the workbook; the UI shows them and offers a separate run report.
    """
    safe_frame = visible_columns(frame).copy()
    for column in safe_frame.columns:
        # Keep text that looks like a formula as text (prevents formula injection).
        safe_frame[column] = safe_frame[column].map(lambda v: "'" + v if isinstance(v, str) and v.startswith("=") else v)
    width_cols = len(safe_frame.columns)
    rows = len(safe_frame)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        safe_frame.to_excel(writer, sheet_name="Audit Results", index=False, startrow=4)
        ws = writer.book["Audit Results"]
        ws.sheet_view.showGridLines = False
        ws.freeze_panes = "A6"
        ws["A1"], ws["A2"], ws["A3"] = audit_name, f"Processed: {date.today():{DATE_FMT}}", f"Source: {source_name} | Findings: {rows:,}"
        ws["A1"].font, ws["A1"].fill = Font(size=16, bold=True, color="FFFFFF"), PatternFill("solid", fgColor="17365D")
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(1, width_cols))
        for cell in ws[5]:
            cell.fill, cell.font = PatternFill("solid", fgColor="1F4E78"), Font(color="FFFFFF", bold=True)
        if rows and width_cols:
            table = Table(displayName="AuditResultsTable", ref=f"A5:{get_column_letter(width_cols)}{5 + rows}")
            table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
            ws.add_table(table)
            if "Audit Finding" in safe_frame.columns:
                col = list(safe_frame.columns).index("Audit Finding") + 1
                for row in range(6, 6 + rows):
                    cell = ws.cell(row, col)
                    if clean_text(cell.value) == "Hours Match":
                        cell.fill, cell.font = PatternFill("solid", fgColor="E2F0D9"), Font(color="375623")
                    else:
                        cell.fill, cell.font = PatternFill("solid", fgColor="FCE4D6"), Font(color="9C0006")
        elif width_cols:
            ws["A6"] = "No findings matched the audit rules."
        thin = Side(style="thin", color="D9E1F2")
        for col_idx, name in enumerate(safe_frame.columns, 1):
            values = [str(name)] + [clean_text(v) for v in safe_frame.iloc[:, col_idx - 1].head(250)]
            width = min(max(max(map(len, values)) + 2, 12), 42)
            ws.column_dimensions[get_column_letter(col_idx)].width = width
            for row in range(6, 6 + rows):
                ws.cell(row, col_idx).alignment = Alignment(vertical="top", wrap_text=width >= 30)
                ws.cell(row, col_idx).border = Border(bottom=thin)
        ws.auto_filter.ref = None
    return output.getvalue()


def styled_sheets_bytes(sheets: dict[str, pd.DataFrame], header_color: str = "0F766E") -> bytes:
    """Plain multi-sheet workbook with bold coloured headers, filters and frozen top row.

    String values starting with '=' are stored as text, never as formulas.
    """
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        for name, frame in sheets.items():
            frame.to_excel(writer, sheet_name=name, index=False)
        for ws in writer.book.worksheets:
            for row in ws.iter_rows(min_row=2):
                for cell in row:
                    if cell.data_type == "f":
                        cell.data_type = "s"
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for cell in ws[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor=header_color)
            for i, col in enumerate(ws.columns, 1):
                width = max((len(str(c.value)) if c.value is not None else 0 for c in list(col)[:250]), default=0) + 2
                ws.column_dimensions[get_column_letter(i)].width = min(max(width, 12), 55)
    return bio.getvalue()
