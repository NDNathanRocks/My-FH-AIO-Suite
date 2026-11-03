# generator_page.py

from __future__ import annotations

import os
import csv
import re
import shutil
import tempfile
from typing import Dict, Tuple, Optional, List, Set

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import column_index_from_string
from openpyxl.worksheet.worksheet import Worksheet

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout,
    QToolBar, QApplication,
    QLabel, QLineEdit, QPushButton,
    QComboBox, QCheckBox,
    QProgressBar, QMessageBox,
    QDialog, QTextBrowser, QFileDialog,
    QSizePolicy
)
from PySide6.QtGui import (
    QAction, QPalette, QColor, QDesktopServices, QKeySequence, QShortcut, QGuiApplication
)
from PySide6.QtCore import Qt, QUrl, QTimer, QThread, QObject, Signal

from utils.csv_utils import EditCSVDialog
from utils.paths_config import DEFAULT_RFPS_FOLDER, DEFAULT_INVOICE_TRACKER
from utils.app_utils import (
    get_tracker_path, get_company_csv_path,
    get_template_path, load_config, save_config_merge,
    process_csv_and_fill_pdfs,
    _force_excel_recalc, run_excel_recalc_subprocess, refresh_common_paths, ensure_tracker_closed_interactive,
    save_workbook_safe,
)
from utils.logging_utils import get_logger, new_run_id, log_timing

logger = get_logger(__name__)

# ──────────────────────────────────────────────────────────────────────────────
# Constants
# ──────────────────────────────────────────────────────────────────────────────
EXCLUDED_SHEETS: Tuple[str, ...] = ("Summary", "Vendor Credits")
HEADER_ROW_INDEX: int = 3  # 1‑based header row index in tracker sheets

# ──────────────────────────────────────────────────────────────────────────────
# Helpers (pure / testable)
# ──────────────────────────────────────────────────────────────────────────────

def _header_map(ws: Worksheet, header_row: int = HEADER_ROW_INDEX) -> Dict[str, str]:
    """Map header text → column letter for the specified header row."""
    return {
        (cell.value or "").strip(): get_column_letter(cell.column)
        for cell in ws[header_row]
        if cell.value and str(cell.value).strip()
    }


def _read_payee_info(csv_path: str) -> Dict[str, Dict[str, str]]:
    """Return mapping CompanyName → full CSV row (keys trimmed)."""
    info: Dict[str, Dict[str, str]] = {}
    with open(csv_path, encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            clean = {str(k).strip(): (str(v).strip() if v is not None else "") for k, v in row.items()}
            key = clean.get("CompanyName") or clean.get("Company Name")
            if key:
                info[key] = clean
    return info


def _find_vendor_sheet(
    tracker_path: str,
    vendor_id: str,
    *,
    preferred_sheet: str = "",
    run_id: str = "-",
) -> Optional[str]:
    """Find the first sheet that contains vendor_id in column 'Vendor ID'."""

    def _norm_sheet_key(name: str) -> str:
        return re.sub(r"[^a-z0-9]+", "", str(name or "").lower())

    wb = load_workbook(tracker_path, data_only=True, read_only=True)  # read_only mode
    try:
        vendor_sheets = [s for s in wb.sheetnames if s not in EXCLUDED_SHEETS]
        ordered_sheets = list(vendor_sheets)

        if preferred_sheet:
            pref_norm = _norm_sheet_key(preferred_sheet)
            exact = [s for s in vendor_sheets if _norm_sheet_key(s) == pref_norm]
            contains = [
                s
                for s in vendor_sheets
                if s not in exact and (pref_norm in _norm_sheet_key(s) or _norm_sheet_key(s) in pref_norm)
            ]
            ordered_sheets = exact + contains + [s for s in vendor_sheets if s not in exact and s not in contains]

        for name in ordered_sheets:
            ws = wb[name]
            hdr = _header_map(ws)
            col_vid = hdr.get("Vendor ID")
            if not col_vid:
                logger.debug(
                    "Generator vendor-sheet scan skip | sheet=%s | reason=missing_vendor_id_column",
                    name,
                    extra={"run_id": run_id},
                )
                continue

            try:
                vid_col_idx = column_index_from_string(col_vid)
            except Exception:
                logger.warning(
                    "Generator vendor-sheet scan skip | sheet=%s | reason=invalid_vendor_id_column | column=%s",
                    name,
                    col_vid,
                    extra={"run_id": run_id},
                )
                continue

            scanned_rows = 0
            blank_streak = 0
            for values in ws.iter_rows(
                min_row=HEADER_ROW_INDEX + 1,
                min_col=vid_col_idx,
                max_col=vid_col_idx,
                values_only=True,
            ):
                scanned_rows += 1
                current_vid = str((values[0] if values else "") or "").strip()

                if not current_vid:
                    blank_streak += 1
                    # Avoid scanning huge formatted tails in some sheets.
                    if scanned_rows >= 5000 and blank_streak >= 5000:
                        logger.debug(
                            "Generator vendor-sheet scan early-stop | sheet=%s | reason=blank_tail | scanned_rows=%s",
                            name,
                            scanned_rows,
                            extra={"run_id": run_id},
                        )
                        break
                else:
                    blank_streak = 0

                if current_vid == vendor_id:
                    row_num = HEADER_ROW_INDEX + scanned_rows
                    logger.info(
                        "Generator vendor-sheet matched | vendor_id=%s | sheet=%s | row=%s | scanned_rows=%s",
                        vendor_id,
                        name,
                        row_num,
                        scanned_rows,
                        extra={"run_id": run_id},
                    )
                    return name

        logger.warning(
            "Generator vendor-sheet not found | vendor_id=%s | sheets_scanned=%s | preferred_sheet=%s",
            vendor_id,
            len(ordered_sheets),
            preferred_sheet,
            extra={"run_id": run_id},
        )
        return None
    finally:
        wb.close()


def _dump_sheet_to_temp_csv(tracker_path: str, sheet_name: str, *, run_id: str = "-") -> str:
    """Read the sheet as all-text and write a temp CSV preserving two blank rows + header."""
    # Use read_only mode and ensure cleanup
    # Use ExcelFile as a context manager so the underlying openpyxl ZipFile
    # is explicitly closed before we return (prevents Windows file-lock leaks).
    with pd.ExcelFile(tracker_path, engine="openpyxl") as xf:
        df = pd.read_excel(
            xf,
            sheet_name=sheet_name,
            header=HEADER_ROW_INDEX - 1,  # 0-based header row
            dtype=str,
        )
    logger.info(
        "Generator sheet dumped to CSV source | sheet=%s | rows=%s | cols=%s | headers=%s",
        sheet_name,
        len(df),
        len(df.columns),
        [str(c) for c in df.columns[:8]],
        extra={"run_id": run_id},
    )
    
    tf = tempfile.NamedTemporaryFile(delete=False, suffix=".csv", mode='w', newline='', encoding='utf-8')
    try:
        writer = csv.writer(tf)
        writer.writerow([])
        writer.writerow([])
        writer.writerow(df.columns.tolist())
        for _, row in df.iterrows():
            vals = [(row[col] if pd.notna(row[col]) else "") for col in df.columns]
            writer.writerow(vals)
        csv_path = tf.name
    finally:
        tf.close()
    
    # Explicitly clear DataFrame to release file handles
    del df
    
    return csv_path


def _pick_col_letter_case_insensitive(hdr: Dict[str, str], *aliases: str) -> Optional[str]:
    if not hdr:
        return None
    for a in aliases:
        if a in hdr:
            return hdr[a]
    lower = {str(k).strip().lower(): v for k, v in hdr.items()}
    for a in aliases:
        col = lower.get(str(a).strip().lower())
        if col:
            return col
    return None


def _formula_cache_health(
    tracker_path: str,
    sheet_name: str,
    *,
    only_new: bool = False,
    run_id: str = "-",
) -> Dict[str, object]:
    """Check whether formula cells in PDF-relevant columns have cached values available.

    When only_new is True, check only rows where "RFP Prepared (Y/N)" == "N".
    """
    wb_formula = load_workbook(tracker_path, data_only=False, read_only=True)
    wb_values = load_workbook(tracker_path, data_only=True, read_only=True)
    try:
        if sheet_name not in wb_formula.sheetnames or sheet_name not in wb_values.sheetnames:
            return {
                "sheet": sheet_name,
                "formula_cells": 0,
                "missing_cached": 0,
                "sample_missing": [],
                "columns_checked": [],
            }

        ws_f = wb_formula[sheet_name]
        ws_v = wb_values[sheet_name]
        hdr = _header_map(ws_f)

        inv_col = _pick_col_letter_case_insensitive(hdr, "Invoice #", "Invoice Number", "Invoice#")
        target_cols = {
            "GST Subtotals": _pick_col_letter_case_insensitive(hdr, "GST Subtotals", "GST Subtotal"),
            "Overall Invoice Bill Totals": _pick_col_letter_case_insensitive(hdr, "Overall Invoice Bill Totals", "Overall Invoice Bill Totals "),
            "Services Subtotal": _pick_col_letter_case_insensitive(hdr, "Services Subtotal", "Services Subtotals"),
            "Travel/Accom Reimbursements": _pick_col_letter_case_insensitive(hdr, "Travel/Accom Reimbursements"),
        }
        target_cols = {k: v for k, v in target_cols.items() if v}
        rfp_col = _pick_col_letter_case_insensitive(hdr, "RFP Prepared (Y/N)", "RFP Prepared", "RFP Prepared Y/N")

        formula_cells = 0
        missing_cached = 0
        sample_missing: List[Dict[str, str]] = []
        rows_checked = 0
        rows_skipped = 0

        for r in range(HEADER_ROW_INDEX + 1, ws_f.max_row + 1):
            if only_new and rfp_col:
                rfp_state = str(ws_f[f"{rfp_col}{r}"].value or "").strip().upper()
                if rfp_state != "N":
                    rows_skipped += 1
                    continue

            rows_checked += 1
            inv = str(ws_f[f"{inv_col}{r}"].value or "").strip() if inv_col else ""
            for col_name, col_letter in target_cols.items():
                formula_raw = ws_f[f"{col_letter}{r}"].value
                if isinstance(formula_raw, str) and formula_raw.startswith("="):
                    formula_cells += 1
                    cached_val = ws_v[f"{col_letter}{r}"].value
                    if cached_val is None or str(cached_val).strip() == "":
                        missing_cached += 1
                        if len(sample_missing) < 10:
                            sample_missing.append({
                                "row": str(r),
                                "invoice": inv,
                                "column": col_name,
                                "formula": formula_raw,
                            })

        logger.info(
            "Generator formula cache health | sheet=%s | only_new=%s | rfp_col=%s | rows_checked=%s | rows_skipped=%s | formula_cells=%s | missing_cached=%s | columns_checked=%s",
            sheet_name,
            only_new,
            rfp_col or "",
            rows_checked,
            rows_skipped,
            formula_cells,
            missing_cached,
            list(target_cols.keys()),
            extra={"run_id": run_id},
        )
        if sample_missing:
            logger.warning(
                "Generator formula cache missing sample | sheet=%s | sample=%s",
                sheet_name,
                sample_missing,
                extra={"run_id": run_id},
            )

        return {
            "sheet": sheet_name,
            "only_new": only_new,
            "rows_checked": rows_checked,
            "rows_skipped": rows_skipped,
            "formula_cells": formula_cells,
            "missing_cached": missing_cached,
            "sample_missing": sample_missing,
            "columns_checked": list(target_cols.keys()),
        }
    finally:
        try:
            wb_formula.close()
        except Exception:
            pass
        try:
            wb_values.close()
        except Exception:
            pass

def _get_new_row_numbers(tracker_path: str, sheet_name: str, *, run_id: str = "-") -> List[int]:
    """Return 1-based Excel row numbers where RFP Prepared (Y/N) == 'N'."""
    wb = load_workbook(tracker_path, data_only=True, read_only=True)
    try:
        if sheet_name not in wb.sheetnames:
            return []
        ws = wb[sheet_name]
        hdr = _header_map(ws)
        rfp_col = _pick_col_letter_case_insensitive(hdr, "RFP Prepared (Y/N)", "RFP Prepared", "RFP Prepared Y/N")
        if not rfp_col:
            logger.warning(
                "Generator new-row scan: RFP Prepared column not found | sheet=%s",
                sheet_name,
                extra={"run_id": run_id},
            )
            return []
        rfp_col_idx = column_index_from_string(rfp_col)
        new_rows: List[int] = []
        for row_cells in ws.iter_rows(min_row=HEADER_ROW_INDEX + 1, min_col=rfp_col_idx, max_col=rfp_col_idx):
            cell = row_cells[0]
            if str(cell.value or "").strip().upper() == "N":
                new_rows.append(cell.row)
        logger.info(
            "Generator new-row scan | sheet=%s | new_rows=%s",
            sheet_name,
            new_rows,
            extra={"run_id": run_id},
        )
        return new_rows
    finally:
        wb.close()


def _default_output_for_company(company: str) -> str:
    return os.path.join(DEFAULT_RFPS_FOLDER, company)

# ──────────────────────────────────────────────────────────────────────────────
# Tracker update helpers for "RFP Prepared (Y/N)"
# ──────────────────────────────────────────────────────────────────────────────

def _strict_cell_key(v) -> str:
    """
    Strict key for invoice comparison:
      - preserve case and all characters (spaces, '_' and '-' included)
      - trim only outer whitespace
      - coerce whole-number floats like 6058.0 to '6058'
    """
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v)).strip()
    return str(v).strip()

def _normalize_invoice_strict(s: str) -> str:
    """(Kept name for compatibility) Strict key: trim only; no uppercasing/normalization."""
    return _strict_cell_key(s)

def _scan_generated_invoices_from_folder(out_dir: str) -> Set[str]:
    """
    Recursively scan for generated PDFs and extract invoice numbers.
    Accepts either:
      - '... RFP <INV> - ... .pdf'   (captures <INV> up to ' - ')
      - '<INV> - ... .pdf'          (captures <INV> up to ' - ')
    <INV> may include letters/digits/spaces/_/-
    """
    invs: Set[str] = set()
    # RFP … <INV> - …
    pat_rfp   = re.compile(r'(?i)\bRFP\b\s+(.+?)\s-\s')
    # <INV> - …
    pat_plain = re.compile(r'^\s*([A-Za-z0-9 _-]+?)\s+-\s+')

    try:
        for root, _, files in os.walk(out_dir):
            for fn in files:
                if not fn.lower().endswith(".pdf"):
                    continue
                base = os.path.splitext(fn)[0]
                m = pat_rfp.search(base) or pat_plain.search(base)
                if m:
                    invs.add(_normalize_invoice_strict(m.group(1)))
    except Exception:
        pass

    return invs

def _update_tracker_rfp_prepared(tracker_path: str, sheet_name: str, invoice_numbers: Set[str], skip_update: bool = False, *, run_id: str = "-") -> Tuple[int, int, List[str], Optional[pd.DataFrame]]:
    """
    Update tracker with RFP Prepared = Y for given invoices.
    
    Args:
        tracker_path: Path to tracker Excel
        sheet_name: Sheet name to update
        invoice_numbers: Set of invoice numbers that were generated
        skip_update: If True, don't write to tracker; return DataFrame instead
        
    Returns:
        (updated_count, total_count, not_found_list, skipped_dataframe)
    """
    if not invoice_numbers:
        logger.info("Generator tracker update skipped | reason=no_generated_invoices", extra={"run_id": run_id})
        return 0, 0, [], None

    # Use read_only+data_only for preview; full open for write (data_only=True on a saved
    # workbook strips Table metadata and causes Excel to repair table1.xml on next open).
    wb = load_workbook(tracker_path, read_only=skip_update, data_only=skip_update)
    try:
        if sheet_name not in wb.sheetnames:
            logger.warning(
                "Generator tracker update failed | reason=sheet_not_found | sheet=%s",
                sheet_name,
                extra={"run_id": run_id},
            )
            return 0, len(invoice_numbers), sorted(list(invoice_numbers)), None

        ws = wb[sheet_name]

        hdr_map = _find_col_letter_from_row(ws[HEADER_ROW_INDEX])
        inv_col = _get_hdr_col(hdr_map, "Invoice #", "Invoice#", "Invoice Number")
        rfp_col = _get_hdr_col(hdr_map, "RFP Prepared (Y/N)", "RFP Prepared", "RFP Prepared Y/N")
        if not inv_col or not rfp_col:
            logger.warning(
                "Generator tracker update failed | reason=missing_required_columns | invoice_col=%s | rfp_col=%s",
                inv_col,
                rfp_col,
                extra={"run_id": run_id},
            )
            return 0, len(invoice_numbers), sorted(list(invoice_numbers)), None

        # Build strict index (trim only; preserve case/symbols; coerce 6058.0 -> '6058')
        index: Dict[str, int] = {}
        for r in range(HEADER_ROW_INDEX + 1, ws.max_row + 1):
            key = _normalize_invoice_strict(ws[f"{inv_col}{r}"].value)
            if key:
                index[key] = r

        updated = 0
        not_found: List[str] = []
        skipped_rows: List[Dict[str, any]] = []
        
        for inv in sorted(invoice_numbers):
            row = index.get(_normalize_invoice_strict(inv))
            if row:
                try:
                    if skip_update:
                        # Collect row data for export
                        row_data = {}
                        # Read-only rows contain EmptyCell objects without .column; use position instead.
                        for col_idx, cell in enumerate(ws[row], start=1):
                            col_letter = get_column_letter(getattr(cell, "column", None) or col_idx)
                            # Find header for this column
                            header_cell = ws[f"{col_letter}{HEADER_ROW_INDEX}"]
                            header = str(header_cell.value or "").strip()
                            if header:
                                row_data[header] = cell.value
                        skipped_rows.append(row_data)
                    else:
                        ws[f"{rfp_col}{row}"].value = "Y"
                    updated += 1
                except Exception:
                    not_found.append(inv)
            else:
                not_found.append(inv)

        logger.info(
            "Generator tracker update result | sheet=%s | requested=%s | updated=%s | not_found=%s | skip_update=%s",
            sheet_name,
            len(invoice_numbers),
            updated,
            len(not_found),
            skip_update,
            extra={"run_id": run_id},
        )

        # Only save if not skipping update
        if not skip_update:
            save_workbook_safe(wb, tracker_path)  # saves, closes, gc, then os.replace
            wb = None  # already closed inside save_workbook_safe
            
        # Convert skipped rows to DataFrame
        skipped_df = pd.DataFrame(skipped_rows) if skipped_rows else None
        
        return updated, len(invoice_numbers), not_found, skipped_df
    finally:
        if wb is not None:
            try:
                wb.close()
            except Exception:
                pass

def _norm_hdr_key(s: str) -> str:
    """Lowercase header key, remove whitespace/punct except '#', to match variants."""
    return re.sub(r'[^a-z0-9#]+', '', str(s).lower())

def _find_col_letter_from_row(row_cells) -> Dict[str, str]:
    """Return normalized header -> column letter for the given header row."""
    mapping: Dict[str, str] = {}
    for cell in row_cells:
        if cell.value and str(cell.value).strip():
            mapping[_norm_hdr_key(cell.value)] = get_column_letter(cell.column)
    return mapping

def _get_hdr_col(hmap: Dict[str, str], *aliases: str) -> Optional[str]:
    """Try multiple alias strings for a header; return the first matching column letter."""
    for a in aliases:
        col = hmap.get(_norm_hdr_key(a))
        if col:
            return col
    return None

# ──────────────────────────────────────────────────────────────────────────────
# Worker Thread for Generation (prevents UI freeze on large sheets)
# ──────────────────────────────────────────────────────────────────────────────

class GenerationWorker(QObject):
    """Worker for RFP generation. Runs in background thread to prevent UI freeze."""
    progress = Signal(int, int)  # (current_index, total)
    status = Signal(str)         # status message
    error = Signal(str)          # error message
    finished = Signal(dict)      # completion data: {out_dir, generated_count, sheet, generated_invoices}

    def __init__(self, tracker_path: str, template_path: str, company: str, selected_vid: str,
                 payee_info: dict, output_dir: str, only_new: bool, skip_update: bool, run_id: str):
        super().__init__()
        self.tracker_path = tracker_path
        self.template_path = template_path
        self.company = company
        self.selected_vid = selected_vid
        self.payee_info = payee_info
        self.output_dir = output_dir
        self.only_new = only_new
        self.skip_update = skip_update
        self.run_id = run_id
        self.invoice_csv = None
        self.recalc_copy_path = None

    def run(self):
        """Main worker thread logic. Emits signals for progress and completion."""
        tracker_for_read = self.tracker_path
        generated_invoices: Set[str] = set()
        total = 0
        out = None

        try:
            # Step 1: Locate vendor sheet
            self.status.emit("Locating vendor sheet…")
            with log_timing(logger, "Worker locate vendor sheet", run_id=self.run_id):
                sheet = _find_vendor_sheet(
                    self.tracker_path,
                    self.selected_vid,
                    preferred_sheet=self.company,
                    run_id=self.run_id,
                )
            if not sheet:
                self.error.emit(f"No sheet found for VendorID {self.selected_vid}")
                return

            # Step 2: Recalculate formula cells.
            # For new-only generation: always recalc the specific new rows so formula
            # values are guaranteed fresh regardless of cached state. Fast (~5-6 rows).
            # For full generation: check formula cache health first; only recalc if needed.
            if self.only_new:
                self.status.emit("Locating new invoice rows…")
                with log_timing(logger, "Worker locate new rows", run_id=self.run_id):
                    target_rows = _get_new_row_numbers(self.tracker_path, sheet, run_id=self.run_id)

                if target_rows:
                    self.status.emit("Refreshing target rows…")
                    fd, self.recalc_copy_path = tempfile.mkstemp(suffix=".~gen_recalc.xlsx")
                    os.close(fd)
                    os.unlink(self.recalc_copy_path)
                    shutil.copy2(self.tracker_path, self.recalc_copy_path)
                    tracker_for_read = self.recalc_copy_path

                    with log_timing(logger, "Worker Excel row recalc (only_new)", run_id=self.run_id):
                        # Run recalc in a separate process to avoid blocking the GUI
                        run_excel_recalc_subprocess(tracker_for_read, sheet_name=sheet, refresh_external=False,
                                                    row_numbers=target_rows)
                    logger.info(
                        "Worker targeted recalc complete | sheet=%s | rows=%s",
                        sheet, target_rows,
                        extra={"run_id": self.run_id},
                    )

                    with log_timing(logger, "Worker verify cache after targeted recalc", run_id=self.run_id):
                        cache_after_targeted = _formula_cache_health(
                            tracker_for_read,
                            sheet,
                            only_new=True,
                            run_id=self.run_id,
                        )

                    if int(cache_after_targeted.get("missing_cached", 0)) > 0:
                        logger.warning(
                            "Worker targeted recalc incomplete; escalating to sheet recalc | sheet=%s | missing_cached=%s",
                            sheet,
                            cache_after_targeted.get("missing_cached", 0),
                            extra={"run_id": self.run_id},
                        )
                        self.status.emit("Refreshing target sheet formulas…")
                        with log_timing(logger, "Worker Excel sheet recalc fallback (only_new)", run_id=self.run_id):
                            run_excel_recalc_subprocess(tracker_for_read, sheet_name=sheet, refresh_external=False)

                        with log_timing(logger, "Worker verify cache after sheet recalc", run_id=self.run_id):
                            cache_after_sheet = _formula_cache_health(
                                tracker_for_read,
                                sheet,
                                only_new=True,
                                run_id=self.run_id,
                            )

                        if int(cache_after_sheet.get("missing_cached", 0)) > 0:
                            logger.warning(
                                "Worker sheet recalc incomplete; escalating to full refresh | sheet=%s | missing_cached=%s",
                                sheet,
                                cache_after_sheet.get("missing_cached", 0),
                                extra={"run_id": self.run_id},
                            )
                            self.status.emit("Performing full refresh fallback…")
                            with log_timing(logger, "Worker Excel full recalc fallback (only_new)", run_id=self.run_id):
                                run_excel_recalc_subprocess(tracker_for_read, refresh_external=True)
                else:
                    logger.warning(
                        "Worker new-row scan returned none; forcing full sheet recalc fallback | sheet=%s | tracker=%s",
                        sheet,
                        self.tracker_path,
                        extra={"run_id": self.run_id},
                    )
                    fd, self.recalc_copy_path = tempfile.mkstemp(suffix=".~gen_recalc.xlsx")
                    os.close(fd)
                    os.unlink(self.recalc_copy_path)
                    shutil.copy2(self.tracker_path, self.recalc_copy_path)
                    tracker_for_read = self.recalc_copy_path

                    self.status.emit("Refreshing tracker formulas…")
                    with log_timing(logger, "Worker Excel full recalc fallback (no_new_rows)", run_id=self.run_id):
                        run_excel_recalc_subprocess(tracker_for_read, sheet_name=sheet, refresh_external=True)

                    with log_timing(logger, "Worker verify cache after full recalc fallback", run_id=self.run_id):
                        cache_after_fallback = _formula_cache_health(
                            tracker_for_read,
                            sheet,
                            only_new=True,
                            run_id=self.run_id,
                        )

                    logger.info(
                        "Worker fallback recalc complete | sheet=%s | missing_cached=%s",
                        sheet,
                        cache_after_fallback.get("missing_cached", 0),
                        extra={"run_id": self.run_id},
                    )
            else:
                self.status.emit("Checking formula cache…")
                with log_timing(logger, "Worker formula cache check", run_id=self.run_id):
                    cache_health = _formula_cache_health(
                        self.tracker_path,
                        sheet,
                        only_new=False,
                        run_id=self.run_id,
                    )

                if int(cache_health.get("missing_cached", 0)) > 0:
                    logger.warning(
                        "Worker formula cache refresh required | sheet=%s | missing=%s",
                        sheet,
                        cache_health.get("missing_cached", 0),
                        extra={"run_id": self.run_id},
                    )
                    self.status.emit("Refreshing tracker calculations…")
                    fd, self.recalc_copy_path = tempfile.mkstemp(suffix=".~gen_recalc.xlsx")
                    os.close(fd)
                    os.unlink(self.recalc_copy_path)
                    shutil.copy2(self.tracker_path, self.recalc_copy_path)
                    tracker_for_read = self.recalc_copy_path

                    with log_timing(logger, "Worker Excel full recalc", run_id=self.run_id):
                        _force_excel_recalc(tracker_for_read, refresh_external=True)

            # Step 3: Prepare CSV
            self.status.emit("Preparing invoice CSV…")
            with log_timing(logger, "Worker dump sheet to CSV", run_id=self.run_id):
                self.invoice_csv = _dump_sheet_to_temp_csv(tracker_for_read, sheet, run_id=self.run_id)

            # Step 4: Generate PDFs
            self.status.emit("Processing…")
            out = self.output_dir
            os.makedirs(out, exist_ok=True)

            with log_timing(logger, "Worker build PDFs", run_id=self.run_id):
                for tick in process_csv_and_fill_pdfs(
                    self.invoice_csv, self.template_path, out, self.company,
                    self.payee_info,
                    only_new=self.only_new,
                    run_id=self.run_id,
                ):
                    inv = None
                    if isinstance(tick, tuple):
                        if len(tick) >= 2:
                            i, total = tick[0], tick[1]
                            if len(tick) >= 3:
                                inv = tick[2]
                        else:
                            i = tick[0]; total = max(total, i)
                    elif isinstance(tick, dict):
                        i = tick.get("i") or tick.get("index") or 0
                        total = tick.get("total") or total
                        inv = tick.get("invoice") or tick.get("Invoice #")
                    else:
                        try:
                            i = int(tick)
                        except Exception:
                            i = 0

                    self.progress.emit(int(i), int(total) if total else 0)

                    if inv:
                        generated_invoices.add(_normalize_invoice_strict(inv))

            if not generated_invoices:
                generated_invoices = _scan_generated_invoices_from_folder(out)

            # Signal completion
            self.finished.emit({
                "out_dir": out,
                "generated_count": total,
                "sheet": sheet,
                "generated_invoices": generated_invoices,
            })

        except Exception as e:
            logger.exception("Worker error", extra={"run_id": self.run_id})
            self.error.emit(str(e))

        finally:
            # Cleanup temp recalc copy and the temp tracker CSV (contains invoice data)
            for temp_path in (self.recalc_copy_path, self.invoice_csv):
                if temp_path and os.path.exists(temp_path):
                    try:
                        os.unlink(temp_path)
                    except OSError:
                        pass

# ──────────────────────────────────────────────────────────────────────────────
# UI Class
# ──────────────────────────────────────────────────────────────────────────────

class GeneratorPage(QWidget):
    """RFP Generator page.
    Allows selecting a company, generating RFPs from the tracker,
    and updating the tracker with "RFP Prepared (Y/N)".
    """

    def __init__(self, controller):
        super().__init__()
        self.ctrl = controller

        # Runtime state
        self.invoice_csv: str = ""
        self.selected_vid: str = ""
        # Read rfp_count directly from the latest config
        self.rfp_count: int = int(load_config().get("rfp_count", 0))

        # Paths (filled on first show; initialize empty to avoid stale values)
        self.tracker_file: str = ""
        self.company_csv: str = ""
        self.template: str = ""

        # Track last-loaded CSV to reload dropdown only when it changes
        self._last_company_csv_loaded: str = ""

        # Worker thread for generation (prevents UI freeze)
        self._generation_thread: Optional[QThread] = None
        self._generation_worker: Optional[GenerationWorker] = None

        # UI setup
        self._apply_palette()
        self._apply_style()
        self._build_ui()

        # First-time refresh (also happens every time page is shown)
        self._refresh_paths_from_config()

        # ESC → Back
        esc = QShortcut(QKeySequence(Qt.Key_Escape), self)
        esc.activated.connect(self._on_back)

        # Initial load of company list
        if os.path.isfile(self.company_csv):
            try:
                self._load_payee_info_into_ui()
            except Exception:
                QMessageBox.warning(self, "Error", "Invalid Company CSV.")
        self._check_ready()

    # ────────────────────────── UI construction ──────────────────────────
    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(50, 40, 50, 40)
        root.setSpacing(20)

        # Toolbar (Back)
        toolbar = QToolBar()
        back_act = QAction("← Back", self)
        back_act.triggered.connect(self._on_back)
        toolbar.addAction(back_act)
        root.addWidget(toolbar)

        # Title
        title = QLabel("🤖 RFP Generator")
        title.setAlignment(Qt.AlignHCenter)
        title.setStyleSheet("font-size:32px; font-weight:bold; color:#3498db;")
        root.addWidget(title)

        # Counter
        self.lbl_counter = QLabel(f"Total RFPs generated: {self.rfp_count}")
        self.lbl_counter.setAlignment(Qt.AlignHCenter)
        root.addWidget(self.lbl_counter)

        # Settings
        box = QWidget(); bl = QVBoxLayout(box); bl.setSpacing(15)
        # Template
        tpl_row = QHBoxLayout()
        tpl_row.addWidget(QLabel("Template PDF:"))
        self.tpl_le = QLineEdit(os.path.basename(self.template)); self.tpl_le.setReadOnly(True)
        tpl_row.addWidget(self.tpl_le)
        bl.addLayout(tpl_row)

        # Company CSV
        csv_row = QHBoxLayout()
        csv_row.addWidget(QLabel("Company CSV:"))
        self.csv_le = QLineEdit(os.path.basename(self.company_csv)); self.csv_le.setReadOnly(True)
        csv_row.addWidget(self.csv_le)
        btn_edit = QPushButton("Edit…"); btn_edit.clicked.connect(self.on_edit_payee)
        csv_row.addWidget(btn_edit)
        bl.addLayout(csv_row)
        root.addWidget(box)

        # Tracker selector
        track_box = QWidget(); tl = QHBoxLayout(track_box)
        tl.addWidget(QLabel("Tracker Excel:"))
        self.track_le = QLineEdit(self.tracker_file); self.track_le.setReadOnly(True)
        tl.addWidget(self.track_le)
        root.addWidget(track_box)

        # Invoice CSV selector + checkboxes
        inv_box = QWidget(); il = QHBoxLayout(inv_box)
        self.chk_only_new = QCheckBox("Only new RFPs"); self.chk_only_new.setChecked(True)
        il.addWidget(self.chk_only_new)
        self.chk_skip_tracker = QCheckBox("Skip tracker update"); self.chk_skip_tracker.setChecked(False)
        il.addWidget(self.chk_skip_tracker)
        root.addWidget(inv_box)

        # Company dropdown + Generate
        proc_box = QWidget(); pl = QHBoxLayout(proc_box)
        pl.addWidget(QLabel("Agency Name:"))
        self.cb_company = QComboBox(); pl.addWidget(self.cb_company)
        self.btn_generate = QPushButton("Generate RFPs"); self.btn_generate.setEnabled(False)
        self.btn_generate.clicked.connect(self.generate)
        pl.addWidget(self.btn_generate)
        root.addWidget(proc_box)

        # Progress + status
        self.progress = QProgressBar(); root.addWidget(self.progress)
        self.lbl_status = QLabel(""); root.addWidget(self.lbl_status)

    # ────────────────────────── Styling / palette ──────────────────────────
    def _apply_palette(self) -> None:
        pal = QPalette()
        pal.setColor(QPalette.Window, QColor(245, 245, 245))
        pal.setColor(QPalette.WindowText, QColor(33, 33, 33))
        pal.setColor(QPalette.Button, QColor(238, 238, 238))
        pal.setColor(QPalette.ButtonText, QColor(33, 33, 33))
        pal.setColor(QPalette.Highlight, QColor(52, 152, 219))
        pal.setColor(QPalette.HighlightedText, QColor(255, 255, 255))
        self.setPalette(pal)

    def _apply_style(self) -> None:
        self.setStyleSheet(
            """
            QPushButton { background:#2980b9; color:white; border:none; border-radius:6px;
                          padding:8px 16px; font-size:14px; }
            QPushButton:disabled { background:#bdc3c7; color:#7f8c7f; }
            QPushButton:hover:enabled { background:#3498db; }
            QLineEdit, QComboBox { border:1px solid #bdc3c7; border-radius:4px; padding:4px; }
            QCheckBox { spacing:4px; font-size:13px; }
            """
        )

    # ────────────────────────── Busy helpers ──────────────────────────
    def _set_controls_enabled(self, enabled: bool) -> None:
        for w in [self.btn_generate, self.cb_company, self.chk_only_new, self.chk_skip_tracker,
                  self.tpl_le, self.csv_le, self.track_le]:
            if w is not None:
                w.setEnabled(enabled)

    def _begin_busy(self, msg: str) -> None:
        self.lbl_status.setText(msg)
        QGuiApplication.setOverrideCursor(Qt.WaitCursor)
        self._set_controls_enabled(False)
        QGuiApplication.processEvents()

    def _update_busy(self, msg: str) -> None:
        self.lbl_status.setText(msg)
        QGuiApplication.processEvents()

    def _end_busy(self) -> None:
        self.lbl_status.clear()
        QGuiApplication.restoreOverrideCursor()
        self._set_controls_enabled(True)
        QGuiApplication.processEvents()

    # ────────────────────────── Data / parsing ──────────────────────────
    def _load_payee_info_into_ui(self) -> None:
        self.payee_info = _read_payee_info(self.company_csv)
        agency_names = sorted(
            {name for name in self.payee_info.keys() if str(name).strip()},
            key=lambda name: str(name).casefold(),
        )
        self.cb_company.blockSignals(True)
        self.cb_company.clear()
        self.cb_company.addItem("Select an Agency…")
        self.cb_company.addItems(agency_names)
        self.cb_company.blockSignals(False)
        # Wire signals (idempotent)
        try:
            self.cb_company.currentTextChanged.disconnect(self._on_company_changed)
        except Exception:
            pass
        try:
            self.cb_company.currentTextChanged.disconnect(self._check_ready)
        except Exception:
            pass
        self.cb_company.currentTextChanged.connect(self._on_company_changed)
        self.cb_company.currentTextChanged.connect(self._check_ready)

    # ────────────────────────── Selectors ──────────────────────────
    def on_edit_payee(self) -> None:
        dlg = EditCSVDialog(self.company_csv, self)
        if dlg.exec():
            try:
                self._load_payee_info_into_ui()
            except Exception:
                QMessageBox.warning(self, "Error", "Invalid Company CSV.")

    # ────────────────────────── Readiness ──────────────────────────
    def _check_ready(self) -> None:
        """Enable Generate once template, company, and VendorID exist."""
        ok = (
            os.path.isfile(self.template)
            and os.path.isfile(self.company_csv)
            and bool(self.cb_company.currentText())
            and bool(self.selected_vid)
        )
        self.btn_generate.setEnabled(ok)

    def _on_company_changed(self, company: str) -> None:
        """Store selected company's VendorID."""
        self.selected_vid = self.payee_info.get(company, {}).get("VendorID", "").strip()
        self._check_ready()

    def _on_back(self) -> None:
        """Return to the home view (wired to toolbar Back + Esc)."""
        self.ctrl.show_home()

    def showEvent(self, e):
        super().showEvent(e)
        # Every time the page becomes visible, re-pull paths from the config file
        self._refresh_paths_from_config()

    def _refresh_paths_from_config(self) -> None:
        """Pull latest paths from disk and refresh UI + dependent data."""
        refresh_common_paths(self)   # sets: tracker_file, company_csv, template and line-edits

        # If company CSV changed, reload companies
        if self.company_csv and os.path.isfile(self.company_csv):
            if self.company_csv != self._last_company_csv_loaded:
                try:
                    self._load_payee_info_into_ui()
                    self._last_company_csv_loaded = self.company_csv
                except Exception:
                    QMessageBox.warning(self, "Error", "Invalid Company CSV.")

        self._check_ready()

    # ────────────────────────── Generate flow ──────────────────────────
    def generate(self) -> None:
        """Start RFP generation in a background thread to prevent UI freeze."""
        run_id = new_run_id()
        self._run_id = run_id
        comp = self.cb_company.currentText()
        
        logger.info(
            "RFP generation started | company=%s | vendor_id=%s | only_new=%s | skip_tracker_update=%s | tracker=%s | template=%s",
            comp,
            self.selected_vid,
            self.chk_only_new.isChecked(),
            self.chk_skip_tracker.isChecked(),
            self.tracker_file,
            self.template,
            extra={"run_id": run_id},
        )
        
        # Validation
        if not comp or comp == "Select an Agency…":
            QMessageBox.warning(self, "Missing Company", "Please select a company.")
            return

        if not os.path.isfile(self.tracker_file):
            QMessageBox.warning(self, "Missing Tracker Excel", "Please select a valid Invoices Tracker Excel file.")
            return

        # Ask user for output folder before starting worker
        default_out = _default_output_for_company(comp)
        if QMessageBox.question(
            self, "Use Default Folder?", f"Save to:\n{default_out}", QMessageBox.Yes | QMessageBox.No
        ) == QMessageBox.Yes:
            out = default_out
        else:
            out_dir = QFileDialog.getExistingDirectory(self, "Select Output Folder")
            if not out_dir:
                return
            out = os.path.join(out_dir, comp)
        
        # Stop any existing worker
        if self._generation_thread and self._generation_thread.isRunning():
            self._generation_thread.quit()
            self._generation_thread.wait()

        # Create worker and thread
        self._generation_worker = GenerationWorker(
            tracker_path=self.tracker_file,
            template_path=self.template,
            company=comp,
            selected_vid=self.selected_vid,
            payee_info=self.payee_info[comp],
            output_dir=out,
            only_new=self.chk_only_new.isChecked(),
            skip_update=self.chk_skip_tracker.isChecked(),
            run_id=run_id,
        )
        
        self._generation_thread = QThread()
        self._generation_worker.moveToThread(self._generation_thread)

        # Connect signals
        self._generation_worker.status.connect(self._on_generation_status)
        self._generation_worker.progress.connect(self._on_generation_progress)
        self._generation_worker.error.connect(self._on_generation_error)
        self._generation_worker.finished.connect(self._on_generation_finished)
        self._generation_thread.started.connect(self._generation_worker.run)

        # Start UI busy state and thread
        self._begin_busy("Locating vendor sheet…")
        self._generation_thread.start()

    def _on_generation_status(self, msg: str) -> None:
        """Update status message from worker."""
        self._update_busy(msg)

    def _on_generation_progress(self, i: int, total: int) -> None:
        """Update progress bar from worker."""
        self.progress.setMaximum(total if total else 0)
        self.progress.setValue(i)
        self.lbl_status.setText(f"{i}/{total}" if total else f"{i}")

    def _on_generation_error(self, error_msg: str) -> None:
        """Handle error from worker."""
        logger.exception("Generation worker error | msg=%s", error_msg, extra={"run_id": self._run_id})
        QMessageBox.critical(self, "Error", error_msg)
        self._end_busy()

    def _on_generation_finished(self, result: dict) -> None:
        """Handle successful completion from worker."""
        try:
            out = result.get("out_dir", "")
            total = result.get("generated_count", 0)
            sheet = result.get("sheet", "")
            generated_invoices = result.get("generated_invoices", set())
            skip_update = self.chk_skip_tracker.isChecked()

            QMessageBox.information(self, "Done", f"RFPs saved to {out}")
            QDesktopServices.openUrl(QUrl.fromLocalFile(out))

            # Update counter
            try:
                self.rfp_count += int(total)
                save_config_merge({"rfp_count": self.rfp_count})
                self.lbl_counter.setText(f"Total RFPs generated: {self.rfp_count}")
            except Exception:
                pass

            # Update tracker if invoices generated
            if generated_invoices:
                if skip_update:
                    self._update_busy("Exporting rows that would be updated…")
                else:
                    self._update_busy("Updating tracker (RFP Prepared = Y)…")
                    
                updated, wanted, not_found, skipped_df = _update_tracker_rfp_prepared(
                    tracker_path=self.tracker_file,
                    sheet_name=sheet,
                    invoice_numbers=generated_invoices,
                    skip_update=skip_update,
                    run_id=self._run_id,
                )
                
                if skip_update:
                    if skipped_df is not None and not skipped_df.empty:
                        export_path = os.path.join(out, f"Tracker_Update_Preview_{self.cb_company.currentText()}.xlsx")
                        skipped_df.to_excel(export_path, index=False, engine='openpyxl')
                        self.lbl_status.setText(f"Preview: {updated} rows would be updated.")
                        QMessageBox.information(
                            self, "Tracker Update Skipped",
                            f"{updated} row(s) would have been updated.\n\n"
                            f"Preview exported to:\n{export_path}"
                        )
                    else:
                        self.lbl_status.setText("No rows found to update in tracker")
                else:
                    self.lbl_status.setText(f"Tracker updated: {updated}/{wanted} rows set to Y")
                    if not_found:
                        logger.warning(
                            "Generator tracker invoice mismatch | sheet=%s | not_found_count=%s",
                            sheet,
                            len(not_found),
                            extra={"run_id": self._run_id},
                        )
                        QMessageBox.information(
                            self, "Tracker Notice",
                            f"{len(not_found)} invoice(s) not found in '{sheet}'."
                        )

        except Exception as e:
            logger.exception("Generation completion handler error", extra={"run_id": self._run_id})
            QMessageBox.critical(self, "Error", str(e))
        finally:
            logger.info("RFP generation finished", extra={"run_id": self._run_id})
            self._run_id = "-"
            self._end_busy()
            
            # Cleanup thread
            if self._generation_thread:
                self._generation_thread.quit()
                self._generation_thread.wait()