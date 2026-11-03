# Authored By: Nathan Nicholas D Souza
# Team: QI Team - Staffing Services at Fraser Health
# Contact: nathan.dsouza@fraserhealth.ca

# app_utils.py

"""
Unified RFP Generator & Merger Utilities

• Configuration management (load/save)
• Filename & currency formatting
• PDF form-filling with PyMuPDF
• CSV processing for RFP generation
• Invoice-number extraction & normalization for PDF merging
"""

import os
import sys
import re
import csv
import json
import shutil
import stat
from datetime import datetime

try:
    import pymupdf as fitz          # PyMuPDF (preferred; avoids third-party fitz package collision)
except Exception:
    import fitz                     # PyMuPDF fallback
from PyPDF2 import PdfMerger
import pandas as pd
import xlwings as xw

from typing import Dict, Any
from utils.paths_config import DEFAULT_INVOICE_TRACKER
from utils.logging_utils import get_logger

logger = get_logger(__name__)

# Canonical config keys
CONF_TRACKER          = "tracker"
CONF_TRACKER_2        = "tracker_2"
CONF_COMPANY_CSV      = "company_csv"
CONF_TEMPLATE         = "template"
CONF_SENT_INVOICES    = "sent_invoices"
CONF_PENDING_INVOICES  = "pending_invoices"

_ALLOWED_PATH_KEYS = {
    CONF_TRACKER, CONF_TRACKER_2, CONF_COMPANY_CSV, CONF_TEMPLATE, CONF_SENT_INVOICES, CONF_PENDING_INVOICES
}

# def _migrate_config_keys(cfg: Dict) -> Dict:
#     """
#     One-time rename of older keys to new canonical names.
#     Safe to run every load; only saves when a change occurs.
#     """
#     mapping = {
#         "tracker_file":      CONF_TRACKER,
#         "payee_csv":         CONF_COMPANY_CSV,
#         "template":          CONF_TEMPLATE,
#         "sent_folder":       CONF_SENT_INVOICES,
#         "pending_output_dir":CONF_PENDING_INVOICES,  # already same, just keep
#     }
#     changed = False
#     for old, new in mapping.items():
#         if old in cfg and new not in cfg:
#             cfg[new] = cfg.pop(old)
#             changed = True
#     if changed:
#         save_config(cfg)
#     return cfg

# -------------------------------------------------------------------
# 1. Application directories & bundled defaults (robust for onedir/onefile)
# -------------------------------------------------------------------
if getattr(sys, "frozen", False):
    # Running as a bundled executable
    APP_DIR = os.path.dirname(sys.executable)             # folder with the .exe (onedir)
    MEI     = getattr(sys, "_MEIPASS", APP_DIR)           # temp unpack dir (onefile) or APP_DIR fallback

    # Prefer ./resources next to the exe; fall back to MEIPASS/resources
    RESOURCES = os.path.join(APP_DIR, "resources")
    if not os.path.isdir(RESOURCES):
        RESOURCES = os.path.join(MEI, "resources")

    CONFIG_PATH  = os.path.join(APP_DIR, "rfp_config.json")

else:
    # Running from source
    THIS_DIR   = os.path.dirname(os.path.abspath(__file__))
    APP_DIR    = THIS_DIR
    BASE_PATH  = os.path.dirname(THIS_DIR)                # project root
    RESOURCES  = os.path.join(BASE_PATH, "resources")

    CONFIG_PATH  = os.path.join(APP_DIR, "rfp_config.json")

ICON_PATH = os.path.join(RESOURCES, "app_icon.ico")
WAV_PATH  = os.path.join(RESOURCES, "ting.wav")

# -------------------------------------------------------------------
# 2. Configuration load/save
# -------------------------------------------------------------------
def load_config() -> dict:
    """Load JSON config from disk, return {} on error."""
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}

def save_config(cfg: dict) -> None:
    """Persist config dict as pretty-printed JSON."""
    try:
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=2)
    except Exception:
        pass


def get_paths() -> Dict[str, str]:
    # cfg = _migrate_config_keys(load_config())
    cfg = load_config()
    # Always return strings (\"\" if missing)
    return {
        "tracker":           str(cfg.get(CONF_TRACKER, "") or ""),
        "tracker_2":         str(cfg.get(CONF_TRACKER_2, "") or ""),
        "company_csv":       str(cfg.get(CONF_COMPANY_CSV, "") or ""),
        "template":          str(cfg.get(CONF_TEMPLATE, "") or ""),
        "sent_invoices":     str(cfg.get(CONF_SENT_INVOICES, "") or ""),
        "pending_invoices":  str(cfg.get(CONF_PENDING_INVOICES, "") or ""),
    }


# Typed getters
def get_tracker_path() -> str:          return get_paths()["tracker"]
def get_tracker_path_2() -> str:        return get_paths()["tracker_2"]
def get_company_csv_path() -> str:      return get_paths()["company_csv"]
def get_template_path() -> str:         return get_paths()["template"]
def get_sent_invoices_dir() -> str:     return get_paths()["sent_invoices"]
def get_pending_invoices_dir() -> str:  return get_paths()["pending_invoices"]


# Settings-only setter (use ONLY from SettingsPage)
def set_path(key: str, value: str) -> None:
    if key not in _ALLOWED_PATH_KEYS:
        raise KeyError(f"Unsupported config key: {key}")
    cfg = load_config()
    cfg[key] = value or ""
    save_config(cfg)


def save_config_merge(updates: Dict[str, Any]) -> None:
    """Merge updates into latest on-disk config and save."""
    cfg = load_config()
    cfg.update(updates or {})
    save_config(cfg)


def refresh_common_paths(page: Any) -> None:
    """Idempotently refresh common config paths and reflect them in line-edits if present."""
    try:
        page.tracker_file = get_tracker_path()
    except Exception:
        page.tracker_file = getattr(page, "tracker_file", "")
    try:
        page.tracker_file_2 = get_tracker_path_2()
    except Exception:
        page.tracker_file_2 = getattr(page, "tracker_file_2", "")
    try:
        page.company_csv = get_company_csv_path()
    except Exception:
        page.company_csv = getattr(page, "company_csv", "")
    try:
        page.pending_output_dir = get_pending_invoices_dir()
    except Exception:
        page.pending_output_dir = getattr(page, "pending_invoices", "")
    try:
        page.template = get_template_path()
    except Exception:
        page.template = getattr(page, "template", "")
    try:
        page.sent_root = get_sent_invoices_dir()
    except Exception:
        page.sent_root = getattr(page, "sent_invoices", "")

    if hasattr(page, "track_le"):
        page.track_le.setText(page.tracker_file or "")
    if hasattr(page, "track2_le"):
        page.track2_le.setText(page.tracker_file_2 or "")
    if hasattr(page, "csv_le"):
        page.csv_le.setText(os.path.basename(page.company_csv) if page.company_csv else "")
    if hasattr(page, "out_dir_le"):
        page.out_dir_le.setText(page.pending_output_dir if page.pending_output_dir else "")
    if hasattr(page, "tpl_le"):
        page.tpl_le.setText(os.path.basename(page.template) if page.template else "")
    if hasattr(page, "le_sent"):
        page.le_sent.setText(page.sent_root if page.sent_root else "")

# -------------------------------------------------------------------
# 3. File & currency formatting
# -------------------------------------------------------------------
def clean_filename(text: str) -> str:
    """Strip characters illegal in filenames."""
    return re.sub(r'[\\/*?:"<>|]', "", text.strip())

def clean_currency(value: str) -> str:
    """Remove '$' and commas, return '0' if empty."""
    return value.replace("$", "").replace(",", "").strip() if value else "0"

def format_number(value: str) -> str:
    """Convert numeric string to plain decimal, '0' on error."""
    try:
        amt = f"{float(value):.2f}"
        return str(amt)
    except Exception:
        return "0"

def format_currency(value: str) -> str:
    """Format numeric string as currency, e.g. '$1,234.56'."""
    try:
        amt = float(clean_currency(value))
        return f"${amt:,.2f}"
    except Exception:
        return "$0.00"


def _is_usable_invoice_id(value: str) -> bool:
    """Return True only for values that look like real invoice IDs."""
    v = str(value or "").strip()
    if not v:
        return False

    lower = v.lower()
    invalid_literals = {
        "invoice",
        "invoice #",
        "invoice#",
        "total",
        "subtotal",
        "grand total",
        "nan",
        "none",
        "null",
        "n/a",
        "na",
        "0",
        "0.0",
        "-",
        "--",
    }
    if lower in invalid_literals:
        return False

    # Must contain at least one alphanumeric character.
    return bool(re.search(r"[A-Za-z0-9]", v))

# def _force_excel_recalc(path):
#     app = xw.App(visible=False)
#     wb  = app.books.open(path)
#     wb.api.RefreshAll()
#     app.api.CalculateUntilAsyncQueriesDone()
#     wb.save()
#     wb.close()
#     app.quit()

import win32com.client
import pythoncom
import gc
import os
import signal
import psutil
import time
import subprocess

def _force_excel_recalc(path, sheet_name=None, refresh_external=True, row_numbers=None):
    """
    Uses pywin32 COM to automate Excel directly (Windows only).
    Requires Excel to be installed.

    Args:
        path: Workbook path.
        sheet_name: If provided, calculate this sheet only (faster for targeted rows).
        refresh_external: If True, run RefreshAll + async wait before calculation.
        row_numbers: Optional list of 1-based row numbers. When provided with sheet_name,
                     only those specific rows are recalculated instead of the whole sheet.
    """
    pythoncom.CoInitialize()  # Initialize COM

    try:
        excel = win32com.client.Dispatch("Excel.Application")
        # Keep Excel silent/hidden for speed and less disruption.
        excel.Visible = False
        excel.DisplayAlerts = False

        # Try to reduce Excel process priority to avoid overwhelming the machine
        # (best-effort; may affect other Excel instances). This helps packaged
        # apps avoid driving the system to "Not Responding" during heavy recalc.
        try:
            for p in psutil.process_iter(['name']):
                nm = (p.info.get('name') or "").lower()
                if nm.startswith('excel'):
                    try:
                        p.nice(psutil.IDLE_PRIORITY_CLASS)
                    except Exception:
                        pass
        except Exception:
            pass

        wb = excel.Workbooks.Open(path)

        if refresh_external:
            # Refresh external data connections only when requested (expensive).
            wb.RefreshAll()
            excel.CalculateUntilAsyncQueriesDone()
            time.sleep(5)

        if sheet_name:
            try:
                ws_obj = wb.Worksheets(sheet_name)
                if row_numbers:
                    # Calculate only the specific rows — much faster than recalculating the whole sheet.
                    for row_num in row_numbers:
                        ws_obj.Rows(f"{row_num}:{row_num}").Calculate()
                        # brief sleep to yield CPU and allow OS to stay responsive
                        time.sleep(0.01)
                    logger.info("Excel COM targeted row recalc | sheet=%s | rows=%s", sheet_name, row_numbers)
                else:
                    ws_obj.Calculate()
            except Exception:
                logger.warning("Excel COM sheet calculate fallback to full | sheet=%s", sheet_name)
                excel.Calculate()
        else:
            excel.Calculate()
        
        try:
            wb.Save()
        except Exception as e:
            logger.exception("Excel COM save failed")

        try:
            wb.Close()
        except Exception as e:
            logger.exception("Excel COM close workbook failed")

        try:
            excel.Quit()
        except Exception as e:
            logger.exception("Excel COM quit failed")

        wb = None
        excel = None
        del wb, excel
        gc.collect()
        
    finally:
        pythoncom.CoUninitialize()


def run_excel_recalc_subprocess(path, sheet_name=None, refresh_external=True, row_numbers=None, timeout: int = 600):
    """Run Excel recalc in a separate process by invoking `utils/recalc_subprocess.py`.

    Falls back to in-process `_force_excel_recalc` on failure.
    """
    try:
        script_dir = os.path.dirname(__file__)
        script_path = os.path.join(script_dir, "recalc_subprocess.py")

        cmd = [sys.executable, script_path, "--path", path]
        if sheet_name:
            cmd += ["--sheet", sheet_name]
        cmd += ["--refresh-external", "1" if refresh_external else "0"]
        if row_numbers:
            cmd += ["--rows", ",".join(str(int(x)) for x in row_numbers)]

        proc = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout)
        if proc.returncode != 0:
            logger.warning(
                "Subprocess recalc failed, rc=%s, stderr=%s", proc.returncode, proc.stderr
            )
            # Fallback to in-process recalc
            _force_excel_recalc(path, sheet_name=sheet_name, refresh_external=refresh_external, row_numbers=row_numbers)
        else:
            logger.info("Subprocess recalc completed | path=%s | sheet=%s", path, sheet_name)
    except Exception as e:
        logger.exception("Subprocess recalc failed, falling back to in-process: %s", e)
        try:
            _force_excel_recalc(path, sheet_name=sheet_name, refresh_external=refresh_external, row_numbers=row_numbers)
        except Exception:
            logger.exception("In-process recalc also failed.")

# -------------------------------------------------------------------
# 3b. Safe workbook save (no lingering file handles on Windows)
# -------------------------------------------------------------------
import gc
import tempfile

def save_workbook_safe(wb, path: str) -> None:
    """
    Save an openpyxl Workbook to *path* without leaving a file lock on Windows.

    Strategy:
      1. Write to a sibling temp file on the same drive (avoids cross-drive rename).
      2. Close the workbook and force a GC cycle to release any internal ZipFile handles.
      3. Atomically replace the original via os.replace().

    This prevents openpyxl's ZipFile writer from ever touching the original path,
    so even if CPython's cyclic GC hasn't yet freed the workbook object the
    original file stays unlocked and openable in Excel immediately.
    """
    path = os.path.abspath(path)
    dir_ = os.path.dirname(path)
    fd, tmp = tempfile.mkstemp(dir=dir_, suffix=".~tmp")
    os.close(fd)
    saved = False
    try:
        wb.save(tmp)
        saved = True
    finally:
        try:
            wb.close()
        except Exception:
            pass
        gc.collect()          # release cyclic ZipFile references
        if not saved:
            try:
                os.unlink(tmp)
            except OSError:
                pass
    os.replace(tmp, path)   # atomic on same drive/share

# -------------------------------------------------------------------
# 4. PDF form-filling with PyMuPDF
# -------------------------------------------------------------------
def fill_pdf_from_row_fitz(
    input_pdf: str,
    output_pdf: str,
    mapping: dict,
    row: dict,
    run_id: str = "-",
) -> None:
    """
    Open input_pdf, fill its form fields according to mapping(row),
    repair broken xrefs, and save to output_pdf.
    """
    def _norm_pdf_field_name(name: str) -> str:
        return re.sub(r"[^a-z0-9]+", "", str(name or "").lower())

    doc = fitz.open(input_pdf)
    mapping_norm = {_norm_pdf_field_name(k): k for k in mapping.keys() if k}
    widget_names = []
    matched = 0
    non_empty = 0

    for page in doc:
        widgets = page.widgets() or []
        for widget in widgets:
            key = widget.field_name
            if key:
                widget_names.append(str(key))

            if not key:
                continue

            source_key = key if key in mapping else mapping_norm.get(_norm_pdf_field_name(key))
            if not source_key:
                continue

            src = mapping[source_key]
            raw = src(row) if callable(src) else row.get(src, "")
            if source_key == "000TOTAL":
                val = format_currency(clean_currency(raw))
            elif source_key in ("Text25", "Text5", "Text18"):
                val = format_number(clean_currency(raw))
            else:
                val = raw or ""
            if val == "" and source_key in ("Text25", "Text5"):
                val = "0"

            val_str = str(val)
            try:
                widget.field_value = val_str
                widget.update()
                matched += 1
                if val_str.strip():
                    non_empty += 1
            except Exception:
                logger.exception(
                    "PDF widget update failed | field=%s | source_key=%s | output=%s",
                    key,
                    source_key,
                    os.path.basename(output_pdf),
                    extra={"run_id": run_id},
                )

    # Help some viewers render values consistently.
    try:
        doc.need_appearances(True)
    except Exception:
        pass

    if matched == 0:
        logger.warning(
            "PDF fill matched no fields | template=%s | output=%s | widget_count=%s | widget_sample=%s | mapping_keys_sample=%s",
            input_pdf,
            os.path.basename(output_pdf),
            len(widget_names),
            widget_names[:12],
            list(mapping.keys())[:12],
            extra={"run_id": run_id},
        )
    else:
        logger.debug(
            "PDF fill summary | output=%s | widget_count=%s | matched=%s | non_empty=%s",
            os.path.basename(output_pdf),
            len(widget_names),
            matched,
            non_empty,
            extra={"run_id": run_id},
        )

    # repair any broken xrefs so save() succeeds
    for xref in range(1, doc.xref_length()):
        try:
            _ = doc.xref_object(xref)
        except RuntimeError:
            doc.update_object(xref, "<<>>")
    doc.save(output_pdf)

# -------------------------------------------------------------------
# 5. CSV processing & RFP generation
# -------------------------------------------------------------------
def process_csv_and_fill_pdfs(
    inv_csv: str,
    tmpl_pdf: str,
    out_dir: str,
    company: str,
    payee_info: dict,
    only_new: bool = True,
    run_id: str = "-",
):
    """
    Read invoice CSV, fill one PDF per row via fill_pdf_from_row_fitz().
    If only_new, only process rows with 'RFP Prepared (Y/N)' == 'N'.
    Yields (current_index, total_rows) for progress.
    """
    os.makedirs(out_dir, exist_ok=True)

    with open(inv_csv, encoding="utf-8-sig", errors="replace", newline="") as f:
        raw_rows = list(csv.reader(f))

    if len(raw_rows) <= 2:
        logger.warning(
            "RFP CSV parse failed | reason=too_few_rows | csv=%s | rows=%s",
            inv_csv,
            len(raw_rows),
            extra={"run_id": run_id},
        )
        raise ValueError("No valid rows in Invoice CSV.")

    rows = [r for r in raw_rows[2:] if any(str(c).strip() for c in r)]
    if not rows:
        raise ValueError("No valid rows in Invoice CSV.")

    # Some tracker sheets include a leading blank/unnamed column while others do not.
    header_row = rows[0]
    lead = str(header_row[0]).strip() if header_row else ""
    drop_first_col = (not lead) or lead.lower().startswith("unnamed")
    if drop_first_col:
        rows = [r[1:] for r in rows if len(r) > 1]

    if not rows:
        raise ValueError("No valid rows in Invoice CSV.")

    headers = [str(h or "").strip() for h in rows[0]]
    norm_headers = [re.sub(r"[^a-z0-9#]", "", h.lower()) for h in headers]
    norm_to_header = {k: h for h, k in zip(headers, norm_headers) if k}

    def pick_header(*aliases: str) -> str:
        for alias in aliases:
            key = re.sub(r"[^a-z0-9#]", "", alias.lower())
            col = norm_to_header.get(key)
            if col:
                return col
        return ""

    invoice_aliases = {"invoice#", "invoicenumber", "invoice"}
    invoice_idx = next((i for i, h in enumerate(norm_headers) if h in invoice_aliases), None)
    # detect RFP Prepared column index (if present) so we can report when rows
    # marked 'N' are skipped due to missing Invoice # values.
    rfp_aliases = {re.sub(r"[^a-z0-9#]", "", a.lower()) for a in ("RFP Prepared (Y/N)", "RFP Prepared", "RFP Prepared Y/N")}
    rfp_idx = next((i for i, h in enumerate(norm_headers) if h in rfp_aliases), None)
    if invoice_idx is None:
        logger.warning(
            "RFP CSV parse failed | reason=missing_invoice_column | headers=%s",
            headers[:20],
            extra={"run_id": run_id},
        )
        raise ValueError("Invoice column not found in Invoice CSV.")

    data_rows = rows[1:]
    invoice_rows = []
    blank_invoice_rows = 0
    invalid_invoice_rows = 0
    blank_invoice_rows_marked_N = 0
    for raw in data_rows:
        # Align row width to headers so dict(zip(...)) is stable.
        if len(raw) < len(headers):
            raw = raw + [""] * (len(headers) - len(raw))
        elif len(raw) > len(headers):
            raw = raw[:len(headers)]

        inv_val = str(raw[invoice_idx] or "").strip()
        col_b_val = str(raw[1] if len(raw) > 1 else "").strip() if raw else ""

        # Treat a blank invoice cell or an explicit Total marker as the end of the sheet.
        # This prevents trailing summary/footer rows from producing an extra RFP.
        is_end_marker = (not inv_val) or col_b_val.lower() == "total"
        if is_end_marker:
            if not inv_val:
                blank_invoice_rows += 1
            if col_b_val.lower() == "total":
                logger.info(
                    "RFP CSV end marker reached | csv=%s | column_b=%s | invoice_col=%s",
                    inv_csv,
                    col_b_val,
                    headers[invoice_idx],
                    extra={"run_id": run_id},
                )
            else:
                logger.info(
                    "RFP CSV blank invoice end marker reached | csv=%s | invoice_col=%s",
                    inv_csv,
                    headers[invoice_idx],
                    extra={"run_id": run_id},
                )
            break

        if not _is_usable_invoice_id(inv_val):
            invalid_invoice_rows += 1
            # If this row was marked 'N' for RFP Prepared and we're in only_new
            # mode the row would otherwise be a candidate — record it so we can
            # warn the user that 'N' rows with missing Invoice # are skipped.
            try:
                if only_new and rfp_idx is not None:
                    rfp_val = str(raw[rfp_idx] or "").strip().upper()
                    if rfp_val == "N":
                        blank_invoice_rows_marked_N += 1
            except Exception:
                pass
            continue
        invoice_rows.append(dict(zip(headers, raw)))

    if not invoice_rows:
        logger.warning(
            "RFP CSV parse empty after invoice filtering | csv=%s | blank_invoice_rows=%s | invalid_invoice_rows=%s | blank_invoice_rows_marked_N=%s",
            inv_csv,
            blank_invoice_rows,
            invalid_invoice_rows,
            blank_invoice_rows_marked_N,
            extra={"run_id": run_id},
        )
        if blank_invoice_rows_marked_N:
            # Provide a clearer error when the only rows were 'N' but lacked invoice numbers.
            raise ValueError(f"No valid rows in Invoice CSV: {blank_invoice_rows_marked_N} rows marked 'N' had no Invoice #.")
        raise ValueError("No valid rows in Invoice CSV.")

    logger.info(
        "RFP CSV parsed | company=%s | csv=%s | rows_raw=%s | rows_non_empty=%s | rows_with_invoice=%s | blank_invoice_rows=%s | invalid_invoice_rows=%s | dropped_leading_col=%s | invoice_col=%s",
        company,
        inv_csv,
        len(raw_rows),
        len(data_rows),
        len(invoice_rows),
        blank_invoice_rows,
        invalid_invoice_rows,
        drop_first_col,
        headers[invoice_idx],
        extra={"run_id": run_id},
    )

    if only_new:
        rfp_col = pick_header("RFP Prepared (Y/N)", "RFP Prepared", "RFP Prepared Y/N") or "RFP Prepared (Y/N)"

        kept = 0
        skipped = 0
        state_counts: Dict[str, int] = {}
        filtered_rows = []
        for row in invoice_rows:
            state = str(row.get(rfp_col, "")).strip().upper()
            state_counts[state or "<blank>"] = state_counts.get(state or "<blank>", 0) + 1
            if state == "N":
                filtered_rows.append(row)
                kept += 1
            else:
                skipped += 1
        # If no explicit 'N' rows were found but there are blank RFP cells,
        # treat blanks as 'N' as a fallback (helpful when trackers don't mark 'N'),
        # but ONLY if they also have a valid Invoice #.
        if kept == 0 and state_counts.get("<blank>", 0) > 0:
            blank_rows_with_invoice = [
                row for row in invoice_rows
                if not str(row.get(rfp_col, "")).strip() and _is_usable_invoice_id(row.get(invoice_col_name, ""))
            ]
            if blank_rows_with_invoice:
                filtered_rows = blank_rows_with_invoice
                logger.warning(
                    "RFP only_new fallback: no explicit 'N' found, using blank RFP cells (with Invoice #) as new | company=%s | blank_rows_with_invoice=%s | total_blanks=%s",
                    company,
                    len(blank_rows_with_invoice),
                    state_counts.get("<blank>", 0),
                    extra={"run_id": run_id},
                )
                for idx, row in enumerate(blank_rows_with_invoice[:5], start=1):
                    logger.debug(
                        "RFP blank-as-new sample | idx=%s | invoice=%s | site=%s | unit=%s",
                        idx,
                        row.get(invoice_col_name, ""),
                        row.get(site_col_name, ""),
                        row.get(unit_col_name, ""),
                        extra={"run_id": run_id},
                    )
            else:
                logger.warning(
                    "RFP only_new: no explicit 'N' found, and blank RFP cells had no Invoice #, skipping generation | company=%s | total_blanks=%s",
                    company,
                    state_counts.get("<blank>", 0),
                    extra={"run_id": run_id},
                )

        invoice_rows = [row for row in filtered_rows]
        logger.info(
            "RFP only_new filter applied | company=%s | filter_column=%s | kept_N=%s | skipped_non_N=%s | states=%s",
            company,
            rfp_col,
            kept,
            skipped,
            state_counts,
            extra={"run_id": run_id},
        )

    total = len(invoice_rows)
    logger.info(
        "RFP PDF generation starting | company=%s | total_rows=%s | output_dir=%s",
        company,
        total,
        out_dir,
        extra={"run_id": run_id},
    )

    invoice_col_name = headers[invoice_idx]
    site_col_name = pick_header("Site") or "Site"
    unit_col_name = pick_header("Unit") or "Unit"
    due_col_name = pick_header("Date Invoice Due") or "Date Invoice Due"
    total_col_name = pick_header("Overall Invoice Bill Totals", "Overall Invoice Bill Totals ") or "Overall Invoice Bill Totals "
    gst_col_name = pick_header("GST Subtotals", "GST Subtotal") or "GST Subtotals"
    services_col_name = pick_header("Services Subtotal", "Services Subtotals") or "Services Subtotal"
    travel_col_name = pick_header("Travel/Accom Reimbursements") or "Travel/Accom Reimbursements"
    agency_personnel_col_name = pick_header("Agency Personnel Name") or "Agency Personnel Name"
    cost_centre_col_name = pick_header("Cost Centre") or "Cost Centre"
    manager_name_col_name = pick_header("Manager Name") or "Manager Name"

    logger.info(
        "RFP mapped source columns | company=%s | invoice=%s | site=%s | unit=%s | due=%s | total=%s | gst=%s | services=%s | travel=%s | personnel=%s | cost_centre=%s | manager=%s",
        company,
        invoice_col_name,
        site_col_name,
        unit_col_name,
        due_col_name,
        total_col_name,
        gst_col_name,
        services_col_name,
        travel_col_name,
        agency_personnel_col_name,
        cost_centre_col_name,
        manager_name_col_name,
        extra={"run_id": run_id},
    )

    for idx, row in enumerate(invoice_rows, start=1):
        inv   = clean_filename(str(row.get(invoice_col_name, "") or "NoInvoice"))
        site  = clean_filename(str(row.get(site_col_name, "") or "NoSite"))
        unit  = clean_filename(str(row.get(unit_col_name, "") or "NoUnit"))
        out_name = f"{company} RFP {inv} - {site} {unit}.pdf"
        out_path = os.path.join(out_dir, out_name)

        # mapping field→CSV-column or function
        mapping = {
            "Text17":       lambda r: datetime.today().strftime("%d/%m/%y"),
            "Invoice No":   invoice_col_name,
            "DDMMYY": lambda r: (
                pd.to_datetime(r.get(due_col_name, ""), dayfirst=False, errors="coerce")
                .strftime("%d/%m/%y")
                if pd.notna(pd.to_datetime(r.get(due_col_name, ""), dayfirst=False, errors="coerce"))
                else ""
            ),
            "DD/MM/YY":     lambda r: "",
            "Text2":        cost_centre_col_name,
            "Text18":       gst_col_name,
            "000TOTAL":     total_col_name,
            "Name of Requestor":  lambda r: "Ashley Kirkham",
            "PositionTitle":      lambda r: "Coordinator",
            'Requestor Signature': lambda r: "A. Kirkham",
            "Date DDMMYY":        lambda r: datetime.today().strftime("%d/%m/%y"),
            "Name of Approver":   manager_name_col_name,
        }
        mapping["PositionTitle_2"] = (
            (lambda r: "Manager")
            if row.get(unit_col_name) == "Medical Imaging"
            else (lambda r: "Clinical Manager")
        )

        payee_map = {
            "Must match vendor listed on invoice": "Payable To",
            "Street Address":   "Street Address",
            "City":             "City",
            "Province":         "Province",
            "Country":          "Country",
            "Postal Code":      "Postal Code",
        }
        for pdf_field, col_name in payee_map.items():
            val = payee_info.get(col_name, "")
            mapping[pdf_field] = (lambda v: (lambda r: v))(val)

        mapping.update({
            "Describe Payment":       lambda r: f"Payment for {r.get(agency_personnel_col_name,'')} hours",
            "Text25":                 services_col_name,
            "Text16":                 lambda r: "3509030 UPP",
            "Describe paymentRow1":   lambda r: f"Payment for {r.get(agency_personnel_col_name,'')} travel/accom",
            "Text5":                  travel_col_name,
            "Text4":                  lambda r: "3509200 UPP",
            "Text3":                  cost_centre_col_name,
        })

        if idx <= 3:
            logger.debug(
                "RFP source row snapshot | company=%s | index=%s/%s | invoice=%s | due=%s | total=%s | gst=%s | services=%s | travel=%s | personnel=%s | manager=%s",
                company,
                idx,
                total,
                row.get(invoice_col_name, ""),
                row.get(due_col_name, ""),
                row.get(total_col_name, ""),
                row.get(gst_col_name, ""),
                row.get(services_col_name, ""),
                row.get(travel_col_name, ""),
                row.get(agency_personnel_col_name, ""),
                row.get(manager_name_col_name, ""),
                extra={"run_id": run_id},
            )

        fill_pdf_from_row_fitz(tmpl_pdf, out_path, mapping, row, run_id=run_id)
        logger.debug(
            "RFP row generated | company=%s | index=%s/%s | invoice=%s | output=%s",
            company,
            idx,
            total,
            row.get("Invoice #", ""),
            os.path.basename(out_path),
            extra={"run_id": run_id},
        )
        yield idx, total, row.get(invoice_col_name, "")

    logger.info(
        "RFP PDF generation complete | company=%s | generated=%s",
        company,
        total,
        extra={"run_id": run_id},
    )

# -------------------------------------------------------------------
# 6. Invoice-number extraction & normalization (Merger)
# -------------------------------------------------------------------
def extract_invoice_number(filename: str) -> str | None:
    """
    1) If the basename starts with "Invoice {inv} -", return exactly {inv}
       (now allows underscores, e.g., "FHA_HOT_RE001").
    1b) If the 2nd word is "RFP" and there's a literal " - " later, return
        the substring between "RFP" and " - " (trimmed).
    2) Otherwise split on whitespace only (NOT underscores), pick the token
       with ≥3 digits; keep underscores/hyphens. If the chosen token has
       extra suffix segments (e.g., "_WITH", "_RECEIPT"), trim everything
       after the last segment that contains a digit.
    """
    base = os.path.splitext(os.path.basename(filename))[0]

    # 1) Exact Invoice pattern: allow underscores and hyphens in {inv}
    #    e.g. "Invoice FH_25_012 - ..." -> "FH_25_012"
    m = re.match(r'(?i)^Invoice\s+([A-Za-z0-9_-]+)\s*-\s*', base)
    if m:
        return m.group(1)

    # 1b) "{anything} RFP {INV} - {anything}" where RFP is the 2nd word
    #     Extract the substring between "RFP" and the literal " - "
    #     (underscores/hyphens preserved)
    m = re.match(r'(?i)^\S+\s+RFP\s+(.+?)\s-\s', base)
    if m:
        return m.group(1).strip()

    # 2) Fallback heuristic — split on whitespace ONLY (keep underscores)
    tokens = re.split(r'\s+', base)
    candidates = []
    for tok in tokens:
        tok_clean = tok.strip(" -.,()")  # keep internal '_' and '-'
        if not tok_clean:
            continue

        # skip tokens that are literally 'invoice...'
        if re.match(r'(?i)^invoice', tok_clean):
            continue

        if not re.search(r'\d', tok_clean):
            continue

        # strip optional 'ADJ' suffix
        core = re.sub(r'(?i)ADJ$', "", tok_clean)
        digit_count = len(re.sub(r'\D', "", core))
        if digit_count >= 3:
            candidates.append((tok_clean, digit_count))

    if not candidates:
        return None

    # pick the token with the most digits; tiebreaker: longest token
    best = max(candidates, key=lambda x: (x[1], len(x[0])))[0]

    # If token has underscores, trim trailing non-numeric segments:
    #   "FHA_HOT_RE001_WITH_RECEIPT" -> "FHA_HOT_RE001"
    if "_" in best:
        parts = best.split("_")
        idx_with_digit = [i for i, p in enumerate(parts) if re.search(r'\d', p)]
        if idx_with_digit:
            best = "_".join(parts[: idx_with_digit[-1] + 1])

    return best

def normalize_key(s: str) -> str:
    """Lowercase & strip non-alphanumerics for substring matching."""
    return re.sub(r'[^A-Za-z0-9]', "", s).lower()

# -------------------------------------------------------------------
# 7. Excel tracker open/close detection & interactive prompt
# -------------------------------------------------------------------

def tracker_lockfile_path(xlsx_path: str) -> str:
    """Excel creates a hidden lock file '~$<name>.xlsx' while the workbook is open."""
    d, n = os.path.split(xlsx_path or "")
    return os.path.join(d, f"~${n}") if d and n else ""

def tracker_appears_open(xlsx_path: str) -> bool:
    """Check if tracker is open in Excel. Automatically removes stale lock files."""
    p = (xlsx_path or "").strip()
    if not p or not os.path.isfile(p):
        return False
    lock = tracker_lockfile_path(p)
    if not os.path.exists(lock):
        return False
    # Lock file present — verify Excel is actually running.
    # If no EXCEL.EXE process exists the lock file is stale (e.g. after a crash or kill).
    try:
        excel_running = any(
            proc.info['name'] == 'EXCEL.EXE'
            for proc in psutil.process_iter(['name'])
        )
    except Exception:
        excel_running = True  # can't tell; assume real
    if not excel_running:
        try:
            os.remove(lock)
        except OSError:
            pass
        return False
    return True

def ensure_tracker_closed_interactive(parent, tracker_path: str) -> bool:
    """
    If the tracker looks open, bring it to the front and prompt the user to save/close it.
    Returns True to proceed, False to abort.
    """
    if not tracker_appears_open(tracker_path):
        return True

    # Lazy imports to avoid forcing Qt on non-UI callers
    from PySide6.QtWidgets import QMessageBox
    from PySide6.QtGui import QDesktopServices
    from PySide6.QtCore import QUrl

    # Bring workbook forward in Excel
    try:
        QDesktopServices.openUrl(QUrl.fromLocalFile(tracker_path))
    except Exception:
        pass  # harmless: we still show the dialog

    # First prompt
    box = QMessageBox(parent)
    box.setIcon(QMessageBox.Warning)
    box.setWindowTitle("Tracker appears to be open")
    box.setText("Your tracker workbook looks open in Excel.\nPlease save and close it before continuing.")
    box.setInformativeText("Click Continue after you've closed it, or Cancel to stop.")
    btn_continue = box.addButton("Continue", QMessageBox.AcceptRole)
    btn_cancel   = box.addButton("Cancel",   QMessageBox.RejectRole)
    box.setDefaultButton(btn_continue)
    box.exec()

    if box.clickedButton() is btn_cancel:
        return False

    # Re-check once more
    if tracker_appears_open(tracker_path):
        resp = QMessageBox.question(
            parent,
            "Tracker still open?",
            "The tracker still appears to be open. Continue anyway?\n"
            "(Recommended: Cancel and close it first.)",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        return resp == QMessageBox.Yes

    return True