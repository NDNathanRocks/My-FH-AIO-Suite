# check_page.py

from __future__ import annotations

import os
import pickle
import shutil
import tempfile
import time
from dataclasses import dataclass
from datetime import datetime, timedelta
from typing import Dict, List, Tuple, Set, Optional
from pathlib import Path

import pandas as pd
from collections import defaultdict
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QToolBar, QLabel, QLineEdit, QPushButton,
    QFileDialog, QMessageBox, QProgressBar, QSizePolicy, QHBoxLayout, QCheckBox,
    QScrollArea, QFrame
)
from PySide6.QtGui import (
    QAction, QPalette, QColor, QDesktopServices, QKeySequence, QShortcut, QGuiApplication
)
from PySide6.QtCore import Qt, QUrl

from utils.app_utils import (
    normalize_key, get_tracker_path, get_tracker_path_2, get_company_csv_path, get_pending_invoices_dir, 
    refresh_common_paths,
    save_workbook_safe,
    APP_DIR,
)
from utils.logging_utils import get_logger, log_timing, new_run_id
from utils.paths_config import DEFAULT_INVOICE_TRACKER

from utils.app_utils import _force_excel_recalc as force_excel_recalc

logger = get_logger(__name__)

HA_LOOKUP_CACHE_VERSION = "v2_payment_line_amount"
TRACKER_RECALC_CACHE_VERSION = "v1"

# Constants
EXCLUDED_SHEETS: Tuple[str, ...] = ("Summary", "Vendor Credits")
HEADER_ROW_INDEX: int = 3
REQUIRED_COLS: Tuple[str, ...] = (
    "Invoice #","Vendor ID","Site","Unit","Invoice Status",
    "Overall Invoice Bill Totals","Date Invoice Due","Date Sent to Manager",
    "Remittance ID","Last Overdue Stage Actioned",
)

PENDING_ORDER: List[str] = [
    "Tracker ID","Vendor Name","Site","Unit","Invoice #","Date Due","Days Overdue",
    "Last Overdue Stage Actioned","Date Sent to Manager","Status","Amount Due",
    "Outstanding","Vendor ID","Primary Emails","Secondary Emails",
]

PRIMARY_SHEET_COLS: List[str] = [
    "Vendor Name", "Site", "Unit", "Invoice #",
    "Date Sent to Manager",
    "Primary Emails", "Secondary Emails",
]

# Email template constants
EMAIL_SUBJECT = "ACTION REQUIRED: {initials} - OVERDUE INVOICES"
EMAIL_BODY = """<html>
<body style="font-family: Calibri, sans-serif; font-size: 11pt; line-height: 1.6; color: #333;">
<p>Hello {first_name},</p>

<p>I hope you're doing well.</p>

<p>As part of our year-end reconciliation, a recent audit has identified outstanding invoices that remain unpaid. To ensure timely closure, we require your prompt assistance in reviewing and confirming the status of these items.</p>

<p>Please review the attached list and confirm by {reachby_date} whether each invoice has been re-submitted or is currently in process for payment. Timely confirmation is important to avoid further escalation and ensure accurate financial reporting.</p>

<p>If any invoices or supporting RFP documentation are required, please let me know and I will provide them right away.</p>

<p>Additionally, please feel free to forward this email to any other administrative assistants who should be included in this review.</p>

<p>Thank you in advance for your prompt attention to this matter. I look forward to your confirmation.</p>

<hr style="border: none; border-top: 1px solid #ccc; margin: 20px 0;">
<p style="font-size: 11pt; color: #666; font-style: italic;">
This report was generated automatically on {generation_date}
</p>
</body>
<body style="font-family:Calibri, sans-serif; font-size:10pt; line-height:1;">
    </span></p><p style="margin: 0;"class="MsoNormal"><b><span lang="EN-US" style="font-size:10.0pt;font-family:&quot;Arial&quot;,sans-serif;color:#31261D;mso-ligatures:standardcontextual">Thank you,</span></b><b><span lang="EN-US" style="font-size:10.0pt;font-family:&quot;Arial&quot;,sans-serif;color:#31261D"></span></b></p><p style="margin: 0;"class="MsoNormal"><b><span style="font-size:10.0pt;font-family:&quot;Arial&quot;,sans-serif;color:#005293;mso-ligatures:standardcontextual"><o:p>&nbsp;</o:p></span></b></p><p style="margin: 0;"class="MsoNormal"><b><span style="font-size:10.0pt;font-family:&quot;Arial&quot;,sans-serif;color:#005293;mso-ligatures:standardcontextual">Surbhi Singla (she/her/hers)</span></b><span style="font-size:11.0pt;color:black;mso-ligatures:standardcontextual"></span></p><p style="margin: 0;"class="MsoNormal"><b><span style="font-size:10.0pt;font-family:&quot;Arial&quot;,sans-serif;color:#31261D;mso-ligatures:standardcontextual">HR Assistant | Staffing Services</span></b><b><span lang="EN-US" style="font-size:10.0pt;font-family:&quot;Arial&quot;,sans-serif;color:#31261D;mso-ligatures:standardcontextual"></span></b></p><p style="margin: 0;"class="MsoNormal"><span lang="EN-US" style="font-size:10.0pt;font-family:&quot;Arial&quot;,sans-serif;color:#31261D;mso-ligatures:standardcontextual"><o:p>&nbsp;</o:p></span></p><p style="margin: 0;"class="MsoNormal"><b><span lang="EN-US" style="font-size:10.0pt;font-family:&quot;Arial&quot;,sans-serif;color:#31261D;mso-ligatures:standardcontextual">Fraser Health | Better Health. Best in health care.</span></b></p><p style="margin: 0;"class="MsoNormal"><b><span lang="EN-US" style="font-size:10.0pt;font-family:&quot;Arial&quot;,sans-serif;color:#31261D;mso-ligatures:standardcontextual"><o:p>&nbsp;</o:p></span></b></p><p style="margin: 0;"class="MsoNormal"><span style="font-size:10.0pt;font-family:&quot;Arial&quot;,sans-serif;color:#31261D;mso-ligatures:standardcontextual">Suite 500, Central City Tower, 13450 - 102 Avenue, Surrey, B.C. V3T 0A3</span></p><p style="margin: 0;"class="MsoNormal"><span lang="EN-US" style="font-size:10.0pt;font-family:&quot;Arial&quot;,sans-serif;color:#31261D;mso-ligatures:standardcontextual"><o:p>&nbsp;</o:p></span></p><p style="margin: 0;"class="MsoNormal"><span style="mso-ligatures:standardcontextual"><a href="https://can01.safelinks.protection.outlook.com/?url=http%3A%2F%2Fwww.fraserhealth.ca%2F&amp;data=05%7C02%7CFHAgencystaffing%40fraserhealth.ca%7Cf1f6a844c56746e6b62d08ddcbae0cf7%7C31f660a5192a4db392baca424f1b259e%7C0%7C0%7C638890671172338966%7CUnknown%7CTWFpbGZsb3d8eyJFbXB0eU1hcGkiOnRydWUsIlYiOiIwLjAuMDAwMCIsIlAiOiJXaW4zMiIsIkFOIjoiTWFpbCIsIldUIjoyfQ%3D%3D%7C0%7C%7C%7C&amp;sdata=ZLcOzCnXH2Sf97M8hasBHo7vQKE7mSmKriu%2F0KMAwPA%3D&amp;reserved=0" originalsrc="http://www.fraserhealth.ca/"><span lang="EN-US" style="font-size:10.0pt;font-family:&quot;Arial&quot;,sans-serif;color:#0563C1">www.fraserhealth.ca</span></a></span><span lang="EN-US" style="font-size:10.0pt;font-family:&quot;Arial&quot;,sans-serif;color:#31261D;mso-ligatures:standardcontextual"></span></p><p style="margin: 0;"class="MsoNormal"><span lang="EN-US" style="font-size:10.0pt;font-family:&quot;Arial&quot;,sans-serif;color:#31261D;mso-ligatures:standardcontextual"><o:p>&nbsp;</o:p></span></p><p style="margin: 0;"class="MsoNormal"><i><span style="font-size:10.0pt;font-family:&quot;Arial&quot;,sans-serif;mso-ligatures:standardcontextual">We recognize that Fraser Health provides care on the traditional, ancestral and unceded lands of the Coast Salish and Nlaka pamux Nations and is home to 32 First Nations within the Fraser Salish region. Fraser Health is dedicated to serving all Indigenous people, and honours the unique cultures of the First Nations, M tis and Inuit living within the Fraser Salish region.</span></i></p><p style="margin: 0;"class="MsoNormal"><span lang="EN-US" style="font-size:10.0pt;font-family:&quot;Arial&quot;,sans-serif;color:#31261D;mso-ligatures:standardcontextual"><o:p>&nbsp;</o:p></span></p><p style="margin: 0;"class="MsoNormal"><span lang="EN-US" style="font-size:10.0pt;font-family:&quot;Arial&quot;,sans-serif;color:#31261D;mso-ligatures:standardcontextual">CONFIDENTIALITY NOTICE: This email is intended only for the person(s) or entity to which it</span></p><p style="margin: 0;"class="MsoNormal"><span lang="EN-US" style="font-size:10.0pt;font-family:&quot;Arial&quot;,sans-serif;color:#31261D;mso-ligatures:standardcontextual">is addressed and may contain confidential information. Any review, distribution, copying, printing</span></p><p style="margin: 0;"class="MsoNormal"><span lang="EN-US" style="font-size:10.0pt;font-family:&quot;Arial&quot;,sans-serif;color:#31261D;mso-ligatures:standardcontextual">or other use of this email by anyone other than the named recipient is prohibited. If you have</span></p><p style="margin: 0;"class="MsoNormal"><span lang="EN-US" style="font-size:10.0pt;font-family:&quot;Arial&quot;,sans-serif;color:#31261D;mso-ligatures:standardcontextual">received this email in error or are not the named recipient, please notify the sender immediately</span></p><p style="margin: 0;"class="MsoNormal"><span lang="EN-US" style="font-size:10.0pt;font-family:&quot;Arial&quot;,sans-serif;color:#31261D;mso-ligatures:standardcontextual">and permanently delete this email and all copies of it. Thank you.</span></p><p style="margin: 0;"class="MsoNormal"><span style="font-size:11.0pt;font-family:&quot;Calibri&quot;,sans-serif"><o:p>&nbsp;</o:p></span></p><p><span style="font-size:11.0pt;font-family:&quot;Calibri&quot;,sans-serif"><o:p>&nbsp;</o:p></span>
</body>
</html>"""

def header_map(ws: Worksheet, header_row: int = HEADER_ROW_INDEX) -> Dict[str, str]:
    """Create mapping of header names to column letters."""
    return {
        (cell.value or "").strip(): get_column_letter(cell.column)
        for cell in ws[header_row]
        if cell.value and str(cell.value).strip()
    }

def sanitize_sheet_name(name: str) -> str:
    """Clean and truncate sheet name for Excel compatibility."""
    cleaned = "".join(ch for ch in str(name) if ch not in '[]*/\\?:').strip()
    name = cleaned.split('@')[0]
    return (name or "Email")[:31]

def clean_email_list(email_str: str) -> List[str]:
    """Parse and clean comma-separated email addresses."""
    if not email_str or str(email_str).lower() in ['nan', 'none', '']:
        return []
    return [email.strip() for email in str(email_str).split(',') if email.strip()]

def read_company_map(csv_path: str) -> Dict[str, str]:
    """Load vendor ID to company name mapping."""
    comp_df = pd.read_csv(csv_path, dtype=str)
    if not {"VendorID","CompanyName"}.issubset(comp_df.columns):
        raise ValueError("Company Info CSV must have 'VendorID' and 'CompanyName'.")
    return dict(zip(comp_df["VendorID"], comp_df["CompanyName"]))


@dataclass(frozen=True)
class HaLookup:
    raw_lookup: Dict[Tuple[str, str], Tuple[float, str]]
    norm_lookup: Dict[Tuple[str, str], Tuple[float, str]]
    norm_collisions: Set[Tuple[str, str]]
    source_rows: int


def _ha_fingerprint(ha_xlsx: str) -> str:
    """
    Fingerprint for caching HA lookups.
    Uses stat info (fast) rather than hashing the file contents (slow).
    """
    p = os.path.abspath(ha_xlsx or "")
    st = os.stat(p)
    return f"{Path(p).name}|{st.st_size}|{int(st.st_mtime)}"


def _tracker_recalc_marker_path(tracker_xlsx: str) -> Path:
    cache_dir = Path(APP_DIR) / ".invoice_check_cache"
    cache_dir.mkdir(parents=True, exist_ok=True)
    p = os.path.abspath(tracker_xlsx or "")
    st = os.stat(p)
    fp = f"{Path(p).name}_{st.st_size}_{int(st.st_mtime)}".replace(" ", "_")
    return cache_dir / f"tracker_recalc_{TRACKER_RECALC_CACHE_VERSION}_{fp}.ok"


def _is_tracker_recalc_fresh(tracker_xlsx: str, *, run_id: str) -> bool:
    try:
        return _tracker_recalc_marker_path(tracker_xlsx).is_file()
    except Exception:
        logger.exception(
            "Tracker recalc marker read failed | tracker=%s",
            tracker_xlsx,
            extra={"run_id": run_id},
        )
        return False


def _mark_tracker_recalc_fresh(tracker_xlsx: str, *, run_id: str) -> None:
    try:
        marker = _tracker_recalc_marker_path(tracker_xlsx)
        marker.touch(exist_ok=True)
    except Exception:
        logger.exception(
            "Tracker recalc marker write failed | tracker=%s",
            tracker_xlsx,
            extra={"run_id": run_id},
        )


def _ha_cache_path(tracker_xlsx: str, ha_xlsx: str) -> Path:
    cache_dir = Path(APP_DIR) / ".invoice_check_cache"
    cache_dir.mkdir(parents=True, exist_ok=True)
    safe_name = Path(ha_xlsx).stem.replace(" ", "_")
    fp = _ha_fingerprint(ha_xlsx).replace("|", "_")
    return cache_dir / f"ha_lookup_{HA_LOOKUP_CACHE_VERSION}_{safe_name}_{fp}.pkl"


def _load_ha_lookup_cache(tracker_xlsx: str, ha_xlsx: str, *, run_id: str) -> Optional[HaLookup]:
    cache_path = _ha_cache_path(tracker_xlsx, ha_xlsx)
    try:
        if cache_path.is_file():
            with open(cache_path, "rb") as f:
                obj = pickle.load(f)
            if isinstance(obj, HaLookup):
                logger.info("HA lookup cache hit | path=%s", str(cache_path), extra={"run_id": run_id})
                return obj
    except Exception:
        logger.exception("HA lookup cache read failed | path=%s", str(cache_path), extra={"run_id": run_id})
    return None


def _save_ha_lookup_cache(tracker_xlsx: str, ha_xlsx: str, lookup: HaLookup, *, run_id: str) -> None:
    cache_path = _ha_cache_path(tracker_xlsx, ha_xlsx)
    tmp = cache_path.with_suffix(cache_path.suffix + ".tmp")
    try:
        with open(tmp, "wb") as f:
            pickle.dump(lookup, f, protocol=pickle.HIGHEST_PROTOCOL)
        os.replace(tmp, cache_path)
        logger.info("HA lookup cache saved | path=%s", str(cache_path), extra={"run_id": run_id})
    except Exception:
        logger.exception("HA lookup cache write failed | path=%s", str(cache_path), extra={"run_id": run_id})
        try:
            if tmp.exists():
                tmp.unlink()
        except Exception:
            pass


def build_ha_lookup(ha_xlsx: str, *, run_id: str = "-") -> HaLookup:
    """Build lookup for HA payments by vendor ID and BOTH raw and normalized invoice number."""
    ha_df = pd.read_excel(ha_xlsx, dtype=str, engine="openpyxl")
    raw_paid = defaultdict(lambda: {"paid": 0.0, "rem": set()})
    norm_paid = defaultdict(lambda: {"paid": 0.0, "rem": set()})
    norm_to_raw = defaultdict(set)

    amount_col = "Payment Line Amount" if "Payment Line Amount" in ha_df.columns else "Payment Total Amount"
    if amount_col != "Payment Line Amount":
        logger.warning(
            "HA column 'Payment Line Amount' not found; falling back to '%s'",
            amount_col,
            extra={"run_id": run_id},
        )
    
    for _, r in ha_df.iterrows():
        vid = str(r.get("Vendor ID","")).strip()
        inv_raw = str(r.get("Invoice Number","")).strip()
        inv_raw_key = inv_raw.strip()
        inv_norm = normalize_key(inv_raw)
        amt_s = str(r.get(amount_col, "0")).replace("$", "").replace(",", "").strip()
        
        try:
            paid = float(amt_s) if amt_s else 0.0
        except (ValueError, TypeError):
            paid = 0.0
            
        rem = str(r.get("Last Issued Check Num","")).strip()
        if inv_raw_key:
            raw_paid[(vid, inv_raw_key)]["paid"] += paid
            if rem:
                raw_paid[(vid, inv_raw_key)]["rem"].add(rem)
        if inv_norm:
            norm_paid[(vid, inv_norm)]["paid"] += paid
            if rem:
                norm_paid[(vid, inv_norm)]["rem"].add(rem)
            if inv_raw_key:
                norm_to_raw[(vid, inv_norm)].add(inv_raw_key)

    norm_collisions: Set[Tuple[str, str]] = set()
    for (vid, inv_norm), raw_values in norm_to_raw.items():
        unique_raw = {v for v in raw_values if v}
        if len(unique_raw) > 1:
            norm_collisions.add((vid, inv_norm))
            logger.warning(
                "HA normalize collision | vendor_id=%s | normalized_invoice=%s | raw_invoices=%s",
                vid,
                inv_norm,
                sorted(unique_raw),
                extra={"run_id": run_id},
            )

    logger.info(
        "HA lookup built | source_rows=%s | raw_keys=%s | norm_keys=%s | normalize_collisions=%s | amount_column=%s",
        len(ha_df),
        len(raw_paid),
        len(norm_paid),
        len(norm_collisions),
        amount_col,
        extra={"run_id": run_id},
    )
    
    raw_lookup = {k: (round(v["paid"], 2), ",".join(sorted(v["rem"]))) for k, v in raw_paid.items()}
    norm_lookup = {k: (round(v["paid"], 2), ",".join(sorted(v["rem"]))) for k, v in norm_paid.items()}
    return HaLookup(
        raw_lookup=raw_lookup,
        norm_lookup=norm_lookup,
        norm_collisions=norm_collisions,
        source_rows=len(ha_df),
    )

def available_vendor_sheets(wb) -> List[str]:
    """Get list of vendor sheets (excluding summary sheets)."""
    return [s for s in wb.sheetnames if s not in EXCLUDED_SHEETS]

def valid_headers(hdr: Dict[str, str]) -> bool:
    """Check if worksheet has all required column headers."""
    return all(col in hdr for col in REQUIRED_COLS)

def safe_float_conversion(value, default=0.0) -> float:
    """Safely convert value to float."""
    if isinstance(value, (int, float)):
        return round(float(value), 2)
    try:
        txt = str(value or "").replace("$", "").replace(",", "").strip()
        return round(float(txt), 2) if txt else default
    except (ValueError, TypeError):
        return default

def extract_name_info(email: str) -> Tuple[str, str]:
    """Extract first name and initials from email format fname.lname@domain.com"""
    try:
        local_part = email.split('@')[0]
        name_parts = local_part.split('.')
        
        if len(name_parts) >= 2:
            first_name = name_parts[0].capitalize()
            last_name = name_parts[1].capitalize()
            initials = f"{first_name[0]}{last_name[0]}".upper()
            return first_name, initials
        else:
            first_name = name_parts[0].capitalize()
            initials = first_name[0].upper()
            return first_name, initials
            
    except (IndexError, AttributeError):
        return "there", "UN"

def create_individual_excel_file(df_subset: pd.DataFrame, primary_email: str, temp_dir: str) -> str:
    """Create temporary Excel file for a specific primary email group."""
    first_name, initials = extract_name_info(primary_email)
    
    filename = f"{initials} - OVERDUE INVOICES.xlsx"
    filepath = Path(temp_dir) / filename
    
    with pd.ExcelWriter(str(filepath), engine="xlsxwriter", 
                       datetime_format="mm/dd/yyyy", date_format="mm/dd/yyyy") as writer:
        df_subset.to_excel(writer, sheet_name="Pending Invoices", index=False)
        
        wb_x = writer.book
        ws_x = writer.sheets["Pending Invoices"]
        
        rows, cols = df_subset.shape
        if rows > 0 and cols > 0:
            ws_x.add_table(0, 0, rows, cols - 1, {
                "columns": [{"header": c} for c in df_subset.columns],
                "header_row": True,
                "style": "Table Style Medium 9",
            })
            
            col_idx = {c: i for i, c in enumerate(df_subset.columns)}
            widths = {
                "Vendor Name": 24, "Site": 10, "Unit": 30, "Invoice #": 24,
                "Last Overdue Stage Actioned": 26, "Date Sent to Manager": 30,
                "Primary Emails": 36, "Secondary Emails": 36,
            }
            
            fmt_date = wb_x.add_format({"num_format": "mm/dd/yyyy", "align": "center"})
            
            for col_name, width in widths.items():
                if col_name in col_idx:
                    col_letter = col_idx[col_name]
                    ws_x.set_column(col_letter, col_letter, width)
                    if col_name == "Date Sent to Manager":
                        ws_x.set_column(col_letter, col_letter, None, fmt_date)
    
    return str(filepath)

def create_outlook_draft(primary_emails: str, secondary_emails: str, attachment_path: str) -> bool:
    """Create Outlook draft email with specified recipients and attachment."""
    try:
        import win32com.client as win32
        
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        
        to_emails = '; '.join(clean_email_list(primary_emails))
        cc_emails = '; '.join(clean_email_list(secondary_emails))
        cc_emails = f"{cc_emails}; fhagencystaffing@fraserhealth.ca" if cc_emails else "fhagencystaffing@fraserhealth.ca"
        
        first_email = clean_email_list(primary_emails)[0] if clean_email_list(primary_emails) else ""
        first_name, initials = extract_name_info(first_email) if first_email else ("there", "UN")
        
        if to_emails:
            mail.To = to_emails
        if cc_emails:
            mail.CC = cc_emails

        mail.SentOnBehalfOfName = "fhagencystaffing@fraserhealth.ca"
        mail.Subject = EMAIL_SUBJECT.format(initials=initials)

        def add_business_days(start_date, business_days):
            current_date = start_date
            while business_days > 0:
                current_date += timedelta(days=1)
                if current_date.weekday() < 5:
                    business_days -= 1
            return current_date

        todays_date = datetime.now()
        reachby_date = add_business_days(todays_date, 3)
        reachby_date = reachby_date.strftime("%B %d")
        
        mail.HTMLBody = EMAIL_BODY.format(
            first_name=first_name,
            reachby_date=reachby_date,
            generation_date=datetime.now().strftime('%B %d, %Y at %I:%M %p')
        )
        
        if Path(attachment_path).exists():
            mail.Attachments.Add(attachment_path)
        
        mail.Display(True)
        return True
        
    except ImportError:
        QMessageBox.warning(None, "Missing Dependency", 
                          "pywin32 package is required for Outlook integration.\n"
                          "Install with: pip install pywin32")
        return False
    except Exception as e:
        logger.exception("Outlook draft creation failed")
        QMessageBox.warning(None, "Outlook Error", 
                          f"Failed to create Outlook draft:\n{str(e)}")
        return False


class InvoiceCheckPage(QWidget):
    """Generate Pending Invoices report with email functionality."""

    def __init__(self, controller):
        super().__init__()
        self.ctrl = controller
        
        self.tracker_file = get_tracker_path()
        self.tracker_file_2 = get_tracker_path_2()
        self.company_csv = get_company_csv_path()
        self.pending_output_dir = get_pending_invoices_dir()
        self.ha_file = ""
        
        self.email_buttons_data = []
        self.temp_files = []
        
        # *** FIX: Initialize email lookup ***
        self.email_lookup = {}
        self.email_table = None
        
        self._setup_ui()
        self._refresh_paths_from_config()
        
        esc = QShortcut(QKeySequence(Qt.Key_Escape), self)
        esc.activated.connect(self.ctrl.show_home)

    def _setup_ui(self) -> None:
        """Initialize the user interface."""
        self._apply_palette()
        self._apply_style()
        self._build_ui()

    def _build_ui(self) -> None:
        """Build the main UI layout."""
        root = QVBoxLayout(self)
        root.setContentsMargins(50, 40, 50, 40)
        root.setSpacing(20)

        self._add_toolbar(root)
        self._add_title(root)
        self._add_file_inputs(root)
        self._add_options(root)
        self._add_controls(root)
        self._add_email_section(root)

    def _add_toolbar(self, layout: QVBoxLayout) -> None:
        """Add navigation toolbar."""
        toolbar = QToolBar()
        back = QAction("← Back", self)
        back.triggered.connect(self.ctrl.show_home)
        toolbar.addAction(back)
        layout.addWidget(toolbar)

    def _add_title(self, layout: QVBoxLayout) -> None:
        """Add page title."""
        title = QLabel("📝 Invoice Check")
        title.setAlignment(Qt.AlignHCenter)
        title.setStyleSheet("font-size:32px; font-weight:bold; color:#3498db;")
        layout.addWidget(title)

    def _add_file_inputs(self, layout: QVBoxLayout) -> None:
        """Add file input controls."""
        self.track_le = QLineEdit(self.tracker_file)
        self.track_le.setReadOnly(True)
        row_t = QHBoxLayout()
        row_t.addWidget(QLabel("Tracker Excel:"), 1)
        row_t.addWidget(self.track_le, 5)
        layout.addLayout(row_t)

        row_t2 = QHBoxLayout()
        self.chk_use_tracker2 = QCheckBox("Include Tracker 2")
        self.chk_use_tracker2.setChecked(False)
        self.track2_le = QLineEdit(self.tracker_file_2)
        self.track2_le.setReadOnly(True)
        row_t2.addWidget(self.chk_use_tracker2, 1)
        row_t2.addWidget(self.track2_le, 5)
        layout.addLayout(row_t2)
        self.chk_use_tracker2.toggled.connect(self._sync_tracker2_ui)
        self._sync_tracker2_ui()

        self.out_dir_le = QLineEdit(self.pending_output_dir)
        self.out_dir_le.setReadOnly(True)
        row_o = QHBoxLayout()
        row_o.addWidget(QLabel("Pending Invoices Folder:"), 1)
        row_o.addWidget(self.out_dir_le, 5)
        layout.addLayout(row_o)

        self.ha_le = QLineEdit()
        self.ha_le.setReadOnly(True)
        self.btn_ha = QPushButton("Browse…")
        self.btn_ha.clicked.connect(lambda: self._choose_file(self.ha_le, "Select HA Payment Excel"))
        row_h = QHBoxLayout()
        row_h.addWidget(QLabel("HA Payment Excel:"), 1)
        row_h.addWidget(self.ha_le, 5)
        row_h.addWidget(self.btn_ha, 1)
        layout.addLayout(row_h)

    def _add_options(self, layout: QVBoxLayout) -> None:
        """Add option checkboxes."""
        chk_row = QHBoxLayout()
        self.chk_generate = QCheckBox("Generate Pending Invoices")
        self.chk_generate.setChecked(True)
        self.chk_generate.setToolTip("When checked, writes the Excel report to the Output Folder.")
        chk_row.addWidget(self.chk_generate)
        chk_row.addStretch(1)
        layout.addLayout(chk_row)

    def _add_controls(self, layout: QVBoxLayout) -> None:
        """Add run button and progress indicators."""
        self.run_btn = QPushButton("Run Check")
        self.run_btn.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        self.run_btn.clicked.connect(self.on_run)
        layout.addWidget(self.run_btn, alignment=Qt.AlignCenter)

        self.progress = QProgressBar()
        self.progress.setVisible(False)
        layout.addWidget(self.progress)

        status_row = QHBoxLayout()
        self.status_lbl = QLabel("")
        self.status_lbl.setStyleSheet("color:#555; font-style:italic;")
        self.status_lbl.setVisible(False)
        
        self.busy_bar = QProgressBar()
        self.busy_bar.setRange(0, 0)
        self.busy_bar.setVisible(False)
        
        status_row.addWidget(self.status_lbl, 5)
        status_row.addWidget(self.busy_bar, 3)
        layout.addLayout(status_row)

    def _add_email_section(self, layout: QVBoxLayout) -> None:
        """Add email functionality section."""
        self.email_section = QFrame()
        self.email_section.setStyleSheet(
            "QFrame { border: 1px solid #bdc3c7; border-radius: 6px; "
            "padding: 10px; background: #f8f9fa; }"
        )
        self.email_section.setVisible(False)
        
        email_layout = QVBoxLayout(self.email_section)
        
        self.email_scroll = QScrollArea()
        self.email_scroll.setWidgetResizable(True)
        self.email_scroll.setMaximumHeight(200)
        
        self.email_content = QWidget()
        self.email_content_layout = QVBoxLayout(self.email_content)
        self.email_scroll.setWidget(self.email_content)
        
        email_layout.addWidget(self.email_scroll)
        layout.addWidget(self.email_section)

    def _apply_palette(self) -> None:
        """Apply color palette to the widget."""
        pal = QPalette()
        pal.setColor(QPalette.Window, QColor(245, 245, 245))
        pal.setColor(QPalette.WindowText, QColor(33, 33, 33))
        pal.setColor(QPalette.Button, QColor(238, 238, 238))
        pal.setColor(QPalette.ButtonText, QColor(33, 33, 33))
        pal.setColor(QPalette.Highlight, QColor(52, 152, 219))
        pal.setColor(QPalette.HighlightedText, QColor(255, 255, 255))
        self.setPalette(pal)

    def _apply_style(self) -> None:
        """Apply CSS styling to the widget."""
        self.setStyleSheet("""
            QPushButton { 
                background:#2980b9; color:white; border:none; border-radius:6px;
                padding:8px 16px; font-size:14px; 
            }
            QPushButton:hover { background:#3498db; }
            QPushButton.email-btn { 
                background:#27ae60; padding:6px 12px; font-size:12px; 
            }
            QPushButton.email-btn:hover { background:#2ecc71; }
            QLineEdit { 
                border:1px solid #bdc3c7; border-radius:4px; padding:4px; 
            }
            QProgressBar { 
                border:1px solid #7f8c8d; border-radius:5px; height:20px; 
            }
        """)

    def showEvent(self, e):
        """Handle widget show event."""
        super().showEvent(e)
        self._refresh_paths_from_config()

    def _refresh_paths_from_config(self) -> None:
        """Refresh file paths from configuration."""
        refresh_common_paths(self)
        self._sync_tracker2_ui()

    def _sync_tracker2_ui(self) -> None:
        if hasattr(self, "track2_le") and hasattr(self, "chk_use_tracker2"):
            self.track2_le.setEnabled(self.chk_use_tracker2.isChecked())

    def _cleanup_temp_files(self):
        """Clean up temporary email attachment files."""
        for file_path in self.temp_files:
            try:
                if Path(file_path).exists():
                    Path(file_path).unlink()
            except Exception:
                pass
        self.temp_files.clear()

    def _create_email_buttons(self, pending_df: pd.DataFrame):
        """Create email draft buttons for each unique primary email."""
        for i in reversed(range(self.email_content_layout.count())): 
            self.email_content_layout.itemAt(i).widget().setParent(None)
        
        self._cleanup_temp_files()
        self.email_buttons_data.clear()
        
        if "Primary Emails" not in pending_df.columns or pending_df.empty:
            self.email_section.setVisible(False)
            return
        
        temp_dir = tempfile.mkdtemp(prefix="invoice_emails_")
        
        primary_emails = sorted(pending_df["Primary Emails"].astype(str).str.strip().unique())
        
        for primary_email in primary_emails:
            if not primary_email or primary_email.lower() in ['nan', 'none', '']:
                continue
                
            email_data = pending_df[pending_df["Primary Emails"].astype(str).str.strip() == primary_email]
            email_subset = email_data[[c for c in PRIMARY_SHEET_COLS if c in email_data.columns]].copy()
            
            if email_subset.empty:
                continue
            
            secondary_emails: Set[str] = set()
            for _, row in email_data.iterrows():
                sec_emails = clean_email_list(row.get("Secondary Emails", ""))
                secondary_emails.update(sec_emails)
            
            secondary_emails_combined = ', '.join(sorted(secondary_emails))
            
            try:
                temp_file = create_individual_excel_file(email_subset, primary_email, temp_dir)
                self.temp_files.append(temp_file)
                
                btn_text = f"📧 Email to: {primary_email} ({len(email_subset)} invoices)"
                email_btn = QPushButton(btn_text)
                email_btn.setProperty("class", "email-btn")
                email_btn.clicked.connect(
                    lambda checked, prim=primary_email, sec=secondary_emails_combined, file=temp_file: 
                    self._send_email_draft(prim, sec, file)
                )
                
                self.email_content_layout.addWidget(email_btn)
                
            except Exception as e:
                logger.exception("Error creating temp file for email | recipient=%s", primary_email)
                continue
        
        self.email_section.setVisible(self.email_content_layout.count() > 0)

    def _send_email_draft(self, primary_emails: str, secondary_emails: str, attachment_path: str):
        """Create Outlook draft for specific email group."""
        if not Path(attachment_path).exists():
            QMessageBox.warning(self, "File Not Found", "Attachment file no longer exists.")
            return
            
        success = create_outlook_draft(primary_emails, secondary_emails, attachment_path)
        if success:
            to_display = primary_emails[:50] + "..." if len(primary_emails) > 50 else primary_emails
            QMessageBox.information(self, "Email Draft Created", 
                                  f"Outlook draft created for {to_display}")

    def _set_controls_enabled(self, enabled: bool) -> None:
        """Enable/disable UI controls during processing."""
        controls = [
            self.run_btn, self.track_le, self.ha_le, self.out_dir_le, self.chk_generate,
            getattr(self, "track2_le", None), getattr(self, "chk_use_tracker2", None),
            getattr(self, "btn_track", None), getattr(self, "btn_ha", None), 
            getattr(self, "btn_outdir", None)
        ]
        
        for control in controls:
            if control is not None:
                control.setEnabled(enabled)

    def _begin_busy(self, msg: str) -> None:
        """Start busy state with loading message."""
        self.status_lbl.setText(msg)
        self.status_lbl.setVisible(True)
        self.busy_bar.setVisible(True)
        QGuiApplication.setOverrideCursor(Qt.WaitCursor)
        self._set_controls_enabled(False)
        QGuiApplication.processEvents()

    def _update_busy(self, msg: str) -> None:
        """Update busy state message."""
        self.status_lbl.setText(msg)
        QGuiApplication.processEvents()

    def _end_busy(self) -> None:
        """End busy state and restore normal UI."""
        self.status_lbl.clear()
        self.status_lbl.setVisible(False)
        self.busy_bar.setVisible(False)
        QGuiApplication.restoreOverrideCursor()
        self._set_controls_enabled(True)
        QGuiApplication.processEvents()

    def _choose_file(self, line_edit: QLineEdit, title: str) -> None:
        """Open file dialog and update line edit."""
        path, _ = QFileDialog.getOpenFileName(
            self, title, str(Path.home() / "Downloads"),
            "Excel Files (*.xlsx *.xls)"
        )
        if path:
            line_edit.setText(path)

    def _selected_tracker_paths(self) -> List[Tuple[int, str]]:
        paths: List[Tuple[int, str]] = [(1, self.tracker_file)]
        if getattr(self, "chk_use_tracker2", None) and self.chk_use_tracker2.isChecked():
            tracker2_path = (self.track2_le.text().strip() if getattr(self, "track2_le", None) else self.tracker_file_2)
            paths.append((2, tracker2_path))
        return paths

    def _validate_inputs(self) -> bool:
        """Validate required input files and directories."""
        self.tracker_file = (self.track_le.text().strip() or self.tracker_file)
        self.tracker_file_2 = (self.track2_le.text().strip() if getattr(self, "track2_le", None) else self.tracker_file_2)

        required_files: List[Tuple[str, str]] = [("Tracker", self.tracker_file), ("HA Payment", self.ha_file)]
        if getattr(self, "chk_use_tracker2", None) and self.chk_use_tracker2.isChecked():
            required_files.append(("Tracker 2", self.tracker_file_2))

        for label, filepath in required_files:
            if not Path(filepath).is_file():
                QMessageBox.warning(self, f"Missing {label} File", 
                                  f"Please select a valid {label} file.")
                return False
        
        if self.chk_generate.isChecked() and not Path(self.pending_output_dir).is_dir():
            QMessageBox.warning(self, "Invalid Output Folder", 
                              "Please select a valid output folder.")
            return False
        
        return True

    def _process_sheet_data(self, ws_val: Worksheet, ws_edit: Worksheet, 
                          hdr: Dict[str, str], ha_lookup: HaLookup, vendor_map: Dict,
                          pending_records: List[Dict], *, run_id: str, tracker_id: int) -> Dict[str, int]:
        """Process a single worksheet for pending invoices."""
        today = datetime.today().date()
        
        col_map = {key: hdr[key] for key in REQUIRED_COLS if key in hdr}
        stats = {
            "rows_seen": 0,
            "rows_with_invoice": 0,
            "status_changed": 0,
            "due0_count": 0,
            "pending_included": 0,
            "ambiguous_norm_hits": 0,
        }
        
        for r in range(HEADER_ROW_INDEX + 1, ws_edit.max_row + 1):
            stats["rows_seen"] += 1
            row_data = self._process_invoice_row(
                r, ws_val, ws_edit, col_map, ha_lookup, vendor_map, today, tracker_id=tracker_id
            )
            
            if not row_data:
                continue

            stats["rows_with_invoice"] += 1
            if row_data.get("status_changed"):
                stats["status_changed"] += 1
            if row_data.get("due0"):
                stats["due0_count"] += 1
            if row_data.get("ambiguous_norm_hit"):
                stats["ambiguous_norm_hits"] += 1
            if row_data.get("should_include_in_pending"):
                stats["pending_included"] += 1
                pending_records.append(row_data["record"])

        logger.info(
            "Sheet scanned | sheet=%s | rows_seen=%s | rows_with_invoice=%s | status_changed=%s | due0=%s | pending_included=%s | ambiguous_norm_hits=%s",
            ws_edit.title,
            stats["rows_seen"],
            stats["rows_with_invoice"],
            stats["status_changed"],
            stats["due0_count"],
            stats["pending_included"],
            stats["ambiguous_norm_hits"],
            extra={"run_id": run_id},
        )
        return stats

    def _process_invoice_row(self, row_num: int, ws_val: Worksheet, ws_edit: Worksheet,
                           col_map: Dict[str, str], ha_lookup: HaLookup, vendor_map: Dict, 
                           today, *, tracker_id: int) -> Dict:
        """Process a single invoice row."""
        site_val = str(ws_edit[f"{col_map['Site']}{row_num}"].value or "").strip()
        unit_val = str(ws_edit[f"{col_map['Unit']}{row_num}"].value or "").strip()
        inv_no = str(ws_edit[f"{col_map['Invoice #']}{row_num}"].value or "").strip()
        vid = str(ws_edit[f"{col_map['Vendor ID']}{row_num}"].value or "").strip()
        overdue_stage = str(ws_edit[f"{col_map['Last Overdue Stage Actioned']}{row_num}"].value or "").strip()
        
        if not inv_no:
            return None
        
        amt_due = safe_float_conversion(ws_val[f"{col_map['Overall Invoice Bill Totals']}{row_num}"].value)
        due_raw = ws_val[f"{col_map['Date Invoice Due']}{row_num}"].value
        sent_raw = ws_val[f"{col_map['Date Sent to Manager']}{row_num}"].value
        
        due_date = pd.to_datetime(due_raw, errors="coerce").date() if due_raw else None
        sent_date = pd.to_datetime(sent_raw, errors="coerce").date() if sent_raw else None
        days_overdue = (today - sent_date).days if sent_date else None
        
        prev_status = str(ws_edit[f"{col_map['Invoice Status']}{row_num}"].value or "").strip()

        # Prefer exact invoice matching first to avoid false 'Paid' via normalization collisions.
        inv_raw_key = inv_no.strip()
        inv_norm = normalize_key(inv_no)
        paid_amt = 0.0
        rem_str = ""
        lookup_hit = False
        lookup_source = "none"
        ambiguous_norm_hit = False

        if vid and inv_raw_key:
            raw_key = (vid, inv_raw_key)
            if raw_key in ha_lookup.raw_lookup:
                paid_amt, rem_str = ha_lookup.raw_lookup.get(raw_key, (0.0, ""))
                lookup_hit = True
                lookup_source = "raw"
            elif inv_norm:
                norm_key = (vid, inv_norm)
                if norm_key in ha_lookup.norm_lookup:
                    if norm_key in ha_lookup.norm_collisions:
                        # Ambiguous: multiple raw invoices share a normalized key.
                        ambiguous_norm_hit = True
                        lookup_hit = False
                        lookup_source = "ambiguous_norm"
                    else:
                        paid_amt, rem_str = ha_lookup.norm_lookup.get(norm_key, (0.0, ""))
                        lookup_hit = True
                        lookup_source = "norm"
        
        tol = 0.05
        remaining = round(amt_due - paid_amt, 2)

        # For positive invoices, any zero-or-negative remaining balance is considered paid.
        if abs(amt_due) <= tol:
            new_status = "Due 0"
        elif amt_due > tol:
            if paid_amt <= tol:
                new_status = "Not paid"
            elif remaining <= tol:
                new_status = "Paid"
            else:
                new_status = "Partially paid"
        else:
            # Credit notes: keep status behavior conservative and tolerance-based.
            if abs(remaining) <= tol:
                new_status = "Paid"
            elif abs(paid_amt) > tol:
                new_status = "Partially paid"
            else:
                new_status = "Not paid"
        
        try:
            ws_edit[f"{col_map['Invoice Status']}{row_num}"].value = new_status
            if col_map.get('Remittance ID') and rem_str:
                ws_edit[f"{col_map['Remittance ID']}{row_num}"].value = rem_str
        except Exception:
            pass
        
        # Report unresolved magnitude for both debit invoices and credit notes.
        outstanding_amt = round(abs(remaining), 2) if new_status in {"Not paid", "Partially paid"} else 0.0
        primary, secondary = self.email_lookup.get((site_val, unit_val), ("", ""))
        
        should_include = (
            days_overdue is not None
            and days_overdue > 30
            and outstanding_amt >= 1.00
            and new_status in {"Not paid", "Partially paid"}
        )

        status_changed = (prev_status.strip() != new_status.strip())
        due0 = (new_status == "Due 0")

        # Only emit row-level debug logs when it helps investigation.
        if logger.isEnabledFor(10) and (should_include or status_changed or ambiguous_norm_hit):
            logger.debug(
                "Row evaluated | sheet=%s | row=%s | vendor_id=%s | invoice=%s | inv_norm=%s | lookup_hit=%s | lookup_source=%s | ambiguous_norm=%s | amount_due=%.2f | paid_amount=%.2f | outstanding=%.2f | days_overdue=%s | prev_status=%s | new_status=%s",
                ws_edit.title,
                row_num,
                vid,
                inv_no,
                inv_norm,
                lookup_hit,
                lookup_source,
                ambiguous_norm_hit,
                amt_due,
                paid_amt,
                outstanding_amt,
                days_overdue,
                prev_status,
                new_status,
                extra={"run_id": getattr(self, "_run_id", "-")},
            )
        
        return {
            "should_include_in_pending": should_include,
            "status_changed": status_changed,
            "due0": due0,
            "ambiguous_norm_hit": ambiguous_norm_hit,
            "record": {
                "Tracker ID": tracker_id,
                "Vendor Name": vendor_map.get(vid, ""),
                "Site": site_val, "Unit": unit_val,
                "Invoice #": inv_no, "Date Due": due_date, "Days Overdue": days_overdue,
                "Last Overdue Stage Actioned": overdue_stage, "Date Sent to Manager": sent_date,
                "Status": new_status, "Amount Due": amt_due, "Outstanding": outstanding_amt,
                "Vendor ID": vid, "Primary Emails": primary, "Secondary Emails": secondary,
            } if should_include else None
        }

    def _write_pending_report(self, pending_df: pd.DataFrame, output_path: str) -> None:
        """Write the pending invoices Excel report with formatting."""
        with pd.ExcelWriter(output_path, engine="xlsxwriter", 
                           datetime_format="mm/dd/yyyy", date_format="mm/dd/yyyy") as writer:
            
            pending_df.to_excel(writer, sheet_name="Pending Overdue", index=False)
            wb_x = writer.book
            ws_x = writer.sheets["Pending Overdue"]
            
            rows, cols = pending_df.shape
            if rows > 0:
                ws_x.add_table(0, 0, rows, cols - 1, {
                    "columns": [{"header": c} for c in pending_df.columns],
                    "header_row": True,
                    "style": "Table Style Medium 9",
                })
                
                self._format_main_sheet(ws_x, wb_x, pending_df)
            
            self._create_email_sheets(writer, pending_df, wb_x)

    def _format_main_sheet(self, ws_x, wb_x, pending_df: pd.DataFrame) -> None:
        """Apply formatting to the main pending invoices sheet."""
        col_idx = {c: i for i, c in enumerate(pending_df.columns)}
        
        widths = {
            "Tracker ID": 10,
            "Vendor Name": 26, "Site": 16, "Unit": 14, "Invoice #": 18,
            "Date Due": 14, "Days Overdue": 14, "Last Overdue Stage Actioned": 26,
            "Date Sent to Manager": 26, "Status": 14,
            "Amount Due": 14, "Outstanding": 14, "Vendor ID": 12,
            "Primary Emails": 28, "Secondary Emails": 32,
        }
        
        fmt_cur = wb_x.add_format({"num_format": "$* #,##0.00", "align": "right"})
        fmt_date = wb_x.add_format({"num_format": "mm/dd/yyyy", "align": "center"})
        
        for col_name, width in widths.items():
            if col_name in col_idx:
                col_idx_num = col_idx[col_name]
                ws_x.set_column(col_idx_num, col_idx_num, width)
                
                if col_name in ("Amount Due", "Outstanding"):
                    ws_x.set_column(col_idx_num, col_idx_num, None, fmt_cur)
                elif col_name in ("Date Due", "Date Sent to Manager"):
                    ws_x.set_column(col_idx_num, col_idx_num, None, fmt_date)

    def _create_email_sheets(self, writer, pending_df: pd.DataFrame, wb_x) -> None:
        """Create individual sheets for each primary email group."""
        if "Primary Emails" not in pending_df.columns or pending_df.empty:
            return
        
        seen_names: Set[str] = set()
        fmt_date = wb_x.add_format({"num_format": "mm/dd/yyyy", "align": "center"})
        
        for primary_email in sorted(pending_df["Primary Emails"].astype(str).str.strip().unique()):
            if not primary_email or primary_email.lower() in ['nan', 'none', '']:
                continue
                
            email_subset = pending_df[pending_df["Primary Emails"].astype(str).str.strip() == primary_email]
            email_data = email_subset[[c for c in PRIMARY_SHEET_COLS if c in email_subset.columns]].copy()
            
            if email_data.empty:
                continue
            
            sheet_name = self._generate_unique_sheet_name(primary_email, seen_names, writer.sheets)
            seen_names.add(sheet_name)
            
            email_data.to_excel(writer, sheet_name=sheet_name, index=False)
            ws_email = writer.sheets[sheet_name]
            
            self._format_email_sheet(ws_email, email_data, fmt_date)

    def _generate_unique_sheet_name(self, primary_email: str, seen_names: Set[str], existing_sheets: Dict) -> str:
        """Generate a unique Excel sheet name."""
        base = sanitize_sheet_name(primary_email)
        sheet_name, counter = base, 1
        
        while sheet_name in seen_names or sheet_name in existing_sheets:
            suffix = f"_{counter}"
            max_base_len = 31 - len(suffix)
            sheet_name = base[:max_base_len] + suffix
            counter += 1
            
        return sheet_name

    def _format_email_sheet(self, ws_email, email_data: pd.DataFrame, fmt_date) -> None:
        """Apply formatting to individual email sheets."""
        rows, cols = email_data.shape
        if rows > 0:
            ws_email.add_table(0, 0, rows, cols - 1, {
                "columns": [{"header": c} for c in email_data.columns],
                "header_row": True,
                "style": "Table Style Medium 9",
            })
            
            email_col_idx = {c: i for i, c in enumerate(email_data.columns)}
            email_widths = {
                "Vendor Name": 24, "Site": 10, "Unit": 30, "Invoice #": 24,
                "Last Overdue Stage Actioned": 26, "Date Sent to Manager": 30,
                "Primary Emails": 36, "Secondary Emails": 36,
            }
            
            for col_name, width in email_widths.items():
                if col_name in email_col_idx:
                    col_idx = email_col_idx[col_name]
                    ws_email.set_column(col_idx, col_idx, width)
                    if col_name == "Date Sent to Manager":
                        ws_email.set_column(col_idx, col_idx, None, fmt_date)

    def _load_email_table(self):
        """*** FIX: Load email table WITHOUT opening a new workbook ***"""
        try:
            df = pd.read_excel(
                self.tracker_file or DEFAULT_INVOICE_TRACKER,
                sheet_name="Summary", 
                header=None, 
                engine="openpyxl"
            )
            
            col0 = df.iloc[:, 0].astype(str).str.strip()
            idx = col0[col0 == "Site Information"]
            
            if idx.empty:
                self.email_table = pd.DataFrame(columns=[
                    "Site","Unit","Manager Email","Secondary Email"
                ])
                self.email_lookup = {}
                return

            start = idx.index[0] + 1
            headers = [str(h).strip() for h in df.iloc[start].tolist()]

            rows = []
            r = start + 1
            while r < len(df) and str(df.iat[r, 0]).strip():
                rows.append(df.iloc[r, :len(headers)].tolist())
                r += 1

            et = pd.DataFrame(rows, columns=headers)
            et = et.loc[:, ~et.columns.duplicated(keep="first")]
            et.columns = [c.strip() for c in et.columns]

            def pick(colnames, needle):
                needle = needle.lower()
                for c in colnames:
                    if c.lower() == needle:
                        return c
                for c in colnames:
                    if needle in c.lower():
                        return c
                return None

            col_site = pick(et.columns, "site")
            col_unit = pick(et.columns, "unit")
            col_mgr = pick(et.columns, "manager email")
            col_sec = pick(et.columns, "secondary email")

            for need, nm in [("Site", col_site), ("Unit", col_unit)]:
                if not nm:
                    self.email_table = pd.DataFrame(columns=[
                        "Site","Unit","Manager Email","Secondary Email"
                    ])
                    self.email_lookup = {}
                    return

            keep = [col_site, col_unit]
            if col_mgr: keep.append(col_mgr)
            if col_sec: keep.append(col_sec)
            et = et[keep].copy()

            ren = {}
            ren[col_site] = "Site"
            ren[col_unit] = "Unit"
            if col_mgr: ren[col_mgr] = "Manager Email"
            if col_sec: ren[col_sec] = "Secondary Email"
            et.rename(columns=ren, inplace=True)

            self.email_table = et

            self.email_lookup = {}
            for _, row in et.iterrows():
                key = (str(row.get("Site","")).strip(), str(row.get("Unit","")).strip())

                if str(row.get("Secondary Email","") or "").strip() == "nan":
                    secondaryEmail = ""
                else:
                    secondaryEmail = str(row.get("Secondary Email","") or "").strip()

                self.email_lookup[key] = (
                    str(row.get("Manager Email","") or "").strip(),
                    secondaryEmail
                )
        except Exception as e:
            logger.exception("Error loading email table")
            self.email_lookup = {}
            self.email_table = None

    def on_run(self) -> None:
        """Main processing with optional dual-tracker support."""
        run_id = new_run_id()
        self._run_id = run_id
        t0 = time.perf_counter()
        self.ha_file = self.ha_le.text().strip()
        self.pending_output_dir = (self.out_dir_le.text().strip() or str(Path(self.tracker_file).parent))

        tracker_jobs: List[Tuple[int, str]] = []
        
        # Load email table before tracker scans.
        self._load_email_table()
        
        if not self._validate_inputs():
            return

        tracker_jobs = self._selected_tracker_paths()
        logger.info(
            "Invoice check started | trackers=%s | ha_file=%s | output_dir=%s | generate_report=%s",
            tracker_jobs,
            self.ha_file,
            self.pending_output_dir,
            self.chk_generate.isChecked(),
            extra={"run_id": run_id},
        )

        try:
            self._begin_busy("Loading files…")

            self._update_busy("Loading company map…")
            with log_timing(logger, "Load company map", run_id=run_id):
                vendor_map = read_company_map(self.company_csv)
            logger.info("Company map loaded | vendors=%s", len(vendor_map), extra={"run_id": run_id})

            pending_records = []

            for tracker_id, tracker_path in tracker_jobs:
                self._update_busy(f"Indexing HA payments for Tracker {tracker_id}…")
                ha_lookup = _load_ha_lookup_cache(tracker_path, self.ha_file, run_id=run_id)
                if ha_lookup is None:
                    with log_timing(logger, f"Build HA lookup (Tracker {tracker_id})", run_id=run_id):
                        ha_lookup = build_ha_lookup(self.ha_file, run_id=run_id)
                    _save_ha_lookup_cache(tracker_path, self.ha_file, ha_lookup, run_id=run_id)

                logger.info(
                    "HA lookup ready | tracker_id=%s | raw_keys=%s | norm_keys=%s | collisions=%s",
                    tracker_id,
                    len(ha_lookup.raw_lookup),
                    len(ha_lookup.norm_lookup),
                    len(ha_lookup.norm_collisions),
                    extra={"run_id": run_id},
                )

                self._process_single_tracker(
                    tracker_path=tracker_path,
                    tracker_id=tracker_id,
                    ha_lookup=ha_lookup,
                    vendor_map=vendor_map,
                    pending_records=pending_records,
                    run_id=run_id,
                )
            
            pending_df = self._create_pending_dataframe(pending_records)
            total_ms = int((time.perf_counter() - t0) * 1000)
            logger.info(
                "Invoice check complete | pending_records=%s | total_elapsed_ms=%s",
                len(pending_df),
                total_ms,
                extra={"run_id": run_id},
            )
            
            self._create_email_buttons(pending_df)
            
            if self.chk_generate.isChecked():
                self._generate_report(pending_df)
            else:
                self._show_summary_only(pending_df)
                
        except Exception as e:
            self.progress.setVisible(False)
            logger.exception("Invoice check failed", extra={"run_id": run_id})
            QMessageBox.critical(self, "Error", f"Unexpected error: {e}")
        finally:
            self._end_busy()
            self._run_id = "-"

    def _process_single_tracker(self, tracker_path: str, tracker_id: int, ha_lookup: HaLookup,
                                vendor_map: Dict, pending_records: List[Dict], *, run_id: str) -> None:
        """Process one tracker file and append pending records with tracker_id."""
        wb_data = None
        wb_edit = None
        stage_path = None

        try:
            tracker_dir = str(Path(tracker_path).parent)
            fd, stage_path = tempfile.mkstemp(dir=tracker_dir, suffix=".~staged.xlsx")
            os.close(fd)
            os.unlink(stage_path)

            with log_timing(logger, f"Stage tracker copy ({tracker_id})", run_id=run_id):
                shutil.copy2(tracker_path, stage_path)

            if _is_tracker_recalc_fresh(tracker_path, run_id=run_id):
                logger.info(
                    "Excel recalc (pre-scan) skipped | tracker_id=%s | reason=fresh_tracker_cache",
                    tracker_id,
                    extra={"run_id": run_id},
                )
            else:
                self._update_busy(f"Refreshing tracker calculations (Tracker {tracker_id})…")
                with log_timing(logger, f"Excel recalc pre-scan ({tracker_id})", run_id=run_id):
                    force_excel_recalc(stage_path)

            self._update_busy(f"Opening Tracker {tracker_id}…")
            with log_timing(logger, f"Open tracker workbooks ({tracker_id})", run_id=run_id):
                wb_data = load_workbook(stage_path, data_only=True, read_only=True)
                wb_edit = load_workbook(stage_path, data_only=False)

            sheets = available_vendor_sheets(wb_data)
            logger.info(
                "Tracker opened | tracker_id=%s | tracker=%s | sheets=%s",
                tracker_id,
                tracker_path,
                len(sheets),
                extra={"run_id": run_id},
            )

            with log_timing(logger, f"Scan tracker sheets ({tracker_id})", run_id=run_id):
                self._process_tracker_sheets(
                    wb_data,
                    wb_edit,
                    sheets,
                    ha_lookup,
                    vendor_map,
                    pending_records,
                    run_id=run_id,
                    tracker_id=tracker_id,
                )

            wb_data.close()
            wb_data = None

            self._save_tracker_updates(wb_edit, stage_path, run_id=run_id)
            wb_edit = None

            self._update_busy(f"Finalizing tracker calculations (Tracker {tracker_id})…")
            with log_timing(logger, f"Excel recalc post-save ({tracker_id})", run_id=run_id):
                force_excel_recalc(stage_path)

            os.replace(stage_path, tracker_path)
            _mark_tracker_recalc_fresh(tracker_path, run_id=run_id)
            logger.info(
                "Tracker replaced from staged file | tracker_id=%s | tracker=%s",
                tracker_id,
                tracker_path,
                extra={"run_id": run_id},
            )
            stage_path = None

        finally:
            try:
                if wb_data is not None:
                    wb_data.close()
                if wb_edit is not None:
                    wb_edit.close()
            except Exception as close_error:
                logger.warning(
                    "Error while closing workbooks | tracker_id=%s | error=%s",
                    tracker_id,
                    close_error,
                    extra={"run_id": run_id},
                )

            if stage_path and os.path.exists(stage_path):
                try:
                    os.unlink(stage_path)
                except Exception:
                    pass

    def _process_tracker_sheets(self, wb_data, wb_edit, sheets: List[str], 
                              ha_lookup: HaLookup, vendor_map: Dict, pending_records: List, *, run_id: str, tracker_id: int) -> None:
        """Process all tracker sheets for pending invoices."""
        total_sheets = len(sheets)
        self.progress.setRange(0, total_sheets)
        self.progress.setValue(0)
        self.progress.setVisible(True)
        
        self._update_busy("Scanning tracker sheets…")

        totals = {
            "rows_seen": 0,
            "rows_with_invoice": 0,
            "status_changed": 0,
            "due0_count": 0,
            "pending_included": 0,
            "ambiguous_norm_hits": 0,
        }
        
        for idx, sheet_name in enumerate(sheets, start=1):
            ws_val = wb_data[sheet_name]
            ws_edit = wb_edit[sheet_name]
            
            hdr = header_map(ws_val)
            if valid_headers(hdr):
                sheet_stats = self._process_sheet_data(
                    ws_val, ws_edit, hdr, ha_lookup, vendor_map, pending_records,
                    run_id=run_id, tracker_id=tracker_id
                )
                for k, v in sheet_stats.items():
                    totals[k] = totals.get(k, 0) + v
            
            self.progress.setValue(idx)
            if idx % 3 == 0:
                self._update_busy(f"Scanning tracker sheets… {idx}/{total_sheets}")

        logger.info(
            "Scan summary | sheets=%s | rows_seen=%s | rows_with_invoice=%s | status_changed=%s | due0=%s | pending_included=%s | ambiguous_norm_hits=%s",
            total_sheets,
            totals["rows_seen"],
            totals["rows_with_invoice"],
            totals["status_changed"],
            totals["due0_count"],
            totals["pending_included"],
            totals["ambiguous_norm_hits"],
            extra={"run_id": run_id},
        )

    def _save_tracker_updates(self, wb_edit, save_path: str, *, run_id: str) -> None:
        """Save updates to a workbook path (closes wb_edit on return)."""
        try:
            logger.info("Saving staged tracker workbook | path=%s", save_path, extra={"run_id": run_id})
            with log_timing(logger, "Save staged tracker workbook", run_id=run_id):
                save_workbook_safe(wb_edit, save_path)
            logger.info("Staged tracker workbook saved | path=%s", save_path, extra={"run_id": run_id})
        except Exception as e:
            logger.exception("Tracker save/refresh failed | path=%s", save_path, extra={"run_id": run_id})
            QMessageBox.warning(self, "Save Warning", f"Tracker save/refresh issue:\n{e}")

    def _create_pending_dataframe(self, pending_records: List[Dict]) -> pd.DataFrame:
        """Create and format the pending invoices DataFrame."""
        pending_df = pd.DataFrame(pending_records)
        
        if not pending_df.empty:
            pending_df = pending_df.loc[:, ~pending_df.columns.duplicated(keep="first")]
            ordered_cols = [c for c in PENDING_ORDER if c in pending_df.columns]
            pending_df = pending_df[ordered_cols] if ordered_cols else pending_df
        
        return pending_df

    def _generate_report(self, pending_df: pd.DataFrame) -> None:
        """Generate the Excel report file."""
        self._update_busy("Writing Pending Invoices report…")
        
        report_name = f"Pending Invoices - {datetime.today().strftime('%B %d, %Y')}.xlsx"
        report_path = Path(self.pending_output_dir) / report_name
        
        self._write_pending_report(pending_df, str(report_path))
        
        # QDesktopServices.openUrl(QUrl.fromLocalFile(str(report_path)))
        self.progress.setVisible(False)
        QMessageBox.information(self, "Done", f"Pending Invoices written to: {report_path}")

    def _show_summary_only(self, pending_df: pd.DataFrame) -> None:
        """Show summary without generating report file."""
        self.progress.setVisible(False)
        QMessageBox.information(
            self, "Scan Complete",
            f"Found {len(pending_df)} pending invoice(s).\n"
            f"(Report generation was skipped.)"
        )