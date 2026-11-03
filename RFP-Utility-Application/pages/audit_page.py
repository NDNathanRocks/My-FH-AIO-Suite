# audit_page.py

from __future__ import annotations

import os
import re
import difflib
import math
from datetime import datetime
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

import pandas as pd
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.utils import get_column_letter

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QToolBar, QLabel, QLineEdit, QPushButton, QStackedLayout,
    QHeaderView, QFileDialog, QMessageBox, QProgressBar, QSizePolicy, QHBoxLayout,
    QTableWidget, QTableWidgetItem, QGridLayout, QCheckBox
)
from PySide6.QtGui import (
    QAction, QPalette, QColor, QDesktopServices, QKeySequence, QShortcut, QGuiApplication
)
from PySide6.QtCore import Qt, QUrl, QTimer

from utils.app_utils import normalize_key, load_config, save_config, _force_excel_recalc, get_tracker_path, get_tracker_path_2, get_company_csv_path, refresh_common_paths, ensure_tracker_closed_interactive


# ──────────────────────────────────────────────────────────────────────────────
# Constants
# ──────────────────────────────────────────────────────────────────────────────
EXCLUDED_SHEETS: Tuple[str, ...] = ("Summary", "Vendor Credits")
HEADER_ROW_INDEX: int = 3  # 1‑based row number containing headers in tracker sheets

RESULT_HEADERS: List[str] = [
    "Vendor Name", "Invoice Number",
    "Site", "Unit",
    "Date Invoice Due", "Date Sent to Manager",
    "Days Overdue", "Amount Due",
    "Amount Paid", "Outstanding Amount",
    "Remittance ID",
]

SITE_HEADERS: List[str] = RESULT_HEADERS.copy()

SINGLE_LOOKUP_FIELDS: List[str] = [
    "Vendor Name", "Invoice Number", "Site", "Unit",
    "Date Invoice Due", "Date Sent to Manager",
    "Days Overdue", "Amount Due", "Amount Paid",
    "Outstanding Amount", "Remittance ID",
]

# Aliases
VENDOR_ALIASES  = ["Vendor ID", "VendorID", "Vendor Id", "Vendor", "VID", "Vid"]
INVOICE_ALIASES = ["Invoice Number", "Invoice #", "Invoice", "Inv #", "Inv", "InvoiceNumber", "Invoice No"]

# ──────────────────────────────────────────────────────────────────────────────
# Helpers
# ──────────────────────────────────────────────────────────────────────────────

def sanitize_sheet_name(name: str) -> str:
    cleaned = re.sub(r"[\[\]\*\/\\\?:]", "", name)
    return cleaned[:31]

def read_company_map(csv_path: str) -> Dict[str, str]:
    comp_df = pd.read_csv(csv_path, dtype=str)
    required = {"VendorID", "CompanyName"}
    if not required.issubset(comp_df.columns):
        raise ValueError("Company Info CSV must have 'VendorID' and 'CompanyName'.")
    return dict(zip(comp_df["VendorID"], comp_df["CompanyName"]))

def build_ha_lookup(ha_xlsx: str) -> Dict[Tuple[str, str], Tuple[float, str]]:
    ha_df = pd.read_excel(ha_xlsx, dtype=str, engine="openpyxl")
    lookup: Dict[Tuple[str, str], Tuple[float, str]] = {}
    for _, r in ha_df.iterrows():
        vid = str(r.get("Vendor ID", "")).strip()
        inv_norm = normalize_key(str(r.get("Invoice Number", "")).strip())
        amt_str = str(r.get("Payment Total Amount", "0")).replace("$", "").replace(",", "").strip()
        try:
            paid_amt = float(amt_str) if amt_str else 0.0
        except Exception:
            paid_amt = 0.0
        paid_amt = round(paid_amt, 2)
        rem_id = r.get("Last Issued Check Num", "")
        if pd.isna(rem_id) or str(rem_id).strip().lower() in ("nan", "none", "nat"):  # blank cell (D5)
            rem_id = ""
        lookup[(vid, inv_norm)] = (paid_amt, rem_id)
    return lookup

def open_tracker_workbook(path: str):
    _force_excel_recalc(path)
    return load_workbook(path, data_only=True)

def available_vendor_sheets(wb) -> List[str]:
    return [s for s in wb.sheetnames if s not in EXCLUDED_SHEETS]

def best_sheet_match(vendor_name: str, sheet_names: Sequence[str]) -> Optional[str]:
    if not sheet_names:
        return None
    best = difflib.get_close_matches(vendor_name.lower(), [s.lower() for s in sheet_names], n=1, cutoff=0.5)
    if not best:
        return None
    return sheet_names[[s.lower() for s in sheet_names].index(best[0])]

def header_map(ws: Worksheet, header_row: int = HEADER_ROW_INDEX) -> Dict[str, str]:
    row = ws[header_row]
    mapping = {
        (cell.value or "").strip(): get_column_letter(cell.column)
        for cell in row if cell.value and str(cell.value).strip()
    }
    return mapping

def find_invoice_row(ws: Worksheet, inv_col_letter: str, invoice_number: str) -> Optional[int]:
    inv_key = normalize_key(invoice_number)
    for r in range(HEADER_ROW_INDEX + 1, ws.max_row + 1):
        cell_val = ws[f"{inv_col_letter}{r}"].value
        if normalize_key(str(cell_val or "")) == inv_key:
            return r
    return None

def to_datetime_or_none(value) -> Optional[datetime]:
    return value if isinstance(value, datetime) else None

def resolve_first_column(columns, aliases):
    cmap = {str(c).strip().lower(): c for c in columns}
    for a in aliases:
        key = a.strip().lower()
        if key in cmap:
            return cmap[key]
    return None

# ──────────────────────────────────────────────────────────────────────────────
# UI Class
# ──────────────────────────────────────────────────────────────────────────────

class AuditPage(QWidget):
    """Wizard-style Audit page: Files → (Lookup | Quarterly Report)."""

    def __init__(self, controller):
        super().__init__()
        self.ctrl = controller

        # File placeholders
        self.tracker_file       = get_tracker_path()
        self.tracker_file_2     = get_tracker_path_2()
        self.company_csv        = get_company_csv_path()
        self.audit_file: str = ""
        self.ha_file: str = ""

        # UI refs
        self.root_stack: QStackedLayout | None = None
        self.progress: QProgressBar | None = None
        self.status_lbl: QLabel | None = None
        self.busy_bar: QProgressBar | None = None

        # Step 0 (files) widgets
        self.track_le: QLineEdit | None = None
        self.track2_le: QLineEdit | None = None
        self.chk_use_tracker2: QCheckBox | None = None
        self.ha_le: QLineEdit | None = None

        # Build
        self._apply_palette()
        self._apply_style()
        self._build_ui()

        # First-time refresh (also happens every time page is shown)
        self._refresh_paths_from_config()

        # ESC → home
        esc = QShortcut(QKeySequence(Qt.Key_Escape), self)
        esc.activated.connect(self.ctrl.show_home)

    # ────────────────────────── UI builders ──────────────────────────
    def showEvent(self, e):
        super().showEvent(e)
        # Every time the page becomes visible, re-pull paths from the config file
        self._refresh_paths_from_config()

    def _refresh_paths_from_config(self) -> None:
        """Pull latest paths from disk and refresh UI + dependent data."""
        refresh_common_paths(self)   # sets: tracker_file, company_csv, template and line-edits
        self._sync_tracker_selection_ui()

    def _selected_tracker_path(self) -> str:
        if self.chk_use_tracker2 and self.chk_use_tracker2.isChecked():
            return (self.track2_le.text().strip() if self.track2_le else self.tracker_file_2)
        return (self.track_le.text().strip() if self.track_le else self.tracker_file)

    def _sync_tracker_selection_ui(self) -> None:
        if self.track2_le is None or self.chk_use_tracker2 is None:
            return
        using_tracker2 = self.chk_use_tracker2.isChecked()
        self.track2_le.setEnabled(using_tracker2)

    def _build_ui(self) -> None:
        main = QVBoxLayout(self)
        main.setContentsMargins(40, 20, 40, 20)
        main.setSpacing(15)

        # Toolbar
        toolbar = QToolBar()
        back = QAction("← Back", self)
        back.triggered.connect(self.ctrl.show_home)
        toolbar.addAction(back)
        main.addWidget(toolbar)

        # Title
        title = QLabel("📊 Audit & Lookup")
        title.setAlignment(Qt.AlignHCenter)
        title.setStyleSheet("font-size:32px; font-weight:bold; color:#3498db;")
        main.addWidget(title)

        # Root stack: 0 = files, 1 = lookup, 2 = quarterly
        self.root_stack = QStackedLayout()
        main.addLayout(self.root_stack)

        # Global status row
        status_row = QHBoxLayout()
        self.status_lbl = QLabel("")
        self.status_lbl.setStyleSheet("color:#555; font-style:italic;")
        self.status_lbl.setVisible(False)
        self.busy_bar = QProgressBar(); self.busy_bar.setRange(0, 0); self.busy_bar.setVisible(False)
        status_row.addWidget(self.status_lbl, 5)
        status_row.addWidget(self.busy_bar, 3)
        main.addLayout(status_row)

        # Build pages
        self._build_files_page()       # index 0
        self._build_lookup_page()      # index 1
        self._build_quarterly_page()   # index 2

        self.root_stack.setCurrentIndex(0)

    def _build_files_page(self) -> None:
        """Select Tracker + HA Payment, centered UX with wide path fields."""
        w = QWidget()
        outer = QVBoxLayout(w)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        # Inner container holds the actual UI
        container = QWidget()
        container.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)

        lo = QVBoxLayout(container)
        lo.setSpacing(16)
        # add generous side padding so the controls "reach" the edges nicely
        lo.setContentsMargins(40, 0, 40, 0)

        intro = QLabel("Select your Tracker and HA Payment files")
        intro.setAlignment(Qt.AlignHCenter)
        intro.setStyleSheet("font-size:18px; color:#34495e;")
        lo.addWidget(intro)

        # ── Tracker row (path expands, button optional) ────────────────
        self.track_le = QLineEdit(self.tracker_file)
        lo.addWidget(self._build_file_row(
            "Tracker Excel:",
            self.track_le, with_button=False
        ))

        tracker2_row = QWidget()
        tracker2_row_layout = QHBoxLayout(tracker2_row)
        tracker2_row_layout.setContentsMargins(0, 0, 0, 0)
        tracker2_row_layout.setSpacing(8)
        self.chk_use_tracker2 = QCheckBox("Use Tracker 2")
        self.track2_le = QLineEdit(self.tracker_file_2)
        self.track2_le.setReadOnly(True)
        tracker2_row_layout.addWidget(self.chk_use_tracker2, 0)
        tracker2_row_layout.addWidget(self.track2_le, 1)
        lo.addWidget(tracker2_row)
        self.chk_use_tracker2.toggled.connect(self._sync_tracker_selection_ui)
        self._sync_tracker_selection_ui()

        # ── HA row (same treatment) ────────────────────────────────────
        self.ha_le = QLineEdit()
        lo.addWidget(self._build_file_row(
            "HA Payment Excel:",
            self.ha_le,
            lambda: self._choose_file(self.ha_le, "Select HA Payment Excel")
        ))

        # ── Actions (two big buttons) ──────────────────────────────────
        row_actions = QHBoxLayout()
        btn_lookup = QPushButton("Go to Lookup")
        btn_audit  = QPushButton("Go to Quarterly Report")
        for b in (btn_lookup, btn_audit):
            b.setMinimumHeight(40)
            b.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        btn_lookup.clicked.connect(self._go_to_lookup)
        btn_audit.clicked.connect(self._go_to_quarterly)
        row_actions.addWidget(btn_lookup)
        row_actions.addWidget(btn_audit)
        lo.addLayout(row_actions)

        # Vertically center the whole thing; let it fill width
        outer.addStretch(1)
        outer.addWidget(container)  # no side stretches → full width
        outer.addStretch(1)

        self.root_stack.addWidget(w)


    def _build_file_row(self, label_text: str, line_edit: QLineEdit,
                        on_browse=None, with_button: bool = True) -> QWidget:
        row = QWidget()
        row.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        grid = QGridLayout(row)
        grid.setContentsMargins(0, 0, 0, 0)
        grid.setHorizontalSpacing(8)

        lbl = QLabel(label_text)

        line_edit.setReadOnly(True)
        line_edit.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        grid.addWidget(lbl,       0, 0)
        grid.addWidget(line_edit, 0, 1)
        grid.setColumnStretch(0, 0)  # label stays compact
        grid.setColumnStretch(1, 1)  # line edit eats remaining width

        if with_button and on_browse:
            btn = QPushButton("Browse…")
            btn.setFixedWidth(90)
            btn.clicked.connect(on_browse)
            grid.addWidget(btn, 0, 2)
            grid.setColumnStretch(2, 0)

        return row

    def _build_lookup_page(self) -> None:
        page = QWidget(); lo = QVBoxLayout(page); lo.setSpacing(10)

        # Back to files
        nav = QHBoxLayout()
        back = QPushButton("← Files"); back.clicked.connect(lambda: self.root_stack.setCurrentIndex(0))
        nav.addWidget(back, alignment=Qt.AlignLeft)
        lo.addLayout(nav)

        # Title
        t = QLabel("Lookup Invoice Status")
        t.setAlignment(Qt.AlignHCenter)
        t.setStyleSheet("font-size:18px; color:#34495e;")
        lo.addWidget(t)

        # Inputs
        form = QHBoxLayout()
        form.addWidget(QLabel("Vendor ID:"), 1)
        self.single_vid = QLineEdit(); form.addWidget(self.single_vid, 2)
        form.addWidget(QLabel("Invoice #:"), 1)
        self.single_inv = QLineEdit(); form.addWidget(self.single_inv, 2)
        btn = QPushButton("Lookup"); btn.clicked.connect(self._on_single_lookup)
        form.addWidget(btn, 1)
        lo.addLayout(form)

        # Results table
        self.lookup_fields = SINGLE_LOOKUP_FIELDS
        self.result_table = QTableWidget(len(self.lookup_fields), 2)
        self.result_table.setHorizontalHeaderLabels(["Field", "Value"])
        self.result_table.verticalHeader().setVisible(False)
        hdr = self.result_table.horizontalHeader(); hdr.setSectionResizeMode(QHeaderView.Stretch)
        for row, field in enumerate(self.lookup_fields):
            item = QTableWidgetItem(field); item.setFlags(item.flags() ^ Qt.ItemIsEditable)
            self.result_table.setItem(row, 0, item)
        lo.addWidget(self.result_table)

        self.root_stack.addWidget(page)
        
    def _build_quarterly_page(self) -> None:
        # Outer wrapper
        w = QWidget()
        outer = QVBoxLayout(w)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        # Inner container (fills full width)
        container = QWidget()
        container.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)

        lo = QVBoxLayout(container)
        # match files page: generous side padding, full width
        lo.setContentsMargins(40, 0, 40, 0)
        lo.setSpacing(16)

        # Back to files
        nav = QHBoxLayout()
        back = QPushButton("← Files")
        back.clicked.connect(lambda: self.root_stack.setCurrentIndex(0))
        nav.addWidget(back, alignment=Qt.AlignLeft)
        lo.addLayout(nav)

        # Title
        title = QLabel("Quarterly Report")
        title.setAlignment(Qt.AlignHCenter)
        title.setStyleSheet("font-size:18px; color:#34495e;")
        lo.addWidget(title)

        # Audit file selector (uses full-width row builder)
        self.audit_le = QLineEdit()
        self.audit_le.setReadOnly(True)
        lo.addWidget(self._build_file_row(
            "Audit File:",
            self.audit_le,
            lambda: self._choose_file(self.audit_le, "Select Audit Excel")
        ))

        # Run Audit
        self.run_btn = QPushButton("Run Audit")
        self.run_btn.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        self.run_btn.clicked.connect(self.on_run)
        lo.addWidget(self.run_btn, alignment=Qt.AlignCenter)

        # Progress bar
        self.progress = QProgressBar()
        self.progress.setVisible(False)
        lo.addWidget(self.progress)

        # Fill width: no horizontal centering wrapper here
        outer.addStretch(1)
        outer.addWidget(container)   # <— fills the full width
        outer.addStretch(1)

        self.root_stack.addWidget(w)


    # ────────────────────────── Styling ──────────────────────────
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
            QPushButton:hover { background:#3498db; }
            QLineEdit { border:1px solid #bdc3c7; border-radius:4px; padding:4px; }
            QProgressBar { border:1px solid #7f8c8d; border-radius:5px; height:20px; }
            QTableWidget { background:white; alternate-background-color:#f9f9f9; }
            QHeaderView::section { background:#2980b9; color:white; padding:4px; font-weight:bold; border:none; }
            """
        )

    # ────────────────────────── Busy helpers ──────────────────────────
    def _set_controls_enabled(self, enabled: bool) -> None:
        # keep it simple — main interactive widgets
        for w in [self.run_btn, self.audit_le, self.track_le, self.track2_le, self.chk_use_tracker2, self.ha_le]:
            if w is not None:
                w.setEnabled(enabled)

    def _begin_busy(self, msg: str = "Working…") -> None:
        self.status_lbl.setText(msg)
        self.status_lbl.setVisible(True)
        self.busy_bar.setVisible(True)
        QGuiApplication.setOverrideCursor(Qt.WaitCursor)
        self._set_controls_enabled(False)
        QGuiApplication.processEvents()

    def _update_busy(self, msg: str) -> None:
        self.status_lbl.setText(msg)
        QGuiApplication.processEvents()

    def _end_busy(self) -> None:
        self.status_lbl.clear()
        self.status_lbl.setVisible(False)
        self.busy_bar.setVisible(False)
        QGuiApplication.restoreOverrideCursor()
        self._set_controls_enabled(True)
        QGuiApplication.processEvents()

    # ────────────────────────── File selectors / navigation ──────────────────────────
    def _choose_file(self, line_edit: QLineEdit, title: str) -> None:
        path, _ = QFileDialog.getOpenFileName(self, title, os.path.join(os.path.expanduser("~"), "Downloads"), "Excel Files (*.xlsx *.xls)")
        if path:
            line_edit.setText(path)

    def _ensure_files_selected(self) -> bool:
        # Read current values from the step-0 line edits into member paths
        if self.track_le and self.track_le.text().strip():
            self.tracker_file = self.track_le.text().strip()
        if self.track2_le and self.track2_le.text().strip():
            self.tracker_file_2 = self.track2_le.text().strip()
        if self.ha_le and self.ha_le.text().strip():
            self.ha_file = self.ha_le.text().strip()

        selected_tracker = self._selected_tracker_path()
        self.tracker_file = selected_tracker

        # Validate
        if not os.path.isfile(selected_tracker):
            QMessageBox.warning(self, "Missing Tracker Excel", "Please select a valid selected tracker Excel file.")
            return False
        if not os.path.isfile(self.ha_file):
            QMessageBox.warning(self, "Missing HA Payment File", "Please select a valid HA Payment Excel file.")
            return False
        return True

    def _go_to_lookup(self) -> None:
        if self._ensure_files_selected():
            self.root_stack.setCurrentIndex(1)

    def _go_to_quarterly(self) -> None:
        if self._ensure_files_selected():
            self.root_stack.setCurrentIndex(2)

    # ────────────────────────── Single lookup ──────────────────────────
    def _on_single_lookup(self) -> None:

        vid = self.single_vid.text().strip()
        inv = self.single_inv.text().strip()
        if not vid or not inv:
            QMessageBox.warning(self, "Input Required", "Enter both VendorID & Invoice#.")
            return
        try:
            self._begin_busy("Looking up invoice…")
            tracker_path = self._selected_tracker_path()
            if not os.path.isfile(tracker_path):
                QMessageBox.warning(self, "Missing Tracker Excel", "Please select a valid selected tracker Excel file.")
                return

            # Vendor → Name
            comp_map = read_company_map(self.company_csv)
            vendor_name = comp_map.get(vid)
            if not vendor_name:
                QMessageBox.warning(self, "Unknown Vendor", f"No company for {vid}")
                return

            # HA payments
            if not os.path.isfile(self.ha_file):
                QMessageBox.warning(self, "Missing HA File", "Select HA Payment file on the previous step.")
                return
            self._update_busy("Loading HA payments…")
            ha_lookup = build_ha_lookup(self.ha_file)

            # Tracker
            self._update_busy("Loading tracker workbook…")
            wb = open_tracker_workbook(tracker_path)
            sheets = available_vendor_sheets(wb)
            sheet = best_sheet_match(vendor_name, sheets)
            if not sheet:
                QMessageBox.warning(self, "No Sheet", f"No sheet for {vendor_name}")
                return
            ws = wb[sheet]

            hdr = header_map(ws)
            needed = ("Invoice #", "Site", "Unit", "Date Invoice Due", "Date Sent to Manager", "Overall Invoice Bill Totals")
            if not all(c in hdr for c in needed):
                QMessageBox.warning(self, "Bad Sheet", "Tracker missing columns.")
                return

            self._update_busy("Finding invoice in tracker…")
            row_idx = find_invoice_row(ws, hdr["Invoice #"], inv)
            if not row_idx:
                QMessageBox.information(self, "Not Found", f"{inv} not in {vendor_name}")
                return

            s = hdr
            site = ws[f"{s['Site']}{row_idx}"].value or ""
            unit = ws[f"{s['Unit']}{row_idx}"].value or ""
            due = ws[f"{s['Date Invoice Due']}{row_idx}"].value
            sent = ws[f"{s['Date Sent to Manager']}{row_idx}"].value
            tot = ws[f"{s['Overall Invoice Bill Totals']}{row_idx}"].value or 0
            amt_due = round(float(tot), 2)

            paid_amt, rem = ha_lookup.get((vid, normalize_key(inv)), (0.00, ""))
            out_amt = round(amt_due - paid_amt, 2)

            days_overdue = ""
            if isinstance(due, datetime):
                days_overdue = (datetime.today().date() - due.date()).days

            due_str = due.strftime("%m/%d/%Y") if isinstance(due, datetime) else ""
            sent_str = sent.strftime("%m/%d/%Y") if isinstance(sent, datetime) else ""

            if paid_amt == 0.0:
                paid_amt = "0.00"
            if out_amt == 0.0:
                out_amt = "0.00"

            values = [
                vendor_name, inv, site, unit, due_str, sent_str,
                days_overdue, amt_due, paid_amt, out_amt, rem,
            ]

            for row, val in enumerate(values):
                item = QTableWidgetItem(str(val or ""))
                item.setFlags(item.flags() ^ Qt.ItemIsEditable)
                self.result_table.setItem(row, 1, item)

            self._update_busy("Done.")

        except Exception as e:
            QMessageBox.critical(self, "Lookup Error", str(e))
        finally:
            self._end_busy()

    # ────────────────────────── Quarterly audit ──────────────────────────
    def on_run(self) -> None:
        
        self.audit_file = self.audit_le.text().strip()
        tracker_path = self._selected_tracker_path()

        if not ensure_tracker_closed_interactive(self, tracker_path):
            return False

        validations = [
            (self.audit_file, "Missing Audit File", "Please select a valid Audit Excel file."),
            (tracker_path, "Missing Tracker Excel", "Please select a valid selected tracker Excel file (Step 1)."),
            (self.ha_file, "Missing HA Payment File", "Please select a valid HA Payment Excel file (Step 1)."),
            (self.company_csv, "Missing Company Info", "Cannot find Company Info CSV (with VendorID column)."),
        ]
        for path, title, msg in validations:
            if not os.path.isfile(path):
                QMessageBox.warning(self, title, msg)
                return

        try:
            self._begin_busy("Loading files…")

            # 1) Load Audit
            audit_df = pd.read_excel(self.audit_file, dtype=str)
            vendor_col  = resolve_first_column(audit_df.columns, VENDOR_ALIASES)
            invoice_col = resolve_first_column(audit_df.columns, INVOICE_ALIASES)

            # 2) Vendor map
            self._update_busy("Loading company map…")
            vendor_map = read_company_map(self.company_csv)

            # 3) Tracker workbook + sheets
            self._update_busy("Opening tracker workbook…")
            wb_data = open_tracker_workbook(tracker_path)
            sheet_names = available_vendor_sheets(wb_data)

            # 4) HA lookup
            self._update_busy("Indexing HA payments…")
            ha_lookup = build_ha_lookup(self.ha_file)

            # 5) Prepare result dataframe
            result_df = pd.DataFrame(columns=RESULT_HEADERS)

            # Progress
            total_rows = len(audit_df)
            self.progress.setRange(0, total_rows)
            self.progress.setValue(0)
            self.progress.setVisible(True)

            # 6) Iterate audit rows
            self._update_busy("Processing rows…")
            for idx, row in audit_df.iterrows():
                vendor_id = str(row.get(vendor_col, "")).strip()
                inv_raw   = str(row.get(invoice_col, "")).strip()
                vendor_name = vendor_map.get(vendor_id, "")

                default_row = [
                    vendor_name, inv_raw,
                    "Invoice Not in Tracker", "",
                    None, None,
                    "", 0.0, 0.0, 0.0,
                    "",
                ]

                if not vendor_name:
                    result_df.loc[idx] = pd.Series(dict(zip(RESULT_HEADERS, default_row)))
                    self.progress.setValue(idx + 1)
                    if (idx + 1) % 25 == 0:
                        self._update_busy(f"Processing rows… {idx + 1}/{total_rows}")
                    continue

                sheet = best_sheet_match(vendor_name, sheet_names)
                if not sheet:
                    result_df.loc[idx] = default_row
                    self.progress.setValue(idx + 1)
                    if (idx + 1) % 25 == 0:
                        self._update_busy(f"Processing rows… {idx + 1}/{total_rows}")
                    continue

                ws = wb_data[sheet]
                hdr = header_map(ws)
                inv_col = hdr.get("Invoice #")
                site_col = hdr.get("Site")
                unit_col = hdr.get("Unit")
                due_col = hdr.get("Date Invoice Due")
                sent_col = hdr.get("Date Sent to Manager")
                total_col = hdr.get("Overall Invoice Bill Totals")
                if None in (inv_col, due_col, sent_col, total_col):
                    result_df.loc[idx] = default_row
                    self.progress.setValue(idx + 1)
                    if (idx + 1) % 25 == 0:
                        self._update_busy(f"Processing rows… {idx + 1}/{total_rows}")
                    continue

                match_row_idx = find_invoice_row(ws, inv_col, inv_raw)
                if match_row_idx is None:
                    result_df.loc[idx] = default_row
                    self.progress.setValue(idx + 1)
                    if (idx + 1) % 25 == 0:
                        self._update_busy(f"Processing rows… {idx + 1}/{total_rows}")
                    continue

                # Extract row values
                site_cell = ws[f"{site_col}{match_row_idx}"].value or ""
                unit_cell = ws[f"{unit_col}{match_row_idx}"].value or ""
                due_cell = ws[f"{due_col}{match_row_idx}"].value
                sent_cell = ws[f"{sent_col}{match_row_idx}"].value
                tot_val = ws[f"{total_col}{match_row_idx}"].value
                amt_due = round(float(tot_val or 0), 2)

                paid_amt, rem_id = ha_lookup.get((vendor_id, normalize_key(inv_raw)), (0.0, ""))
                out_amt = round((amt_due - paid_amt), 2)

                row_dict = {
                    "Vendor Name": vendor_name,
                    "Invoice Number": inv_raw,
                    "Site": site_cell,
                    "Unit": unit_cell,
                    "Date Invoice Due": to_datetime_or_none(due_cell),
                    "Date Sent to Manager": to_datetime_or_none(sent_cell),
                    "Days Overdue": "",  # formula later
                    "Amount Due": amt_due,
                    "Amount Paid": paid_amt,
                    "Outstanding Amount": out_amt,
                    "Remittance ID": rem_id,
                }
                result_df.loc[idx] = pd.Series(row_dict)

                self.progress.setValue(idx + 1)
                if (idx + 1) % 25 == 0:
                    self._update_busy(f"Processing rows… {idx + 1}/{total_rows}")

            # Normalize date dtypes
            for date_col in ("Date Invoice Due", "Date Sent to Manager"):
                result_df[date_col] = (
                    pd.to_datetime(result_df[date_col], errors="coerce")
                    .apply(lambda ts: ts.to_pydatetime() if not pd.isna(ts) else None)
                )

            # 7) Write Excel output
            self._update_busy("Writing Excel output…")
            output_path = self._write_output_workbook(self.audit_file, result_df)

            QDesktopServices.openUrl(QUrl.fromLocalFile(output_path))
            self.progress.setVisible(False)
            QMessageBox.information(self, "Done", f"Audit saved to: {output_path}")

        except Exception as e:
            self.progress.setVisible(False)
            QMessageBox.critical(self, "Error", f"An unexpected error occurred: {e}")
        finally:
            self._end_busy()

    # ────────────────────────── Excel writers ──────────────────────────
    def _write_output_workbook(self, audit_file: str, result_df: pd.DataFrame) -> str:
        audit_dir = os.path.dirname(audit_file)
        base = os.path.splitext(os.path.basename(audit_file))[0]
        today_str = datetime.today().strftime("%B %d, %Y")
        out_name = f"{base} - {today_str}.xlsx"
        output_path = os.path.join(audit_dir, out_name)

        wb = xlsxwriter.Workbook(output_path)
        try:
            self._write_main_audit_sheet(wb, result_df)
            self._write_quarterly_report_sheet(wb, result_df)
            self._write_per_site_sheets(wb, result_df)
        finally:
            wb.close()
        return output_path

    def _formats(self, wb: xlsxwriter.Workbook) -> Dict[str, xlsxwriter.format.Format]:
        return {
            "date": wb.add_format({"num_format": "mm/dd/yyyy", "align": "center"}),
            "cur": wb.add_format({"num_format": "$* #,##0.00", "align": "right"}),
            "center": wb.add_format({"align": "center"}),
            "green": wb.add_format({"bg_color": "#ccffcc"}),
            "red": wb.add_format({"bg_color": "#ffcccc"}),
            "title": wb.add_format({"bold": True, "font_size": 14, "color": "#3498db"}),
        }

    def _write_main_audit_sheet(self, wb: xlsxwriter.Workbook, result_df: pd.DataFrame) -> None:
        ws = wb.add_worksheet("Audit")
        fmts = self._formats(wb)
        headers = RESULT_HEADERS

        nrows, ncols = result_df.shape
        data = result_df.values.tolist()

        cols = []
        for h in headers:
            cd = {"header": h}
            if h == "Vendor Name":
                cd["total_string"] = "Totals"
            if h == "Invoice Number":
                cd["total_function"] = "count"
            if h in ("Date Invoice Due", "Date Sent to Manager"):
                cd["format"] = fmts["date"]
            if h == "Days Overdue":
                cd["format"] = fmts["center"]
            if h in ("Amount Due", "Amount Paid", "Outstanding Amount"):
                cd.update({"format": fmts["cur"], "total_function": "sum"})
            cols.append(cd)

        tbl = []
        due_idx = headers.index("Date Invoice Due")
        days_idx = headers.index("Days Overdue")
        out_idx = headers.index("Outstanding Amount")
        amt_idx = headers.index("Amount Due")
        paid_idx = headers.index("Amount Paid")
        for r in range(nrows):
            row = data[r].copy()
            due_ref = xl_rowcol_to_cell(r + 1, due_idx)
            row[days_idx] = f"=TODAY()-{due_ref}"
            amt_ref = xl_rowcol_to_cell(r + 1, amt_idx)
            paid_ref = xl_rowcol_to_cell(r + 1, paid_idx)
            row[out_idx] = f"={amt_ref}-{paid_ref}"
            tbl.append(row)

        for r, row in enumerate(tbl):
            for c, val in enumerate(row):
                if isinstance(val, pd._libs.tslibs.timestamps.Timestamp):
                    row[c] = val.to_pydatetime()
                elif pd.isna(val):
                    row[c] = None

        ws.add_table(
            0, 0, nrows + 1, ncols - 1,
            {
                "data": tbl,
                "columns": cols,
                "header_row": True,
                "total_row": True,
                "style": "Table Style Medium 9",
                "banded_rows": True,
                "autofilter": True,
            },
        )

        days_col_letter = xl_rowcol_to_cell(0, days_idx)[0]
        days_range = f"{days_col_letter}2:{days_col_letter}{nrows + 1}"
        out_col_letter = xl_rowcol_to_cell(0, out_idx)[0]
        out_range = f"{out_col_letter}2:{out_col_letter}{nrows + 1}"

        ws.conditional_format(days_range, {"type": "cell", "criteria": "<", "value": 0, "format": fmts["green"]})
        ws.conditional_format(days_range, {"type": "cell", "criteria": ">=", "value": 0, "format": fmts["red"]})
        ws.conditional_format(out_range, {"type": "cell", "criteria": "=", "value": 0, "format": fmts["green"]})
        ws.conditional_format(out_range, {"type": "cell", "criteria": ">", "value": 0, "format": fmts["red"]})

        for i, h in enumerate(headers):
            ws.set_column(i, i, 25 if h == "Vendor Name" else 15)

    def _write_quarterly_report_sheet(self, wb: xlsxwriter.Workbook, result_df: pd.DataFrame) -> None:
        ws = wb.add_worksheet("Quarterly Report")
        fmts = self._formats(wb)

        df = result_df.copy()
        df["Site"] = df["Site"].astype(str).str.strip()

        site_sums = (
            df[(df["Site"] != "") & (df["Outstanding Amount"] > 0)]
            .groupby("Site", as_index=False)["Outstanding Amount"].sum()
        )

        cols = [
            {"header": "Site", "total_string": "Grand Total"},
            {"header": "Sum of Remaining Outstanding Amount", "format": fmts["cur"], "total_function": "sum"},
        ]

        ws.add_table(
            0, 0, len(site_sums) + 1, 1,
            {
                "data": site_sums.values.tolist(),
                "columns": cols,
                "header_row": True,
                "total_row": True,
                "style": "Table Style Medium 9",
                "banded_rows": True,
                "autofilter": True,
            },
        )

        ws.set_column(0, 0, 20)
        ws.set_column(1, 1, 40)

    def _write_per_site_sheets(self, wb: xlsxwriter.Workbook, result_df: pd.DataFrame) -> None:
        fmts = self._formats(wb)
        site_headers = SITE_HEADERS

        for site in sorted(result_df["Site"].unique()):
            subset = result_df[(result_df["Site"] == site) & (result_df["Outstanding Amount"] > 0)]
            if subset.empty:
                continue

            sheet_name = sanitize_sheet_name(site)
            ws = wb.add_worksheet(sheet_name)

            ws.write(0, 0, site, fmts["title"])

            cols = []
            for h in site_headers:
                cd = {"header": h}
                if h == "Vendor Name":
                    cd["total_string"] = "Total"
                if h == "Invoice Number":
                    cd["total_function"] = "count"
                if h in ("Date Invoice Due", "Date Sent to Manager"):
                    cd["format"] = fmts["date"]
                if h == "Days Overdue":
                    cd["format"] = fmts["center"]
                if h in ("Amount Due", "Amount Paid", "Outstanding Amount"):
                    cd.update({"format": fmts["cur"], "total_function": "sum"})
                cols.append(cd)

            tbl_rows: List[List[object]] = []
            for r, rec in enumerate(subset.to_dict("records")):
                row = [
                    rec["Vendor Name"], rec["Invoice Number"], rec["Site"], rec["Unit"],
                    rec["Date Invoice Due"], rec["Date Sent to Manager"], "",
                    rec["Amount Due"], rec["Amount Paid"], rec["Outstanding Amount"],
                    rec["Remittance ID"]
                ]
                due_ref = xl_rowcol_to_cell(2 + r + 1, site_headers.index("Date Invoice Due"))
                row[site_headers.index("Days Overdue")] = f"=TODAY()-{due_ref}"
                tbl_rows.append(row)

            for row in tbl_rows:
                for i, v in enumerate(row):
                    if isinstance(v, pd._libs.tslibs.timestamps.Timestamp):
                        row[i] = v.to_pydatetime()
                    elif pd.isna(v):
                        row[i] = None

            ws.add_table(
                2, 0, 2 + len(tbl_rows) + 1, len(site_headers) - 1,
                {
                    "data": tbl_rows,
                    "columns": cols,
                    "header_row": True,
                    "total_row": True,
                    "style": "Table Style Medium 9",
                    "banded_rows": True,
                    "autofilter": True,
                },
            )

            days_idx = site_headers.index("Days Overdue")
            days_col = xl_rowcol_to_cell(0, days_idx)[0]
            days_range = f"{days_col}4:{days_col}{len(tbl_rows) + 3}"
            ws.conditional_format(days_range, {"type": "cell", "criteria": "<", "value": 0, "format": fmts["green"]})
            ws.conditional_format(days_range, {"type": "cell", "criteria": ">=", "value": 0, "format": fmts["red"]})

            for i, h in enumerate(site_headers):
                ws.set_column(i, i, 25 if h == "Vendor Name" else 15)
