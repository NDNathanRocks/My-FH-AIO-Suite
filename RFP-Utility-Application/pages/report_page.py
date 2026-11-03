# report_page.py

from __future__ import annotations

import os
import math
import zipfile
from datetime import datetime
from typing import Iterable, List, Tuple

try:
    import pymupdf as fitz      # PyMuPDF (preferred; avoids third-party fitz package collision)
except Exception:
    import fitz                 # PyMuPDF fallback
import pandas as pd
import xlsxwriter
import win32com.client          # Outlook
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QToolBar, QLabel,
    QLineEdit, QPushButton, QDateEdit, QProgressBar,
    QMessageBox, QFileDialog, QSizePolicy
)
from PySide6.QtGui import (
    QAction, QPalette, QColor, QKeySequence, QShortcut, QGuiApplication
)
from PySide6.QtCore import Qt, QDate, QSize, QTimer

from utils.paths_config import DEFAULT_INVOICE_TRACKER
from utils.app_utils import get_tracker_path, get_sent_invoices_dir, refresh_common_paths, ensure_tracker_closed_interactive

class ReportPage(QWidget):
    """
    Builds a per-period statement from merged RFP PDFs, zips the files,
    drafts Outlook emails, then copies the input {site}/{unit} folder
    into the Sent base under a date-range folder.

    UX buffering:
    - status label + indeterminate progress bar for long phases
    - determinate progress while scanning PDFs
    - disables controls + wait cursor while running
    """

    def __init__(self, controller):
        super().__init__()
        self.ctrl = controller
        self.email_table = None
        self.pdfcount = 0

        self.tracker_file = get_tracker_path()
        self.sent_root = get_sent_invoices_dir()

        # ── UI scaffold ────────────────────────────────────────────────
        esc = QShortcut(QKeySequence(Qt.Key_Escape), self)
        esc.activated.connect(self.ctrl.show_home)

        self._apply_palette()
        self._apply_style()
        self._build_ui()

        # First-time refresh (also happens every time page is shown)
        self._refresh_paths_from_config()

    # ─────────────────────────── UI builders ───────────────────────────
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
        self.setStyleSheet("""
            QPushButton { background:#2980b9; color:white; border:none; border-radius:6px;
                          padding:8px 16px; font-size:14px; }
            QPushButton:hover { background:#3498db; }
            QLineEdit, QDateEdit {
                background:white; border:1px solid #bdc3c7; border-radius:4px; padding:4px; font-size:13px;
            }
            QLabel { font-size:13px; }
            QProgressBar { border:1px solid #7f8c8d; border-radius:5px; height:20px; }
            QProgressBar::chunk { background:#2ecc71; border-radius:5px; }
        """)

    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(50, 40, 50, 40)
        root.setSpacing(16)

        # Toolbar
        toolbar = QToolBar()
        toolbar.setIconSize(QSize(16, 16))
        back = QAction("← Back", self)
        back.triggered.connect(self.ctrl.show_home)
        toolbar.addAction(back)
        root.addWidget(toolbar)

        # Title
        title = QLabel("📆 Statement Report\n+ Email")
        title.setAlignment(Qt.AlignHCenter)
        title.setStyleSheet("font-size:32px; font-weight:bold; color:#3498db;")
        root.addWidget(title)

        # Tracker Excel (unchanged behavior; kept for parity with other pages)
        track_row = QHBoxLayout()
        track_row.addWidget(QLabel("Tracker Excel:"), 1)
        self.track_le = QLineEdit(self.tracker_file); self.track_le.setReadOnly(True)
        track_row.addWidget(self.track_le, 5)
        root.addLayout(track_row)

        # Sent base folder 
        sent_row = QHBoxLayout()
        sent_row.addWidget(QLabel("Sent Invoices Folder:"), 1)
        self.le_sent = QLineEdit(self.sent_root); self.le_sent.setReadOnly(True)
        sent_row.addWidget(self.le_sent, 5)
        root.addLayout(sent_row)

        # Input folder (merged PDFs)
        input_row = QHBoxLayout()
        input_row.addWidget(QLabel("Merged Unit Folder:"), 1)
        self.le_folder = QLineEdit(); self.le_folder.setReadOnly(True)
        input_row.addWidget(self.le_folder, 5)
        btn_browse = QPushButton("Browse…")
        btn_browse.clicked.connect(self._choose_input_folder)
        input_row.addWidget(btn_browse, 1)
        root.addLayout(input_row)

        # Date range (defaults to current week's Monday → Sunday; user can change)
        dates = QHBoxLayout()
        dates.addWidget(QLabel("From:"), 1)
        self.date_from = QDateEdit(calendarPopup=True)
        self.date_to   = QDateEdit(calendarPopup=True)
        today  = QDate.currentDate()
        monday = today.addDays(1 - today.dayOfWeek())  # Qt: Monday=1 … Sunday=7
        sunday = monday.addDays(6)
        self.date_from.setDate(monday)
        self.date_to.setDate(sunday)
        dates.addWidget(self.date_from, 3)
        dates.addWidget(QLabel("To:"), 1)
        dates.addWidget(self.date_to, 3)
        root.addLayout(dates)

        # Run
        self.btn_run = QPushButton("Generate Statement")
        self.btn_run.setEnabled(False)
        self.btn_run.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        self.btn_run.clicked.connect(self.on_generate)
        root.addWidget(self.btn_run, alignment=Qt.AlignCenter)

        # Progress (determinate) + status (indeterminate)
        self.progress = QProgressBar(); self.progress.setVisible(False)
        root.addWidget(self.progress)

        status_row = QHBoxLayout()
        self.status_lbl = QLabel(""); self.status_lbl.setStyleSheet("color:#555; font-style:italic;")
        self.status_lbl.setVisible(False)
        self.busy_bar = QProgressBar(); self.busy_bar.setRange(0, 0); self.busy_bar.setVisible(False)
        status_row.addWidget(self.status_lbl, 5)
        status_row.addWidget(self.busy_bar, 3)
        root.addLayout(status_row)

    # ─────────────────────────── File pickers ───────────────────────────
    def _choose_input_folder(self) -> None:
        folder = QFileDialog.getExistingDirectory(self, "Select Merged Invoices and RFPs Folder",
                                                  r"M:\Staffing Services Agency Intake")
        if folder:
            self.le_folder.setText(folder)
            self._update_run_enabled()

    def _update_run_enabled(self) -> None:
        self.btn_run.setEnabled(bool(self.le_folder.text().strip()) and bool(self.le_sent.text().strip()))

    def showEvent(self, e):
        super().showEvent(e)
        # Every time the page becomes visible, re-pull paths from the config file
        self._refresh_paths_from_config()

    def _refresh_paths_from_config(self) -> None:
        """Pull latest paths from disk and refresh UI + dependent data."""
        refresh_common_paths(self)   # sets: tracker_file, company_csv, template and line-edits

    # ─────────────────────────── Busy helpers ───────────────────────────
    def _set_controls_enabled(self, enabled: bool) -> None:
        for w in (self.btn_run, self.track_le, self.le_folder, self.le_sent):
            if w:
                w.setEnabled(enabled)

    def _begin_busy(self, msg: str) -> None:
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

    # ─────────────────────────── Main action ───────────────────────────
    def on_generate(self) -> None:

        folder = self.le_folder.text().strip()
        if not os.path.isdir(folder):
            QMessageBox.warning(self, "Invalid Folder", "Please select a valid folder.")
            return

        if not ensure_tracker_closed_interactive(self, self.tracker_file):
            return False

        # {site}/{unit} from the folder path
        unit_name = os.path.basename(os.path.normpath(folder))
        site_name = os.path.basename(os.path.dirname(os.path.normpath(folder)))

        try:
            self._begin_busy("Scanning PDFs…")

            self._load_email_table()

            # 1) Gather PDFs in the selected folder
            pdfs = [f for f in os.listdir(folder) if f.lower().endswith(".pdf")]
            if not pdfs:
                QMessageBox.information(self, "No PDFs", "No PDF files found in that folder.")
                return

            # 2) Parse PDFs → DataFrame (progress visible)
            df = self._build_report_dataframe(folder, pdfs)

            # 3) Chunk & Zip (≤ 20 MB each)
            self._update_busy("Creating ZIP file(s)…")
            zip_paths = self._chunk_and_zip(folder, pdfs)

            # 4) Copy {site}/{unit} into Sent/date-range (no overwrite)
            self._update_busy("Copying Site/Unit to Sent folder…")
            date_label = self._date_range_label()
            dest_unit_dir = self._copy_unit_to_sent(folder, site_name, unit_name, date_label)

            # 5) Save Statement report (default to Sent/date/site/unit)
            self._update_busy("Writing Statement report…")
            report_path = self._write_report(df, default_dir=dest_unit_dir)

            # 6) Draft Outlook emails (attach ZIPs + report)
            self._update_busy("Preparing Outlook drafts…")
            self._draft_emails(folder, zip_paths, report_path)

            self.progress.setVisible(False)
            QMessageBox.information(self, "Done",
                f"Statement written to:\n{report_path}\n\n"
                f"Copied to Sent:\n{dest_unit_dir}"
            )

        except Exception as e:
            self.progress.setVisible(False)
            QMessageBox.critical(self, "Error", f"Unexpected error:\n{e}")
        finally:
            self._end_busy()

    # ─────────────────────────── Data / report ───────────────────────────
    def _build_report_dataframe(self, folder: str, pdfs: List[str]) -> pd.DataFrame:
        """Extract fields from RFP PDFs in folder → DataFrame (with progress)."""
        records = []
        total = len(pdfs)
        self.progress.setRange(0, total)
        self.progress.setValue(0)
        self.progress.setVisible(True)

        for idx, fn in enumerate(pdfs, start=1):
            path = os.path.join(folder, fn)

            # Site/Unit hint from filename: "RFP #### - {site} {unit}.pdf"
            base = os.path.splitext(fn)[0]
            if " - " in base:
                suffix = base.split(" - ", 1)[1]
                tokens = suffix.split()
                site = tokens[0] if tokens else ""
                unit = " ".join(tokens[1:]) if len(tokens) > 1 else ""
            else:
                site, unit = "", ""

            # Read first page widgets
            try:
                doc = fitz.open(path)
                widgets = doc[0].widgets() or []
                vals = {w.field_name: w.field_value for w in widgets}
                records.append({
                    "Vendor":     vals.get("Must match vendor listed on invoice", ""),
                    "Site":       site,
                    "Unit":       unit,
                    "Invoice #":  str(vals.get("Invoice No", "")),
                    "Date Due":   vals.get("DDMMYY", ""),
                    "Amount Due": float(str(vals.get("000TOTAL", "")).replace("$", "").replace(",", "") or 0.0),
                    "Status":     ""
                })
                doc.close()
            except Exception:
                # ignore unreadable PDFs; keep processing
                pass

            self.progress.setValue(idx)

        self.pdfcount = len(records)
        return pd.DataFrame(records)

    def _write_report(self, df: pd.DataFrame, default_dir: str) -> str:
        """Write the statement Excel; default to the provided directory."""
        fr = self.date_from.date().toString("MMM d")
        to = self.date_to.date().toString("MMM d, yyyy")
        unit = os.path.basename(os.path.normpath(self.le_folder.text().strip()))
        default_name = f"Statement Report - {fr} - {to} ({unit}).xlsx"

        # default save path in the new Sent destination
        save_path = os.path.join(default_dir, default_name)

        headers = df.columns.tolist()
        data = df.values.tolist()
        nrows, ncols = len(data), len(headers)
        amt_idx = headers.index("Amount Due") if "Amount Due" in headers else -1

        wb = xlsxwriter.Workbook(save_path)
        ws = wb.add_worksheet("Sheet1")

        # Title
        fmt_title = wb.add_format({'bold': True, 'font_size': 18})
        ws.write('A1', os.path.splitext(os.path.basename(save_path))[0], fmt_title)

        # Table
        cols = []
        for h in headers:
            cd = {'header': h}
            if h == "Amount Due":
                cd['total_function'] = "sum"
            if h == "Invoice #":
                cd['total_function'] = "count"
            if h == "Vendor":
                cd['total_string'] = "Totals"
            cols.append(cd)

        ws.add_table(
            2, 0, 2 + nrows + 1, ncols - 1,
            {'data': data, 'columns': cols, 'header_row': True, 'total_row': True,
             'style': 'Table Style Medium 9', 'banded_rows': True}
        )

        # Formats
        cen = wb.add_format({'align': 'center'})
        cur = wb.add_format({'num_format': '$* #,##0.00', 'align': 'right'})
        for i in range(ncols):
            if i == amt_idx:
                ws.set_column(i, i, 15, cur)
            else:
                ws.set_column(i, i, 20, cen)

        wb.close()
        return save_path

    # ─────────────────────────── ZIP + Outlook ───────────────────────────
    def _chunk_and_zip(self, folder: str, pdfs: List[str]) -> List[str]:
        """Group files into ≤20MB chunks and zip each. Returns list of zip paths."""
        files = [os.path.join(folder, f) for f in pdfs]
        limit = 20 * 1024 * 1024

        groups: List[List[str]] = []
        cur: List[str] = []
        cur_size = 0

        for fpath in files:
            sz = os.path.getsize(fpath)
            if cur and (cur_size + sz > limit):
                groups.append(cur)
                cur, cur_size = [], 0
            cur.append(fpath)
            cur_size += sz
        if cur:
            groups.append(cur)

        zip_paths: List[str] = []
        base = os.path.basename(folder)

        for idx, grp in enumerate(groups, start=1):
            zname = f"{base}" + (f"_part{idx}" if len(groups) > 1 else "")
            zpath = os.path.join(folder, f"{zname}.zip")
            with zipfile.ZipFile(zpath, 'w', compression=zipfile.ZIP_DEFLATED) as zf:
                for fp in grp:
                    zf.write(fp, arcname=os.path.basename(fp))
            zip_paths.append(zpath)

        return zip_paths

    def _clean_email(self, value) -> str:
        """
        Coerce Excel/pandas values to a safe email string.
        Turns None/NaN/'nan'/'None' into ''.
        """
        if value is None:
            return ""
        s = str(value).strip()
        return "" if s.lower() in ("nan", "none", "nat") else s
    
    def _load_email_table(self):
        """*** FIX: Load email table from tracker file with proper closure ***"""
        try:
            # Use pandas read_excel which properly closes the file
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
                    "Site", "Unit", "Cost Centre", "Manager Name",
                    "Manager Email", "Secondary Contact(s)", "Secondary Email"
                ])
                return
                
            start = idx.index[0] + 1
            headers = df.iloc[start].tolist()
            
            rows = []
            r = start + 1
            while r < len(df) and str(df.iat[r, 0]).strip():
                rows.append(df.iloc[r, :len(headers)].tolist())
                r += 1
                
            self.email_table = pd.DataFrame(rows, columns=headers)
            
        except Exception as e:
            print(f"Error loading email table: {e}")
            self.email_table = pd.DataFrame(columns=[
                "Site", "Unit", "Cost Centre", "Manager Name",
                "Manager Email", "Secondary Contact(s)", "Secondary Email"
            ])

    def _draft_emails(self, folder: str, zip_paths: List[str], report_path: str) -> None:
        """Draft Outlook emails with zip(s) and the statement attached."""
        # 1) Load email table lazily
        if self.email_table is None:
            self._load_email_table()

        # 2) Attempt to extract a Cost Centre from one of the PDFs
        cost_centre = ""
        try:
            any_pdf = next((p for p in os.listdir(folder) if p.lower().endswith(".pdf")), None)
            if any_pdf:
                fp = os.path.join(folder, any_pdf)
                doc = fitz.open(fp)
                vals = {w.field_name: w.field_value for w in (doc[0].widgets() or [])}
                cost_centre = (vals.get("Text2", "") or vals.get("Cost Centre", "") or "").strip()
                doc.close()
        except Exception:
            pass

        # 3) Lookup emails (sanitize to avoid NaN/float issues)
        to_addr, cc_addr = "", ""
        try:
            row = self.email_table[self.email_table['Cost Centre'].astype(str).str.strip() == cost_centre]
            if not row.empty:
                to_addr = self._clean_email(row.iloc[0].get('Manager Email'))
                cc_addr = self._clean_email(row.iloc[0].get('Secondary Email'))
        except Exception:
            pass

        # Always CC AP fallback; only include non-empty strings
        cc_parts = []
        if cc_addr:
            cc_parts.append(cc_addr)
        cc_parts.append("FHAgencyStaffing@fraserhealth.ca")
        cc_all = ";".join(cc_parts)

        # 4) Draft one mail per zip
        ol = win32com.client.Dispatch("Outlook.Application")
        date_str = datetime.now().strftime("%B %d, %Y")
        base = os.path.basename(folder)
        fname = to_addr.split('@')[0].split('.')[0].capitalize() if to_addr else ""

        for idx, zpath in enumerate(zip_paths, start=1):
            zname = os.path.splitext(os.path.basename(zpath))[0]
            mail = ol.CreateItem(0)  # olMailItem
            mail.SentOnBehalfOfName = "FHAgencyStaffing@fraserhealth.ca"
            mail.To = to_addr or ""              # must be a string
            mail.CC = cc_all                     # already a safe string
            mail.Subject = f"Invoice(s) for processing - {date_str} - {zname}"
            mail.BodyFormat = 2  # HTML
            mail.HTMLBody = self._email_html_body(fname)
            mail.Attachments.Add(Source=zpath)
            mail.Attachments.Add(Source=report_path)
            mail.Display(True)

    # ─────────────────────────── Sent folder sync ───────────────────────────
    def _date_range_label(self) -> str:
        """Human-friendly date range like 'July 1 - July 31' or 'July 25 - August 2'."""
        a = self.date_from.date()
        b = self.date_to.date()
        left = a.toString("MMMM d")
        right = b.toString("MMMM d")
        return f"{left} - {right}"

    def _copy_unit_to_sent(self, src_unit_dir: str, site: str, unit: str, date_label: str) -> str:
        """
        Copy the selected {site}/{unit} folder into:
          {sent_root}/{date_label}/{site}/{unit}
        - Creates missing folders
        - Adds files without overwriting (suffixes ... (1), (2), ...)
        Returns the destination unit folder path.
        """
        # Build destination
        dest_date_dir = os.path.join(self.sent_root, date_label)
        dest_site_dir = os.path.join(dest_date_dir, site or "_UnknownSite")
        dest_unit_dir = os.path.join(dest_site_dir, unit or "_UnknownUnit")
        os.makedirs(dest_unit_dir, exist_ok=True)

        # Copy (merge) all files preserving relative structure
        for root, dirs, files in os.walk(src_unit_dir):
            rel = os.path.relpath(root, src_unit_dir)
            rel = "" if rel == "." else rel
            target_dir = os.path.join(dest_unit_dir, rel)
            os.makedirs(target_dir, exist_ok=True)

            for fname in files:
                src = os.path.join(root, fname)
                base, ext = os.path.splitext(fname)
                dest = os.path.join(target_dir, fname)

                # Avoid overwriting: append (n)
                if os.path.exists(dest):
                    n = 1
                    while True:
                        cand = os.path.join(target_dir, f"{base} ({n}){ext}")
                        if not os.path.exists(cand):
                            dest = cand
                            break
                        n += 1

                try:
                    # copy2 to preserve timestamps
                    from shutil import copy2
                    copy2(src, dest)
                except Exception:
                    # continue with others; don't fail the whole run
                    pass

        return dest_unit_dir

    # ─────────────────────────── Email body template ───────────────────────────
    def _email_html_body(self, fname: str) -> str:
        """Return the static HTML body (unchanged), with an optional greeting name."""
        return f"""
<html>
  <body style="font-family:Calibri, sans-serif; font-size:11pt; line-height:1.4;">
    <p>Hello {fname},</p>
    <p>I hope this finds you well.</p>
    <p>Please find attached a compressed zip file containing 
       <strong><span style="color:#FF0000;">{self.pdfcount}</span></strong> request for payment forms for your <strong>wet</strong> signature. 
       Attached is a statement report for your reference.</p>
    <p><strong>Instructions for processing payments:</strong></p>
    <ol>
      <li>Please <strong>print, sign and date</strong> the attached request for payment forms
        <ol type="disc">
          <li><em>Please sign for approver section unless indicated otherwise*</em></li>
        </ol>
      </li>
      <li>Please <strong>mail the physical form(s) and supporting documents</strong>  to:
        <ol type="disc">
          <li>
            <strong>c/o:</strong> Accounts Payable<br/>
            8521 198a Street<br/>
            Langley BC V2Y 0A1
          </li>
        </ol>
      </li>
    </ol>
    <p>Instructions on how to extract files from zipped folders are available 
      <a href="https://support.microsoft.com/en-us/windows/zip-and-unzip-files-8d28fa72-f2f9-712f-67df-f80cf89fd4e5">here</a>.
    </p>
    <p>If you have any questions or concerns, please do not hesitate to contact us at 
      <a href="mailto:FHAgencyStaffing@fraserhealth.ca">FHAgencyStaffing@fraserhealth.ca</a>.
    </p>
  </body>
  <body style="font-family:Calibri, sans-serif; font-size:10pt; line-height:1;">
      </span></p><p style="margin: 0;"class="MsoNormal"><b><span lang="EN-US" style="font-size:10.0pt;font-family:&quot;Arial&quot;,sans-serif;color:#31261D;mso-ligatures:standardcontextual">Thank you,</span></b><b><span lang="EN-US" style="font-size:10.0pt;font-family:&quot;Arial&quot;,sans-serif;color:#31261D"></span></b></p><p style="margin: 0;"class="MsoNormal"><b><span style="font-size:10.0pt;font-family:&quot;Arial&quot;,sans-serif;color:#005293;mso-ligatures:standardcontextual"><o:p>&nbsp;</o:p></span></b></p><p style="margin: 0;"class="MsoNormal"><b><span style="font-size:10.0pt;font-family:&quot;Arial&quot;,sans-serif;color:#005293;mso-ligatures:standardcontextual">Surbhi Singla (she/her/hers)</span></b><span style="font-size:11.0pt;color:black;mso-ligatures:standardcontextual"></span></p><p style="margin: 0;"class="MsoNormal"><b><span style="font-size:10.0pt;font-family:&quot;Arial&quot;,sans-serif;color:#31261D;mso-ligatures:standardcontextual">HR Assistant | Staffing Services</span></b><b><span lang="EN-US" style="font-size:10.0pt;font-family:&quot;Arial&quot;,sans-serif;color:#31261D;mso-ligatures:standardcontextual"></span></b></p><p style="margin: 0;"class="MsoNormal"><span lang="EN-US" style="font-size:10.0pt;font-family:&quot;Arial&quot;,sans-serif;color:#31261D;mso-ligatures:standardcontextual"><o:p>&nbsp;</o:p></span></p><p style="margin: 0;"class="MsoNormal"><b><span lang="EN-US" style="font-size:10.0pt;font-family:&quot;Arial&quot;,sans-serif;color:#31261D;mso-ligatures:standardcontextual">Fraser Health | Better Health. Best in health care.</span></b></p><p style="margin: 0;"class="MsoNormal"><b><span lang="EN-US" style="font-size:10.0pt;font-family:&quot;Arial&quot;,sans-serif;color:#31261D;mso-ligatures:standardcontextual"><o:p>&nbsp;</o:p></span></b></p><p style="margin: 0;"class="MsoNormal"><span style="font-size:10.0pt;font-family:&quot;Arial&quot;,sans-serif;color:#31261D;mso-ligatures:standardcontextual">Suite 500, Central City Tower, 13450 - 102 Avenue, Surrey, B.C. V3T 0A3</span></p><p style="margin: 0;"class="MsoNormal"><span lang="EN-US" style="font-size:10.0pt;font-family:&quot;Arial&quot;,sans-serif;color:#31261D;mso-ligatures:standardcontextual"><o:p>&nbsp;</o:p></span></p><p style="margin: 0;"class="MsoNormal"><span style="mso-ligatures:standardcontextual"><a href="https://can01.safelinks.protection.outlook.com/?url=http%3A%2F%2Fwww.fraserhealth.ca%2F&amp;data=05%7C02%7CFHAgencystaffing%40fraserhealth.ca%7Cf1f6a844c56746e6b62d08ddcbae0cf7%7C31f660a5192a4db392baca424f1b259e%7C0%7C0%7C638890671172338966%7CUnknown%7CTWFpbGZsb3d8eyJFbXB0eU1hcGkiOnRydWUsIlYiOiIwLjAuMDAwMCIsIlAiOiJXaW4zMiIsIkFOIjoiTWFpbCIsIldUIjoyfQ%3D%3D%7C0%7C%7C%7C&amp;sdata=ZLcOzCnXH2Sf97M8hasBHo7vQKE7mSmKriu%2F0KMAwPA%3D&amp;reserved=0" originalsrc="http://www.fraserhealth.ca/"><span lang="EN-US" style="font-size:10.0pt;font-family:&quot;Arial&quot;,sans-serif;color:#0563C1">www.fraserhealth.ca</span></a></span><span lang="EN-US" style="font-size:10.0pt;font-family:&quot;Arial&quot;,sans-serif;color:#31261D;mso-ligatures:standardcontextual"></span></p><p style="margin: 0;"class="MsoNormal"><span lang="EN-US" style="font-size:10.0pt;font-family:&quot;Arial&quot;,sans-serif;color:#31261D;mso-ligatures:standardcontextual"><o:p>&nbsp;</o:p></span></p><p style="margin: 0;"class="MsoNormal"><i><span style="font-size:10.0pt;font-family:&quot;Arial&quot;,sans-serif;mso-ligatures:standardcontextual">We recognize that Fraser Health provides care on the traditional, ancestral and unceded lands of the Coast Salish and Nlaka pamux Nations and is home to 32 First Nations within the Fraser Salish region. Fraser Health is dedicated to serving all Indigenous people, and honours the unique cultures of the First Nations, M tis and Inuit living within the Fraser Salish region.</span></i></p><p style="margin: 0;"class="MsoNormal"><span lang="EN-US" style="font-size:10.0pt;font-family:&quot;Arial&quot;,sans-serif;color:#31261D;mso-ligatures:standardcontextual"><o:p>&nbsp;</o:p></span></p><p style="margin: 0;"class="MsoNormal"><span lang="EN-US" style="font-size:10.0pt;font-family:&quot;Arial&quot;,sans-serif;color:#31261D;mso-ligatures:standardcontextual">CONFIDENTIALITY NOTICE: This email is intended only for the person(s) or entity to which it</span></p><p style="margin: 0;"class="MsoNormal"><span lang="EN-US" style="font-size:10.0pt;font-family:&quot;Arial&quot;,sans-serif;color:#31261D;mso-ligatures:standardcontextual">is addressed and may contain confidential information. Any review, distribution, copying, printing</span></p><p style="margin: 0;"class="MsoNormal"><span lang="EN-US" style="font-size:10.0pt;font-family:&quot;Arial&quot;,sans-serif;color:#31261D;mso-ligatures:standardcontextual">or other use of this email by anyone other than the named recipient is prohibited. If you have</span></p><p style="margin: 0;"class="MsoNormal"><span lang="EN-US" style="font-size:10.0pt;font-family:&quot;Arial&quot;,sans-serif;color:#31261D;mso-ligatures:standardcontextual">received this email in error or are not the named recipient, please notify the sender immediately</span></p><p style="margin: 0;"class="MsoNormal"><span lang="EN-US" style="font-size:10.0pt;font-family:&quot;Arial&quot;,sans-serif;color:#31261D;mso-ligatures:standardcontextual">and permanently delete this email and all copies of it. Thank you.</span></p><p style="margin: 0;"class="MsoNormal"><span style="font-size:11.0pt;font-family:&quot;Calibri&quot;,sans-serif"><o:p>&nbsp;</o:p></span></p><p><span style="font-size:11.0pt;font-family:&quot;Calibri&quot;,sans-serif"><o:p>&nbsp;</o:p></span>
    </body>
</html>
"""