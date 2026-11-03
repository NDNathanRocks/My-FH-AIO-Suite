# merger_page.py

from __future__ import annotations

import os
import re
import tempfile
import traceback
import shutil
from typing import Dict, List, Tuple, Optional
from datetime import date, timedelta
import pandas as pd
from PIL import Image
from PyPDF2 import PdfMerger, PdfReader

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QToolBar,
    QLabel, QLineEdit, QPushButton, QProgressBar,
    QTextEdit, QFileDialog, QMessageBox, QSizePolicy, QApplication, QCheckBox
)
from PySide6.QtGui import (
    QAction, QPalette, QColor, QKeySequence, QShortcut, QDesktopServices, QGuiApplication
)
from PySide6.QtCore import Qt, QUrl, QTimer

from utils.app_utils import extract_invoice_number, normalize_key, get_tracker_path, get_company_csv_path, refresh_common_paths, ensure_tracker_closed_interactive, save_workbook_safe
from utils.paths_config import (
    DEFAULT_INVOICES_FOLDER,
    DEFAULT_RFPS_FOLDER,
    DEFAULT_MERGED_FOLDER,
    RECYCLE_BIN_FOLDER,
    DEFAULT_INVOICE_TRACKER
)

# Prefer shared recalc util; fallback to local if missing
try:
    from utils.app_utils import _force_excel_recalc as force_excel_recalc
except Exception:  # pragma: no cover
    def force_excel_recalc(path: str) -> None:
        try:
            import xlwings as xw
        except Exception:
            return
        app = xw.App(visible=False)
        try:
            wb = app.books.open(path)
            wb.api.RefreshAll()
            app.api.CalculateUntilAsyncQueriesDone()
            wb.save(); wb.close()
        finally:
            app.quit()

IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff"}

# Junk / OS files we never want to touch
JUNK_FILENAMES = {"thumbs.db", "desktop.ini", ".ds_store"}

def _is_junk_file(path: str) -> bool:
    name = os.path.basename(path).lower()
    return name in JUNK_FILENAMES or name.startswith("~$")

def _purge_junk_files(root: str) -> int:
    """Optionally delete junk OS files under a root. Safe to skip if undesired."""
    removed = 0
    for r, _, files in os.walk(root):
        for f in files:
            p = os.path.join(r, f)
            if _is_junk_file(p):
                try:
                    os.remove(p)
                    removed += 1
                except Exception:
                    pass
    return removed

class MergerPage(QWidget):
    """Merge invoice files with their matching RFP PDFs.

    What's new:
    - Modular helpers for scanning, matching, merging, moving, and logging
    - Clear buffering states: status label + indeterminate/determinate progress
    - Controls lock + wait cursor while busy; UI stays responsive
    - **Dry run** mode to simulate actions without changing files
    - **Skip tracker update** option to preview changes without modifying tracker
    - Failed merges are no longer moved anywhere; they remain in place
    - Successful merges are moved to Recycle Bin *preserving input subfolders*
    - After a real run, we prune empty Site/Unit folders from the input Invoices directory
    - **Generates single Excel file with two sheets**: Successfully Merged and Failed to Merge
    """

    def __init__(self, controller):
        super().__init__()
        self.ctrl = controller

        self.tracker_file = get_tracker_path()
        self.company_csv = get_company_csv_path()

        # Directories (defaults can be changed in the UI)
        self.invoice_dir = DEFAULT_INVOICES_FOLDER
        self.rfp_dir = DEFAULT_RFPS_FOLDER
        self.output_dir = DEFAULT_MERGED_FOLDER

        # Build UI / theme
        self._apply_palette()
        self._apply_style()
        self._build_ui()

        # First-time refresh (also happens every time page is shown)
        self._refresh_paths_from_config()

        # ESC → Back
        esc = QShortcut(QKeySequence(Qt.Key_Escape), self)
        esc.activated.connect(self.on_back)

    # ────────────────────────────── UI ──────────────────────────────
    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(50, 40, 50, 40)
        root.setSpacing(20)

        # Toolbar with Back
        toolbar = QToolBar()
        back_act = QAction("← Back", self)
        back_act.triggered.connect(self.on_back)
        toolbar.addAction(back_act)
        root.addWidget(toolbar)

        # Title
        title = QLabel("📑 Invoice ↔ RFP\nMerger")
        title.setAlignment(Qt.AlignHCenter)
        title.setStyleSheet("font-size:32px; font-weight:bold; color:#3498db;")
        root.addWidget(title)

        # Folders group
        folder_box = QWidget(); fl = QVBoxLayout(folder_box); fl.setSpacing(15)

        # Tracker
        self.track_le = QLineEdit(self.tracker_file); self.track_le.setReadOnly(True)
        fl.addLayout(self._row("Tracker Excel:", self.track_le))

        # Invoices
        self.inv_le = QLineEdit(self.invoice_dir); self.inv_le.setReadOnly(True)
        self.inv_btn = QPushButton("Browse…"); self.inv_btn.clicked.connect(lambda: self._choose_folder(self.inv_le))
        fl.addLayout(self._row("Invoices Folder:", self.inv_le, self.inv_btn))

        # RFPs
        self.rfp_le = QLineEdit(self.rfp_dir); self.rfp_le.setReadOnly(True)
        self.rfp_btn = QPushButton("Browse…"); self.rfp_btn.clicked.connect(lambda: self._choose_folder(self.rfp_le))
        fl.addLayout(self._row("RFPs Folder:", self.rfp_le, self.rfp_btn))

        # Output
        self.out_le = QLineEdit(self.output_dir); self.out_le.setReadOnly(True)
        self.out_btn = QPushButton("Browse…"); self.out_btn.clicked.connect(lambda: self._choose_folder(self.out_le))
        fl.addLayout(self._row("Merged RFPs and Invoices Folder:", self.out_le, self.out_btn))

        root.addWidget(folder_box)

        # Options checkboxes
        options_box = QWidget(); ol = QHBoxLayout(options_box)
        self.chk_dry_run = QCheckBox("Dry run (no changes)")
        self.chk_dry_run.setChecked(True)
        ol.addWidget(self.chk_dry_run)
        self.chk_skip_tracker = QCheckBox("Skip tracker update")
        self.chk_skip_tracker.setChecked(False)
        ol.addWidget(self.chk_skip_tracker)
        root.addWidget(options_box)
        
        # Merge button
        self.merge_btn = QPushButton("Run Merge"); self.merge_btn.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        self.merge_btn.clicked.connect(self.on_merge)
        root.addWidget(self.merge_btn, alignment=Qt.AlignCenter)

        # Progress bars / status
        self.progress = QProgressBar(); self.progress.setVisible(False)   # determinate
        root.addWidget(self.progress)
        self.busy_bar = QProgressBar(); self.busy_bar.setRange(0, 0); self.busy_bar.setVisible(False)  # indeterminate
        root.addWidget(self.busy_bar)
        self.status_lbl = QLabel(""); self.status_lbl.setStyleSheet("color:#555; font-style:italic;")
        root.addWidget(self.status_lbl)

        # Log area
        self.log = QTextEdit(); self.log.setReadOnly(True)
        root.addWidget(self.log)

    def showEvent(self, e):
        super().showEvent(e)
        self.ctrl.setMinimumSize(850, 800)
        self.ctrl.resize(850, 800)
        # Every time the page becomes visible, re-pull paths from the config file
        self._refresh_paths_from_config()

    def _refresh_paths_from_config(self) -> None:
        """Pull latest paths from disk and refresh UI + dependent data."""
        refresh_common_paths(self)   # sets: tracker_file, company_csv, template and line-edits


    def on_back(self) -> None:
        self.ctrl.setMinimumSize(500, 500)
        self.ctrl.resize(850, 650)
        self.ctrl.show_home()

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
            QProgressBar::chunk { background:#2ecc71; }
            QTextEdit { border:1px solid #bdc3c7; border-radius:4px; padding:4px; }
            QCheckBox { spacing:4px; font-size:13px; }
            """
        )

    def _row(self, label: str, widget: QLineEdit, button: Optional[QPushButton] = None) -> QHBoxLayout:
        layout = QHBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)

        layout.addWidget(QLabel(label), 1)
        if button is not None:
            layout.addWidget(widget, 5)
            layout.addWidget(button, 1)
        else:
            # no button → give the widget the full remaining space
            layout.addWidget(widget, 6)  # 5+1 combined

        return layout

    def _choose_folder(self, line_edit: QLineEdit) -> None:
        path = QFileDialog.getExistingDirectory(self, "Select Folder", "M:\\Staffing Services Agency Intake")
        if path:
            line_edit.setText(path)

    # ────────────────────────── Busy helpers ──────────────────────────
    def _set_controls_enabled(self, enabled: bool) -> None:
        for w in [self.merge_btn, self.inv_le, self.inv_btn, self.rfp_le, self.rfp_btn, 
                  self.out_le, self.out_btn, self.chk_dry_run, self.chk_skip_tracker]:
            if w is not None:
                w.setEnabled(enabled)

    def _begin_busy(self, msg: str) -> None:
        self.status_lbl.setText(msg)
        self.busy_bar.setVisible(True)
        self.progress.setVisible(False)
        QGuiApplication.setOverrideCursor(Qt.WaitCursor)
        self._set_controls_enabled(False)
        QGuiApplication.processEvents()

    def _update_busy(self, msg: str) -> None:
        self.status_lbl.setText(msg)
        QGuiApplication.processEvents()

    def _begin_progress(self, msg: str, maximum: int) -> None:
        self.status_lbl.setText(msg)
        self.busy_bar.setVisible(False)
        self.progress.setVisible(True)
        self.progress.setRange(0, maximum)
        self.progress.setValue(0)
        QGuiApplication.processEvents()

    def _end_busy(self) -> None:
        self.busy_bar.setVisible(False)
        self.progress.setVisible(False)
        self.status_lbl.clear()
        QGuiApplication.restoreOverrideCursor()
        self._set_controls_enabled(True)
        QGuiApplication.processEvents()

    def _log(self, text: str) -> None:
        self.log.append(text)
        # Auto-scroll to bottom
        self.log.verticalScrollBar().setValue(self.log.verticalScrollBar().maximum())

    # ────────────────────────── Helpers ──────────────────────────

    def canon_delims(s: str) -> str:
        """Treat '_' and '-' as equivalent by mapping both to '-' (for a lenient 2nd pass)."""
        return norm_id_keep_delims(s).replace('_', '-')

    def _collect_rfp_pdfs(self, rfp_dir: str) -> List[Dict[str, str]]:
        rfp_files: List[Dict[str, str]] = []
        for root, _, files in os.walk(rfp_dir):
            for f in files:
                if _is_junk_file(f):
                    continue
                if f.lower().endswith(".pdf"):
                    rfp_files.append({"name": f, "path": os.path.join(root, f)})
        return rfp_files


    def _collect_site_map(self, inv_dir: str) -> Dict[str, List[str]]:
        site_map: Dict[str, List[str]] = {}
        for root, _, files in os.walk(inv_dir):
            rel = os.path.relpath(root, inv_dir)
            for f in files:
                if _is_junk_file(f):
                    continue
                site_map.setdefault(rel, []).append(os.path.join(root, f))
        return site_map


    @staticmethod
    def _invoice_from_filename(filename: str) -> str | None:
        """
        Prefer any text wrapped in backticks (`...`) as the invoice number.
        Examples:
          "`INV BV 233`.pdf"                      -> "INV BV 233"
          "`INV BV 233` - Hello 21-21-21.pdf"     -> "INV BV 233"
        Fallback: use extract_invoice_number() for legacy patterns.
        """
        base = os.path.splitext(os.path.basename(filename))[0]
        m = re.search(r'`([^`]+)`', base)
        if m:
            return m.group(1).strip() or None
        return extract_invoice_number(filename)

    def _group_by_invoice(self, files: List[str]) -> Dict[str, List[str]]:
        groups: Dict[str, List[str]] = {}
        for p in files:
            if _is_junk_file(p):
                continue
            num = self._invoice_from_filename(os.path.basename(p))
            if not num:
                self._log(f"❌ Could not extract invoice# from {os.path.basename(p)}")
                groups.setdefault("__UNMATCHED__", []).append(p)
            else:
                groups.setdefault(num, []).append(p)
        if "__UNMATCHED__" in groups:
            pass
        return {k: v for k, v in groups.items() if k != "__UNMATCHED__"}


    def _extract_agency_name(self, rfp_path: str) -> str:
        try:
            reader = PdfReader(rfp_path)
            fields = reader.get_fields() or {}
            field = fields.get("Must match vendor listed on invoice")
            if field:
                return field.get("/V", "") or field.get("value", "") or ""
        except Exception:
            return ""
        return ""

    def _build_output_name(self, rfp_name: str, invoice_num: str) -> str:
        base, _ = os.path.splitext(rfp_name)
        idx = base.rfind(" - ")
        if idx >= 0:
            suffix = base[idx + 3 :].split()
            site = suffix[0] if suffix else ""
            unit = " ".join(suffix[1:]) if len(suffix) > 1 else ""
            return f"RFP {invoice_num} - {site} {unit}.pdf" if unit else f"RFP {invoice_num} - {site}.pdf"
        return f"RFP {invoice_num}.pdf"

    def _merge_files(self, rfp_pdf: str, attachments: List[str], out_path: str) -> Tuple[bool, List[str]]:
        attachments = [p for p in attachments if not _is_junk_file(p)]  # final safeguard
        temp_pdfs: List[str] = []
        try:
            merger = PdfMerger()
            # a) RFP first
            merger.append(rfp_pdf)
            # b) invoice PDFs, then other files
            pdfs = sorted((p for p in attachments if p.lower().endswith(".pdf")), key=os.path.getmtime)
            others = sorted((p for p in attachments if not p.lower().endswith(".pdf")), key=os.path.getmtime)
            for p in pdfs + others:
                if p.lower().endswith(".pdf"):
                    merger.append(p)
                else:
                    ext = os.path.splitext(p)[1].lower()
                    if ext in IMAGE_EXTS:
                        img = Image.open(p).convert("RGB")
                        tf = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
                        img.save(tf.name, format="PDF")
                        temp_pdfs.append(tf.name)
                        merger.append(tf.name)
                    else:
                        self._log(f"   ⚠️ Skipping unsupported file: {os.path.basename(p)}")
            # write out
            with open(out_path, "wb") as fout:
                merger.write(fout)
            merger.close()
            return True, attachments
        except Exception as e:
            self._log(f"   ❌ Merge error: {e}")
            for line in traceback.format_exc().splitlines():
                self._log("      " + line)
            return False, []
        finally:
            for tf in temp_pdfs:
                try:
                    os.remove(tf)
                except Exception:
                    pass

    def _move_to_recycle_preserve_tree(self, paths: List[str], base_dir: str, recycle_root: str) -> None:
        """Move each file to recycle_root/<Week Range>/..., preserving its relative path under base_dir.
        Week folder is named 'Month D - Month D', e.g. 'August 25 - August 31'
        """
        # Compute current week's Monday → Sunday range
        today = date.today()
        monday = today - timedelta(days=today.weekday())  # Monday = 0
        sunday = monday + timedelta(days=6)

        # Format: 'August 25 - August 31'
        week_label = f"{monday.strftime('%B')} {monday.day} - {sunday.strftime('%B')} {sunday.day}"

        week_root = os.path.join(recycle_root, week_label)

        for p in paths:
            try:
                rel = os.path.relpath(p, base_dir)
            except ValueError:
                # If file isn't under base_dir (edge case), drop into root
                rel = os.path.basename(p)

            target_dir = os.path.join(week_root, os.path.dirname(rel))
            os.makedirs(target_dir, exist_ok=True)

            dest = os.path.join(target_dir, os.path.basename(p))

            try:
                shutil.move(p, dest)
            except Exception as e:
                self._log(f"❌ Failed to move {os.path.basename(p)} → {dest}: {e}")

    def _extract_site_unit_from_path(self, file_path: str, base_dir: str) -> Tuple[str, str]:
        """
        Extract Site and Unit from file path relative to base directory.
        Example: if file is "base_dir/Site1/Unit2/invoice.pdf", returns ("Site1", "Unit2")
        """
        try:
            rel_path = os.path.relpath(os.path.dirname(file_path), base_dir)
            parts = rel_path.split(os.sep)
            
            # Filter out '.' which represents the base directory itself
            parts = [p for p in parts if p != '.']
            
            if len(parts) >= 2:
                return parts[0], parts[1]
            elif len(parts) == 1:
                return parts[0], ""
            else:
                return "", ""
        except Exception:
            return "", ""

    def _generate_merge_report(self, merged_data: List[Dict[str, str]], 
                               failed_data: List[Dict[str, str]], 
                               output_dir: str) -> str:
        """
        Generate an Excel file with two sheets: Successfully Merged and Failed to Merge.
        Returns the path to the generated file.
        """
        timestamp = date.today().strftime("%Y%m%d")
        report_path = os.path.join(output_dir, f"Merge_Report_{timestamp}.xlsx")
        
        # Create DataFrames
        merged_df = pd.DataFrame(merged_data) if merged_data else pd.DataFrame(columns=["Invoice #", "Vendor ID", "Site", "Unit"])
        failed_df = pd.DataFrame(failed_data) if failed_data else pd.DataFrame(columns=["Invoice #", "Vendor ID", "Site", "Unit"])
        
        # Write to Excel with two sheets
        with pd.ExcelWriter(report_path, engine='openpyxl') as writer:
            merged_df.to_excel(writer, sheet_name='Successfully Merged', index=False)
            failed_df.to_excel(writer, sheet_name='Failed to Merge', index=False)
        
        return report_path

    # ────────────────────────── Core actions ──────────────────────────
    def on_merge(self) -> None:

        inv_dir = self.inv_le.text().strip()
        rfp_dir = self.rfp_le.text().strip()
        out_dir = self.out_le.text().strip()
        dry_run = self.chk_dry_run.isChecked()
        skip_tracker = self.chk_skip_tracker.isChecked()

        if not os.path.isfile(self.tracker_file):
            QMessageBox.warning(self, "Missing Tracker Excel", "Please select a valid Invoices Tracker Excel file.")
            return

        # 1) Validate folders
        for d in (inv_dir, rfp_dir, out_dir):
            if not os.path.isdir(d):
                QMessageBox.warning(self, "Invalid Folder", f"Folder not found: {d}")
                return

        self.merge_btn.setEnabled(False)
        self.log.clear()

        removed = _purge_junk_files(inv_dir) + _purge_junk_files(rfp_dir)
        if removed:
            self._log(f"🧹 Removed {removed} junk OS file(s) (e.g., thumbs.db, Desktop.ini).")

        # Lists to collect merge results for reporting
        merged_report_data: List[Dict[str, str]] = []
        failed_report_data: List[Dict[str, str]] = []

        try:
            # 2) Scan RFPs
            self._begin_busy("Scanning RFP PDFs…")
            rfp_files = self._collect_rfp_pdfs(rfp_dir)
            self._update_busy(f"Found {len(rfp_files)} RFP PDFs. Scanning invoice folders…")

            # 3) Scan invoices grouped by subfolder
            site_map = self._collect_site_map(inv_dir)
            total_invoices = sum(len(files) for files in site_map.values())
            
            mode_msg = "Dry run: evaluating…" if dry_run else "Merging…"
            self._begin_progress(mode_msg, total_invoices)

            merged_paths: List[str] = []
            unmerged_paths: List[str] = []
            merged_records: List[Dict[str, str]] = []  # (Agency Name, Invoice Number)
            processed = 0

            for sub, files in site_map.items():
                dest = out_dir if sub in (".", "") else os.path.join(out_dir, sub)
                if not dry_run:
                    os.makedirs(dest, exist_ok=True)

                # 4) Group by invoice #
                groups = self._group_by_invoice(files)

                # 5) Merge per invoice #
                for num, grp_files in groups.items():
                    inv_key = self._strict_invoice_key(num)

                    strict_matches = []
                    for r in rfp_files:
                        ident = self._extract_rfp_identifier(r["name"])
                        if ident and self._strict_invoice_key(ident) == inv_key:
                            strict_matches.append(r)

                    # Extract Site and Unit from first file in group
                    site, unit = self._extract_site_unit_from_path(grp_files[0], inv_dir) if grp_files else ("", "")

                    if not strict_matches:
                        self._log(f"🚩 No RFP match for invoice '{num}' in '{sub}'. Skipped.")
                        # Add to failed report
                        failed_report_data.append({
                            "Invoice #": num,
                            "Vendor ID": "",  # We don't have vendor ID without RFP match
                            "Site": site,
                            "Unit": unit
                        })
                        unmerged_paths.extend(grp_files)
                        processed += len(grp_files)
                        self.progress.setValue(processed)
                        QApplication.processEvents()
                        continue

                    if len(strict_matches) > 1:
                        self._log(f"⚠️ Multiple strict matches for '{num}'. Using first: {strict_matches[0]['name']}")

                    rfp = strict_matches[0]
                    self._log(f"↪ {'Would merge' if dry_run else 'Merging'} RFP '{rfp['name']}' + {len(grp_files)} attachments")

                    # Extract Agency Name (best effort) & record
                    agency = self._extract_agency_name(rfp["path"]) or ""
                    merged_records.append({"Agency Name": agency, "Invoice Number": num})

                    # Output file name and merge
                    out_name = self._build_output_name(rfp["name"], num)
                    out_path = os.path.join(dest, out_name)

                    if dry_run:
                        self._log(f"   ↪ Would save as {os.path.join(sub, out_name)}")
                        merged_paths.extend(grp_files)
                        # Add to successful merge report
                        merged_report_data.append({
                            "Invoice #": num,
                            "Vendor ID": agency,
                            "Site": site,
                            "Unit": unit
                        })
                    else:
                        ok, _used = self._merge_files(rfp["path"], grp_files, out_path)
                        if ok:
                            self._log(f"   ✔ Saved {os.path.join(sub, out_name)}")
                            merged_paths.extend(grp_files)
                            # Add to successful merge report
                            merged_report_data.append({
                                "Invoice #": num,
                                "Vendor ID": agency,
                                "Site": site,
                                "Unit": unit
                            })
                        else:
                            self._log(f"   ❌ Failed merging invoice {num}: see log above")
                            unmerged_paths.extend(grp_files)
                            # Add to failed merge report
                            failed_report_data.append({
                                "Invoice #": num,
                                "Vendor ID": agency,
                                "Site": site,
                                "Unit": unit
                            })

                    processed += len(grp_files)
                    self.progress.setValue(processed)
                    QApplication.processEvents()

            # 6) Generate merge report Excel file
            self._update_busy("Generating merge report…")
            report_path = self._generate_merge_report(merged_report_data, failed_report_data, out_dir)
            self._log(f"📊 Merge report generated: {os.path.basename(report_path)}")

            # 7) Post-process files
            #    • Failed merges: leave in place (no move)
            self._update_busy("Finalizing…")
            if unmerged_paths:
                self._log(f"⚠️  {len(unmerged_paths)} files (invoices & attachments) failed — left in original folders.")

            #    • Successful merges: move to recycle bin preserving input subfolders
            if merged_paths:
                good_root = os.path.join(RECYCLE_BIN_FOLDER, "Invoices that Merged Successfully")
                if dry_run:
                    self._log(f"🧪 Dry run: would move {len(merged_paths)} merged invoices to {good_root} (preserving site/unit folders)")
                else:
                    self._update_busy("Moving merged invoices to Recycle-Bin…")
                    self._move_to_recycle_preserve_tree(merged_paths, base_dir=inv_dir, recycle_root=good_root)
                    self._log(f"✅ {len(merged_paths)} merged invoices moved → {good_root}")

            # 8) Update tracker: set Merged (Y/N) = "Y" for invoices we merged
            if merged_records:
                if dry_run:
                    self._log("🧪 Dry run: would update tracker Merged (Y/N) to 'Y' for merged invoices")
                elif skip_tracker:
                    self._log("⏭️ Skip tracker update: exporting excel file instead…")
                else:
                    self._update_busy("Updating tracker (Merged = Y)…")
                
                if skip_tracker:
                    pass
                    # Export preview without updating tracker
                    # updated, not_found, preview_df = self._update_tracker_merges(
                    #     tracker_path=self.tracker_file,
                    #     company_csv_path=self.company_csv,
                    #     records=merged_records,
                    #     skip_update=True
                    # )
                    
                    # if preview_df is not None and not preview_df.empty:
                    #     # export_path = os.path.join(os.path.dirname(out_dir), f"Tracker_Update_Preview_Merger.xlsx")
                    #     # preview_df.to_excel(export_path, index=False, engine='openpyxl')
                    #     self._log(f"📝 Preview: {updated} row(s) would be updated. Exported to: {os.path.basename(export_path)}")
                    #     QMessageBox.information(
                    #         self, "Tracker Update Skipped",
                    #         f"Tracker update was skipped.\n\n"
                    #         f"{updated} row(s) would have been updated.\n\n"
                    #         # f"Preview exported to:\n{export_path}"
                    #     )
                    # else:
                    #     self._log("ℹ️ No rows found to update in tracker")
                        
                    # if not_found:
                    #     self._log(f"⚠️ {len(not_found)} invoice(s) not found in tracker:")
                    #     for row in not_found[:20]:
                    #         self._log(f"   • [{row.get('CompanyName','?')}] Invoice '{row.get('Invoice Number','?')}'")
                    #     if len(not_found) > 20:
                    #         self._log(f"   …and {len(not_found)-20} more")
                            
                elif not dry_run:
                    # Actually update the tracker
                    updated, not_found, _ = self._update_tracker_merges(
                        tracker_path=self.tracker_file,
                        company_csv_path=self.company_csv,
                        records=merged_records,
                        skip_update=False
                    )
                    self._log(f"📝 Tracker updated: {updated} row(s) set to Y.")
                    if not_found:
                        self._log(f"⚠️ {len(not_found)} invoice(s) not found in tracker:")
                        for row in not_found[:20]:
                            self._log(f"   • [{row.get('CompanyName','?')}] Invoice '{row.get('Invoice Number','?')}'")
                        if len(not_found) > 20:
                            self._log(f"   …and {len(not_found)-20} more")

            completion_msg = "Dry run complete." if dry_run else "All invoices processed."
            QMessageBox.information(self, "Done", f"{completion_msg}\n\nMerge report: {os.path.basename(report_path)}")

        except Exception as e:
            # Log full traceback while keeping the message box concise
            self._log("".join(traceback.format_exc().splitlines()))
            QMessageBox.critical(self, "Error", str(e))
        finally:
            self.merge_btn.setEnabled(True)
            self._end_busy()

    # ────────────────────────── ID helpers ──────────────────────────
    @staticmethod
    def _extract_rfp_identifier(name: str) -> str | None:
        """
        Return the token immediately after 'RFP' in an RFP filename.
        Strict-friendly:
        - Prefer '... RFP <ID> - ...' → capture <ID> up to ' - '
        - Fallback: '... RFP <ID>' to end of string (allows spaces/_/-)
        Examples:
        'Solutions RFP 82260-584 - LMH Emergency.pdf' -> '82260-584'
        'Sympatico RFP FHA_HOT_005.pdf'               -> 'FHA_HOT_005'
        'Acme RFP 123 45.pdf'                         -> '123 45'
        """
        base = os.path.splitext(os.path.basename(name))[0]
        # Prefer "… RFP <ID> - …"
        m = re.search(r'(?i)\bRFP\b\s+(.+?)\s-\s', base)
        if m:
            return m.group(1).strip()
        # Fallback: "… RFP <ID>" to end of string; allow spaces/_/-
        m = re.search(r'(?i)\bRFP\b\s+([A-Za-z0-9 _-]+)$', base)
        return m.group(1).strip() if m else None


    @staticmethod
    def _norm_id_keep_delims(s: str) -> str:
        """Uppercase and keep only letters, digits, '_' and '-'. (No spaces or other punctuation.)"""
        return re.sub(r'[^A-Za-z0-9_-]', '', str(s)).upper()

    @staticmethod
    def _strict_invoice_key(v) -> str:
        """
        Strict key for invoice comparison:
        - trim only outer whitespace
        - preserve case and characters (spaces, '_' and '-' are all allowed)
        - coerce whole-number floats (e.g., 6058.0) to '6058'
        """
        if v is None:
            return ""
        if isinstance(v, float) and v.is_integer():
            return str(int(v)).strip()
        return str(v).strip()

    # ────────────────────────── Tracker update helpers ──────────────────────────
    
    def _read_company_map(self, csv_path: str) -> Dict[str, str]:
        """
        Return a mapping: normalized 'Payable To' -> 'CompanyName'
        Used to convert Agency (PDF) -> CompanyName (tracker sheet).
        """
        df = pd.read_csv(csv_path, dtype=str)
        for col in ("Payable To", "CompanyName"):
            if col not in df.columns:
                raise ValueError(f"Company CSV must have '{col}' column.")
        def norm(s: str) -> str:
            return (str(s or "").strip().upper())
        return {norm(row["Payable To"]): str(row["CompanyName"]).strip()
                for _, row in df.iterrows()}

    def _sheet_by_company(self, wb, company_name: str):
        """Find sheet by company name (exact first, then case-insensitive)."""
        if company_name in wb.sheetnames:
            return wb[company_name]
        # case-insensitive fallback
        low = company_name.lower()
        for nm in wb.sheetnames:
            if nm.lower() == low:
                return wb[nm]
        return None

    @staticmethod
    def _header_map_row3(ws) -> Dict[str, str]:
        """
        Build header map from row 3, starting B3 onward (but safe to read entire row).
        Returns dict: header_text -> column_letter
        """
        from openpyxl.utils import get_column_letter
        hdr = {}
        for cell in ws[3]:  # whole row
            val = str(cell.value or "").strip()
            if val:
                hdr[val] = get_column_letter(cell.column)
        return hdr

    @staticmethod
    def _norm_invoice_value(s: str) -> str:
        """Strict compare key (trim only; preserve case; keep spaces/_/-; coerce 6058.0 → '6058')."""
        return MergerPage._strict_invoice_key(s)


    def _update_tracker_merges(self, tracker_path: str, company_csv_path: str,
                        records: List[Dict[str, str]], skip_update: bool = False) -> Tuple[int, List[Dict[str, str]], Optional[pd.DataFrame]]:
        """
        For each {'Agency Name','Invoice Number'} in records:
        1) map Agency -> CompanyName via CSV (Payable To → CompanyName)
        2) open that sheet, locate 'Invoice #' and 'Merged (Y/N)' (headers in row 3)
        3) find row where 'Invoice #' == invoice number (strict: trimmed, case-insensitive)
        4) set 'Merged (Y/N)' = 'Y' (unless skip_update is True)
        
        Returns (count_updated, not_found_list, preview_dataframe)
        If skip_update is True, returns a DataFrame of rows that would be updated
        """
        updated = 0
        not_found: List[Dict[str, str]] = []
        preview_rows: List[Dict[str, any]] = []

        if not os.path.isfile(tracker_path):
            raise FileNotFoundError(f"Tracker not found: {tracker_path}")
        if not os.path.isfile(company_csv_path):
            raise FileNotFoundError(f"Company CSV not found: {company_csv_path}")

        # 1) build Agency->Company map (Payable To → CompanyName)
        pay_to_to_company = self._read_company_map(company_csv_path)

        # 2) bucket invoices by company name
        buckets: Dict[str, List[str]] = {}
        unresolved: List[Dict[str, str]] = []
        for rec in records:
            agency = (rec.get("Agency Name") or "").strip()
            inv    = (rec.get("Invoice Number") or "").strip()
            if not agency or not inv:
                continue
            company = pay_to_to_company.get(agency.strip().upper())
            if not company:
                unresolved.append({"Agency Name": agency, "Invoice Number": inv, "Reason": "Agency not in CSV"})
                continue
            buckets.setdefault(company, []).append(inv)

        if unresolved:
            for row in unresolved:
                self._log(f"⚠️ Skip update: {row['Agency Name']} not found in CSV (invoice {row['Invoice Number']})")

        if not buckets:
            return 0, unresolved, None

        # 3) open tracker once; update per company sheet
        # Use read_only+data_only for preview; full open for write (data_only=True on a saved
        # workbook strips Table metadata and causes Excel to repair table1.xml on next open).
        wb = load_workbook(tracker_path, read_only=skip_update, data_only=skip_update)
        try:
            for company_name, inv_list in buckets.items():
                ws = self._sheet_by_company(wb, company_name)
                if ws is None:
                    for inv in inv_list:
                        not_found.append({"CompanyName": company_name, "Invoice Number": inv, "Reason": "Sheet not found"})
                    continue

                hdr = self._header_map_row3(ws)
                inv_col = hdr.get("Invoice #")
                merged_col = hdr.get("Merged (Y/N)")
                if not inv_col or not merged_col:
                    for inv in inv_list:
                        not_found.append({"CompanyName": company_name, "Invoice Number": inv, "Reason": "Headers missing"})
                    continue

                # Build index from invoice value -> row number (strict compare: trimmed + uppercase)
                index: Dict[str, int] = {}
                for r in range(4, ws.max_row + 1):  # data starts at row 4
                    v = self._norm_invoice_value(ws[f"{inv_col}{r}"].value)
                    if v:
                        index[v] = r

                # Update rows (or collect for preview)
                for inv in inv_list:
                    key = self._norm_invoice_value(inv)
                    row = index.get(key)
                    if row:
                        try:
                            if skip_update:
                                # Collect row data for preview
                                row_data = {"CompanyName": company_name, "Invoice Number": inv}
                                for col_name, col_letter in hdr.items():
                                    row_data[col_name] = ws[f"{col_letter}{row}"].value
                                preview_rows.append(row_data)
                            else:
                                ws[f"{merged_col}{row}"].value = "Y"
                            updated += 1
                        except Exception:
                            not_found.append({"CompanyName": company_name, "Invoice Number": inv, "Reason": "Write failed"})
                    else:
                        not_found.append({"CompanyName": company_name, "Invoice Number": inv, "Reason": "Invoice not found"})

            # save once (only if not skipping update)
            if not skip_update:
                save_workbook_safe(wb, tracker_path)  # saves, closes, gc, then os.replace
                wb = None  # already closed inside save_workbook_safe

        finally:
            if wb is not None:
                try:
                    wb.close()
                except Exception:
                    pass

        # Convert preview rows to DataFrame
        preview_df = pd.DataFrame(preview_rows) if preview_rows else None

        return updated, not_found, preview_df