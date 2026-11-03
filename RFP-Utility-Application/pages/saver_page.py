# saver_page.py

from __future__ import annotations

import os
import shutil
import tempfile
from typing import List, Dict, Any

try:
	import pymupdf as fitz 		# PyMuPDF (preferred; avoids third-party fitz package collision)
except Exception:
	import fitz 				# PyMuPDF fallback
import extract_msg 	  # pip install extract-msg

from PySide6.QtWidgets import (
	QWidget, QVBoxLayout, QHBoxLayout, QToolBar, QGroupBox, QFormLayout,
	QLabel, QLineEdit, QPushButton, QComboBox, QProgressBar, QTextEdit,
	QFileDialog, QMessageBox, QScrollArea, QInputDialog, QDialog,
	QGraphicsView, QGraphicsScene, QGraphicsPixmapItem 	# NEW: Added GraphicsView components
)
from PySide6.QtGui import (
	QAction, QPalette, QColor, QPixmap, QImage, QDesktopServices,
	QKeySequence, QGuiApplication, QCursor, QShortcut, QPainter  # NEW: Added QPainter
)
from PySide6.QtCore import Qt, QUrl, QTimer

from utils.paths_config import DEFAULT_INVOICES_FOLDER
from utils.app_utils import extract_invoice_number, WAV_PATH

# ──────────────────────────────────────────────────────────────────────────────
# Constants / helpers
# ( ... no changes in this section ... )
# ──────────────────────────────────────────────────────────────────────────────

MIN_IMAGE_SIZE = 10 * 1024  # ignore tiny inline icons, etc.

RECYCLE_BIN_DEFAULT = (
	r"M:\Staffing Services Agency Intake\Recycle Bin\Emails Processed Successfully"
)

DEFAULT_EMAILS_FOLDER = r"M:\Staffing Services Agency Intake\Emails"

def ensure_folder(path: str) -> None:
	"""Create folder if it doesn't exist."""
	os.makedirs(path, exist_ok=True)
	

def unique_path_in_dir(target_dir: str, filename: str) -> str:
	"""Return a non-colliding path in target_dir by appending (n) if needed."""
	base, ext = os.path.splitext(filename)
	candidate = os.path.join(target_dir, filename)
	n = 1
	while os.path.exists(candidate):
		candidate = os.path.join(target_dir, f"{base} ({n}){ext}")
		n += 1
	return candidate

def get_msg_subject(msg_path: str) -> str:
	"""Read the subject from a .msg file and close it promptly."""
	msg = extract_msg.Message(msg_path)
	try:
		return msg.subject or "(no subject)"
	except Exception:
		return "(no subject)"
	finally:
		try:
			msg.close()
		except Exception:
			pass

def extract_pdf_attachments_from_msg(msg_path: str, temp_dir: str) -> List[str]:
	"""
	Extract PDF/image attachments from a .msg into temp_dir and return paths.
	We allow large images too (convertible elsewhere if needed by caller).
	"""
	saved: List[str] = []
	msg = extract_msg.Message(msg_path)
	try:
		for att in msg.attachments:
			# Determine a filename
			name = att.longFilename or att.shortFilename or "attachment"
			base, ext = os.path.splitext(name)
			ext = ext.lower()
			data = att.data

			if ext == ".pdf":
				out_path = os.path.join(temp_dir, base + ext)
				with open(out_path, "wb") as f:
					f.write(data)
				saved.append(out_path)

			elif ext in {".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tiff"} and len(data) >= MIN_IMAGE_SIZE:
				out_path = os.path.join(temp_dir, base + ext)
				with open(out_path, "wb") as f:
					f.write(data)
				saved.append(out_path)
		return saved
	finally:
		try:
			msg.close()
		except Exception:
			pass

def discover_sites(root: str) -> List[str]:
	if not os.path.isdir(root):
		return []
	return sorted(d for d in os.listdir(root) if os.path.isdir(os.path.join(root, d)))


def discover_units(root: str, site: str) -> List[str]:
	path = os.path.join(root, site)
	if not os.path.isdir(path):
		return []
	return sorted(d for d in os.listdir(path) if os.path.isdir(os.path.join(path, d)))

# ──────────────────────────────────────────────────────────────────────────────
# Toast class
# ( ... no changes in this section ... )
# ──────────────────────────────────────────────────────────────────────────────
class _ToastPopup(QDialog):
	"""Small, auto-closing, non-blocking toast message (bottom-right of parent)."""
	def __init__(self, parent: QWidget, text: str, duration_ms: int = 1600):
		super().__init__(parent,
						 Qt.Tool | Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)
		self.setAttribute(Qt.WA_TransparentForMouseEvents, True)
		self.setAttribute(Qt.WA_ShowWithoutActivating, True)
		self.setFocusPolicy(Qt.NoFocus)

		box = QVBoxLayout(self)
		box.setContentsMargins(14, 10, 14, 10)
		lbl = QLabel(text, self)
		lbl.setStyleSheet("color:white;")
		box.addWidget(lbl)

		self.setStyleSheet("""
			QDialog {
				background: rgba(40, 167, 69, 220);  /* green w/ alpha */
				border-radius: 10px;
			}
		""")

		self.adjustSize()
		self._place_bottom_right(parent)
		self.show()
		QTimer.singleShot(duration_ms, self.close)

	def _place_bottom_right(self, parent: QWidget) -> None:
		# position near bottom-right of the parent window
		if parent is None:
			return
		pg = parent.frameGeometry()
		self.adjustSize()
		x = pg.right() - self.width() - 24
		y = pg.bottom() - self.height() - 40
		self.move(max(0, x), max(0, y))

# ──────────────────────────────────────────────────────────────────────────────
# UI class
# ──────────────────────────────────────────────────────────────────────────────

class SaverPage(QWidget):
	"""
	Loads Outlook .msg files from a folder, extracts PDF/image attachments to a temp area,
	shows a live preview, and lets the user save each attachment into {output}/{site}/{unit}.

	UX buffering:
	- A status label + indeterminate progress bar when scanning/loading.
	- Determinate progress bar while stepping through attachments for the current email.
	- A log window for step-by-step feedback.
	- A quick notification (label/log/beep) when an email is fully processed.
	"""

	def __init__(self, controller):
		super().__init__()
		self.ctrl = controller

		# Paths / state
		self.in_dir_path: str = DEFAULT_EMAILS_FOLDER
		self.out_dir_path: str = DEFAULT_INVOICES_FOLDER
		self.recycle_bin: str = RECYCLE_BIN_DEFAULT

		# Temp directory for extracted attachments
		self.temp_dir = tempfile.TemporaryDirectory()

		# Email/PDF processing state
		self.emails: List[Dict[str, Any]] = []  # [{'source_path', 'subject', 'attachments': [paths...], 'type': 'email'|'pdf'}]
		self.emails_total: int = 0
		self.emails_done: int = 0

		self.current_email: int = 0  # index into self.emails
		self.current_attach: int = 0    # index into self.emails[i]['attachments']
		self.processed_count: int = 0   # attachments processed within CURRENT email

		self.last_site: str = ""
		self.last_unit: str = ""
		self.last_invoice: str = ""

		# Ensure recycle-bin exists
		ensure_folder(self.recycle_bin)

		# Theme & layout
		self._apply_theme()
		self._build_ui()

		# Quick escape to home
		esc = QShortcut(QKeySequence(Qt.Key_Escape), self)
		esc.activated.connect(self._on_back)

	# ────────────────────────── UI building / styling ──────────────────────────

	def _apply_theme(self) -> None:
		# ( ... no changes in this section ... )
		pal = QPalette()
		pal.setColor(QPalette.Window, QColor(245, 245, 245))
		pal.setColor(QPalette.WindowText, QColor(33, 33, 33))
		pal.setColor(QPalette.Button, QColor(238, 238, 238))
		pal.setColor(QPalette.ButtonText, QColor(33, 33, 33))
		pal.setColor(QPalette.Highlight, QColor(52, 152, 219))
		pal.setColor(QPalette.HighlightedText, QColor(255, 255, 255))
		self.setPalette(pal)

		self.setStyleSheet("""
			QGroupBox { border:2px solid #3498db; border-radius:5px; margin-top:10px; padding-top:18px; }
			QGroupBox::title { background:#3498db; color:white; padding:1px 6px; }
			QPushButton { background:#2980b9; color:white; padding:6px 12px; border:none; border-radius:4px; }
			QPushButton:disabled { background:#bdc3c7; color:#7f8c7f; }
			QLabel, QLineEdit, QComboBox { font-size:13px; }
			QTextEdit { background:white; border:1px solid #bdc3c7; border-radius:4px; }
		""")

	def _build_ui(self) -> None:
		root = QVBoxLayout(self)
		root.setContentsMargins(10, 10, 10, 10)
		root.setSpacing(10)

		# Toolbar
		toolbar = QToolBar()
		back_act = QAction("← Back", self)
		back_act.triggered.connect(self._on_back)
		toolbar.addAction(back_act)
		root.addWidget(toolbar)

		main = QHBoxLayout()
		main.setSpacing(15)

		# ────────────────── LEFT PANEL ──────────────────
		left = QWidget()
		lv = QVBoxLayout(left)
		lv.setSpacing(12)

		# Folders
		fg = QGroupBox("📁 Folders")
		form = QFormLayout(fg)

		self.le_in = QLineEdit(readOnly=True)
		self.le_in.setText(self.in_dir_path)
		btn_in = QPushButton("Browse…")
		btn_in.clicked.connect(lambda: self._browse("in"))
		form.addRow("Emails Folder:", self._hline(self.le_in, btn_in))

		self.le_out = QLineEdit(readOnly=True)
		self.le_out.setText(self.out_dir_path)
		btn_out = QPushButton("Browse…")
		btn_out.clicked.connect(lambda: self._browse("out"))
		form.addRow("Invoices Folder:", self._hline(self.le_out, btn_out))

		lv.addWidget(fg)

		# Load attachments
		self.btn_load = QPushButton("Load Attachments")
		self.btn_load.clicked.connect(self.load_queue)
		lv.addWidget(self.btn_load)

		# (Navigation group MOVED to right column)

		# Email subject
		subj_box = QGroupBox("✉️ Email Subject")
		sb = QHBoxLayout(subj_box)
		self.lbl_subject = QLabel("(none)")
		self.lbl_subject.setTextInteractionFlags(Qt.TextSelectableByMouse)
		btn_preview = QPushButton("Preview Email")
		btn_preview.clicked.connect(self.preview_email)
		sb.addWidget(self.lbl_subject, 1)
		sb.addWidget(btn_preview, 0)
		lv.addWidget(subj_box)

		# Current + Next 2
		qg = QGroupBox("📋 Current + Next 2")
		qv = QVBoxLayout(qg)
		self.lbl_current = QLabel("Current: (none)"); self.lbl_current.setStyleSheet("font-size:16px;")
		self.lbl_next1   = QLabel("Next: (none)")
		self.lbl_next2   = QLabel("")
		for lbl in (self.lbl_current, self.lbl_next1, self.lbl_next2):
			lbl.setTextInteractionFlags(Qt.TextSelectableByMouse)
			qv.addWidget(lbl)
		lv.addWidget(qg)

		# Details (invoice/site/unit)
		dg = QGroupBox("✏️ Details")
		df = QFormLayout(dg)
		self.le_invoice = QLineEdit()
		self.lbl_last_invoice = QLabel("")
		self.lbl_last_invoice.setStyleSheet("color:#555;")
		self.lbl_last_invoice.setTextInteractionFlags(Qt.TextSelectableByMouse)
		df.addRow("Last Invoice Processed:", self.lbl_last_invoice)
		df.addRow("Invoice #:", self.le_invoice)

		self.cb_site = QComboBox()
		self.cb_site.currentTextChanged.connect(self._on_site_change)
		btn_add_site = QPushButton("Add Site…"); btn_add_site.clicked.connect(self.add_site)
		df.addRow("Site:", self._hline(self.cb_site, btn_add_site))

		self.cb_unit = QComboBox()
		btn_add_unit = QPushButton("Add Unit…"); btn_add_unit.clicked.connect(self.add_unit)
		df.addRow("Unit:", self._hline(self.cb_unit, btn_add_unit))
		lv.addWidget(dg)

		# Progress + status + log
		self.progress = QProgressBar(); lv.addWidget(self.progress)

		status_row = QHBoxLayout()
		self.status_lbl = QLabel(""); self.status_lbl.setStyleSheet("color:#555; font-style:italic;")
		self.busy_bar = QProgressBar(); self.busy_bar.setRange(0, 0); self.busy_bar.setVisible(False)
		status_row.addWidget(self.status_lbl, 5)
		status_row.addWidget(self.busy_bar, 2)
		lv.addLayout(status_row)

		self.log = QTextEdit(); self.log.setReadOnly(True)
		lv.addWidget(self.log)

		main.addWidget(left, 3)

		# ────────────────── RIGHT COLUMN ──────────────────
		right = QWidget()
		rv = QVBoxLayout(right)
		rv.setSpacing(12)

		# MODIFIED: Preview group
		preview = QGroupBox("🔍 Preview")
		pv = QVBoxLayout(preview)

		# NEW: Create the view and scene
		self.preview_scene = QGraphicsScene(self)
		self.preview_view = QGraphicsView(self.preview_scene)
		
		# NEW: Improve rendering and interaction
		self.preview_view.setRenderHint(QPainter.Antialiasing, True)
		self.preview_view.setRenderHint(QPainter.SmoothPixmapTransform, True)
		self.preview_view.setTransformationAnchor(QGraphicsView.AnchorUnderMouse)
		self.preview_view.setResizeAnchor(QGraphicsView.AnchorUnderMouse)
		self.preview_view.setDragMode(QGraphicsView.ScrollHandDrag) # Pan with mouse drag
		
		# NEW: Add wheel event for zooming
		self.preview_view.wheelEvent = self._on_preview_wheel
		
		pv.addWidget(self.preview_view)

		# NEW: Add zoom/rotate buttons
		preview_toolbar = QHBoxLayout()
		self.btn_zoom_in = QPushButton("Zoom In (+)")
		self.btn_zoom_out = QPushButton("Zoom Out (-)")
		self.btn_zoom_fit = QPushButton("Fit to View")
		self.btn_rot_cw = QPushButton("Rotate CW (→)")
		self.btn_rot_ccw = QPushButton("Rotate CCW (←)")
		
		preview_toolbar.addWidget(self.btn_zoom_in)
		preview_toolbar.addWidget(self.btn_zoom_out)
		preview_toolbar.addWidget(self.btn_zoom_fit)
		preview_toolbar.addWidget(self.btn_rot_cw)
		preview_toolbar.addWidget(self.btn_rot_ccw)
		
		pv.addLayout(preview_toolbar)
		
		# NEW: Connect signals
		self.btn_zoom_in.clicked.connect(self._on_zoom_in)
		self.btn_zoom_out.clicked.connect(self._on_zoom_out)
		self.btn_zoom_fit.clicked.connect(self._on_zoom_fit)
		self.btn_rot_cw.clicked.connect(self._on_rotate_cw)
		self.btn_rot_ccw.clicked.connect(self._on_rotate_ccw)

		rv.addWidget(preview)

		# Navigation group
		nav = QGroupBox("🔀 Navigation")
		nv = QHBoxLayout(nav)
		self.btn_prev   = QPushButton("Previous"); 	   self.btn_prev.clicked.connect(self.go_previous)
		self.btn_next   = QPushButton("Next"); 		   self.btn_next.clicked.connect(self.go_next)
		self.btn_skip   = QPushButton("Skip"); 		   self.btn_skip.clicked.connect(self.skip_current)
		self.btn_confirm= QPushButton("Confirm & Save");  self.btn_confirm.clicked.connect(self.confirm_current)
		for b in (self.btn_prev, self.btn_next, self.btn_skip, self.btn_confirm):
			nv.addWidget(b)

		rv.addWidget(nav)

		main.addWidget(right, 4)

		root.addLayout(main)


	def _hline(self, widget, button) -> QWidget:
		# ( ... no changes in this section ... )
		box = QHBoxLayout()
		box.addWidget(widget, 1)
		box.addWidget(button, 0)
		wrap = QWidget(); wrap.setLayout(box)
		return wrap

	# ────────────────────────── Lifecycle ──────────────────────────

	def showEvent(self, event):
		# ( ... no changes in this section ... )
		self.ctrl.showMaximized()
		super().showEvent(event)

	def _on_back(self):
		# ( ... no changes in this section ... )
		self.ctrl.showNormal()
		self.ctrl.show_home()

	# ────────────────────────── Busy helpers ──────────────────────────

	def _set_controls_enabled(self, enabled: bool) -> None:
		# MODIFIED: Added new preview buttons
		for w in (self.btn_load, self.btn_prev, self.btn_next, self.btn_skip, self.btn_confirm,
				  self.le_in, self.le_out, self.cb_site, self.cb_unit, self.le_invoice,
				  self.btn_zoom_in, self.btn_zoom_out, self.btn_zoom_fit,
				  self.btn_rot_cw, self.btn_rot_ccw):
			if w:
				w.setEnabled(enabled)

	def _begin_busy(self, msg: str) -> None:
		# ( ... no changes in this section ... )
		self.status_lbl.setText(msg)
		self.busy_bar.setVisible(True)
		QGuiApplication.setOverrideCursor(Qt.WaitCursor)
		self._set_controls_enabled(False)
		QGuiApplication.processEvents()

	def _update_busy(self, msg: str) -> None:
		# ( ... no changes in this section ... )
		self.status_lbl.setText(msg)
		QGuiApplication.processEvents()

	def _end_busy(self) -> None:
		# ( ... no changes in this section ... )
		self.status_lbl.clear()
		self.busy_bar.setVisible(False)
		QGuiApplication.restoreOverrideCursor()
		self._set_controls_enabled(True)
		QGuiApplication.processEvents()

	def _log(self, text: str) -> None:
		# ( ... no changes in this section ... )
		self.log.append(text)
		self.log.verticalScrollBar().setValue(self.log.verticalScrollBar().maximum())

	# ──────────────────────── Sound Helpers ──────────────────────────

	def _play_notify_sound(self) -> None:
		# ( ... no changes in this section ... )
		"""Play a positive 'ting' (SystemAsterisk). Falls back gracefully."""
		# 1) Local WAV via Qt (put a small 'ting.wav' in resources/)
		try:
			from PySide6.QtMultimedia import QSoundEffect
			if not hasattr(self, "_ting"):
				self._ting = QSoundEffect(self)
				if os.path.exists(WAV_PATH):
					self._ting.setSource(QUrl.fromLocalFile(WAV_PATH))
					self._ting.setVolume(0.7)  # 0.0–1.0
			if getattr(self, "_ting", None) and self._ting.source().isValid():
				self._ting.play()
				return
		except Exception:
			pass

		# 2) QGui beep
		try:
			QGuiApplication.beep()
			return
		except Exception:
			pass

		# 3) Windows system sounds
		try:
			import winsound  # only on Windows
			for alias in ("SystemAsterisk", "SystemNotification", "SystemExclamation"):
				try:
					winsound.PlaySound(alias, winsound.SND_ALIAS | winsound.SND_ASYNC)
					return
				except Exception:
					continue
		except Exception:
			pass

	def _show_toast(self, message: str) -> None:
		# ( ... no changes in this section ... )
		"""Show a small celebration toast; auto-closes."""
		try:
			_ToastPopup(self, f"🎉 {message}")
		except Exception:
			# Silent fail if windowing system doesn’t allow it
			self._log(message)

	# ────────────────────────── Browsing ──────────────────────────

	def _browse(self, which: str) -> None:
		# ( ... no changes in this section ... )
		path = QFileDialog.getExistingDirectory(self, "Select Folder", os.path.expanduser(r"~/Downloads"))
		if not path:
			return
		if which == "in":
			self.in_dir_path = path
			self.le_in.setText(path)
		else:
			self.out_dir_path = path
			self.le_out.setText(path)

	# ────────────────────────── Queue load / scanning ──────────────────────────

	def load_queue(self) -> None:
		"""Scan input folder for .msg emails and .pdf files, and prime the queue."""
		if not os.path.isdir(self.in_dir_path) or not os.path.isdir(self.out_dir_path):
			QMessageBox.warning(self, "Invalid Folder", "Select both input & output folders.")
			return

		try:
			self._begin_busy("Scanning files…")

			self.emails.clear()
			# Walk all subfolders for .msg and .pdf files
			for root, _, files in os.walk(self.in_dir_path):
				for fn in files:
					fn_lower = fn.lower()
					file_path = os.path.join(root, fn)
					
					if fn_lower.endswith(".msg"):
						# Process email files
						atts = extract_pdf_attachments_from_msg(file_path, self.temp_dir.name)
						if atts:
							subject = get_msg_subject(file_path)
							self.emails.append({
								"source_path": file_path,
								"subject": subject,
								"attachments": atts,
								"type": "email"
							})
								
					elif fn_lower.endswith(".pdf"):
						# Process direct PDF files - treat each as standalone
						# Copy to temp directory to maintain consistent workflow
						temp_pdf_path = os.path.join(self.temp_dir.name, fn)
						shutil.copy(file_path, temp_pdf_path)
						self.emails.append({
							"source_path": file_path,
							"subject": fn,
							"attachments": [temp_pdf_path],
							"type": "pdf"
						})

			if not self.emails:
				self._end_busy()
				QMessageBox.information(self, "No Files", "No emails or PDFs found.")
				return

			# Count types for reporting
			email_count = sum(1 for e in self.emails if e['type'] == 'email')
			pdf_count = sum(1 for e in self.emails if e['type'] == 'pdf')

			# Reset counters and UI for the first item
			self.emails_total = len(self.emails)
			self.emails_done = 0
			self.current_email = 0
			self.current_attach = 0
			self.processed_count = 0

			self._populate_sites_units()
			self._setup_progress_for_current_email()
			self._load_current_email_or_finish()

			self.btn_load.setEnabled(False)
			self._update_busy(f"Loaded {email_count} email(s) and {pdf_count} PDF(s).")

		finally:
			self._end_busy()

	def _populate_sites_units(self) -> None:
		# ( ... no changes in this section ... )
		"""Refresh site & unit combo boxes from the output folder."""
		self.cb_site.clear()
		for site in discover_sites(self.out_dir_path):
			self.cb_site.addItem(site)
		self._on_site_change(self.cb_site.currentText())

	def _setup_progress_for_current_email(self) -> None:
		# ( ... no changes in this section ... )
		"""Configure the determinate progress bar for the current email's attachments."""
		if self.current_email < len(self.emails):
			num_atts = len(self.emails[self.current_email]['attachments'])
			self.progress.setRange(0, max(1, num_atts))
			self.progress.setValue(self.processed_count)
		else:
			self.progress.setRange(0, 1)
			self.progress.setValue(1)

	# ────────────────────────── Email stepping ──────────────────────────

	def _load_current_email_or_finish(self) -> None:
		# ( ... no changes in this section ... )
		"""If emails remain, load current attachment preview; otherwise finish."""
		if self.current_email >= len(self.emails):
			# All done
			self._log("✅ All files processed.")
			QMessageBox.information(self, "Done", "All files processed.")
			self.btn_load.setEnabled(True)
			self.lbl_subject.setText("(none)")
			self.lbl_current.setText("Current: (none)")
			self.lbl_next1.setText("Next: (none)")
			self.lbl_next2.setText("")
			self._clear_preview()
			return

		email = self.emails[self.current_email]
		atts = email['attachments']

		# If this email has no more attachments, wrap it up and move to next
		if not atts:
			self._on_email_finished(email)
			return

		# Show current + next two attachment names
		name = os.path.basename(atts[self.current_attach])
		# Add type tag to subject
		type_tag = "[📧 Email]" if email['type'] == 'email' else "[📄 PDF]"
		self.lbl_subject.setText(f"{type_tag} {email['subject']}")
		self.lbl_current.setText(f"Current: {name}")

		n1 = os.path.basename(atts[self.current_attach + 1]) if self.current_attach + 1 < len(atts) else "(none)"
		n2 = os.path.basename(atts[self.current_attach + 2]) if self.current_attach + 2 < len(atts) else ""
		self.lbl_next1.setText(f"Next: {n1}")
		self.lbl_next2.setText(n2)

		# Auto-fill invoice from filename
		inv = extract_invoice_number(name) or ""
		self.le_invoice.setText(inv)

		# Restore last site/unit selection
		if self.last_site:
			idx = self.cb_site.findText(self.last_site)
			if idx >= 0:
				self.cb_site.setCurrentIndex(idx)
				self._on_site_change(self.cb_site.currentText())
		if self.last_unit:
			idx = self.cb_unit.findText(self.last_unit)
			if idx >= 0:
				self.cb_unit.setCurrentIndex(idx)

		# Render preview
		self._show_preview(atts[self.current_attach])

	def _on_email_finished(self, email: Dict[str, Any]) -> None:
		"""Handle end-of-item: move source file to recycle, notify, advance to next item."""
		# Move the source file (.msg or .pdf) to the recycle bin (unique name, non-destructive)
		try:
			# Ensure we use absolute path for recycle bin
			recycle_path = os.path.abspath(self.recycle_bin)
			ensure_folder(recycle_path)
			fn = os.path.basename(email['source_path'])
			dest = unique_path_in_dir(recycle_path, fn)
			shutil.move(email['source_path'], dest)
		except Exception as e:
			file_type = "file" if email.get('type') == 'pdf' else "email"
			self._log(f"⚠️ Could not move {file_type} to recycle bin: {e}")

		# Remove this email from the queue
		self.emails.pop(self.current_email)

		# Update counters
		self.emails_done += 1
		self.current_attach = 0
		self.processed_count = 0

		# Notify user (non-blocking)
		self._notify_email_done(email['subject'])

		# Setup progress for the next email and continue
		self._setup_progress_for_current_email()
		self._load_current_email_or_finish()

	def _notify_email_done(self, subject: str) -> None:
		# ( ... no changes in this section ... )
		"""Notify that one email finished: status label, log, toast, and sound."""
		msg = f"Email {self.emails_done}/{self.emails_total} complete: {subject}"
		self.status_lbl.setText(msg)
		self._log(f"✅ {msg}")
		self._play_notify_sound()
		self._show_toast(msg)

	# ────────────────────────── Site / Unit management ──────────────────────────

	def _on_site_change(self, site: str) -> None:
		# ( ... no changes in this section ... )
		"""When site changes, repopulate units for that site."""
		self.cb_unit.clear()
		if site:
			for unit in discover_units(self.out_dir_path, site):
				self.cb_unit.addItem(unit)

	def add_site(self) -> None:
		# ( ... no changes in this section ... )
		text, ok = QInputDialog.getText(self, "Add Site", "Enter new Site:")
		if ok and text.strip():
			ensure_folder(os.path.join(self.out_dir_path, text.strip()))
			if self.cb_site.findText(text) == -1:
				self.cb_site.addItem(text)
			self.cb_site.setCurrentText(text.strip())

	def add_unit(self) -> None:
		# ( ... no changes in this section ... )
		site = self.cb_site.currentText().strip()
		if not site:
			QMessageBox.warning(self, "No Site", "Select a Site first.")
			return
		text, ok = QInputDialog.getText(self, "Add Unit", "Enter new Unit:")
		if ok and text.strip():
			ensure_folder(os.path.join(self.out_dir_path, site, text.strip()))
			if self.cb_unit.findText(text) == -1:
				self.cb_unit.addItem(text)
			self.cb_unit.setCurrentText(text.strip())

	# ────────────────────────── Navigation ──────────────────────────
	def go_previous(self) -> None:
		# ( ... no changes in this section ... )
		if self.current_email >= len(self.emails):
			return
		atts = self.emails[self.current_email]['attachments']
		if not atts:
			return
		self.current_attach = (self.current_attach - 1) % len(atts)
		self._load_current_email_or_finish()

	def go_next(self) -> None:
		# ( ... no changes in this section ... )
		if self.current_email >= len(self.emails):
			return
		atts = self.emails[self.current_email]['attachments']
		if not atts:
			return
		self.current_attach = (self.current_attach + 1) % len(atts)
		self._load_current_email_or_finish()

	def skip_current(self) -> None:
		# ( ... no changes in this section ... )
		"""Skip (remove) the current attachment but keep the email in the queue."""
		if self.current_email >= len(self.emails):
			return
		atts = self.emails[self.current_email]['attachments']
		if not atts:
			return

		atts.pop(self.current_attach)
		self.processed_count += 1
		self.progress.setValue(self.processed_count)

		# Advance index safely
		if atts:
			self.current_attach %= len(atts)
		else:
			self.current_attach = 0
		self._load_current_email_or_finish()

	# ────────────────────────── Confirm / save ──────────────────────────
	def confirm_current(self) -> None:
		# ( ... no changes in this section ... )
		"""Copy the current attachment into {output}/{site}/{unit}, optionally prefixing invoice."""
		if self.current_email >= len(self.emails):
			return
		atts = self.emails[self.current_email]['attachments']
		if not atts:
			return

		path = atts[self.current_attach]
		inv = self.le_invoice.text().strip()
		self.last_invoice = inv
		self.lbl_last_invoice.setText(inv or "(none)")

		site = self.cb_site.currentText().strip()
		unit = self.cb_unit.currentText().strip()
		if not (site and unit):
			QMessageBox.warning(self, "Missing", "Select both Site and Unit.")
			return

		dest_dir = os.path.join(self.out_dir_path, site, unit)
		ensure_folder(dest_dir)

		orig = os.path.basename(path)
		
		if inv and inv != extract_invoice_number(orig):
			new_name = f"`{inv}` - {orig}"
		else:
			new_name = orig

		dest_path = os.path.join(dest_dir, new_name)
		try:
			shutil.copy(path, dest_path)
			self._log(f"💾 Saved: {os.path.join(site, unit, new_name)}")
		except Exception as e:
			QMessageBox.critical(self, "Save Error", f"Failed to save attachment:\n{e}")
			return

		# Remember last selections
		self.last_site = site
		self.last_unit = unit

		# Remove saved attachment from queue
		atts.pop(self.current_attach)
		self.processed_count += 1
		self.progress.setValue(self.processed_count)

		# Advance within the email or finish the email
		if atts:
			self.current_attach %= len(atts)
		else:
			self.current_attach = 0

		self._load_current_email_or_finish()

	# ────────────────────────── Preview rendering ──────────────────────────

	# NEW: Zoom/Rotate/Pan handlers
	def _on_zoom_in(self):
		"""Zoom in the preview"""
		self.preview_view.scale(1.2, 1.2)

	def _on_zoom_out(self):
		"""Zoom out the preview"""
		self.preview_view.scale(0.8, 0.8)

	def _on_zoom_fit(self):
		"""Fit the first page to the viewport."""
		if not self.preview_scene.items():
			return  # No items to fit

		# Reset any previous transformations
		self.preview_view.resetTransform()
		
		# Find the first pixmap item (first page)
		first_page_item = None
		for item in reversed(self.preview_scene.items()):  # Reversed to get items in order added
			if isinstance(item, QGraphicsPixmapItem):
				first_page_item = item
				break
		
		if first_page_item:
			# Get the bounding rect of the first page with some margin
			page_rect = first_page_item.boundingRect()
			page_rect.translate(first_page_item.pos())
			page_rect.adjust(-10, -10, 10, 10)
			
			# Fit the first page in the view
			self.preview_view.fitInView(page_rect, Qt.KeepAspectRatio)

	def _on_rotate_cw(self):
		"""Rotate preview 90 degrees clockwise"""
		self.preview_view.rotate(90)

	def _on_rotate_ccw(self):
		"""Rotate preview 90 degrees counter-clockwise"""
		self.preview_view.rotate(-90)
		
	def _on_preview_wheel(self, event):
		"""Override wheel event for zooming"""
		# Check for Ctrl key if you want to only zoom with Ctrl+Wheel
		if event.modifiers() & Qt.ControlModifier:
			factor = 1.15 if event.angleDelta().y() > 0 else (1 / 1.15)
			self.preview_view.scale(factor, factor)
			# We consume the event, so it doesn't try to scroll
			event.accept()
		else:
			QGraphicsView.wheelEvent(self.preview_view, event)

	# MODIFIED: to use QGraphicsScene
	def _clear_preview(self) -> None:
		self.preview_scene.clear()
		# Reset view transform on new item
		self.preview_view.resetTransform()

	def _show_preview(self, path: str) -> None:
		"""Render a high-quality scrollable preview of the file (PDF or image)."""
		self._clear_preview()
		ext = os.path.splitext(path)[1].lower()

		# Keep track of vertical position for multi-page
		y_offset = 0.0

		# Use high DPI scaling for crisp text rendering (4x = ~288 DPI)
		SCALE_FACTOR = 4.0

		try:
			if ext == ".pdf":
				doc = fitz.open(path)
				try:
					for page_num, page in enumerate(doc):
						# Render at high resolution for crisp text
						pix = page.get_pixmap(matrix=fitz.Matrix(SCALE_FACTOR, SCALE_FACTOR))
						fmt = QImage.Format_RGBA8888 if pix.alpha else QImage.Format_RGB888
						img = QImage(pix.samples, pix.width, pix.height, pix.stride, fmt)
						pm = QPixmap.fromImage(img)

						# Add page separator for multi-page PDFs
						if page_num > 0:
							# Add spacing
							y_offset += 30
							# Add text separator
							separator_text = f"─────── Page {page_num + 1} ───────"
							text_item = self.preview_scene.addText(separator_text)
							text_item.setDefaultTextColor(QColor("#888888"))
							font = text_item.font()
							font.setPointSize(12)
							text_item.setFont(font)
							text_item.setPos(0, y_offset)
							y_offset += text_item.boundingRect().height() + 20

						# Add the full-resolution pixmap to the scene
						item = self.preview_scene.addPixmap(pm)
						item.setPos(0, y_offset)
						
						# Update y_offset for next page
						y_offset += pm.height() + 30
				finally:
					doc.close()
			else:
				# Load full-resolution image
				pm = QPixmap(path)
				if not pm.isNull():
					# Add the full-resolution pixmap to the scene
					item = self.preview_scene.addPixmap(pm)
					item.setPos(0, 0)
				else:
					raise Exception("Image file is null or invalid.")

			# Set the scene rect to match the content (with some padding)
			scene_rect = self.preview_scene.itemsBoundingRect()
			scene_rect.adjust(-20, -20, 20, 20)
			self.preview_scene.setSceneRect(scene_rect)
			
			# After loading, fit to view
			self._on_zoom_fit()

		except Exception as e:
			self._log(f"❌ Preview render error: {e}")
			# Add a text item to the scene
			text_item = self.preview_scene.addText(f"Preview not available:\n{e}")
			text_item.setDefaultTextColor(QColor("red"))

	# ────────────────────────── Misc ──────────────────────────

	def preview_email(self) -> None:
		"""Open the source file in the default handler (Outlook for .msg, PDF viewer for .pdf)."""
		if self.current_email < len(self.emails):
			source_path = self.emails[self.current_email]['source_path']
			if os.path.exists(source_path):
				QDesktopServices.openUrl(QUrl.fromLocalFile(source_path))