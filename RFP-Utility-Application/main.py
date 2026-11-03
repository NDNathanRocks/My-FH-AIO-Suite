# Authored By: Nathan Nicholas D Souza
# Team: QI Team - Staffing Services at Fraser Health
# Contact: nathan.dsouza@fraserhealth.ca

# main.py
# -----------------------------------------------------------------------------
# MyAgencyPro - Application entrypoint & window routing
#
# - Sets a stable AppUserModelID on Windows so the taskbar icon is reliable.
# - Applies light "dark mode" adjustments when the system palette is dark.
# - Creates a QMainWindow with a QStackedWidget and all app pages.
# - Exposes simple page-switch methods used by child pages.
# -----------------------------------------------------------------------------

import os
import sys

from utils.logging_utils import configure_app_logging, get_logger, install_qmessagebox_logging

LOG_PATH = configure_app_logging()
logger = get_logger(__name__)

# --- Ensure Windows taskbar shows our custom icon -----------------------------
# Must be done BEFORE importing/creating any Qt application objects.
if sys.platform.startswith("win"):
	try:
		import ctypes
		ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
			"fraserhealth.myagencypro"  # Company.App.Product (any stable unique string)
		)
	except Exception:
		# Non-fatal: if it fails, the app still works, icon may not pin correctly
		pass

from PySide6.QtWidgets import QApplication, QMainWindow, QStackedWidget
from PySide6.QtGui import QIcon, QPalette, QColor

from utils.app_utils import ICON_PATH

# Pages
from pages.home_page import HomePage
from pages.generator_page import GeneratorPage
from pages.merger_page import MergerPage
from pages.saver_page import SaverPage
from pages.report_page import ReportPage
from pages.audit_page import AuditPage
from pages.check_page import InvoiceCheckPage
from pages.statistics_page import StatisticsPage
from pages.settings_page import SettingsPage
# from pages.update_page import UpdatePage


class MainWindow(QMainWindow):
	"""Main window housing a QStackedWidget for all app pages."""

	def __init__(self):
		super().__init__()
		self.setWindowTitle("MyAgencyPro")

		# Window icon (also set on QApplication below for taskbar/dialogs)
		if os.path.isfile(ICON_PATH):
			self.setWindowIcon(QIcon(ICON_PATH))

		# Central stacked widget acts as a simple page router
		self.stack = QStackedWidget()
		self.setCentralWidget(self.stack)

		# Instantiate pages (pass self so pages can call show_* methods)
		self.home = HomePage(self)
		self.generator = GeneratorPage(self)
		self.merger = MergerPage(self)
		self.saver = SaverPage(self)
		self.report = ReportPage(self)
		self.audit = AuditPage(self)
		self.check = InvoiceCheckPage(self)
		self.settings = SettingsPage(self)
		self.statistics = StatisticsPage(self)
		# self.update_data = UpdatePage(self)

		# Register pages with the stack (order doesn’t matter for navigation)
		for page in (
			self.home,
			self.generator,
			self.merger,
			self.saver,
			self.report,
			self.audit,
			self.check,
			self.settings,
			self.statistics,
			# self.update_data,
		):
			self.stack.addWidget(page)

		# Start on Home
		self.show_home()

	# --- Light "dark mode" adaptation ----------------------------------------
	@staticmethod
	def adapt_for_dark_mode():
		"""
		If the system palette background is dark, switch common text roles to white.
		This keeps your existing theme but improves contrast in dark environments.
		"""
		app = QApplication.instance()
		if not app:
			return
		pal = app.palette()

		# Perceptual luminance of the Window background (0=black, 1=white)
		bg = pal.color(QPalette.Window)
		luminance = 0.2126 * bg.redF() + 0.7152 * bg.greenF() + 0.0722 * bg.blueF()

		if luminance < 0.5:
			pal.setColor(QPalette.WindowText, QColor("white"))
			pal.setColor(QPalette.ButtonText, QColor("white"))
			pal.setColor(QPalette.Text, QColor("white"))
			pal.setColor(QPalette.HighlightedText, QColor("white"))
			app.setPalette(pal)

	# --- Simple page navigation API (called by child pages) -------------------
	def show_home(self): self.stack.setCurrentWidget(self.home)
	def show_generator(self): self.stack.setCurrentWidget(self.generator)
	def show_merger(self): self.stack.setCurrentWidget(self.merger)
	def show_saver(self): self.stack.setCurrentWidget(self.saver)
	def show_report(self): self.stack.setCurrentWidget(self.report)
	def show_audit(self): self.stack.setCurrentWidget(self.audit)
	def show_check(self): self.stack.setCurrentWidget(self.check)
	def show_settings(self): self.stack.setCurrentWidget(self.settings)
	def show_statistics(self): self.stack.setCurrentWidget(self.statistics)
	# def show_update_data(self): self.stack.setCurrentWidget(self.update_data)


# --- Application bootstrap ----------------------------------------------------
if __name__ == "__main__":
	if os.name == "nt":
		try:
			import ctypes
			ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("FraserHealth.MyAgencyPro")  # any stable string
		except Exception:
			pass

	app = QApplication(sys.argv)
	install_qmessagebox_logging()
	logger.info("Application bootstrap started | log_file=%s", LOG_PATH)

	# Apply dark-mode contrast tweaks early
	MainWindow.adapt_for_dark_mode()

	# Set application icon (taskbar, dialogs, etc.). Window icon is also set in MainWindow.
	if os.path.isfile(ICON_PATH):app.setWindowIcon(QIcon(ICON_PATH))

	win = MainWindow()
	if os.path.isfile(ICON_PATH):win.setWindowIcon(QIcon(ICON_PATH))
	win.resize(850, 650)
	win.setMinimumSize(500, 500)
	win.show()

	exit_code = app.exec()
	logger.info("Application exit | code=%s", exit_code)
	sys.exit(exit_code)