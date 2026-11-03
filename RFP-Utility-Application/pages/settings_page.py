# settings_page.py

import os
from PySide6.QtWidgets import (
    QWidget, QFormLayout, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QFileDialog, QToolBar
)
from PySide6.QtGui import QAction, QPalette, QColor, QShortcut, QKeySequence
from PySide6.QtCore import Qt
from utils.app_utils import (
    load_config, save_config, set_path, get_paths,
    CONF_TRACKER, CONF_TRACKER_2, CONF_COMPANY_CSV, CONF_TEMPLATE,
    CONF_SENT_INVOICES, CONF_PENDING_INVOICES
)

class SettingsPage(QWidget):
    def __init__(self, controller):
        super().__init__()
        self.ctrl = controller
        self.cfg = load_config()  # raw, for initial display only

        # ── Palette & Style ────────────────────────────────────────
        pal = QPalette()
        pal.setColor(QPalette.Window, QColor(245,245,245))
        pal.setColor(QPalette.WindowText, QColor(33,33,33))
        pal.setColor(QPalette.Button, QColor(238,238,238))
        pal.setColor(QPalette.ButtonText, QColor(33,33,33))
        pal.setColor(QPalette.Highlight, QColor(52,152,219))
        pal.setColor(QPalette.HighlightedText, QColor(255,255,255))
        self.setPalette(pal)
        self.setStyleSheet("""
            QLabel { font-size:14px; }
            QLineEdit { border:1px solid #bdc3c7;
                        border-radius:4px; padding:4px;
                        background:white; }
            QPushButton { background:#2980b9; color:white;
                          border:none; border-radius:6px;
                          padding:6px 12px; font-size:13px; }
            QPushButton:disabled { background:#bdc3c7; color:#7f8c7f; }
            QPushButton:hover { background:#3498db; }
        """)

        root = QVBoxLayout(self); root.setContentsMargins(50,40,50,40); root.setSpacing(20)

        tb = QToolBar()
        back = QAction("← Back", self); back.triggered.connect(self.ctrl.show_home)
        tb.addAction(back); root.addWidget(tb)

        title = QLabel("⚙️ Settings"); title.setAlignment(Qt.AlignHCenter)
        title.setStyleSheet("font-size:24px; font-weight:bold; color:#3498db;")
        root.addWidget(title)

        form = QFormLayout(); root.addLayout(form)

        esc = QShortcut(QKeySequence(Qt.Key_Escape), self); esc.activated.connect(self.ctrl.show_home)

        def add_file_field(cfg_key: str, label: str, filter_str: str):
            le = QLineEdit(self.cfg.get(cfg_key, "")); le.setReadOnly(True)
            btn = QPushButton("Browse…")

            if cfg_key == CONF_TEMPLATE or cfg_key == CONF_COMPANY_CSV:
                btn.setEnabled(False)

            def _browse():
                f, _ = QFileDialog.getOpenFileName(self, f"Select {label}", "M:\\Staffing Services Agency Intake", filter_str)
                if f:
                    le.setText(f)
                    set_path(cfg_key, f)   # write via central setter
            btn.clicked.connect(_browse)

            row = QWidget(); hb = QHBoxLayout(row); hb.setContentsMargins(0,0,0,0)
            hb.addWidget(le, 1); hb.addWidget(btn, 0)
            form.addRow(label + ":", row)

        def add_folder_field(cfg_key: str, label: str):
            le = QLineEdit(self.cfg.get(cfg_key, "")); le.setReadOnly(True)
            btn = QPushButton("Browse…")

            def _browse():
                d = QFileDialog.getExistingDirectory(self, f"Select {label}", "M:\\Staffing Services Agency Intake")
                if d:
                    le.setText(d)
                    set_path(cfg_key, d)

            btn.clicked.connect(_browse)

            row = QWidget(); hb = QHBoxLayout(row); hb.setContentsMargins(0,0,0,0)
            hb.addWidget(le, 1); hb.addWidget(btn, 0)
            form.addRow(label + ":", row)

        # Canonical keys & labels
        add_file_field(CONF_TRACKER,         "Tracker Excel", "Excel Files (*.xlsx *.xls)")
        add_file_field(CONF_TRACKER_2,       "Tracker 2 Excel (Optional)", "Excel Files (*.xlsx *.xls)")
        add_folder_field(CONF_SENT_INVOICES, "Sent Invoices Folder")
        add_folder_field(CONF_PENDING_INVOICES, "Pending Invoices Folder")
        add_file_field(CONF_TEMPLATE,        "RFP Template PDF",   "PDF Files (*.pdf)")
        add_file_field(CONF_COMPANY_CSV,     "Company CSV",          "CSV Files (*.csv);;Excel Files (*.xlsx *.xls)")

        root.addStretch()