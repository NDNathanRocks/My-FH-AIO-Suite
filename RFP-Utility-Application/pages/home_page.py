# home_page.py

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QSizePolicy
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QPalette, QColor

class HomePage(QWidget):
    def __init__(self, controller):
        super().__init__()
        self.ctrl = controller
        self.setObjectName("HomePage")
        self._apply_palette()
        self._apply_style()

        # Central layout to center content
        layout = QVBoxLayout(self)
        layout.setContentsMargins(60, 40, 60, 40)
        layout.setSpacing(20)
        layout.addStretch()

        # ── Title ─────────────────────────────────────────────────────
        title_lbl = QLabel("🏠 MyAgencyPro")
        title_lbl.setAlignment(Qt.AlignHCenter)
        title_lbl.setStyleSheet("font-size:32px; font-weight:bold; color:#3498db;")
        layout.addWidget(title_lbl)

        # ── Main buttons ─────────────────────────────────────────────
        btn_group = QWidget()
        btn_layout = QVBoxLayout(btn_group)
        btn_layout.setSpacing(15)
        for text, slot in [
            ("Process Attachments", self.ctrl.show_saver),
            ("Generate RFPs",       self.ctrl.show_generator),
            ("Merge Invoices",      self.ctrl.show_merger),
            ("Prepare Email",       self.ctrl.show_report),
            ("Invoice Check",       self.ctrl.show_check),
            ("Search and Audit",    self.ctrl.show_audit),
            # ("Import/Export",       self.ctrl.show_update_data),
        ]:
            btn = QPushButton(text)
            btn.setMinimumHeight(40)
            btn.clicked.connect(slot)
            btn_layout.addWidget(btn)
        layout.addWidget(btn_group, alignment=Qt.AlignHCenter)

        # ── Bottom row: Settings & Statistics ────────────────────────
        bottom = QWidget()
        bot_layout = QHBoxLayout(bottom)
        bot_layout.setContentsMargins(0, 0, 0, 0)  # no extra margins around the layout
        bot_layout.setAlignment(Qt.AlignCenter)   # center the contents

        btn_settings = QPushButton("Settings")
        btn_settings.setObjectName("smallBtn")
        btn_settings.setFixedWidth(50)
        btn_settings.setMinimumHeight(40)
        btn_settings.clicked.connect(self.ctrl.show_settings)

        btn_stats = QPushButton("Statistics")
        btn_stats.setObjectName("smallBtn")
        btn_stats.setFixedWidth(50)
        btn_stats.setMinimumHeight(40)
        btn_stats.clicked.connect(self.ctrl.show_statistics)

        bot_layout.addWidget(btn_settings)
        bot_layout.addWidget(btn_stats)

        # add the bottom widget centered in the parent layout
        layout.addWidget(bottom, alignment=Qt.AlignHCenter)

        layout.addStretch()

    def _apply_palette(self):
        pal = QPalette()
        pal.setColor(QPalette.Window, QColor(245,245,245))
        pal.setColor(QPalette.WindowText, QColor(33,33,33))
        pal.setColor(QPalette.Button, QColor(238,238,238))
        pal.setColor(QPalette.ButtonText, QColor(33,33,33))
        pal.setColor(QPalette.Highlight, QColor(52,152,219))
        pal.setColor(QPalette.HighlightedText, QColor(255,255,255))
        self.setPalette(pal)

    def _apply_style(self):
        self.setStyleSheet("""
            /* Main buttons stay at 200px */
            QPushButton {
                background:#2980b9; color:white;
                border:none; border-radius:6px;
                padding:8px 16px; font-size:16px;
                width:200px;
            }
            QPushButton:hover {
                background:#3498db;
            }
            QPushButton:disabled {
                background:#bdc3c7; color:#7f8c7f;
            }

            /* Bottom-bow “smallBtn” override */
            QPushButton#smallBtn {
                min-width: 50px;
                max-width: 50px;
                width: 98px;
                padding: 8px;
            }
        """)
