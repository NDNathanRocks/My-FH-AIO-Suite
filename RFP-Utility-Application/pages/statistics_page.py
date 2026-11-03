# statistics_page.py

import os
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QSizePolicy, QToolBar
)
from PySide6.QtGui import QAction, QPalette, QColor, QShortcut, QKeySequence
from PySide6.QtCore import Qt

from utils.app_utils import load_config


class StatisticsPage(QWidget):
    def __init__(self, controller):
        super().__init__()
        self.ctrl = controller
        self.cfg  = load_config()

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
            QLabel#stat { font-size:48px; font-weight:bold; color:#2980b9; }
            QLabel#label{ font-size:16px; }
            QPushButton { background:#2980b9; color:white;
                          border:none; border-radius:6px;
                          padding:8px 16px; font-size:14px; }
            QPushButton:hover { background:#3498db; }
        """)

        # ── Layout ────────────────────────────────────────────────
        root = QVBoxLayout(self)
        root.setContentsMargins(50,40,50,40)
        root.setSpacing(20)

        # Toolbar with Back
        tb = QToolBar()
        back = QAction("← Back", self)
        back.triggered.connect(self.ctrl.show_home)
        tb.addAction(back)
        root.addWidget(tb)

        # Title
        title = QLabel("📈 Statistics")
        title.setAlignment(Qt.AlignHCenter)
        title.setObjectName("label")
        root.addWidget(title)

        # RFPs generated stat
        self.stat_rfps = QLabel("", alignment=Qt.AlignHCenter)
        self.stat_rfps.setObjectName("stat")
        root.addWidget(self.stat_rfps)

        lbl = QLabel("RFPs Generated", alignment=Qt.AlignHCenter)
        lbl.setObjectName("label")
        root.addWidget(lbl)

        # Refresh button
        btn = QPushButton("Refresh")
        btn.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        btn.clicked.connect(self._load_stats)
        root.addWidget(btn, alignment=Qt.AlignHCenter)

        root.addStretch()
        self._load_stats()

        esc = QShortcut(QKeySequence(Qt.Key_Escape), self)
        esc.activated.connect(self.ctrl.show_home)

    def _load_stats(self):
        cfg = load_config()
        count = cfg.get("rfp_count", 0)
        self.stat_rfps.setText(str(count))