# update_page.py

from __future__ import annotations

import os
import sqlite3
from typing import Dict, List, Optional, Tuple
from dataclasses import dataclass
from datetime import datetime

import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout,
    QToolBar, QApplication, QTableWidget,
    QLabel, QLineEdit, QPushButton,
    QComboBox, QProgressBar, QMessageBox,
    QFileDialog, QTableWidgetItem, QHeaderView,
    QAbstractItemView, QSplitter
)
from PySide6.QtGui import (
    QAction, QPalette, QColor, QKeySequence, QShortcut, QGuiApplication
)
from PySide6.QtCore import Qt, Signal, QTimer

from utils.app_utils import load_config, save_config_merge


# ──────────────────────────────────────────────────────────────────────────────
# Constants
# ──────────────────────────────────────────────────────────────────────────────
DEFAULT_DB_FOLDER = "./databases"


# ──────────────────────────────────────────────────────────────────────────────
# Data Classes
# ──────────────────────────────────────────────────────────────────────────────
@dataclass
class SheetInfo:
    """Information about a sheet in the database."""
    name: str
    row_count: int
    column_count: int
    columns: List[str]


# ──────────────────────────────────────────────────────────────────────────────
# Database Handler
# ──────────────────────────────────────────────────────────────────────────────
class ExcelDatabaseHandler:
    """Handle Excel to SQLite conversion and vice versa."""
    
    def __init__(self, db_path: str):
        self.db_path = db_path
        self.conn: Optional[sqlite3.Connection] = None
        self._ensure_db_exists()
    
    def _ensure_db_exists(self) -> None:
        """Create database and metadata table if they don't exist."""
        os.makedirs(os.path.dirname(self.db_path) or ".", exist_ok=True)
        conn = sqlite3.connect(self.db_path)
        try:
            conn.execute("""
                CREATE TABLE IF NOT EXISTS _metadata (
                    sheet_name TEXT PRIMARY KEY,
                    original_name TEXT,
                    row_count INTEGER,
                    column_count INTEGER,
                    columns TEXT,
                    created_at TEXT,
                    modified_at TEXT
                )
            """)
            conn.commit()
        finally:
            conn.close()
    
    def connect(self) -> sqlite3.Connection:
        """Get or create database connection."""
        if self.conn is None:
            self.conn = sqlite3.connect(self.db_path)
            self.conn.row_factory = sqlite3.Row
        return self.conn
    
    def close(self) -> None:
        """Close database connection."""
        if self.conn:
            try:
                self.conn.close()
            except Exception:
                pass
            finally:
                self.conn = None
    
    def __enter__(self):
        self.connect()
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()
    
    @staticmethod
    def _sanitize_table_name(name: str) -> str:
        """Convert sheet name to valid SQLite table name."""
        # Replace spaces and special chars with underscores
        sanitized = "".join(c if c.isalnum() else "_" for c in name)
        # Ensure it starts with a letter
        if sanitized and not sanitized[0].isalpha():
            sanitized = "sheet_" + sanitized
        return sanitized or "sheet_unnamed"
    
    @staticmethod
    def _sanitize_column_name(name: str) -> str:
        """Convert column name to valid SQLite column name."""
        if not name or not str(name).strip():
            return "column_unnamed"
        sanitized = "".join(c if c.isalnum() else "_" for c in str(name))
        if sanitized and sanitized[0].isdigit():
            sanitized = "col_" + sanitized
        return sanitized or "column_unnamed"
    
    def import_excel(self, excel_path: str, progress_callback=None) -> List[str]:
        """
        Import all sheets from Excel file into database.
        
        Args:
            excel_path: Path to Excel file
            progress_callback: Optional callback(current, total, message)
            
        Returns:
            List of imported sheet names
        """
        imported_sheets = []
        wb = None
        
        try:
            wb = load_workbook(excel_path, data_only=True, read_only=True)
            sheet_names = [s for s in wb.sheetnames]
            total_sheets = len(sheet_names)
            
            conn = self.connect()
            
            for idx, sheet_name in enumerate(sheet_names, 1):
                if progress_callback:
                    progress_callback(idx, total_sheets, f"Importing sheet: {sheet_name}")
                
                try:
                    # Read sheet data
                    df = pd.read_excel(
                        excel_path,
                        sheet_name=sheet_name,
                        dtype=str,
                        engine="openpyxl"
                    )
                    
                    # Clean column names
                    df.columns = [self._sanitize_column_name(col) for col in df.columns]
                    
                    # Handle duplicate column names
                    cols = []
                    col_counts: Dict[str, int] = {}
                    for col in df.columns:
                        if col in col_counts:
                            col_counts[col] += 1
                            cols.append(f"{col}_{col_counts[col]}")
                        else:
                            col_counts[col] = 0
                            cols.append(col)
                    df.columns = cols
                    
                    # Create sanitized table name
                    table_name = self._sanitize_table_name(sheet_name)
                    
                    # Drop existing table if it exists
                    conn.execute(f"DROP TABLE IF EXISTS {table_name}")
                    
                    # Write to database
                    df.to_sql(table_name, conn, if_exists="replace", index=False)
                    
                    # Update metadata
                    now = datetime.now().isoformat()
                    conn.execute("""
                        INSERT OR REPLACE INTO _metadata 
                        (sheet_name, original_name, row_count, column_count, columns, created_at, modified_at)
                        VALUES (?, ?, ?, ?, ?, ?, ?)
                    """, (
                        table_name,
                        sheet_name,
                        len(df),
                        len(df.columns),
                        ",".join(df.columns),
                        now,
                        now
                    ))
                    conn.commit()
                    
                    imported_sheets.append(sheet_name)
                    
                    # Clean up DataFrame
                    del df
                    
                except Exception as e:
                    raise Exception(f"Error importing sheet '{sheet_name}': {str(e)}")
            
            return imported_sheets
            
        finally:
            if wb:
                try:
                    wb.close()
                except Exception:
                    pass
    
    def export_excel(self, excel_path: str, progress_callback=None) -> int:
        """
        Export all sheets from database to Excel file.
        
        Args:
            excel_path: Path to output Excel file
            progress_callback: Optional callback(current, total, message)
            
        Returns:
            Number of sheets exported
        """
        sheets = self.get_sheet_list()
        total_sheets = len(sheets)
        
        if total_sheets == 0:
            raise ValueError("No sheets found in database")
        
        wb = Workbook()
        # Remove default sheet
        if "Sheet" in wb.sheetnames:
            del wb["Sheet"]
        
        try:
            conn = self.connect()
            
            for idx, sheet_info in enumerate(sheets, 1):
                if progress_callback:
                    progress_callback(idx, total_sheets, f"Exporting: {sheet_info.name}")
                
                # Read data from database
                df = pd.read_sql_query(f"SELECT * FROM {sheet_info.name}", conn)
                
                # Create worksheet with original name
                ws = wb.create_sheet(title=sheet_info.original_name)
                
                # Write data
                for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), 1):
                    for c_idx, value in enumerate(row, 1):
                        # Convert None to empty string for Excel
                        cell_value = "" if value is None or (isinstance(value, float) and pd.isna(value)) else value
                        ws.cell(row=r_idx, column=c_idx, value=cell_value)
                
                # Clean up DataFrame
                del df
            
            # Save workbook
            wb.save(excel_path)
            return total_sheets
            
        finally:
            try:
                wb.close()
            except Exception:
                pass
    
    def get_sheet_list(self) -> List[SheetInfo]:
        """Get list of all sheets in database."""
        conn = self.connect()
        cursor = conn.execute("""
            SELECT sheet_name, original_name, row_count, column_count, columns
            FROM _metadata
            ORDER BY original_name
        """)
        
        sheets = []
        for row in cursor.fetchall():
            sheets.append(SheetInfo(
                name=row["sheet_name"],
                row_count=row["row_count"],
                column_count=row["column_count"],
                columns=row["columns"].split(",") if row["columns"] else []
            ))
        
        # Add original_name attribute for display
        cursor = conn.execute("SELECT sheet_name, original_name FROM _metadata")
        name_map = {row["sheet_name"]: row["original_name"] for row in cursor.fetchall()}
        
        for sheet in sheets:
            sheet.original_name = name_map.get(sheet.name, sheet.name)
        
        return sheets
    
    def get_sheet_data(self, sheet_name: str, limit: Optional[int] = None) -> pd.DataFrame:
        """
        Get data from a specific sheet.
        
        Args:
            sheet_name: Name of sheet (sanitized table name)
            limit: Optional row limit for large datasets
            
        Returns:
            DataFrame with sheet data
        """
        conn = self.connect()
        query = f"SELECT * FROM {sheet_name}"
        if limit:
            query += f" LIMIT {limit}"
        return pd.read_sql_query(query, conn)
    
    def update_sheet_data(self, sheet_name: str, df: pd.DataFrame) -> None:
        """Update sheet data in database."""
        conn = self.connect()
        
        # Replace table data
        df.to_sql(sheet_name, conn, if_exists="replace", index=False)
        
        # Update metadata
        now = datetime.now().isoformat()
        conn.execute("""
            UPDATE _metadata 
            SET row_count = ?, column_count = ?, columns = ?, modified_at = ?
            WHERE sheet_name = ?
        """, (
            len(df),
            len(df.columns),
            ",".join(df.columns),
            now,
            sheet_name
        ))
        conn.commit()


# ──────────────────────────────────────────────────────────────────────────────
# UI Class
# ──────────────────────────────────────────────────────────────────────────────
class UpdatePage(QWidget):
    """
    Page for converting Excel files to database and vice versa.
    Allows viewing and editing data with sheet-wise organization.
    """
    
    def __init__(self, controller):
        super().__init__()
        self.ctrl = controller
        
        # State
        self.current_db_path: str = ""
        self.db_handler: Optional[ExcelDatabaseHandler] = None
        self.current_sheet: Optional[SheetInfo] = None
        self.current_df: Optional[pd.DataFrame] = None
        
        # UI setup
        self._apply_palette()
        self._apply_style()
        self._build_ui()
        
        # ESC → Back
        esc = QShortcut(QKeySequence(Qt.Key_Escape), self)
        esc.activated.connect(self._on_back)
        
        # Load last database if exists
        self._load_last_database()
    
    def _build_ui(self) -> None:
        """Build the UI layout."""
        root = QVBoxLayout(self)
        root.setContentsMargins(50, 40, 50, 40)
        root.setSpacing(20)
        
        # Toolbar
        toolbar = QToolBar()
        back_act = QAction("← Back", self)
        back_act.triggered.connect(self._on_back)
        toolbar.addAction(back_act)
        root.addWidget(toolbar)
        
        # Title
        title = QLabel("📊 Excel ↔ Database Converter")
        title.setAlignment(Qt.AlignHCenter)
        title.setStyleSheet("font-size:32px; font-weight:bold; color:#3498db;")
        root.addWidget(title)
        
        # File selection row
        file_box = QWidget()
        file_layout = QHBoxLayout(file_box)
        file_layout.setContentsMargins(0, 0, 0, 0)
        
        file_layout.addWidget(QLabel("Database:"))
        self.db_path_le = QLineEdit()
        self.db_path_le.setReadOnly(True)
        self.db_path_le.setPlaceholderText("No database loaded")
        file_layout.addWidget(self.db_path_le, 1)
        
        btn_new_db = QPushButton("New Database")
        btn_new_db.clicked.connect(self._on_new_database)
        file_layout.addWidget(btn_new_db)
        
        btn_open_db = QPushButton("Open Database")
        btn_open_db.clicked.connect(self._on_open_database)
        file_layout.addWidget(btn_open_db)
        
        root.addWidget(file_box)
        
        # Action buttons row
        action_box = QWidget()
        action_layout = QHBoxLayout(action_box)
        action_layout.setContentsMargins(0, 0, 0, 0)
        
        self.btn_import = QPushButton("📥 Import Excel → Database")
        self.btn_import.clicked.connect(self._on_import_excel)
        self.btn_import.setEnabled(False)
        action_layout.addWidget(self.btn_import)
        
        self.btn_export = QPushButton("📤 Export Database → Excel")
        self.btn_export.clicked.connect(self._on_export_excel)
        self.btn_export.setEnabled(False)
        action_layout.addWidget(self.btn_export)
        
        action_layout.addStretch()
        
        root.addWidget(action_box)
        
        # Progress bar
        self.progress = QProgressBar()
        self.progress.setVisible(False)
        root.addWidget(self.progress)
        
        # Status label
        self.lbl_status = QLabel("")
        self.lbl_status.setAlignment(Qt.AlignCenter)
        root.addWidget(self.lbl_status)
        
        # Splitter for sheet list and data view
        splitter = QSplitter(Qt.Horizontal)
        
        # Left panel: Sheet list
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(0, 0, 0, 0)
        
        left_layout.addWidget(QLabel("Sheets:"))
        
        self.sheet_combo = QComboBox()
        self.sheet_combo.currentIndexChanged.connect(self._on_sheet_changed)
        left_layout.addWidget(self.sheet_combo)
        
        self.sheet_info_label = QLabel("")
        self.sheet_info_label.setWordWrap(True)
        self.sheet_info_label.setStyleSheet("padding: 10px; background: #ecf0f1; border-radius: 4px;")
        left_layout.addWidget(self.sheet_info_label)
        
        left_layout.addStretch()
        
        # Right panel: Data table
        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(0, 0, 0, 0)
        
        right_layout.addWidget(QLabel("Data Preview:"))
        
        self.data_table = QTableWidget()
        self.data_table.setEditTriggers(QAbstractItemView.DoubleClicked)
        self.data_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.data_table.horizontalHeader().setStretchLastSection(True)
        self.data_table.setAlternatingRowColors(True)
        right_layout.addWidget(self.data_table)
        
        # Data action buttons
        data_btn_layout = QHBoxLayout()
        
        self.btn_refresh = QPushButton("🔄 Refresh")
        self.btn_refresh.clicked.connect(self._on_refresh_data)
        self.btn_refresh.setEnabled(False)
        data_btn_layout.addWidget(self.btn_refresh)
        
        self.btn_save_changes = QPushButton("💾 Save Changes")
        self.btn_save_changes.clicked.connect(self._on_save_changes)
        self.btn_save_changes.setEnabled(False)
        data_btn_layout.addWidget(self.btn_save_changes)
        
        data_btn_layout.addStretch()
        
        right_layout.addLayout(data_btn_layout)
        
        # Add panels to splitter
        splitter.addWidget(left_panel)
        splitter.addWidget(right_panel)
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 3)
        
        root.addWidget(splitter, 1)
    
    def _apply_palette(self) -> None:
        """Apply color palette."""
        pal = QPalette()
        pal.setColor(QPalette.Window, QColor(245, 245, 245))
        pal.setColor(QPalette.WindowText, QColor(33, 33, 33))
        pal.setColor(QPalette.Button, QColor(238, 238, 238))
        pal.setColor(QPalette.ButtonText, QColor(33, 33, 33))
        pal.setColor(QPalette.Highlight, QColor(52, 152, 219))
        pal.setColor(QPalette.HighlightedText, QColor(255, 255, 255))
        self.setPalette(pal)
    
    def _apply_style(self) -> None:
        """Apply stylesheet."""
        self.setStyleSheet("""
            QPushButton {
                background: #2980b9;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 8px 16px;
                font-size: 14px;
            }
            QPushButton:disabled {
                background: #bdc3c7;
                color: #7f8c8d;
            }
            QPushButton:hover:enabled {
                background: #3498db;
            }
            QLineEdit, QComboBox {
                border: 1px solid #bdc3c7;
                border-radius: 4px;
                padding: 4px;
            }
            QTableWidget {
                border: 1px solid #bdc3c7;
                border-radius: 4px;
                gridline-color: #ecf0f1;
            }
            QTableWidget::item {
                padding: 4px;
            }
            QHeaderView::section {
                background-color: #34495e;
                color: white;
                padding: 6px;
                border: none;
                font-weight: bold;
            }
        """)
    
    def _set_controls_enabled(self, enabled: bool) -> None:
        """Enable/disable controls during operations."""
        for widget in [self.btn_import, self.btn_export, self.btn_refresh, 
                      self.btn_save_changes, self.sheet_combo]:
            widget.setEnabled(enabled and self.db_handler is not None)
    
    def _begin_busy(self, msg: str) -> None:
        """Start busy state."""
        self.lbl_status.setText(msg)
        self.progress.setVisible(True)
        self.progress.setValue(0)
        QGuiApplication.setOverrideCursor(Qt.WaitCursor)
        self._set_controls_enabled(False)
        QApplication.processEvents()
    
    def _update_busy(self, current: int, total: int, msg: str) -> None:
        """Update busy state."""
        self.progress.setMaximum(total)
        self.progress.setValue(current)
        self.lbl_status.setText(f"{msg} ({current}/{total})")
        QApplication.processEvents()
    
    def _end_busy(self, msg: str = "") -> None:
        """End busy state."""
        self.lbl_status.setText(msg)
        self.progress.setVisible(False)
        QGuiApplication.restoreOverrideCursor()
        self._set_controls_enabled(True)
        QApplication.processEvents()
    
    def _load_last_database(self) -> None:
        """Load the last used database from config."""
        config = load_config()
        last_db = config.get("last_excel_database", "")
        
        if last_db and os.path.exists(last_db):
            try:
                self._load_database(last_db)
            except Exception as e:
                QMessageBox.warning(self, "Error", f"Failed to load last database:\n{str(e)}")
    
    def _load_database(self, db_path: str) -> None:
        """Load a database file."""
        try:
            # Close existing handler
            if self.db_handler:
                self.db_handler.close()
            
            # Create new handler
            self.db_handler = ExcelDatabaseHandler(db_path)
            self.current_db_path = db_path
            self.db_path_le.setText(os.path.basename(db_path))
            self.db_path_le.setToolTip(db_path)
            
            # Save to config
            save_config_merge({"last_excel_database": db_path})
            
            # Load sheets
            self._refresh_sheet_list()
            
            # Enable buttons
            self.btn_import.setEnabled(True)
            self.btn_export.setEnabled(True)
            
            self.lbl_status.setText(f"Loaded database: {os.path.basename(db_path)}")
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load database:\n{str(e)}")
            raise
    
    def _refresh_sheet_list(self) -> None:
        """Refresh the sheet dropdown."""
        if not self.db_handler:
            return
        
        self.sheet_combo.blockSignals(True)
        self.sheet_combo.clear()
        
        try:
            sheets = self.db_handler.get_sheet_list()
            
            if sheets:
                for sheet in sheets:
                    display_name = f"{sheet.original_name} ({sheet.row_count} rows)"
                    self.sheet_combo.addItem(display_name, sheet)
                
                self.btn_refresh.setEnabled(True)
                self.btn_save_changes.setEnabled(True)
            else:
                self.sheet_combo.addItem("No sheets available")
                self.btn_refresh.setEnabled(False)
                self.btn_save_changes.setEnabled(False)
                
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to load sheets:\n{str(e)}")
        finally:
            self.sheet_combo.blockSignals(False)
        
        # Load first sheet if available
        if self.sheet_combo.count() > 0:
            self._on_sheet_changed(0)
    
    def _on_new_database(self) -> None:
        """Create a new database."""
        db_path, _ = QFileDialog.getSaveFileName(
            self,
            "Create New Database",
            DEFAULT_DB_FOLDER,
            "SQLite Database (*.db *.sqlite *.sqlite3)"
        )
        
        if db_path:
            # Ensure .db extension
            if not db_path.lower().endswith(('.db', '.sqlite', '.sqlite3')):
                db_path += '.db'
            
            try:
                self._load_database(db_path)
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to create database:\n{str(e)}")
    
    def _on_open_database(self) -> None:
        """Open an existing database."""
        db_path, _ = QFileDialog.getOpenFileName(
            self,
            "Open Database",
            DEFAULT_DB_FOLDER,
            "SQLite Database (*.db *.sqlite *.sqlite3);;All Files (*)"
        )
        
        if db_path:
            try:
                self._load_database(db_path)
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to open database:\n{str(e)}")
    
    def _on_import_excel(self) -> None:
        """Import Excel file into database."""
        if not self.db_handler:
            QMessageBox.warning(self, "No Database", "Please create or open a database first.")
            return
        
        excel_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Excel File to Import",
            "",
            "Excel Files (*.xlsx *.xls);;All Files (*)"
        )
        
        if not excel_path:
            return
        
        # Confirm overwrite
        sheets = self.db_handler.get_sheet_list()
        if sheets:
            reply = QMessageBox.question(
                self,
                "Confirm Import",
                "This will replace all existing data in the database. Continue?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No
            )
            if reply != QMessageBox.Yes:
                return
        
        try:
            self._begin_busy("Importing Excel file...")
            
            imported = self.db_handler.import_excel(
                excel_path,
                progress_callback=self._update_busy
            )
            
            self._end_busy(f"Successfully imported {len(imported)} sheet(s)")
            
            # Refresh UI
            self._refresh_sheet_list()
            
            QMessageBox.information(
                self,
                "Import Complete",
                f"Successfully imported {len(imported)} sheet(s) from:\n{os.path.basename(excel_path)}"
            )
            
        except Exception as e:
            self._end_busy("")
            QMessageBox.critical(self, "Import Error", f"Failed to import Excel file:\n{str(e)}")
    
    def _on_export_excel(self) -> None:
        """Export database to Excel file."""
        if not self.db_handler:
            QMessageBox.warning(self, "No Database", "No database loaded.")
            return
        
        sheets = self.db_handler.get_sheet_list()
        if not sheets:
            QMessageBox.warning(self, "No Data", "Database contains no sheets to export.")
            return
        
        excel_path, _ = QFileDialog.getSaveFileName(
            self,
            "Export to Excel",
            "",
            "Excel Files (*.xlsx);;All Files (*)"
        )
        
        if not excel_path:
            return
        
        # Ensure .xlsx extension
        if not excel_path.lower().endswith('.xlsx'):
            excel_path += '.xlsx'
        
        try:
            self._begin_busy("Exporting to Excel...")
            
            count = self.db_handler.export_excel(
                excel_path,
                progress_callback=self._update_busy
            )
            
            self._end_busy(f"Successfully exported {count} sheet(s)")
            
            QMessageBox.information(
                self,
                "Export Complete",
                f"Successfully exported {count} sheet(s) to:\n{excel_path}"
            )
            
            # Ask to open file
            reply = QMessageBox.question(
                self,
                "Open File?",
                "Would you like to open the exported file?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.Yes
            )
            
            if reply == QMessageBox.Yes:
                from PySide6.QtCore import QUrl
                from PySide6.QtGui import QDesktopServices
                QDesktopServices.openUrl(QUrl.fromLocalFile(excel_path))
            
        except Exception as e:
            self._end_busy("")
            QMessageBox.critical(self, "Export Error", f"Failed to export to Excel:\n{str(e)}")
    
    def _on_sheet_changed(self, index: int) -> None:
        """Handle sheet selection change."""
        if index < 0 or not self.db_handler:
            return
        
        sheet_info = self.sheet_combo.itemData(index)
        if not isinstance(sheet_info, SheetInfo):
            return
        
        self.current_sheet = sheet_info
        
        # Update info label
        info_text = (
            f"<b>Sheet:</b> {sheet_info.original_name}<br>"
            f"<b>Rows:</b> {sheet_info.row_count}<br>"
            f"<b>Columns:</b> {sheet_info.column_count}"
        )
        self.sheet_info_label.setText(info_text)
        
        # Load data
        self._load_sheet_data()
    
    def _load_sheet_data(self, limit: int = 1000) -> None:
        """Load data for current sheet into table."""
        if not self.current_sheet or not self.db_handler:
            return
        
        try:
            # Show loading message for large datasets
            if self.current_sheet.row_count > limit:
                self.lbl_status.setText(f"Loading first {limit} rows (of {self.current_sheet.row_count})...")
                QApplication.processEvents()
            
            df = self.db_handler.get_sheet_data(self.current_sheet.name, limit=limit)
            self.current_df = df
            
            # Populate table
            self.data_table.clear()
            self.data_table.setRowCount(len(df))
            self.data_table.setColumnCount(len(df.columns))
            self.data_table.setHorizontalHeaderLabels(df.columns.tolist())
            
            for row_idx, (_, row) in enumerate(df.iterrows()):
                for col_idx, value in enumerate(row):
                    item = QTableWidgetItem(str(value) if pd.notna(value) else "")
                    self.data_table.setItem(row_idx, col_idx, item)
            
            # Resize columns to content
            self.data_table.resizeColumnsToContents()
            
            if self.current_sheet.row_count > limit:
                self.lbl_status.setText(f"Showing {limit} of {self.current_sheet.row_count} rows")
            else:
                self.lbl_status.setText(f"Loaded {len(df)} rows")
            
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to load sheet data:\n{str(e)}")
            self.lbl_status.setText("Error loading data")
    
    def _on_refresh_data(self) -> None:
        """Refresh current sheet data."""
        if self.current_sheet:
            self._load_sheet_data()
    
    def _on_save_changes(self) -> None:
        """Save changes made to the data table back to database."""
        if not self.current_sheet or not self.db_handler or self.current_df is None:
            return
        
        reply = QMessageBox.question(
            self,
            "Confirm Save",
            f"Save changes to sheet '{self.current_sheet.original_name}'?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply != QMessageBox.Yes:
            return
        
        try:
            # Collect data from table
            rows = []
            for row_idx in range(self.data_table.rowCount()):
                row_data = {}
                for col_idx in range(self.data_table.columnCount()):
                    header = self.data_table.horizontalHeaderItem(col_idx).text()
                    item = self.data_table.item(row_idx, col_idx)
                    value = item.text() if item else ""
                    row_data[header] = value if value else None
                rows.append(row_data)
            
            # Create DataFrame
            df = pd.DataFrame(rows)
            
            # Save to database
            self.db_handler.update_sheet_data(self.current_sheet.name, df)
            
            # Refresh sheet list (row count may have changed)
            self._refresh_sheet_list()
            
            self.lbl_status.setText("Changes saved successfully")
            QMessageBox.information(self, "Success", "Changes saved to database.")
            
        except Exception as e:
            QMessageBox.critical(self, "Save Error", f"Failed to save changes:\n{str(e)}")
    
    def _on_back(self) -> None:
        """Return to home view."""
        self.ctrl.show_home()
    
    def closeEvent(self, event):
        """Clean up when closing."""
        if self.db_handler:
            self.db_handler.close()
        super().closeEvent(event)