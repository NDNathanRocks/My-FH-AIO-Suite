# Authored By: Nathan Nicholas D Souza
# Team: QI Team - Staffing Services at Fraser Health
# Contact: nathan.dsouza@fraserhealth.ca

# csv_utils.py

import csv
import os
import stat
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout,
    QTableWidget, QTableWidgetItem,
    QPushButton, QMessageBox
)

class EditCSVDialog(QDialog):
    """
    A dialog for viewing and editing a CSV file in a QTableWidget.
    """
    def __init__(self, csv_path: str, parent=None):
        super().__init__(parent)
        self.csv_path = csv_path
        self.setWindowTitle("Edit Company Info CSV")
        self.resize(800, 600)

        # Main layout
        layout = QVBoxLayout(self)

        # Table widget
        self.table = QTableWidget()
        layout.addWidget(self.table)

        # Buttons: Add Row, Delete Row, Save, Cancel
        btn_layout = QHBoxLayout()
        add_btn = QPushButton("Add Row")
        add_btn.clicked.connect(self.add_row)
        del_btn = QPushButton("Delete Row")
        del_btn.clicked.connect(self.delete_row)
        save_btn = QPushButton("Save")
        save_btn.clicked.connect(self.save)
        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)
        btn_layout.addWidget(add_btn)
        btn_layout.addWidget(del_btn)
        btn_layout.addStretch()
        btn_layout.addWidget(save_btn)
        btn_layout.addWidget(cancel_btn)
        layout.addLayout(btn_layout)

        # Load CSV data
        self.load()

    def load(self):
        """
        Load CSV file into the table widget, excluding the header row.
        """
        try:
            with open(self.csv_path, newline='', encoding='utf-8-sig') as f:
                reader = csv.reader(f)
                rows = list(reader)
            if not rows:
                return
            headers = rows[0]
            data = rows[1:]

            self.table.setColumnCount(len(headers))
            self.table.setRowCount(len(data))
            self.table.setHorizontalHeaderLabels(headers)

            for i, row in enumerate(data):
                for j, val in enumerate(row):
                    item = QTableWidgetItem(val)
                    self.table.setItem(i, j, item)
        except Exception as e:
            QMessageBox.critical(self, "Error loading CSV", str(e))
            self.reject()

    def add_row(self):
        """
        Insert an empty row at the end of the table.
        """
        self.table.insertRow(self.table.rowCount())

    def delete_row(self):
        """
        Remove the currently selected row.
        """
        row = self.table.currentRow()
        if row >= 0:
            self.table.removeRow(row)

    def save(self):
        """
        Save the table contents back to the CSV file, preserving headers.
        """
        try:
            # Make file writable
            os.chmod(self.csv_path, stat.S_IWRITE)

            # Extract headers
            headers = [
                self.table.horizontalHeaderItem(c).text()
                for c in range(self.table.columnCount())
            ]
            rows = [headers]

            # Extract each data row
            for i in range(self.table.rowCount()):
                row = []
                for j in range(self.table.columnCount()):
                    item = self.table.item(i, j)
                    row.append(item.text() if item else '')
                rows.append(row)

            # Write back to file
            with open(self.csv_path, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerows(rows)

            # Restore read-only permissions
            os.chmod(self.csv_path, stat.S_IREAD | stat.S_IRGRP | stat.S_IROTH)

            self.accept()
        except Exception as e:
            QMessageBox.critical(self, "Error saving CSV", str(e))
