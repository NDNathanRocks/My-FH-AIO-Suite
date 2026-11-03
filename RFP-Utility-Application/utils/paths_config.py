# Authored By: Nathan Nicholas D Souza
# Team: QI Team - Staffing Services at Fraser Health
# Contact: nathan.dsouza@fraserhealth.ca

# paths_config.py

"""
Centralized path configuration for the application.
Edit here to change any of the default folders.
"""

#: Where to look for incoming invoice PDFs
DEFAULT_INVOICES_FOLDER      = r"M:\Staffing Services Agency Intake\Invoices"

#: Where to look for incoming RFP PDFs
DEFAULT_RFPS_FOLDER          = r"M:\Staffing Services Agency Intake\RFPs"

#: Where merged RFP+Invoices are written
DEFAULT_MERGED_FOLDER        = r"M:\Staffing Services Agency Intake\Merged Invoices and RFPs"

#: “Recycle Bin” folder where we’ll stash the merged invoices
RECYCLE_BIN_FOLDER           = r"M:\Staffing Services Agency Intake\Recycle Bin"

#: Where to look for the invoices tracker XLSX file
DEFAULT_INVOICE_TRACKER      = r"M:\Staffing Services Agency Intake\Fiscal Tracking\2025-2026 Fiscal\2025-2026FY Invoices Tracker.xlsx"