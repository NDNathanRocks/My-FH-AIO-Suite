# MyAgencyPro (RFP Utility Application)

Author: Nathan Nicholas D Souza  
Team: QI Team - Staffing Services, Fraser Health  
Contact: nathan.dsouza@fraserhealth.ca

## Overview

MyAgencyPro is a PySide6 desktop app that supports the full invoice and RFP workflow:

1. Process Outlook .msg attachments into organized invoice folders.
2. Generate RFP PDFs from tracker data and company info.
3. Merge invoice files with matching RFPs.
4. Prepare statement reports and Outlook draft emails.
5. Run invoice checks and audit/lookup workflows.

The app is designed for day-to-day operational use with clear UI flows, progress indicators, and non-destructive file handling.

## Recent Updates

- Added and stabilized additional pages beyond the original 4-page flow:
  - Invoice Check
  - Search and Audit
  - Settings
  - Statistics
- Merger now includes:
  - Dry run mode
  - Skip tracker update mode
  - Consolidated Excel output with two sheets: Successfully Merged and Failed to Merge
  - Better handling of failed merges (left in place)
- Statement report flow now includes:
  - ZIP chunking for attachments (<= 20 MB per ZIP)
  - Outlook draft generation
  - Copy-to-Sent workflow by site/unit and date range
- Attachment processor preview upgraded:
  - Graphics-based preview
  - Zoom, fit-to-view, and rotation controls
- Improved tracker/Excel reliability and file-lock handling in key workflows.
- Logging expanded and centralized with uncaught exception capture and message-box logging.
- PyMuPDF import handling updated to avoid package collisions with the unrelated fitz package.

## Main Features By Page

### Home

Navigation hub for:
- Process Attachments
- Generate RFPs
- Merge Invoices
- Prepare Email
- Invoice Check
- Search and Audit
- Settings
- Statistics

### Process Attachments (Saver)

- Scans .msg files for PDF/image attachments.
- Skips tiny inline images.
- Shows current and upcoming attachments.
- Lets user assign Invoice, Site, and Unit before save.
- Supports preview email opening.
- Includes preview tools (zoom/fit/rotate).
- Saves into site/unit structure and keeps unique filenames when needed.

### Generate RFPs

- Uses tracker workbook + company CSV + template PDF.
- Supports filtering and generation workflows from tracker data.
- Writes filled PDFs with progress and status feedback.
- Integrates tracker-safe save patterns and recalculation handling.

### Merge Invoices

- Matches invoices and RFPs by invoice number logic.
- Merges into output folder while preserving site/unit structure.
- Supports dry run (no file changes).
- Supports skip tracker update.
- Produces a merge summary Excel report with success/failure sheets.
- Moves successful source invoices to recycle-bin area while keeping failed items visible for review.

### Prepare Email (Statement Report)

- Reads merged unit folder PDFs.
- Builds statement report Excel.
- Creates ZIP bundles when attachments exceed email-friendly size.
- Drafts Outlook emails using tracker summary mapping.
- Copies processed unit folder into Sent Invoices by date-range folder.

### Invoice Check

- Reconciles tracker rows with HA payment data.
- Updates status logic (paid/partially paid/not paid/due 0 behavior).
- Generates pending invoice outputs.
- Creates Outlook draft buttons per primary recipient with grouped attachments.
- Uses local cache for HA lookup and tracker recalculation markers to improve speed.

### Search and Audit

Wizard-style flow:
1. Select Tracker + HA payment files.
2. Choose either:
   - Single lookup (Vendor ID + Invoice #)
   - Quarterly report generation

Outputs include:
- Detailed audit sheet
- Quarterly summary sheet
- Per-site sheets
- Conditional formatting and totals

### Settings

Central path configuration for:
- Tracker Excel
- Sent Invoices folder
- Pending Invoices folder
- Template PDF
- Company CSV

### Statistics

- Displays generated RFP count from config.
- Manual refresh support.

## Installation

## Option 1: Run executable

If a packaged executable is available, run it directly.

## Option 2: Run from Python

1. Install Python 3.10+ (3.12 is supported in current environment).
2. Install dependencies:

```bash
pip install -r requirements.txt
```

3. Run the app:

```bash
python main.py
```

## Important dependency note

This project uses PyMuPDF. Do not install the unrelated fitz package.

If your environment previously had fitz installed, clean it once:

```bash
pip uninstall -y fitz frontend
pip install -r requirements.txt
```

## Basic Usage (Recommended Order)

1. Open Settings and confirm all path values.
2. Process Attachments to organize incoming files.
3. Generate RFPs from tracker data.
4. Merge Invoices with RFPs.
5. Prepare Email to generate statement + drafts.
6. Run Invoice Check and Audit as needed for reconciliation and reporting.

## Configuration and Data Files

- App config: utils/rfp_config.json
- Defaults: utils/paths_config.py
- Company mapping source: Company Info.csv
- App log file: utils/myagencypro.log

## Logging and Diagnostics

The app initializes rotating file logging and captures:
- General app logs
- Uncaught exceptions
- QMessageBox warning/error/info events

Default log location:
- utils/myagencypro.log

## Project Structure (Current)

- main.py
- pages/
  - home_page.py
  - saver_page.py
  - generator_page.py
  - merger_page.py
  - report_page.py
  - check_page.py
  - audit_page.py
  - settings_page.py
  - statistics_page.py
  - update_page.py (currently not active in navigation)
- utils/
  - app_utils.py
  - csv_utils.py
  - logging_utils.py
  - paths_config.py
- resources/
- Company Info.csv
- requirements.txt

## Workflow contract (do not change without owner approval)

These are relied on daily and are pinned by regression tests in `tests/`:

- Tracker layout: vendor sheets with headers on row 3, data from row 4; `Summary` and `Vendor Credits` are not vendor sheets. Columns used: `Vendor ID`, `Site`, `Unit`, `Invoice #`, `RFP Prepared (Y/N)`, `Merged (Y/N)`, `Invoice Status`, `Remittance ID` and the amount/date columns mapped in `utils/app_utils.py`.
- Generated RFP name: `{CompanyName} RFP {Invoice #} - {Site} {Unit}.pdf`. Merged name: `RFP {invoice} - {site} {unit}.pdf`.
- PDF field mapping in `process_csv_and_fill_pdfs` (requestor, cost codes `3509030 UPP` / `3509200 UPP`, dd/mm/yy dates).
- Only rows with `RFP Prepared (Y/N)` = `N` are generated when "only new" is on; generation stops at a blank `Invoice #` or a `Total` marker.
- Invoice-number extraction from file names (`extract_invoice_number`, back-ticked IDs in the merger).
- Tracker writes only change the target Y/N or status cells; formulas are preserved; saves are atomic (`save_workbook_safe`).

## Tests

```bash
pip install -r requirements-dev.txt
python -m pytest tests
```

Windows COM modules are stubbed, so the tests run on any OS. On Linux, PySide6 needs `libegl1`. One test is marked `xfail` on purpose: it documents a known crash that is waiting on a decision (see root `PLAN.md` D3).

## Known issues awaiting a decision

See root `PLAN.md` (D3–D5): the blank "RFP Prepared" fallback crashes; the "Total" end-marker column when the tracker has a leading blank column; blank HA check numbers recorded as `nan`.

## Troubleshooting

- RuntimeError about static/ during fitz import:
  - Remove conflicting fitz/frontend packages and reinstall requirements.
- Outlook draft creation fails:
  - Install pywin32 and ensure Outlook is available on the machine.
- Tracker save or replace errors on Windows/OneDrive:
  - Close open workbook handles and ensure file is not open in Excel.
- No output from scan/generation:
  - Re-check configured paths in Settings and file naming conventions.

## Support

For issues or enhancement requests, contact:
- nathan.dsouza@fraserhealth.ca
