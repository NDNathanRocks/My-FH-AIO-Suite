# AI Coding Agent Instructions

## Project Overview

**RFP Utility Application**  
Author: Nathan Nicholas D Souza  
Team: QI Team - Staffing Services, Fraser Health  
Contact: nathan.dsouza@fraserhealth.ca

This is a desktop application built with **PySide6** (Qt for Python) that helps process invoices and RFP (Request for Proposal) documents for staffing services. The application provides a GUI-based workflow for:

1. **Email Attachment Processing** - Extract and organize PDF/image attachments from Outlook `.msg` files
2. **RFP Generation** - Auto-fill RFP forms from invoice CSV data
3. **Document Merging** - Merge invoices with their corresponding RFPs
4. **Report Generation** - Create Excel statement reports from processed documents
5. **Audit & Checking Tools** - Verify and validate invoice processing

## Architecture

### Technology Stack
- **Framework**: PySide6 (Qt 6.9.1)
- **PDF Processing**: PyMuPDF (fitz)
- **Email Parsing**: extract_msg
- **Excel Operations**: pandas, openpyxl, XlsxWriter
- **Additional**: PyPDF2 for PDF manipulation

### Project Structure
```
RFP-Utility-Application/
├── main.py                 # Application entry point with QMainWindow
├── pages/                  # Individual page modules for each feature
│   ├── home_page.py       # Landing page/navigation
│   ├── saver_page.py      # Attachment processor (email → organized files)
│   ├── generator_page.py  # RFP form generation from CSV
│   ├── merger_page.py     # Invoice + RFP merging
│   ├── report_page.py     # Statement report generation
│   ├── audit_page.py      # Audit tools
│   ├── check_page.py      # Invoice checking
│   ├── statistics_page.py # Statistics/analytics
│   └── settings_page.py   # Configuration
├── utils/                  # Shared utilities
│   ├── app_utils.py       # Application-wide helpers
│   ├── csv_utils.py       # CSV processing utilities
│   └── paths_config.py    # Default path configurations
├── resources/             # Icons, templates, sounds
├── Company Info.csv       # Company/vendor information
└── requirements.txt       # Python dependencies
```

## Key Design Patterns

### 1. Page-Based Navigation
- Each feature is a separate `QWidget` subclass in `pages/`
- Main window uses `QStackedWidget` for page switching
- Controller pattern: pages receive a `controller` reference to trigger navigation

### 2. UI Structure
- **Toolbar**: Top-level actions (Back, navigation)
- **Groupboxes**: Logical sections with styled titles
- **Dual-column layouts**: Left for controls/input, Right for preview/output
- **Progress feedback**: Status labels, progress bars (determinate and indeterminate)
- **Live previews**: QScrollArea with dynamically rendered content

### 3. File Processing
- Heavy operations run with visual feedback (`_begin_busy`, `_end_busy`)
- Temporary directories for intermediate processing
- "Recycle Bin" pattern: never destructive deletion, always move to recovery folder
- Path safety: `unique_path_in_dir()` prevents overwrites

## Coding Standards

### Style Guidelines
1. **Naming Conventions**:
   - Classes: `PascalCase` (e.g., `SaverPage`)
   - Methods: `snake_case` (e.g., `load_queue`)
   - Private/internal methods: `_leading_underscore` (e.g., `_show_preview`)
   - Constants: `UPPER_SNAKE_CASE` (e.g., `MIN_IMAGE_SIZE`)

2. **Type Hints**: Use type annotations for clarity
   ```python
   def load_queue(self) -> None:
   def discover_sites(root: str) -> List[str]:
   ```

3. **Docstrings**: Provide clear, concise docstrings for non-trivial functions
   ```python
   def extract_pdf_attachments_from_msg(msg_path: str, temp_dir: str) -> List[str]:
       """
       Extract PDF/image attachments from a .msg into temp_dir and return paths.
       We allow large images too (convertible elsewhere if needed by caller).
       """
   ```

4. **Error Handling**: 
   - Use try-finally for resource cleanup
   - User-facing errors via `QMessageBox`
   - Internal logging via `_log()` method
   - Silent degradation for non-critical features

### Qt/PySide6 Best Practices

1. **Layout Management**: Use proper layout hierarchy
   ```python
   layout = QVBoxLayout(self)
   layout.setContentsMargins(10, 10, 10, 10)
   layout.setSpacing(10)
   ```

2. **Signal/Slot Connections**: Connect UI events properly
   ```python
   button.clicked.connect(self.method_name)
   combo.currentTextChanged.connect(self._on_change)
   ```

3. **Widget Lifecycle**:
   - Implement `showEvent` for page initialization
   - Clean up resources in appropriate methods
   - Use `QTimer.singleShot()` for delayed operations

4. **Cursor Management**: Show wait cursor during long operations
   ```python
   QGuiApplication.setOverrideCursor(Qt.WaitCursor)
   # ... operation ...
   QGuiApplication.restoreOverrideCursor()
   ```

5. **Process Events**: Keep UI responsive
   ```python
   QGuiApplication.processEvents()
   ```

### File Path Handling

1. **Always use absolute paths** for file operations
2. **Use `os.path.join()`** for cross-platform compatibility
3. **Validate paths** before operations: `os.path.exists()`, `os.path.isdir()`
4. **Create directories safely**: `os.makedirs(path, exist_ok=True)`

### Preview Rendering (Critical)

When rendering multi-page PDFs or multiple images in a `QScrollArea`:

1. **Clear previous content** before adding new widgets
   ```python
   def _clear_preview(self) -> None:
       for i in reversed(range(self.preview_layout.count())):
           w = self.preview_layout.itemAt(i).widget()
           if w:
               w.setParent(None)
   ```

2. **Set proper sizing for each page/image**:
   - Each QLabel must have proper size hints
   - Use `setScaledContents(False)` to prevent distortion
   - Add spacing between pages: `layout.setSpacing(10)`
   - Ensure labels don't have overlapping geometry

3. **Use proper alignment and constraints**:
   ```python
   label = QLabel()
   label.setAlignment(Qt.AlignCenter)
   label.setPixmap(pixmap)
   label.setMinimumSize(1, 1)  # Allow proper shrinking
   ```

4. **Process PyMuPDF pages correctly**:
   ```python
   for page_num, page in enumerate(doc):
       pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
       # Convert to QImage properly
       # Create QLabel with proper sizing
       # Add to layout
   ```

## Common Patterns in This Codebase

### 1. Busy State Management
```python
def long_operation(self):
    try:
        self._begin_busy("Processing...")
        # ... work ...
        self._update_busy("Almost done...")
        # ... more work ...
    finally:
        self._end_busy()
```

### 2. User Feedback
```python
self._log("✅ Success message")
self._log("⚠️ Warning message")
self._log("❌ Error message")
self._show_toast("Friendly notification")
self._play_notify_sound()
```

### 3. File Discovery
```python
def discover_items(root: str) -> List[str]:
    if not os.path.isdir(root):
        return []
    return sorted(d for d in os.listdir(root) 
                  if os.path.isdir(os.path.join(root, d)))
```

### 4. Safe File Operations
```python
def unique_path_in_dir(target_dir: str, filename: str) -> str:
    """Return a non-colliding path by appending (n) if needed."""
    base, ext = os.path.splitext(filename)
    candidate = os.path.join(target_dir, filename)
    n = 1
    while os.path.exists(candidate):
        candidate = os.path.join(target_dir, f"{base} ({n}){ext}")
        n += 1
    return candidate
```

## Testing & Debugging

### Local Testing
- Test with real `.msg` files from `~/Downloads` or test folders
- Verify folder paths match the organization's structure
- Check all callbacks and signal connections
- Test with both single-page and multi-page PDFs
- Validate CSV parsing with real data

### Error Scenarios to Handle
1. Missing/invalid file paths
2. Corrupted PDF/image files
3. Empty attachments
4. Duplicate invoice numbers
5. Missing CSV columns
6. Permission errors on network drives

### UI Testing Checklist
- [ ] All buttons respond correctly
- [ ] Progress bars update smoothly
- [ ] Preview renders all pages without overlap
- [ ] Log window scrolls to bottom automatically
- [ ] Toast notifications appear and disappear
- [ ] Keyboard shortcuts work (Escape to back, etc.)

## Integration Points

### Default Paths (from `paths_config.py`)
- Invoices: `M:\Staffing Services Agency Intake\Invoices`
- RFPs: `M:\Staffing Services Agency Intake\RFPs`
- Merged: `M:\Staffing Services Agency Intake\Merged Invoices and RFPs`
- Recycle Bin: `M:\Staffing Services Agency Intake\Recycle Bin`

### External Dependencies
- **Outlook**: Opens `.msg` files with `QDesktopServices.openUrl()`
- **Windows Shell**: Uses `SetCurrentProcessExplicitAppUserModelID` for taskbar icon
- **File System**: Relies on network drive access (M:\ drive)

## Known Issues & Solutions

### Issue: PDF Pages Overlaying in Preview
**Problem**: When rendering multi-page PDFs, page 2+ may overlay page 1.

**Root Cause**: QLabel widgets in QVBoxLayout not properly sized or positioned.

**Solution**: 
- Ensure each QLabel has proper minimum height based on pixmap dimensions
- Add explicit spacing between labels
- Set proper size policies for labels
- Clear pixmap references properly to avoid memory issues

### Issue: Memory Usage with Large PDFs
**Solution**: 
- Render at appropriate DPI (Matrix(2, 2) = 144 DPI)
- Close PyMuPDF documents explicitly in finally blocks
- Clear preview widgets before loading new content

## Future Enhancement Guidelines

When adding new features:

1. **Create a new page module** in `pages/` following the pattern
2. **Register in main.py** with the QStackedWidget
3. **Add navigation** from home page or toolbar
4. **Follow the busy/progress feedback pattern**
5. **Use groupboxes** for logical UI sections
6. **Implement `showEvent`** for page-specific initialization
7. **Add Escape key handler** for quick exit
8. **Log all significant operations** for debugging
9. **Use toast notifications** for user celebration moments
10. **Test with real data** before deployment

## Contact & Support

For questions or issues:
- **Developer**: Nathan Nicholas D Souza
- **Email**: nathan.dsouza@fraserhealth.ca
- **Team**: QI Team - Staffing Services, Fraser Health

---

**Last Updated**: January 27, 2026
