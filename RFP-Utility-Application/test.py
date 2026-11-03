try:
    import pymupdf as fitz  # PyMuPDF (preferred; avoids third-party fitz package collision)
except Exception:
    import fitz  # PyMuPDF fallback

def read_pdf_fields(pdf_path):
    """
    Read and display all form fields and their values from a PDF.
    
    Args:
        pdf_path (str): Path to the PDF file
    """
    print(f"\n--- PDF Field Reader ---")
    print(f"Reading: {pdf_path}\n")
    
    try:
        doc = fitz.open(pdf_path)
        print(f"✅ PDF opened successfully\n")
        
        fields_found = []
        
        for page_num, page in enumerate(doc, start=1):
            widgets = page.widgets() or []
            
            if widgets:
                print(f"[Page {page_num}]")
                
                for widget in widgets:
                    field_name = widget.field_name
                    if field_name:
                        field_value = widget.field_value
                        fields_found.append((field_name, field_value))
                        print(f"  Field: {field_name:<35} | Value: {field_value}")
                
                print()
        
        if not fields_found:
            print("⚠️  No form fields found in this PDF.")
        else:
            print(f"\n--- Summary ---")
            print(f"Total fields found: {len(fields_found)}")
            print(f"\nFields with due dates (for pinpointing the issue):")
            for name, value in fields_found:
                if name and value and "/" in str(value):
                    print(f"  {name:<35} = {value}")
        
        doc.close()
        
    except Exception as e:
        print(f"❌ ERROR: {e}")


if __name__ == "__main__":
    # Uncomment and modify one of these as needed:
    
    # Option 1: Check Excel file (original behavior)
    # FILE_TO_CHECK = r"M:\Staffing Services Agency Intake\Testing\2025-2026FY Invoices Tracker.xlsx"
    # CLEANED_FILE_OUTPUT = r"M:\Staffing Services Agency Intake\Testing\2025-2026FY Invoices Tracker reformatted.xlsx"
    # original_dfs = check_excel_file(FILE_TO_CHECK)
    
    # Option 2: Read PDF fields
    # Modify this path to point to one of your generated RFP PDFs
    pdf_to_read = input("Enter the path to the RFP PDF file: ").strip()
    if pdf_to_read:
        read_pdf_fields(pdf_to_read)
    else:
        print("No PDF path provided.")


