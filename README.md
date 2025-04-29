# EXCEL-AUTOMIZATION
I created this at work in order to get rid of a teddious task of opening and looking through multiple different pdfs at one time.

# Monthly Surcharge PDF Matching

This notebook extracts and matches PDF content against descriptions from a monthly surcharge Excel sheet.

## How to Use

1. Upload the monthly Excel file (e.g., `June_2025_Surcharges.xlsx`).
2. Upload PDFs into a folder (e.g., `June_PDFs`).
3. Update variables in the notebook:
   - `pdf_folder`(make a folder and put the PDFs in it)
   - `excel_path` (make sure the excel is also in jupyter)
   - `output_excel_path`(change to June for example for next month)
   - `sheet_name` (Tab on the excel e.g., June 2025
   - `desc_col` (stays the same)
4. Run the notebook.
5. Check `Updated_<month>_Surcharges.xlsx` for output.

## Dependencies

- pandas
- PyMuPDF (`pip install pymupdf`)
