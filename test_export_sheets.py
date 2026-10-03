# -*- coding: utf-8 -*-
import os
import win32com.client
import fitz

# Let's test exporting the current AWS_Hybrid_Architecture_Template.xlsx to PDF sheet by sheet
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

try:
    wb_path = os.path.abspath("AWS_Hybrid_Architecture_Template.xlsx")
    wb = excel.Workbooks.Open(wb_path)
    
    # Export each worksheet to a separate PDF
    for i in range(1, wb.Worksheets.Count + 1):
        ws = wb.Worksheets(i)
        pdf_out = os.path.abspath(f"test_sheet_{i}.pdf")
        if os.path.exists(pdf_out):
            os.remove(pdf_out)
        try:
            ws.ExportAsFixedFormat(0, pdf_out) # 0 = xlTypePDF
            print(f"Sheet {i} ({ws.Name}) exported to {pdf_out}")
            
            # Check with fitz
            doc = fitz.open(pdf_out)
            print(f"  Pages: {len(doc)}")
            pix = doc[0].get_pixmap(dpi=150)
            pix.save(f"test_sheet_{i}_p1.png")
            print(f"  Rendered page 1 to test_sheet_{i}_p1.png")
        except Exception as e:
            print(f"  Failed to export Sheet {i}: {e}")
            
    wb.Close(False)
finally:
    excel.Quit()
