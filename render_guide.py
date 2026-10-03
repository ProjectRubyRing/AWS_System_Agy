# -*- coding: utf-8 -*-
import os
import win32com.client
import fitz

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

try:
    wb = excel.Workbooks.Open(os.path.abspath("test_guide_connected.xlsx"))
    pdf = os.path.abspath("test_guide.pdf")
    if os.path.exists(pdf):
        os.remove(pdf)
    wb.Worksheets(1).ExportAsFixedFormat(0, pdf)
    wb.Close(False)
    
    doc = fitz.open(pdf)
    pix = doc[0].get_pixmap(dpi=150)
    pix.save("test_guide.png")
    print("Exported test_guide.png successfully!")
finally:
    excel.Quit()
