# -*- coding: utf-8 -*-
import os
import win32com.client

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

try:
    wb = excel.Workbooks.Open(os.path.abspath("AWS_Hybrid_Architecture_Template.xlsx"))
    print(f"Workbook: {wb.Name}")
    for ws in wb.Worksheets:
        print(f"\nSheet: {ws.Name} (Shapes count: {ws.Shapes.Count})")
        # Check first 5 shapes
        for i in range(1, min(6, ws.Shapes.Count + 1)):
            shp = ws.Shapes(i)
            print(f"  Shape {i}: Name='{shp.Name}', Type={shp.Type}")
    wb.Close(False)
finally:
    excel.Quit()
