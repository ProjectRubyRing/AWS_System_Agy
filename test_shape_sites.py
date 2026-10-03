# -*- coding: utf-8 -*-
import os
import win32com.client

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

try:
    wb = excel.Workbooks.Add()
    ws = wb.Worksheets(1)
    
    # 1. Add a standard rectangle
    rect = ws.Shapes.AddShape(1, 100, 100, 200, 100) # 1 = msoShapeRectangle
    print(f"Standard rectangle ConnectionSiteCount: {rect.ConnectionSiteCount}")
    
    # 2. Add rounded rectangle
    rrect = ws.Shapes.AddShape(5, 100, 250, 200, 100) # 5 = msoShapeRoundedRectangle
    print(f"Standard rounded rectangle ConnectionSiteCount: {rrect.ConnectionSiteCount}")
    
    wb.Close(False)
finally:
    excel.Quit()
