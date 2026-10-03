# -*- coding: utf-8 -*-
import os
import win32com.client

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

try:
    wb = excel.Workbooks.Open(os.path.abspath("AWS_Hybrid_Architecture_Template.xlsx"))
    ws3 = wb.Worksheets('03_パーツ集_作図図形・コネクタ線')
    hub = ws3.Shapes('Hub_16pt_Demo')
    print(f"Hub: Left={hub.Left}, Top={hub.Top}, Width={hub.Width}, Height={hub.Height}")
    for i in range(1, 17):
        line = ws3.Shapes(f"HubDemo_Line_{i}")
        # Note: Connector line doesn't always have BeginX directly, but Left, Top, Width, Height
        print(f"Line {i}: Left={line.Left:.1f}, Top={line.Top:.1f}, Width={line.Width:.1f}, Height={line.Height:.1f}")
    wb.Close(False)
finally:
    excel.Quit()
