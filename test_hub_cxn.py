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
    print(f"Hub name: {hub.Name}")
    print(f"Hub ConnectionSiteCount: {hub.ConnectionSiteCount}")
    print(f"Hub Left: {hub.Left}, Top: {hub.Top}, Width: {hub.Width}, Height: {hub.Height}")
    
    # Check the lines attached to HubDemo
    for i in range(1, 17):
        try:
            line = ws3.Shapes(f"HubDemo_Line_{i}")
            cf = line.ConnectorFormat
            print(f"Line {i}: Connected={cf.BeginConnected}, Site={cf.BeginConnectionSite}")
        except Exception as e:
            print(f"Line {i} error: {e}")
    wb.Close(False)
finally:
    excel.Quit()
