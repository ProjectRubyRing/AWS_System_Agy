# -*- coding: utf-8 -*-
import os
import win32com.client

excel = win32com.client.DispatchEx("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

wb_path = os.path.abspath("AWS_Hybrid_Architecture_Template.xlsx")
wb = excel.Workbooks.Open(wb_path, ReadOnly=True)

for s_idx in range(1, wb.Sheets.Count + 1):
    ws = wb.Sheets(s_idx)
    print(f"\n--- Sheet {s_idx}: {ws.Name} (Shapes count: {ws.Shapes.Count}) ---")
    pictures = 0
    rectangles = 0
    connectors = 0
    others = 0
    for shp_idx in range(1, min(ws.Shapes.Count + 1, 20)): # sample first 20
        shp = ws.Shapes(shp_idx)
        print(f"  Shape {shp_idx}: Name='{shp.Name}', Type={shp.Type}, Left={shp.Left:.1f}, Top={shp.Top:.1f}, Width={shp.Width:.1f}, Height={shp.Height:.1f}")
    
    for shp_idx in range(1, ws.Shapes.Count + 1):
        t = ws.Shapes(shp_idx).Type
        if t == 13: # msoPicture
            pictures += 1
        elif t == 1: # msoAutoShape
            rectangles += 1
        elif t == 9: # msoLine
            connectors += 1
        else:
            others += 1
    print(f"  Summary: Pictures={pictures}, AutoShapes={rectangles}, Lines={connectors}, Others={others}")

wb.Close(False)
excel.Quit()
