# -*- coding: utf-8 -*-
import os
import zipfile
import shutil
import win32com.client

# Test creating a shape with guide-based connection sites in a simple workbook
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

try:
    wb = excel.Workbooks.Add()
    ws = wb.Worksheets(1)
    # Add a dummy shape so Excel creates drawing1.xml
    ws.Shapes.AddShape(1, 10, 10, 50, 50)
    
    test_xlsx = os.path.abspath("test_cxn_guide.xlsx")
    if os.path.exists(test_xlsx):
        os.remove(test_xlsx)
    wb.SaveAs(test_xlsx)
    wb.Close(False)
    
    # Now unzip and inject custGeom with guide-based connection sites
    temp_dir = "temp_test_guide"
    if os.path.exists(temp_dir):
        shutil.rmtree(temp_dir)
    with zipfile.ZipFile(test_xlsx, "r") as z:
        z.extractall(temp_dir)
        
    drawing_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <xdr:twoCellAnchor editAs="oneCell">
    <xdr:from><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>8</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>8</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:sp macro="" textlink="">
      <xdr:nvSpPr>
        <xdr:cNvPr id="101" name="TestShape"/>
        <xdr:cNvSpPr/>
      </xdr:nvSpPr>
      <xdr:spPr>
        <a:xfrm><a:off x="0" y="0"/><a:ext cx="2000000" cy="1000000"/></a:xfrm>
        <a:custGeom>
          <a:avLst/>
          <a:gdLst>
            <a:gd name="x1" fmla="*/ w 1 5"/>
            <a:gd name="x2" fmla="*/ w 2 5"/>
            <a:gd name="x3" fmla="*/ w 3 5"/>
            <a:gd name="x4" fmla="*/ w 4 5"/>
            <a:gd name="y1" fmla="*/ h 1 5"/>
            <a:gd name="y2" fmla="*/ h 2 5"/>
            <a:gd name="y3" fmla="*/ h 3 5"/>
            <a:gd name="y4" fmla="*/ h 4 5"/>
          </a:gdLst>
          <a:ahLst/>
          <a:cxnLst>
            <!-- Top side: 4 sites -->
            <a:cxn ang="16200000"><a:pos x="x1" y="t"/></a:cxn>
            <a:cxn ang="16200000"><a:pos x="x2" y="t"/></a:cxn>
            <a:cxn ang="16200000"><a:pos x="x3" y="t"/></a:cxn>
            <a:cxn ang="16200000"><a:pos x="x4" y="t"/></a:cxn>
            <!-- Right side: 4 sites -->
            <a:cxn ang="0"><a:pos x="r" y="y1"/></a:cxn>
            <a:cxn ang="0"><a:pos x="r" y="y2"/></a:cxn>
            <a:cxn ang="0"><a:pos x="r" y="y3"/></a:cxn>
            <a:cxn ang="0"><a:pos x="r" y="y4"/></a:cxn>
            <!-- Bottom side: 4 sites -->
            <a:cxn ang="5400000"><a:pos x="x4" y="b"/></a:cxn>
            <a:cxn ang="5400000"><a:pos x="x3" y="b"/></a:cxn>
            <a:cxn ang="5400000"><a:pos x="x2" y="b"/></a:cxn>
            <a:cxn ang="5400000"><a:pos x="x1" y="b"/></a:cxn>
            <!-- Left side: 4 sites -->
            <a:cxn ang="10800000"><a:pos x="l" y="y4"/></a:cxn>
            <a:cxn ang="10800000"><a:pos x="l" y="y3"/></a:cxn>
            <a:cxn ang="10800000"><a:pos x="l" y="y2"/></a:cxn>
            <a:cxn ang="10800000"><a:pos x="l" y="y1"/></a:cxn>
          </a:cxnLst>
          <a:rect l="l" t="t" r="r" b="b"/>
          <a:pathLst>
            <a:path w="2000000" h="1000000">
              <a:moveTo><a:pt x="0" y="0"/></a:moveTo>
              <a:lnTo><a:pt x="2000000" y="0"/></a:lnTo>
              <a:lnTo><a:pt x="2000000" y="1000000"/></a:lnTo>
              <a:lnTo><a:pt x="0" y="1000000"/></a:lnTo>
              <a:close/>
            </a:path>
          </a:pathLst>
        </a:custGeom>
        <a:solidFill><a:srgbClr val="F8FAFC"/></a:solidFill>
        <a:ln w="22225"><a:solidFill><a:srgbClr val="1A202C"/></a:solidFill></a:ln>
      </xdr:spPr>
    </xdr:sp>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>'''
    
    with open(os.path.join(temp_dir, "xl", "drawings", "drawing1.xml"), "w", encoding="utf-8") as f:
        f.write(drawing_xml)
        
    os.remove(test_xlsx)
    with zipfile.ZipFile(test_xlsx, "w", zipfile.ZIP_DEFLATED) as z_out:
        for root, dirs, files in os.walk(temp_dir):
            for file in files:
                full_p = os.path.join(root, file)
                rel_p = os.path.relpath(full_p, temp_dir)
                z_out.write(full_p, rel_p)
                
    # Now open with Excel COM, connect lines, and verify positions!
    wb2 = excel.Workbooks.Open(test_xlsx)
    ws2 = wb2.Worksheets(1)
    shp = ws2.Shapes("TestShape")
    print(f"TestShape ConnectionSiteCount: {shp.ConnectionSiteCount}")
    print(f"Shape: Left={shp.Left}, Top={shp.Top}, Width={shp.Width}, Height={shp.Height}")
    
    # Connect 16 lines and check their coordinates
    for i in range(1, 17):
        cxn = ws2.Shapes.AddConnector(1, shp.Left, shp.Top, shp.Left + 50, shp.Top + 50)
        cxn.ConnectorFormat.BeginConnect(shp, i)
        print(f"Site {i:2d}: BeginConnected={cxn.ConnectorFormat.BeginConnected}, Left={cxn.Left:.1f}, Top={cxn.Top:.1f}, Width={cxn.Width:.1f}, Height={cxn.Height:.1f}")
        
    wb2.SaveAs(os.path.abspath("test_guide_connected.xlsx"))
    wb2.Close(False)
    print("Test completed successfully!")
finally:
    excel.Quit()
