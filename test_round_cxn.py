# -*- coding: utf-8 -*-
import os
import zipfile
import shutil
import win32com.client
import fitz

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

try:
    wb = excel.Workbooks.Add()
    ws = wb.Worksheets(1)
    ws.Shapes.AddShape(1, 10, 10, 50, 50)
    
    test_xlsx = os.path.abspath("test_round_cxn.xlsx")
    if os.path.exists(test_xlsx):
        os.remove(test_xlsx)
    wb.SaveAs(test_xlsx)
    wb.Close(False)
    
    temp_dir = "temp_test_round"
    if os.path.exists(temp_dir):
        shutil.rmtree(temp_dir)
    with zipfile.ZipFile(test_xlsx, "r") as z:
        z.extractall(temp_dir)
        
    # Test rounded rect custGeom with guide formulas
    drawing_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <xdr:twoCellAnchor editAs="oneCell">
    <xdr:from><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>8</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>7</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:sp macro="" textlink="">
      <xdr:nvSpPr>
        <xdr:cNvPr id="101" name="RoundNode"/>
        <xdr:cNvSpPr/>
      </xdr:nvSpPr>
      <xdr:spPr>
        <a:xfrm><a:off x="0" y="0"/><a:ext cx="2000000" cy="1000000"/></a:xfrm>
        <a:custGeom>
          <a:avLst/>
          <a:gdLst>
            <a:gd name="rad" fmla="*/ 8000 w 100000"/>
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
            <a:cxn ang="16200000"><a:pos x="x1" y="t"/></a:cxn>
            <a:cxn ang="16200000"><a:pos x="x2" y="t"/></a:cxn>
            <a:cxn ang="16200000"><a:pos x="x3" y="t"/></a:cxn>
            <a:cxn ang="16200000"><a:pos x="x4" y="t"/></a:cxn>
            <a:cxn ang="0"><a:pos x="r" y="y1"/></a:cxn>
            <a:cxn ang="0"><a:pos x="r" y="y2"/></a:cxn>
            <a:cxn ang="0"><a:pos x="r" y="y3"/></a:cxn>
            <a:cxn ang="0"><a:pos x="r" y="y4"/></a:cxn>
            <a:cxn ang="5400000"><a:pos x="x4" y="b"/></a:cxn>
            <a:cxn ang="5400000"><a:pos x="x3" y="b"/></a:cxn>
            <a:cxn ang="5400000"><a:pos x="x2" y="b"/></a:cxn>
            <a:cxn ang="5400000"><a:pos x="x1" y="b"/></a:cxn>
            <a:cxn ang="10800000"><a:pos x="l" y="y4"/></a:cxn>
            <a:cxn ang="10800000"><a:pos x="l" y="y3"/></a:cxn>
            <a:cxn ang="10800000"><a:pos x="l" y="y2"/></a:cxn>
            <a:cxn ang="10800000"><a:pos x="l" y="y1"/></a:cxn>
          </a:cxnLst>
          <a:rect l="l" t="t" r="r" b="b"/>
          <a:pathLst>
            <a:path w="100000" h="100000">
              <a:moveTo><a:pt x="8000" y="0"/></a:moveTo>
              <a:lnTo><a:pt x="92000" y="0"/></a:lnTo>
              <a:arcTo wR="8000" hR="8000" stAng="16200000" swAng="5400000"/>
              <a:lnTo><a:pt x="100000" y="92000"/></a:lnTo>
              <a:arcTo wR="8000" hR="8000" stAng="0" swAng="5400000"/>
              <a:lnTo><a:pt x="8000" y="100000"/></a:lnTo>
              <a:arcTo wR="8000" hR="8000" stAng="5400000" swAng="5400000"/>
              <a:lnTo><a:pt x="0" y="8000"/></a:lnTo>
              <a:arcTo wR="8000" hR="8000" stAng="10800000" swAng="5400000"/>
              <a:close/>
            </a:path>
          </a:pathLst>
        </a:custGeom>
        <a:solidFill><a:srgbClr val="EFF6FF"/></a:solidFill>
        <a:ln w="22225"><a:solidFill><a:srgbClr val="1E40AF"/></a:solidFill></a:ln>
      </xdr:spPr>
      <xdr:txBody>
        <a:bodyPr vert="horz" lIns="45720" tIns="45720" rIns="45720" bIns="45720" anchor="ctr"/>
        <a:lstStyle/>
        <a:p>
          <a:pPr algn="ctr"/>
          <a:r>
            <a:rPr lang="ja-JP" sz="1000" b="1">
              <a:solidFill><a:srgbClr val="1E40AF"/></a:solidFill>
              <a:latin typeface="Meiryo UI"/>
              <a:ea typeface="Meiryo UI"/>
            </a:rPr>
            <a:t>Cloud Blue ノード</a:t>
          </a:r>
        </a:p>
      </xdr:txBody>
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
                
    wb2 = excel.Workbooks.Open(test_xlsx)
    ws2 = wb2.Worksheets(1)
    shp = ws2.Shapes("RoundNode")
    print(f"RoundNode ConnectionSiteCount: {shp.ConnectionSiteCount}")
    
    # Connect 16 outward lines
    hub_left = shp.Left
    hub_top = shp.Top
    hub_w = shp.Width
    hub_h = shp.Height
    
    for site_idx in range(1, 17):
        if 1 <= site_idx <= 4:
            # Top -> line goes up
            fx = hub_left + hub_w * (0.2 * site_idx)
            fy = hub_top - 50
        elif 5 <= site_idx <= 8:
            # Right -> line goes right
            fx = hub_left + hub_w + 50
            fy = hub_top + hub_h * (0.2 * (site_idx - 4))
        elif 9 <= site_idx <= 12:
            # Bottom -> line goes down
            fx = hub_left + hub_w * (1.0 - 0.2 * (site_idx - 8))
            fy = hub_top + hub_h + 50
        else:
            # Left -> line goes left
            fx = hub_left - 50
            fy = hub_top + hub_h * (1.0 - 0.2 * (site_idx - 12))
            
        cxn = ws2.Shapes.AddConnector(1, fx, fy, fx + 10, fy + 10)
        cxn.ConnectorFormat.BeginConnect(shp, site_idx)
        cxn.ConnectorFormat.EndDisconnect()
        # Set End point to (fx, fy)
        # Note: In Excel COM, to place the end point outward:
        # Actually AddConnector has (Type, BeginX, BeginY, EndX, EndY).
        # When we BeginConnect(shp, site_idx), the Begin point moves to the connection site!
        # The End point remains at (EndX, EndY)!
        # So setting AddConnector(1, fx, fy, fx, fy) means End point is at (fx, fy)!
        cxn.Line.EndArrowheadStyle = 2
        
    pdf = os.path.abspath("test_round_connected.pdf")
    if os.path.exists(pdf):
        os.remove(pdf)
    ws2.ExportAsFixedFormat(0, pdf)
    wb2.Close(False)
    
    doc = fitz.open(pdf)
    pix = doc[0].get_pixmap(dpi=150)
    pix.save("test_round_connected.png")
    print("Exported test_round_connected.png successfully!")
finally:
    excel.Quit()
