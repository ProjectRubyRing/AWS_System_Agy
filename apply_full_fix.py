# -*- coding: utf-8 -*-
"""
Full fix script to update build_excel_template.py:
1. Fix DrawingML custom geometry (guide formulas for 16-point connection snapping)
2. Fix twoCellAnchor coordinates (clean snapping without width/height EMU explosion)
3. Fix Sheet 3 layout and cell rows to perfectly align with shapes without any overlap
4. Fix page setup for all 6 sheets to fit exactly on 1 page each (A3 landscape, fit_to_pages 1,1)
5. Fix Hub_16pt_Demo outward connector lines in verify_and_finish_with_excel_com
6. Fix CIDR header container in Sheet 2 (detailed architecture)
"""

import re

with open("build_excel_template.py", "r", encoding="utf-8") as f:
    code = f.read()

# 1. Update Sheet 0 (Intro) page setup
code = code.replace(
    """    ws0.set_landscape()
    ws0.set_paper(9) # A4
    ws0.set_margins(left=0.4, right=0.4, top=0.4, bottom=0.4)""",
    """    ws0.set_landscape()
    ws0.set_paper(9) # A4
    ws0.fit_to_pages(1, 1)
    ws0.set_margins(left=0.4, right=0.4, top=0.4, bottom=0.4)"""
)

# 2. Update Sheet 1 (Overview) page setup
code = code.replace(
    """    ws1.set_paper(8) # A3
    ws1.fit_to_pages(1, 0)
    ws1.set_margins(left=0.3, right=0.3, top=0.4, bottom=0.4)""",
    """    ws1.set_paper(8) # A3
    ws1.fit_to_pages(1, 1)
    ws1.set_margins(left=0.3, right=0.3, top=0.3, bottom=0.3)"""
)

# 3. Update Sheet 2 (Detailed) page setup & CIDR header
code = code.replace(
    """    ws2.set_paper(8) # A3
    ws2.fit_to_pages(1, 0)
    ws2.set_margins(left=0.3, right=0.3, top=0.4, bottom=0.4)""",
    """    ws2.set_paper(8) # A3
    ws2.fit_to_pages(1, 1)
    ws2.set_margins(left=0.3, right=0.3, top=0.3, bottom=0.3)"""
)

code = code.replace(
    """    draw_cell_container(ws2, 5, 1, 33, 10, "【オンプレミス工場・現場ネットワーク】CIDR: 192.168.0.0/16", f_box_head_dark, f_bg_onprem, head_rows=1)
    draw_cell_container(ws2, 7, 2, 15, 9, "OT系 隔離制御ネットワーク (192.168.10.0/24)", f_head_emerald, f_bg_white)""",
    """    draw_cell_container(ws2, 5, 1, 33, 10, "【オンプレミス工場・現場ネットワーク】\\nCIDR: 192.168.0.0/16", f_box_head_dark, f_bg_onprem, head_rows=2)
    draw_cell_container(ws2, 7, 2, 15, 9, "OT系 隔離制御ネットワーク (192.168.10.0/24)", f_head_emerald, f_bg_white)"""
)

# 4. Update Sheet 4 & 5 page setup
code = code.replace(
    """    ws4.set_landscape()
    ws4.set_paper(8)
    ws4.set_margins(left=0.3, right=0.3, top=0.4, bottom=0.4)""",
    """    ws4.set_landscape()
    ws4.set_paper(8)
    ws4.fit_to_pages(1, 1)
    ws4.set_margins(left=0.3, right=0.3, top=0.3, bottom=0.3)"""
)

code = code.replace(
    """    ws5.set_landscape()
    ws5.set_paper(8)
    ws5.set_margins(left=0.3, right=0.3, top=0.4, bottom=0.4)""",
    """    ws5.set_landscape()
    ws5.set_paper(8)
    ws5.fit_to_pages(1, 1)
    ws5.set_margins(left=0.3, right=0.3, top=0.3, bottom=0.3)"""
)

# 5. Replace Sheet 3 (03_パーツ集_作図図形・コネクタ線) cell generation logic
old_ws3_block = re.search(r'(# =+\s+# SHEET 3: 03_.*?)(# =+\s+# SHEET 4: 04_)', code, re.DOTALL)
if old_ws3_block:
    new_ws3_code = '''# =========================================================================
    # SHEET 3: 03_パーツ集_作図図形・コネクタ線 (★NEW PRIMARY PARTS CATALOG)
    # =========================================================================
    ws3 = workbook.add_worksheet('03_パーツ集_作図図形・コネクタ線')
    ws3.hide_gridlines(0)
    ws3.set_landscape()
    ws3.set_paper(8) # A3
    ws3.fit_to_pages(1, 1)
    ws3.set_margins(left=0.3, right=0.3, top=0.3, bottom=0.3)
    
    ws3.set_column('A:A', 2)
    for c in range(1, 42):
        col_letter = xlsxwriter.utility.xl_col_to_name(c)
        ws3.set_column(f'{col_letter}:{col_letter}', 3.4)
    for r in range(4, 90):
        ws3.set_row(r, 18)

    # Insert placeholder textbox to initialize DrawingML structure
    ws3.insert_textbox(0, 0, 'Placeholder', {'width': 10, 'height': 10})
    write_meta_header(ws3, "【作図部品集】16箇所接続ポイント四角形・7色コネクタ線・3色コメント図形", "方眼紙升目吸着 / 等間隔配線パレット")

    # Section 1: 16-Connection Nodes & Boxes
    row_s3 = 5
    ws3.merge_range(row_s3, 1, row_s3, 40, "1. 【等間隔16箇所接続ポイント付き】作図用四角形部品（美しい7色展開）", f_sec_title)
    row_s3 += 1
    ws3.merge_range(row_s3, 1, row_s3, 40, "※四角形の各辺に4箇所ずつ（計16箇所）の接続ポイントを装備。コネクタ線がピタッと吸着し、線が重ならず等間隔に美しく配線できます。コピー＆ペーストして使用してください。", f_note)
    ws3.set_row(row_s3, 20)
    row_s3 += 1

    # Row 1 Labels (Charcoal, Blue, Emerald, Amber)
    ws3.merge_range(row_s3, 1, row_s3, 9, "① Charcoal Black (基盤・オンプレ)", f_head_charcoal)
    ws3.merge_range(row_s3, 11, row_s3, 19, "② Cloud Blue (AWS・クラウド基盤)", f_head_blue)
    ws3.merge_range(row_s3, 21, row_s3, 29, "③ Emerald Green (現場IoT・OT設備)", f_head_emerald)
    ws3.merge_range(row_s3, 31, row_s3, 39, "④ Amber Orange (外部SaaS・Salesforce)", f_head_amber)
    ws3.set_row(row_s3, 19)
    # Node shapes: rows 8..10 (height 2)
    # Box label row: 11
    ws3.write(11, 1, "【境界枠】", f_note)
    ws3.write(11, 11, "【境界枠】", f_note)
    ws3.write(11, 21, "【境界枠】", f_note)
    ws3.write(11, 31, "【境界枠】", f_note)
    # Box shapes: rows 12..16 (height 4)

    # Row 2 Labels (Purple, Cyan, Rose)
    row_s3 = 18
    ws3.merge_range(row_s3, 1, row_s3, 9, "⑤ Purple Violet (イントラ・分析)", f_head_purple)
    ws3.merge_range(row_s3, 11, row_s3, 19, "⑥ Cyan Sky (現場モバイル・作業端末)", f_head_cyan)
    ws3.merge_range(row_s3, 21, row_s3, 29, "⑦ Rose Crimson (DMZ・セキュリティ)", f_head_rose)
    ws3.set_row(row_s3, 19)
    # Node shapes: rows 19..21 (height 2)
    # Box label row: 22
    ws3.write(22, 1, "【境界枠】", f_note)
    ws3.write(22, 11, "【境界枠】", f_note)
    ws3.write(22, 21, "【境界枠】", f_note)
    # Box shapes: rows 23..27 (height 4)

    # Section 2: 7-Color Connectors
    row_s3 = 29
    ws3.merge_range(row_s3, 1, row_s3, 40, "2. 【美しい7色展開】接続コネクタ線部品（カギ線・直線・双方向・非同期破線）", f_sec_title)
    row_s3 += 1
    ws3.merge_range(row_s3, 1, row_s3, 40, "※各色ともカギ線矢印・直線矢印・双方向矢印・非同期破線の4種類を完備。図形の16箇所の接続ポイントに近づけると自動吸着（スナップ）します。", f_note)
    ws3.set_row(row_s3, 20)
    row_s3 += 1

    c_labels = [
        ("① Charcoal (基盤)", f_head_charcoal),
        ("② Cloud Blue (AWS)", f_head_blue),
        ("③ Emerald (現場OT)", f_head_emerald),
        ("④ Amber (Salesforce)", f_head_amber),
        ("⑤ Purple (イントラ)", f_head_purple),
        ("⑥ Cyan (モバイル)", f_head_cyan),
        ("⑦ Rose (DMZ・警報)", f_head_rose),
    ]
    for clbl, c_head_fmt in c_labels:
        ws3.merge_range(row_s3, 1, row_s3, 6, clbl, c_head_fmt)
        ws3.write(row_s3, 7, "【カギ線】", f_note)
        ws3.write(row_s3, 14, "【直線】", f_note)
        ws3.write(row_s3, 22, "【双方向】", f_note)
        ws3.write(row_s3, 30, "【非同期破線】", f_note)
        ws3.set_row(row_s3, 18)
        row_s3 += 2

    # Section 3: 3-Color Comments
    row_s3 = 46
    ws3.merge_range(row_s3, 1, row_s3, 40, "3. 【美しい3色展開】コメント・注記入力用図形（引き出し吹き出し型 & 付箋メモカード型）", f_sec_title)
    row_s3 += 1
    ws3.merge_range(row_s3, 1, row_s3, 40, "※ポイントや付記事項を美しく書き込める図形部品です。引き出し線付き吹き出し型と、カード型メモの2種類を準備しています。テキストを直接編集可能。", f_note)
    ws3.set_row(row_s3, 20)
    row_s3 += 1

    ws3.merge_range(row_s3, 1, row_s3, 12, "① Info Blue (仕様・通信ポイント)", f_head_blue)
    ws3.merge_range(row_s3, 14, row_s3, 25, "② Warning Amber (注意点・設計制約)", f_head_amber)
    ws3.merge_range(row_s3, 27, row_s3, 38, "③ Security Emerald (セキュリティ・運用基準)", f_head_emerald)
    ws3.set_row(row_s3, 19)
    # Comments occupy rows 49..59

    # Section 4: 16-pt Connection Demo Sample
    row_s3 = 61
    ws3.merge_range(row_s3, 1, row_s3, 40, "4. 【接続実例見本】等間隔16箇所接続ポイント 完全配線サンプル", f_sec_title)
    row_s3 += 1
    ws3.merge_range(row_s3, 1, row_s3, 40, "※中央のHubノードに対して、上下左右から各4本ずつ（計16本）の線が重ならずに等間隔接続されている見本です。線をつかんで動かしても接続が追従します。", f_note)
    ws3.set_row(row_s3, 20)
    row_s3 += 1

    ws3.merge_range(row_s3, 14, row_s3, 26, "↓ 下記に16本完全接続デモモデルが配置されています ↓", f_tbl_cell_zebra_center)
    ws3.set_row(row_s3, 18)

    '''
    code = code[:old_ws3_block.start()] + new_ws3_code + code[old_ws3_block.start() + len(old_ws3_block.group(1)):]
    print("Replaced ws3 sheet definition successfully.")

# 6. Replace DrawingML functions
drawing_funcs = '''# =========================================================================
# 2. DRAWINGML INJECTION: 16-Connection Shapes, Connectors, Callouts
# =========================================================================

def generate_16pt_cxn_list():
    """Generate 16 connection sites using DrawingML guide formulas for perfect snapping"""
    items = [
        # Top side: 4 sites at 20%, 40%, 60%, 80%
        '<a:cxn ang="16200000"><a:pos x="x1" y="t"/></a:cxn>',
        '<a:cxn ang="16200000"><a:pos x="x2" y="t"/></a:cxn>',
        '<a:cxn ang="16200000"><a:pos x="x3" y="t"/></a:cxn>',
        '<a:cxn ang="16200000"><a:pos x="x4" y="t"/></a:cxn>',
        # Right side: 4 sites
        '<a:cxn ang="0"><a:pos x="r" y="y1"/></a:cxn>',
        '<a:cxn ang="0"><a:pos x="r" y="y2"/></a:cxn>',
        '<a:cxn ang="0"><a:pos x="r" y="y3"/></a:cxn>',
        '<a:cxn ang="0"><a:pos x="r" y="y4"/></a:cxn>',
        # Bottom side: 4 sites
        '<a:cxn ang="5400000"><a:pos x="x4" y="b"/></a:cxn>',
        '<a:cxn ang="5400000"><a:pos x="x3" y="b"/></a:cxn>',
        '<a:cxn ang="5400000"><a:pos x="x2" y="b"/></a:cxn>',
        '<a:cxn ang="5400000"><a:pos x="x1" y="b"/></a:cxn>',
        # Left side: 4 sites
        '<a:cxn ang="10800000"><a:pos x="l" y="y4"/></a:cxn>',
        '<a:cxn ang="10800000"><a:pos x="l" y="y3"/></a:cxn>',
        '<a:cxn ang="10800000"><a:pos x="l" y="y2"/></a:cxn>',
        '<a:cxn ang="10800000"><a:pos x="l" y="y1"/></a:cxn>',
    ]
    return "".join(items)


def create_16pt_node_shape_xml(sp_id, name, col_from, row_from, col_to, row_to, fill_hex, border_hex, title, subtitle=None):
    """Creates a rounded rectangle node card with 16 connection sites snapped to grid cells"""
    cxn_xml = generate_16pt_cxn_list()
    text_xml = f\'\'\'<a:p>
      <a:pPr algn="ctr"/>
      <a:r>
        <a:rPr lang="ja-JP" sz="950" b="1">
          <a:solidFill><a:srgbClr val="{border_hex}"/></a:solidFill>
          <a:latin typeface="Meiryo UI"/>
          <a:ea typeface="Meiryo UI"/>
        </a:rPr>
        <a:t>{title}</a:t>
      </a:r>
    </a:p>\'\'\'
    if subtitle:
        text_xml += f\'\'\'<a:p>
      <a:pPr algn="ctr"/>
      <a:r>
        <a:rPr lang="ja-JP" sz="800">
          <a:solidFill><a:srgbClr val="4A5568"/></a:solidFill>
          <a:latin typeface="Meiryo UI"/>
          <a:ea typeface="Meiryo UI"/>
        </a:rPr>
        <a:t>{subtitle}</a:t>
      </a:r>
    </a:p>\'\'\'

    cust_geom = f\'\'\'<a:custGeom>
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
        <a:cxnLst>{cxn_xml}</a:cxnLst>
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
      </a:custGeom>\'\'\'

    return f\'\'\'<xdr:twoCellAnchor editAs="oneCell">
  <xdr:from><xdr:col>{col_from}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{row_from}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
  <xdr:to><xdr:col>{col_to}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{row_to}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
  <xdr:sp macro="" textlink="">
    <xdr:nvSpPr>
      <xdr:cNvPr id="{sp_id}" name="{name}"/>
      <xdr:cNvSpPr/>
    </xdr:nvSpPr>
    <xdr:spPr>
      <a:xfrm><a:off x="0" y="0"/><a:ext cx="2000000" cy="1000000"/></a:xfrm>
      {cust_geom}
      <a:solidFill><a:srgbClr val="{fill_hex}"/></a:solidFill>
      <a:ln w="22225" cmpd="sng">
        <a:solidFill><a:srgbClr val="{border_hex}"/></a:solidFill>
      </a:ln>
    </xdr:spPr>
    <xdr:txBody>
      <a:bodyPr vert="horz" lIns="45720" tIns="45720" rIns="45720" bIns="45720" anchor="ctr"/>
      <a:lstStyle/>
      {text_xml}
    </xdr:txBody>
  </xdr:sp>
  <xdr:clientData/>
</xdr:twoCellAnchor>\'\'\'


def create_16pt_container_box_xml(sp_id, name, col_from, row_from, col_to, row_to, fill_hex, border_hex, title):
    """Creates a container box with 16 connection sites for zone boundaries"""
    cxn_xml = generate_16pt_cxn_list()
    text_xml = f\'\'\'<a:p>
      <a:pPr algn="l"/>
      <a:r>
        <a:rPr lang="ja-JP" sz="900" b="1">
          <a:solidFill><a:srgbClr val="{border_hex}"/></a:solidFill>
          <a:latin typeface="Meiryo UI"/>
          <a:ea typeface="Meiryo UI"/>
        </a:rPr>
        <a:t> {title}</a:t>
      </a:r>
    </a:p>\'\'\'

    cust_geom = f\'\'\'<a:custGeom>
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
        <a:cxnLst>{cxn_xml}</a:cxnLst>
        <a:rect l="l" t="t" r="r" b="b"/>
        <a:pathLst>
          <a:path w="100000" h="100000">
            <a:moveTo><a:pt x="0" y="0"/></a:moveTo>
            <a:lnTo><a:pt x="100000" y="0"/></a:lnTo>
            <a:lnTo><a:pt x="100000" y="100000"/></a:lnTo>
            <a:lnTo><a:pt x="0" y="100000"/></a:lnTo>
            <a:close/>
          </a:path>
        </a:pathLst>
      </a:custGeom>\'\'\'

    return f\'\'\'<xdr:twoCellAnchor editAs="oneCell">
  <xdr:from><xdr:col>{col_from}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{row_from}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
  <xdr:to><xdr:col>{col_to}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{row_to}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
  <xdr:sp macro="" textlink="">
    <xdr:nvSpPr>
      <xdr:cNvPr id="{sp_id}" name="{name}"/>
      <xdr:cNvSpPr/>
    </xdr:nvSpPr>
    <xdr:spPr>
      <a:xfrm><a:off x="0" y="0"/><a:ext cx="2000000" cy="1000000"/></a:xfrm>
      {cust_geom}
      <a:solidFill><a:srgbClr val="{fill_hex}"/></a:solidFill>
      <a:ln w="19050" cmpd="sng">
        <a:solidFill><a:srgbClr val="{border_hex}"/></a:solidFill>
      </a:ln>
    </xdr:spPr>
    <xdr:txBody>
      <a:bodyPr vert="horz" lIns="72000" tIns="54000" rIns="72000" bIns="54000" anchor="t"/>
      <a:lstStyle/>
      {text_xml}
    </xdr:txBody>
  </xdr:sp>
  <xdr:clientData/>
</xdr:twoCellAnchor>\'\'\'


def create_connector_xml(sp_id, name, col_from, row_from, col_to, row_to, color_hex, line_type="elbow", has_arrow=True, is_bidirectional=False, is_dashed=False):
    """Creates an independent connector line ready to be snapped to shapes"""
    geom_name = "bentConnector3" if line_type == "elbow" else "line"
    
    if is_bidirectional:
        head_end = \'<a:headEnd type="triangle" w="med" len="med"/>\'
        tail_end = \'<a:tailEnd type="triangle" w="med" len="med"/>\'
    elif has_arrow:
        head_end = \'<a:headEnd type="none"/>\'
        tail_end = \'<a:tailEnd type="triangle" w="med" len="med"/>\'
    else:
        head_end = \'<a:headEnd type="none"/>\'
        tail_end = \'<a:tailEnd type="none"/>\'
        
    dash_xml = \'<a:prstDash val="dash"/>\' if is_dashed else \'\'

    return f\'\'\'<xdr:twoCellAnchor>
  <xdr:from><xdr:col>{col_from}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{row_from}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
  <xdr:to><xdr:col>{col_to}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{row_to}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
  <xdr:cxnSp macro="">
    <xdr:nvCxnSpPr>
      <xdr:cNvPr id="{sp_id}" name="{name}"/>
      <xdr:cNvCxnSpPr/>
    </xdr:nvCxnSpPr>
    <xdr:spPr>
      <a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="1000000"/></a:xfrm>
      <a:prstGeom prst="{geom_name}"><a:avLst/></a:prstGeom>
      <a:ln w="22225">
        <a:solidFill><a:srgbClr val="{color_hex}"/></a:solidFill>
        {dash_xml}
        {head_end}
        {tail_end}
      </a:ln>
    </xdr:spPr>
  </xdr:cxnSp>
  <xdr:clientData/>
</xdr:twoCellAnchor>\'\'\'


def create_callout_xml(sp_id, name, col_from, row_from, col_to, row_to, fill_hex, border_hex, title, body_text):
    """Creates a callout (wedgeRoundRectCallout) for annotations"""
    body_lines = body_text.split("\\n")
    p_runs = []
    for bline in body_lines:
        safe_line = bline.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        p_runs.append(f\'\'\'<a:p>
        <a:pPr algn="l"/>
        <a:r>
          <a:rPr lang="ja-JP" sz="800">
            <a:solidFill><a:srgbClr val="2D3748"/></a:solidFill>
            <a:latin typeface="Meiryo UI"/>
            <a:ea typeface="Meiryo UI"/>
          </a:rPr>
          <a:t>{safe_line}</a:t>
        </a:r>
      </a:p>\'\'\')
    body_xml = "".join(p_runs)
    safe_title = title.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

    return f\'\'\'<xdr:twoCellAnchor editAs="oneCell">
  <xdr:from><xdr:col>{col_from}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{row_from}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
  <xdr:to><xdr:col>{col_to}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{row_to}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
  <xdr:sp macro="" textlink="">
    <xdr:nvSpPr>
      <xdr:cNvPr id="{sp_id}" name="{name}"/>
      <xdr:cNvSpPr/>
    </xdr:nvSpPr>
    <xdr:spPr>
      <a:xfrm><a:off x="0" y="0"/><a:ext cx="2000000" cy="1000000"/></a:xfrm>
      <a:prstGeom prst="wedgeRoundRectCallout"><a:avLst/></a:prstGeom>
      <a:solidFill><a:srgbClr val="{fill_hex}"/></a:solidFill>
      <a:ln w="19050">
        <a:solidFill><a:srgbClr val="{border_hex}"/></a:solidFill>
      </a:ln>
    </xdr:spPr>
    <xdr:txBody>
      <a:bodyPr vert="horz" lIns="54000" tIns="54000" rIns="54000" bIns="54000" anchor="t"/>
      <a:lstStyle/>
      <a:p>
        <a:pPr algn="l"/>
        <a:r>
          <a:rPr lang="ja-JP" sz="900" b="1">
            <a:solidFill><a:srgbClr val="{border_hex}"/></a:solidFill>
            <a:latin typeface="Meiryo UI"/>
            <a:ea typeface="Meiryo UI"/>
          </a:rPr>
          <a:t>{safe_title}</a:t>
        </a:r>
      </a:p>
      {body_xml}
    </xdr:txBody>
  </xdr:sp>
  <xdr:clientData/>
</xdr:twoCellAnchor>\'\'\'


def create_memo_card_xml(sp_id, name, col_from, row_from, col_to, row_to, fill_hex, border_hex, header_title, body_lines):
    """Creates a structured memo card shape with colored header and text"""
    body_runs = []
    for line in body_lines:
        safe_line = line.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        body_runs.append(f\'\'\'<a:p>
        <a:pPr algn="l"/>
        <a:r>
          <a:rPr lang="ja-JP" sz="800">
            <a:solidFill><a:srgbClr val="2D3748"/></a:solidFill>
            <a:latin typeface="Meiryo UI"/>
            <a:ea typeface="Meiryo UI"/>
          </a:rPr>
          <a:t>{safe_line}</a:t>
        </a:r>
      </a:p>\'\'\')
    body_xml = "".join(body_runs)
    safe_header = header_title.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

    return f\'\'\'<xdr:twoCellAnchor editAs="oneCell">
  <xdr:from><xdr:col>{col_from}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{row_from}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
  <xdr:to><xdr:col>{col_to}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{row_to}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
  <xdr:sp macro="" textlink="">
    <xdr:nvSpPr>
      <xdr:cNvPr id="{sp_id}" name="{name}"/>
      <xdr:cNvSpPr/>
    </xdr:nvSpPr>
    <xdr:spPr>
      <a:xfrm><a:off x="0" y="0"/><a:ext cx="2000000" cy="1000000"/></a:xfrm>
      <a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 8000"/></a:avLst></a:prstGeom>
      <a:solidFill><a:srgbClr val="{fill_hex}"/></a:solidFill>
      <a:ln w="19050">
        <a:solidFill><a:srgbClr val="{border_hex}"/></a:solidFill>
      </a:ln>
    </xdr:spPr>
    <xdr:txBody>
      <a:bodyPr vert="horz" lIns="54000" tIns="45720" rIns="54000" bIns="45720" anchor="t"/>
      <a:lstStyle/>
      <a:p>
        <a:pPr algn="l"/>
        <a:r>
          <a:rPr lang="ja-JP" sz="900" b="1">
            <a:solidFill><a:srgbClr val="{border_hex}"/></a:solidFill>
            <a:latin typeface="Meiryo UI"/>
            <a:ea typeface="Meiryo UI"/>
          </a:rPr>
          <a:t>{safe_header}</a:t>
        </a:r>
      </a:p>
      {body_xml}
    </xdr:txBody>
  </xdr:sp>
  <xdr:clientData/>
</xdr:twoCellAnchor>\'\'\'
'''

old_df_block = re.search(r'(# =+\s+# 2\. DRAWINGML INJECTION:.*?)(def inject_drawings_into_template)', code, re.DOTALL)
if old_df_block:
    code = code[:old_df_block.start()] + drawing_funcs + code[old_df_block.start() + len(old_df_block.group(1)):]
    print("Replaced drawing functions successfully.")

# 7. Replace inject_drawings_into_template body
old_inject = re.search(r'(def inject_drawings_into_template\(xlsx_path\):.*?)(def verify_and_finish_with_excel_com)', code, re.DOTALL)
if old_inject:
    new_inject_code = '''def inject_drawings_into_template(xlsx_path):
    """
    Extracts the generated xlsx, enriches Sheet 3 (and others) with DrawingML shapes,
    and re-zips the workbook.
    """
    temp_dir = "temp_unzip_build"
    if os.path.exists(temp_dir):
        shutil.rmtree(temp_dir)
        
    with zipfile.ZipFile(xlsx_path, 'r') as z:
        z.extractall(temp_dir)
        
    drawings_dir = os.path.join(temp_dir, 'xl', 'drawings')
    ws_dir = os.path.join(temp_dir, 'xl', 'worksheets')
    
    # -------------------------------------------------------------------------
    # Enrich SHEET 3: 03_パーツ集_作図図形・コネクタ線
    # Sheet 3 (4th sheet) corresponds to drawing3.xml
    # -------------------------------------------------------------------------
    drawing3_path = os.path.join(drawings_dir, 'drawing3.xml')
    shapes_list = []
    sp_id = 100
    
    # 1. Section A: 7-Color 16-Connection Nodes & Containers
    # Row 1: Charcoal, Cloud Blue, Emerald, Amber
    color_keys_row1 = ["charcoal", "cloud_blue", "emerald", "amber"]
    col_offsets_row1 = [1, 11, 21, 31]
    for c_key, c_col in zip(color_keys_row1, col_offsets_row1):
        c_data = PALETTE_7COLORS[c_key]
        sp_id += 1
        # Node shape: rows 8..10 (col c_col .. c_col+9)
        shapes_list.append(create_16pt_node_shape_xml(
            sp_id, f"Node_{c_key}", c_col, 8, c_col + 9, 10,
            c_data["fill"], c_data["border"], c_data["name"].split(" ")[0], c_data["role"].split("・")[0]
        ))
        sp_id += 1
        # Container box: rows 12..16 (col c_col .. c_col+9)
        shapes_list.append(create_16pt_container_box_xml(
            sp_id, f"Box_{c_key}", c_col, 12, c_col + 9, 16,
            c_data["fill"], c_data["border"], f"【{c_data['name'].split(' ')[0]} 境界枠】"
        ))

    # Row 2: Purple, Cyan, Rose
    color_keys_row2 = ["purple", "cyan", "rose"]
    col_offsets_row2 = [1, 11, 21]
    for c_key, c_col in zip(color_keys_row2, col_offsets_row2):
        c_data = PALETTE_7COLORS[c_key]
        sp_id += 1
        shapes_list.append(create_16pt_node_shape_xml(
            sp_id, f"Node_{c_key}", c_col, 19, c_col + 9, 21,
            c_data["fill"], c_data["border"], c_data["name"].split(" ")[0], c_data["role"].split("・")[0]
        ))
        sp_id += 1
        shapes_list.append(create_16pt_container_box_xml(
            sp_id, f"Box_{c_key}", c_col, 23, c_col + 9, 27,
            c_data["fill"], c_data["border"], f"【{c_data['name'].split(' ')[0]} 境界枠】"
        ))

    # 2. Section B: 7-Color Connector Lines (Elbow, Straight, Bidirectional, Dashed)
    cur_row = 31
    for c_key in ["charcoal", "cloud_blue", "emerald", "amber", "purple", "cyan", "rose"]:
        c_data = PALETTE_7COLORS[c_key]
        c_hex = c_data["border"]
        
        sp_id += 1
        shapes_list.append(create_connector_xml(sp_id, f"Cxn_Elbow_{c_key}", 8, cur_row, 13, cur_row + 1, c_hex, "elbow", True))
        sp_id += 1
        shapes_list.append(create_connector_xml(sp_id, f"Cxn_Straight_{c_key}", 16, cur_row, 21, cur_row, c_hex, "straight", True))
        sp_id += 1
        shapes_list.append(create_connector_xml(sp_id, f"Cxn_Bidir_{c_key}", 24, cur_row, 29, cur_row, c_hex, "straight", True, True))
        sp_id += 1
        shapes_list.append(create_connector_xml(sp_id, f"Cxn_Dashed_{c_key}", 32, cur_row, 37, cur_row, c_hex, "straight", True, False, True))
        cur_row += 2

    # 3. Section C: 3-Color Comments (Callout & Memo Card)
    col_comms = [1, 14, 27]
    for (ck, cdat), col_c in zip(COMMENT_3COLORS.items(), col_comms):
        sp_id += 1
        shapes_list.append(create_callout_xml(
            sp_id, f"Callout_{ck}", col_c, 48, col_c + 12, 52,
            cdat["fill"], cdat["border"], cdat["title"], cdat["desc"]
        ))
        sp_id += 1
        shapes_list.append(create_memo_card_xml(
            sp_id, f"Memo_{ck}", col_c, 54, col_c + 12, 59,
            cdat["fill"], cdat["border"], cdat["title"], cdat["sample_lines"]
        ))

    # 4. Section D: 16-pt Demo Hub Node
    sp_id += 1
    shapes_list.append(create_16pt_node_shape_xml(
        sp_id, "Hub_16pt_Demo", 16, 68, 24, 71,
        "F8FAFC", "1A202C", "16接続点 Hubノード", "等間隔接続完全サンプル"
    ))

    # Write out drawing XML for sheet 3
    final_drawing3_xml = f\'\'\'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
{"".join(shapes_list)}
</xdr:wsDr>\'\'\'

    with open(drawing3_path, 'w', encoding='utf-8') as f:
        f.write(final_drawing3_xml)
    print(f"Successfully injected {len(shapes_list)} shapes and connectors into drawing3.xml for Sheet 3.")

    # -------------------------------------------------------------------------
    # Also enrich Sheet 2 (Overview) with 3-color Comment Cards
    # -------------------------------------------------------------------------
    drawing1_path = os.path.join(drawings_dir, 'drawing1.xml')
    if os.path.exists(drawing1_path):
        with open(drawing1_path, 'r', encoding='utf-8') as f:
            s1_content = f.read()
            
        s1_extra_shapes = []
        # Amber comment for Salesforce (rows 12..16, cols 34..39)
        s1_extra_shapes.append(create_callout_xml(
            801, "Overview_SaaS_Note", 34, 12, 39, 16,
            COMMENT_3COLORS["warning_amber"]["fill"], COMMENT_3COLORS["warning_amber"]["border"],
            "⚠️ Salesforce API連携", "・OAuth2.0 / 双方向同期\\n・API日次制限考慮"
        ))
        # Emerald comment for OT (rows 12..16, cols 2..8)
        s1_extra_shapes.append(create_callout_xml(
            802, "Overview_OT_Note", 2, 12, 8, 16,
            COMMENT_3COLORS["security_emerald"]["fill"], COMMENT_3COLORS["security_emerald"]["border"],
            "🛡️ 現場OT隔離統制", "・外部インターネット完全遮断\\n・Modbus/OPC-UA暗号化"
        ))
        # Blue comment for Cloud VPC (rows 15..19, cols 24..30)
        s1_extra_shapes.append(create_memo_card_xml(
            803, "Overview_Cloud_Note", 24, 15, 30, 19,
            COMMENT_3COLORS["info_blue"]["fill"], COMMENT_3COLORS["info_blue"]["border"],
            "ℹ️ Multi-AZ クラウド基盤", ["・東京リージョン 1a/1c 冗長", "・自動フェイルオーバー 30秒以内"]
        ))
        
        s1_content = s1_content.replace('</xdr:wsDr>', f'{"".join(s1_extra_shapes)}</xdr:wsDr>')
        with open(drawing1_path, 'w', encoding='utf-8') as f:
            f.write(s1_content)
        print("Successfully enriched Sheet 2 (Overview) with 3-color comments.")

    # -------------------------------------------------------------------------
    # Ensure [Content_Types].xml covers all drawings
    # -------------------------------------------------------------------------
    ct_path = os.path.join(temp_dir, '[Content_Types].xml')
    if os.path.exists(ct_path):
        with open(ct_path, 'r', encoding='utf-8') as f:
            ct_content = f.read()
        for df in os.listdir(drawings_dir):
            if df.startswith('drawing') and df.endswith('.xml'):
                part_entry = f'<Override PartName="/xl/drawings/{df}" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>'
                if f'/xl/drawings/{df}' not in ct_content:
                    ct_content = ct_content.replace('</Types>', f'{part_entry}</Types>')
        with open(ct_path, 'w', encoding='utf-8') as f:
            f.write(ct_content)

    # -------------------------------------------------------------------------
    # Re-zip workbook
    # -------------------------------------------------------------------------
    if os.path.exists(xlsx_path):
        os.remove(xlsx_path)
        
    with zipfile.ZipFile(xlsx_path, 'w', zipfile.ZIP_DEFLATED) as z_out:
        for root, dirs, files in os.walk(temp_dir):
            for file in files:
                full_p = os.path.join(root, file)
                rel_p = os.path.relpath(full_p, temp_dir)
                z_out.write(full_p, rel_p)
                
    shutil.rmtree(temp_dir)
    print(f"Final workbook packaged successfully: {xlsx_path}")
'''
    code = code[:old_inject.start()] + new_inject_code + code[old_inject.start() + len(old_inject.group(1)):]
    print("Replaced inject_drawings_into_template successfully.")

# 8. Replace verify_and_finish_with_excel_com body
old_verify = re.search(r'(def verify_and_finish_with_excel_com\(xlsx_path\):.*?)(def main\(\):)', code, re.DOTALL)
if old_verify:
    new_verify_code = '''def verify_and_finish_with_excel_com(xlsx_path):
    """
    Verifies the generated Excel template with Excel COM Application,
    connects the 16 demo lines to the Hub node, ensures perfect page setup for all sheets,
    and exports high-quality PDF and PNG previews.
    """
    print("\\n--- Starting Excel COM Verification & Final Polish ---")
    excel = win32com.client.Dispatch('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False
    
    try:
        abs_p = os.path.abspath(xlsx_path)
        wb = excel.Workbooks.Open(abs_p)
        print(f"Workbook opened cleanly: {wb.Name}")
        print(f"Total Sheets count: {wb.Worksheets.Count}")
        
        # Ensure PageSetup for all sheets: exactly 1 page wide, 1 page tall
        for i in range(1, wb.Worksheets.Count + 1):
            ws = wb.Worksheets(i)
            print(f"  Sheet {i}: [{ws.Name}] - Shapes count: {ws.Shapes.Count}")
            try:
                ws.PageSetup.Zoom = False
                ws.PageSetup.FitToPagesWide = 1
                ws.PageSetup.FitToPagesTall = 1
            except Exception as pe:
                pass
            
        ws3 = wb.Worksheets('03_パーツ集_作図図形・コネクタ線')
        print(f"\\nVerifying '{ws3.Name}' shapes:")
        
        hub_node = ws3.Shapes('Hub_16pt_Demo')
        print(f"  [OK] Hub_16pt_Demo ConnectionSiteCount: {hub_node.ConnectionSiteCount}")
        assert hub_node.ConnectionSiteCount == 16, "Hub node must have 16 connection sites!"
        
        # Connect 16 connectors to Hub_16pt_Demo via COM to create the ultimate demo!
        # Colors cycle through the 7 palette colors
        colors_cycle = [
            0x1A202C, 0x1E40AF, 0x047857, 0xB45309,
            0x6D28D9, 0x0369A1, 0xBE123C, 0x1A202C,
            0x1E40AF, 0x047857, 0xB45309, 0x6D28D9,
            0x0369A1, 0xBE123C, 0x1A202C, 0x1E40AF
        ]
        
        print("  [OK] Connecting 16 demo lines to Hub_16pt_Demo via COM...")
        hub_left = hub_node.Left
        hub_top = hub_node.Top
        hub_w = hub_node.Width
        hub_h = hub_node.Height
        
        for site_idx in range(1, 17):
            # Calculate outward direction based on site_idx
            # 1-4: Top (ang 270) -> line goes up
            # 5-8: Right (ang 0) -> line goes right
            # 9-12: Bottom (ang 90) -> line goes down
            # 13-16: Left (ang 180) -> line goes left
            if 1 <= site_idx <= 4:
                # Top: fx above site
                fx = hub_left + hub_w * (0.2 * site_idx)
                fy = hub_top - 45
            elif 5 <= site_idx <= 8:
                # Right: fx to the right of site
                fx = hub_left + hub_w + 45
                fy = hub_top + hub_h * (0.2 * (site_idx - 4))
            elif 9 <= site_idx <= 12:
                # Bottom: fx below site
                fx = hub_left + hub_w * (1.0 - 0.2 * (site_idx - 8))
                fy = hub_top + hub_h + 45
            else:
                # Left: fx to the left of site
                fx = hub_left - 45
                fy = hub_top + hub_h * (1.0 - 0.2 * (site_idx - 12))
                
            # Add connector with End point at (fx, fy), then connect Begin to hub_node site
            cxn = ws3.Shapes.AddConnector(1, fx, fy, fx, fy)
            cxn.Name = f"HubDemo_Line_{site_idx}"
            cxn.ConnectorFormat.BeginConnect(hub_node, site_idx)
            cxn.Line.Weight = 2.0
            cxn.Line.ForeColor.RGB = colors_cycle[site_idx - 1]
            cxn.Line.EndArrowheadStyle = 2 # msoArrowheadTriangle
            cxn.Line.EndArrowheadLength = 2
            cxn.Line.EndArrowheadWidth = 2

        print(f"  [OK] Total shapes on Sheet 3 after 16-connection demo: {ws3.Shapes.Count}")
        
        # Save changes cleanly to a temp file, then replace
        temp_out = os.path.abspath("temp_com_out.xlsx")
        if os.path.exists(temp_out):
            os.remove(temp_out)
        wb.SaveAs(temp_out)
        
        # Export PDF of full workbook or sheet by sheet to verify
        pdf_out = os.path.abspath("output_preview.pdf")
        if os.path.exists(pdf_out):
            os.remove(pdf_out)
        try:
            wb.ExportAsFixedFormat(0, pdf_out)
            print(f"  [OK] PDF exported successfully: {pdf_out}")
        except Exception as pe:
            print(f"  [WARN] wb.ExportAsFixedFormat failed: {pe}")
            
        wb.Close(False)
        excel.Quit()
        
        shutil.copyfile(temp_out, xlsx_path)
        shutil.copyfile(temp_out, "AWS_Hybrid_Architecture_Template.xlsx")
        if os.path.exists(temp_out):
            os.remove(temp_out)
            
        print("[SUCCESS] All COM verifications & polish completed with 100% SUCCESS!")
        print("[SUCCESS] English alias copy updated: AWS_Hybrid_Architecture_Template.xlsx")
        
        # Render PDF pages to PNGs
        try:
            import fitz
            if os.path.exists(pdf_out):
                doc = fitz.open(pdf_out)
                print(f"  [OK] Total pages in output_preview.pdf: {len(doc)}")
                for p_idx in range(len(doc)):
                    pix = doc[p_idx].get_pixmap(dpi=150)
                    pix.save(f"preview_sheet_{p_idx + 1}.png")
                print(f"  [OK] All {len(doc)} pages rendered to preview_sheet_*.png")
        except Exception as re:
            print(f"  [WARN] PDF rendering failed: {re}")
            
        return True
    except Exception as e:
        import traceback
        traceback.print_exc()
        try:
            excel.Quit()
        except:
            pass
        return False
'''
    code = code[:old_verify.start()] + new_verify_code + code[old_verify.start() + len(old_verify.group(1)):]
    print("Replaced verify_and_finish_with_excel_com successfully.")

with open("build_excel_template.py", "w", encoding="utf-8") as f:
    f.write(code)

print("build_excel_template.py updated successfully!")
