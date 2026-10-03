# -*- coding: utf-8 -*-
import zipfile
import xml.etree.ElementTree as ET

with zipfile.ZipFile("AWS_Hybrid_Architecture_Template.xlsx", "r") as z:
    ct = z.read("[Content_Types].xml").decode('utf-8')
    print("Content_Types.xml has png?:", "png" in ct)
    
    # check drawing5.xml.rels targets
    d5_rels_xml = z.read("xl/drawings/_rels/drawing5.xml.rels").decode('utf-8')
    root_rels = ET.fromstring(d5_rels_xml)
    rel_map = {}
    for elem in root_rels:
        rId = elem.attrib.get('Id')
        target = elem.attrib.get('Target')
        rel_map[rId] = target
        
    print(f"Total rels in drawing5.xml.rels: {len(rel_map)}")
    
    # check if targets exist in zip
    missing_targets = []
    for rId, target in rel_map.items():
        # Target is like '../media/image62.png' -> 'xl/media/image62.png'
        normalized = target.replace('../', 'xl/')
        if normalized not in z.namelist():
            missing_targets.append(normalized)
    print(f"Missing targets in zip for drawing5: {missing_targets}")

    # check drawing5.xml blip tags
    d5_xml = z.read("xl/drawings/drawing5.xml").decode('utf-8')
    root_d5 = ET.fromstring(d5_xml)
    ns = {'xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
          'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
          'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
    
    blips = root_d5.findall('.//a:blip', ns)
    print(f"Total blip tags in drawing5.xml: {len(blips)}")
    unmatched = []
    for b in blips:
        embed_id = b.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
        if embed_id not in rel_map:
            unmatched.append(embed_id)
    print(f"Unmatched blip embed IDs: {unmatched}")
