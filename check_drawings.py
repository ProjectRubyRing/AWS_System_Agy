# -*- coding: utf-8 -*-
import zipfile
import xml.etree.ElementTree as ET

with zipfile.ZipFile("AWS_Hybrid_Architecture_Template.xlsx", "r") as z:
    for s_idx in range(1, 7):
        rel_path = f"xl/worksheets/_rels/sheet{s_idx}.xml.rels"
        if rel_path in z.namelist():
            tree = ET.fromstring(z.read(rel_path))
            drawings = [elem.attrib.get('Target') for elem in tree if 'drawing' in elem.attrib.get('Type', '')]
            print(f"Sheet {s_idx} -> drawings: {drawings}")
        else:
            print(f"Sheet {s_idx} -> NO RELS")

    # Now let's inspect drawing4 and drawing5 rels
    for d_idx in [4, 5]:
        d_rel = f"xl/drawings/_rels/drawing{d_idx}.xml.rels"
        if d_rel in z.namelist():
            tree = ET.fromstring(z.read(d_rel))
            images = [elem.attrib.get('Target') for elem in tree if 'image' in elem.attrib.get('Type', '')]
            print(f"drawing{d_idx}.xml.rels -> {len(images)} images: first 5 = {images[:5]}")
        else:
            print(f"drawing{d_idx}.xml.rels -> NOT FOUND")
