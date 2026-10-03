# -*- coding: utf-8 -*-
import os
import re

with open('build_excel_template.py', 'r', encoding='utf-8') as f:
    lines = f.readlines()

referenced_pngs = set()
for line in lines:
    matches = re.findall(r"['\"]([a-zA-Z0-9_\-]+\.png)['\"]", line)
    for m in matches:
        referenced_pngs.add(m)

print(f"Total png filenames found in code: {len(referenced_pngs)}")
missing_in_assets = []
missing_in_root = []
for p in sorted(referenced_pngs):
    in_assets = os.path.exists(os.path.join("assets_icons", p))
    in_root = os.path.exists(p)
    if not in_assets and not in_root:
        missing_in_assets.append(p)
    print(f"  {p:30s}: in_assets={in_assets}, in_root={in_root}")

print(f"\nMissing files completely: {len(missing_in_assets)}")
for m in missing_in_assets:
    print(f"  MISSING: {m}")
