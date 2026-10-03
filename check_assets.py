# -*- coding: utf-8 -*-
import os
import re

with open('build_excel_template.py', 'r', encoding='utf-8') as f:
    text = f.read()

print("Assets referenced in build_excel_template.py:")
assets = sorted(list(set(re.findall(r'assets_icons/[a-zA-Z0-9_\.]+', text))))
missing = []
for a in assets:
    exists = os.path.exists(a)
    if not exists:
        missing.append(a)
    print(f"  {a}: {'OK' if exists else 'MISSING'}")

print(f"\nTotal assets referenced: {len(assets)}, Missing: {len(missing)}")
