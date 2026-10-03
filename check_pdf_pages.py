# -*- coding: utf-8 -*-
import fitz

doc = fitz.open("output_preview.pdf")
print("Total pages in PDF:", len(doc))
for i, page in enumerate(doc):
    text = page.get_text()
    lines = [l.strip() for l in text.split('\n') if l.strip()]
    first_lines = " | ".join(lines[:3]) if lines else "(empty)"
    print(f"Page {i+1:2d}: size=({page.rect.width:.1f}, {page.rect.height:.1f}) | {first_lines[:80]}")
