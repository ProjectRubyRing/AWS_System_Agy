# -*- coding: utf-8 -*-
import fitz # PyMuPDF
import os
from PIL import Image

pdf_path = "output_preview.pdf"
print(f"=== Inspecting {pdf_path} ===")
doc = fitz.open(pdf_path)
print(f"Total pages: {len(doc)}")

for i in range(len(doc)):
    page_num = i + 1
    page = doc[i]
    rect = page.rect
    text = page.get_text()
    first_line = text.split("\n")[0] if text else "(No text)"
    
    # Check images on page
    images = page.get_images()
    
    print(f"Page {page_num:2d}: size=({rect.width:.1f}, {rect.height:.1f}), images_count={len(images)}, text_len={len(text)}, first_line={first_line[:40]}")
    
    # Try rendering pixmap to see if zlib error triggers
    try:
        pix = page.get_pixmap(dpi=150)
    except Exception as e:
        print(f"  [ERROR rendering Page {page_num}]: {e}")

print("\n=== Checking preview PNG images ===")
for i in range(1, 19):
    png_path = f"preview_sheet_{i}.png"
    if os.path.exists(png_path):
        try:
            im = Image.open(png_path)
            im.verify() # verify integrity
            # re-open to check bbox
            im = Image.open(png_path)
            bbox = im.getbbox()
            extrema = im.getextrema()
            print(f"  {png_path}: OK, size={im.size}, bbox={bbox}, file_size={os.path.getsize(png_path):,} bytes")
        except Exception as e:
            print(f"  {png_path}: [CORRUPT/ERROR] {e}")
    else:
        print(f"  {png_path}: [NOT FOUND]")
