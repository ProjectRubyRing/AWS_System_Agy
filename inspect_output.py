# -*- coding: utf-8 -*-
import os
import sys
import zipfile
import openpyxl
import win32com.client
from PIL import Image

print("=== 1. Checking Excel Files ===")
excel_files = [
    "AWS_Hybrid_Architecture_Template.xlsx",
    "AWS_システム構成図_テンプレート.xlsx"
]

for name in excel_files:
    if not os.path.exists(name):
        print(f"[MISSING] {name}")
        continue
    size = os.path.getsize(name)
    print(f"\n[FILE] {name} (size: {size:,} bytes)")
    
    # 1. Zip check
    try:
        with zipfile.ZipFile(name, 'r') as z:
            bad = z.testzip()
            if bad:
                print(f"  [ERROR] Zip CRC/corruption in {bad}")
            else:
                print(f"  [OK] Zip integrity verified cleanly")
            drawings = [f for f in z.namelist() if 'drawing' in f]
            print(f"  [INFO] Drawing files: {drawings}")
    except Exception as e:
        print(f"  [ERROR] Zip read error: {e}")

    # 2. openpyxl check
    try:
        wb = openpyxl.load_workbook(name, data_only=True)
        print(f"  [OK] openpyxl loaded workbook: {wb.sheetnames}")
    except Exception as e:
        print(f"  [ERROR] openpyxl load failed: {e}")

# 3. Excel COM check
print("\n=== 2. Checking Excel via COM ===")
try:
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    
    for name in excel_files:
        path = os.path.abspath(name)
        if not os.path.exists(path):
            continue
        try:
            wb = excel.Workbooks.Open(path, ReadOnly=True)
            print(f"[COM OPEN OK] {name}")
            print(f"  Sheet count: {wb.Sheets.Count}")
            for i in range(1, wb.Sheets.Count + 1):
                ws = wb.Sheets(i)
                shapes_count = ws.Shapes.Count
                print(f"    Sheet {i} [{ws.Name}]: {shapes_count} shapes")
            wb.Close(False)
        except Exception as e:
            print(f"[COM ERROR] Could not open {name}: {e}")
    excel.Quit()
except Exception as e:
    print(f"[COM INIT ERROR] {e}")

# 4. Checking PDF
print("\n=== 3. Checking output_preview.pdf ===")
pdf_path = "output_preview.pdf"
if os.path.exists(pdf_path):
    print(f"[FILE] {pdf_path} (size: {os.path.getsize(pdf_path):,} bytes)")
    try:
        import pypdf
        reader = pypdf.PdfReader(pdf_path)
        print(f"  [OK] PDF Page count: {len(reader.pages)}")
        for idx, page in enumerate(reader.pages):
            text = page.extract_text() or ""
            print(f"    Page {idx+1}: text length = {len(text)}, first 40 chars = {repr(text[:40])}")
    except ImportError:
        try:
            import fitz # PyMuPDF
            doc = fitz.open(pdf_path)
            print(f"  [OK] PyMuPDF Page count: {len(doc)}")
        except Exception as e2:
            print(f"  [NOTE] pypdf/fitz not installed: {e2}")
    except Exception as e:
        print(f"  [ERROR] PDF read error: {e}")
else:
    print(f"[MISSING] {pdf_path}")

# 5. Checking PNG Previews
print("\n=== 4. Checking preview_sheet_*.png ===")
pngs = sorted([f for f in os.listdir(".") if f.startswith("preview_sheet_") and f.endswith(".png")],
              key=lambda x: int(x.split("_")[2].split(".")[0]) if x.split("_")[2].split(".")[0].isdigit() else 0)
print(f"Found {len(pngs)} preview pngs: {pngs}")
for p in pngs:
    try:
        im = Image.open(p)
        print(f"  {p}: size={im.size}, mode={im.mode}, file_size={os.path.getsize(p):,} bytes")
    except Exception as e:
        print(f"  {p}: [ERROR] {e}")
