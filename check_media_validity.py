# -*- coding: utf-8 -*-
import zipfile
import io
from PIL import Image

for fname in ["AWS_Hybrid_Architecture_Template.xlsx", "AWS_現場システム統合構成図_テンプレート.xlsx"]:
    print(f"\n================ Checking {fname} ================")
    try:
        with zipfile.ZipFile(fname, "r") as z:
            media_files = [n for n in z.namelist() if n.startswith("xl/media/")]
            print(f"Total media files: {len(media_files)}")
            corrupted = []
            formats = {}
            for m in media_files:
                data = z.read(m)
                try:
                    im = Image.open(io.BytesIO(data))
                    im.verify()
                    formats[im.format] = formats.get(im.format, 0) + 1
                except Exception as e:
                    corrupted.append((m, str(e)))
                    
            print(f"Formats detected: {formats}")
            print(f"Corrupted images count: {len(corrupted)}")
            for m, err in corrupted:
                print(f"  {m}: {err}")
    except Exception as e:
        print(f"Error checking {fname}: {e}")
