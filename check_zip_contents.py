# -*- coding: utf-8 -*-
import zipfile

with zipfile.ZipFile("AWS_Hybrid_Architecture_Template.xlsx", "r") as z:
    names = z.namelist()
    media_files = [n for n in names if n.startswith("xl/media/")]
    rels_files = [n for n in names if "rels" in n]
    print(f"Total files in zip: {len(names)}")
    print(f"Total media files in zip: {len(media_files)}")
    print(f"Total rels files in zip: {len(rels_files)}")
    print("\nMedia files:")
    for m in media_files[:10]:
        info = z.getinfo(m)
        print(f"  {m}: file_size={info.file_size}, compress_size={info.compress_size}")
    if len(media_files) > 10:
        print(f"  ... and {len(media_files)-10} more media files")
        
    print("\nRels files:")
    for r in rels_files:
        print(f"  {r}")
