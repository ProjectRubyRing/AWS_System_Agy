# -*- coding: utf-8 -*-
import os
import zipfile
import win32com.client

# Let's inspect drawing3.xml from AWS_Hybrid_Architecture_Template.xlsx
with zipfile.ZipFile("AWS_Hybrid_Architecture_Template.xlsx", "r") as z:
    for name in z.namelist():
        if "drawing3.xml" in name:
            content = z.read(name).decode("utf-8")
            with open("dump_drawing3.xml", "w", encoding="utf-8") as out:
                out.write(content)
            print(f"Dumped {name}, size: {len(content)}")
