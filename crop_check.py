# -*- coding: utf-8 -*-
from PIL import Image

# Crop EC2 icon area on preview_sheet_15.png
im = Image.open("preview_sheet_15.png")
print("Image size:", im.size)

# Let's crop around EC2 icon
# B. AWS サービス詳細アイコンカタログ [AWS Compute & Containers]
# EC2 is around x: 100~300, y: 1000~1300
crop_area = (50, 1000, 500, 1400)
cropped = im.crop(crop_area)
cropped.save("crop_ec2.png")
print("Saved crop_ec2.png")
