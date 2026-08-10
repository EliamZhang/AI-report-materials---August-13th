# -*- coding: utf-8 -*-
"""精确检测文字块溢出到的最低行，并测出文字实际底界。"""
import sys
sys.stdout.reconfigure(encoding='utf-8')
import numpy as np
from PIL import Image

img = np.array(Image.open(r".claude\work\p3_check\slide3.png").convert("RGB")).astype(int)
H, W, _ = img.shape
SX, SY = W/13.333, H/7.5

def near(c, t, tol=70):
    return all(abs(int(c[i])-t[i])<=tol for i in range(3))

# 右侧文字区域 x 6.9-12.96in，从 2.1 到 4.5in 逐行扫描
x0, x1 = int(6.9*SX), int(12.96*SX)
y_end = int(4.5*SY)
BODY = (61, 60, 58)

last_content_y = 0
for yy in range(int(2.1*SY), y_end):
    row = img[yy, x0:x1]
    dark = sum(near(p, BODY) for p in row)
    if dark > 5:
        last_content_y = yy

print(f"文字块内容最低像素行: y={last_content_y}px = {last_content_y/SY:.2f}in")
print(f"'五类定位原因示例'标题顶: 4.13in")

# 4.13in 之下是否还有文字像素
below = img[int(4.13*SY):int(4.43*SY), x0:x1]
dark_below = sum(near(p, BODY) for p in below.reshape(-1,3))
print(f"标题区域(y4.13-4.43)深色像素: {dark_below} (应≈0)")
