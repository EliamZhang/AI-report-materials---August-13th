# -*- coding: utf-8 -*-
"""检查：标题区域(4.13-4.43)的深色像素分布，区分是标题文字还是溢出的正文。"""
import sys
sys.stdout.reconfigure(encoding='utf-8')
import numpy as np
from PIL import Image

img = np.array(Image.open(r".claude\work\p3_check\slide3.png").convert("RGB")).astype(int)
H, W, _ = img.shape
SX, SY = W/13.333, H/7.5

def near(c, t, tol=80):
    return all(abs(int(c[i])-t[i])<=tol for i in range(3))

BODY = (61, 60, 58)

# 标题"五类定位原因示例"在 (6.80,4.13) 1.21x0.25 —— 左侧 x 6.8-8.0
# 文字块 x 6.9-12.96
for name, x0in, x1in in [("标题区(左)", 6.80, 8.01), ("文字块区(右)", 8.01, 12.96)]:
    x0, x1 = int(x0in*SX), int(x1in*SX)
    total = 0
    for yy in range(int(4.13*SY), int(4.43*SY)):
        row = img[yy, x0:x1]
        total += sum(near(p, BODY) for p in row)
    print(f"{name} (y4.13-4.43): 深色像素={total}")

# 逐行看 4.05-4.55 的内容（整宽）
print("\n逐行(y4.05-4.55):")
for yy in range(int(4.05*SY), int(4.55*SY)):
    row = img[yy, int(6.8*SX):int(12.96*SX)]
    dark = sum(near(p, BODY) for p in row)
    if dark > 2:
        print(f"  y={yy/SY:.2f}in 深色={dark}")
