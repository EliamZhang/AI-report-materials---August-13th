# -*- coding: utf-8 -*-
"""验证：第3页右侧文字块区域——蓝色加粗标签存在、文字未溢出到下方区域。"""
import sys, os
sys.stdout.reconfigure(encoding='utf-8')
import numpy as np
from PIL import Image

img = np.array(Image.open(r".claude\work\p3_check\slide3.png").convert("RGB")).astype(int)
H, W, _ = img.shape
SX, SY = W / 13.333, H / 7.5

def near(c, t, tol=60):
    return all(abs(int(c[i]) - t[i]) <= tol for i in range(3))

BLUE = (31, 78, 121)

# 文字块区域 x 6.90-12.96in, y 2.18-4.0in
x0, x1 = int(6.90*SX), int(12.96*SX)
y0, y1 = int(2.18*SY), int(4.0*SY)
region = img[y0:y1, x0:x1]
blue_n = sum(near(p, BLUE) for p in region.reshape(-1,3))
print(f"文字块区域蓝色标签像素: {blue_n}")

# 检查是否溢出：文字块底(3.97)与下方"五类定位原因示例"(4.13)之间 y 4.0-4.13 应基本空白
y_a, y_b = int(4.0*SY), int(4.13*SY)
gap = img[y_a:y_b, x0:x1]
dark = sum(near(p, (61,60,58), tol=80) for p in gap.reshape(-1,3))
print(f"间隙区(4.0-4.13in)深色像素: {dark} (应为0附近)")

# 深蓝标签行分布：检查每段开头是否都有蓝字
for seg_i, (yy0, yy1) in enumerate([(2.18,2.55),(2.55,2.95),(2.95,3.35),(3.35,3.97)]):
    seg = img[int(yy0*SY):int(yy1*SY), x0:x1]
    blue_seg = sum(near(p, BLUE) for p in seg.reshape(-1,3))
    print(f"段{seg_i+1} (y{yy0}-{yy1}in) 蓝色像素: {blue_seg}")
