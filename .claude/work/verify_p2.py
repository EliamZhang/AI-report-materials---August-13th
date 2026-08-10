# -*- coding: utf-8 -*-
"""像素验证：第2页两个竖标签的浅蓝底+蓝条+竖排文字是否渲染正确。"""
import sys, os
sys.stdout.reconfigure(encoding='utf-8')
import numpy as np
from PIL import Image

img = np.array(Image.open(r".claude\work\p2_check\slide2.png").convert("RGB")).astype(int)
H, W, _ = img.shape
SX, SY = W / 13.333, H / 7.5

def near(c, t, tol=40):
    return all(abs(int(c[i]) - t[i]) <= tol for i in range(3))

BLUE = (31, 78, 121)
BLUE_LIGHT = (227, 235, 244)
TEXT_DARK = (10, 10, 10)

# 左标签 (0.90,5.22) 0.85x1.47 → 像素
def rect_mask(x0, y0, x1, y1):
    px0, py0, px1, py1 = int(x0*SX), int(y0*SY), int(x1*SX), int(y1*SY)
    m = np.zeros((H, W), dtype=bool)
    m[py0:py1, px0:px1] = True
    return m

for name, (x0, y0) in [("左标签", (0.90, 5.22)), ("右标签", (7.07, 5.24))]:
    m = rect_mask(x0, y0, x0+0.85, y0+1.47)
    sub = img[m]
    light = sum(near(p, BLUE_LIGHT) for p in sub)
    # 蓝条区域（左侧 0.06in）
    bar = rect_mask(x0, y0, x0+0.06, y0+1.47)
    barc = img[bar]
    bluen = sum(near(p, BLUE) for p in barc)
    # 文字像素（暗色）
    darkn = sum(near(p, TEXT_DARK, tol=60) for p in sub)
    total = m.sum()
    print(f"{name}: 区域={total}px 浅蓝={light}({light/total:.1%}) 蓝条={bluen}({bluen/len(barc):.1%}) 深色文字={darkn}")
