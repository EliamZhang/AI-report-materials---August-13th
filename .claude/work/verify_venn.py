# -*- coding: utf-8 -*-
"""像素验证：检查 slide4.png 中韦恩图区域是否存在蓝/金色圆形像素。"""
import sys, os
sys.stdout.reconfigure(encoding='utf-8')
import numpy as np
from PIL import Image

img = np.array(Image.open(r".claude\work\venn_check\slide4.png").convert("RGB")).astype(int)
H, W, _ = img.shape
print(f"图片尺寸 {W}x{H}")

def near(c, target, tol=40):
    return all(abs(int(c[i]) - target[i]) <= tol for i in range(3))

BLUE = (31, 78, 121)
GOLD = (214, 160, 0)

# 左栏区域：x 0.3-5.0in, y 1.0-4.1in → 缩放后
def in_px(inch_x, inch_y):
    return int(inch_x / 13.333 * W), int(inch_y / 7.5 * H)

x0, x1 = in_px(0.3, 0)[0], in_px(5.0, 0)[0]
y0, y1 = in_px(0, 1.0)[1], in_px(0, 4.2)[1]
region = img[y0:y1, x0:x1]

blue_mask = np.array([near(p, BLUE) for p in region.reshape(-1, 3)]).reshape(region.shape[:2])
gold_mask = np.array([near(p, GOLD) for p in region.reshape(-1, 3)]).reshape(region.shape[:2])

print(f"左栏区域: x[{x0},{x1}] y[{y0},{y1}]")
print(f"蓝色像素数: {blue_mask.sum()}  占比: {blue_mask.mean():.4f}")
print(f"金色像素数: {gold_mask.sum()}  占比: {gold_mask.mean():.4f}")

# 圆形分布检查：蓝色像素是否呈圆形分布（中心集中）
if blue_mask.sum() > 500:
    ys, xs = np.nonzero(blue_mask)
    cy, cx = (ys.min() + ys.max()) / 2, (xs.min() + xs.max()) / 2
    rx, ry = (xs.max() - xs.min()) / 2, (ys.max() - ys.min()) / 2
    print(f"蓝色像素包围盒: 中心({cx:.0f},{cy:.0f}) 半径x={rx:.0f} y={ry:.0f} 纵横比={rx/ry:.2f}")
    if 0.7 < rx / ry < 1.4:
        print("→ 蓝色形状近似圆形 ✓")
    else:
        print("→ 蓝色形状不是圆形（可能是条形图残留）✗")

if gold_mask.sum() > 500:
    ys, xs = np.nonzero(gold_mask)
    cy, cx = (ys.min() + ys.max()) / 2, (xs.min() + xs.max()) / 2
    rx, ry = (xs.max() - xs.min()) / 2, (ys.max() - ys.min()) / 2
    print(f"金色像素包围盒: 中心({cx:.0f},{cy:.0f}) 半径x={rx:.0f} y={ry:.0f} 纵横比={rx/ry:.2f}")
    if 0.7 < rx / ry < 1.4:
        print("→ 金色形状近似圆形 ✓")
    else:
        print("→ 金色形状不是圆形 ✗")

# 交集区域颜色：蓝金叠加应偏深/偏绿
mixed = region[blue_mask & gold_mask]
print(f"蓝金同时命中的像素: {len(mixed)}")
