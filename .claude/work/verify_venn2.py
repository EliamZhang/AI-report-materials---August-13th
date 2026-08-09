# -*- coding: utf-8 -*-
"""按 alpha 混合后的实际颜色验证韦恩图两个圆及其交集区域。"""
import sys, os
sys.stdout.reconfigure(encoding='utf-8')
import numpy as np
from PIL import Image

img = np.array(Image.open(r".claude\work\venn_check\slide4.png").convert("RGB")).astype(int)
H, W, _ = img.shape
# 英寸 → 像素 (1280/13.333in 宽, 720/7.5in 高)
SX, SY = W / 13.333, H / 7.5

# 理论混合色（alpha 28% 叠加白色背景）
def mix(rgb, alpha):
    return tuple(int(rgb[i] * alpha + 255 * (1 - alpha)) for i in range(3))

BLUE = (31, 78, 121)
GOLD = (214, 160, 0)
BLUE_LIGHT = mix(BLUE, 0.28)   # (192, 205, 218)
GOLD_LIGHT = mix(GOLD, 0.28)   # (244, 228, 184)

# 左圆区域
def circle_mask(cx_in, cy_in, r_in):
    cx, cy, r = int(cx_in * SX), int(cy_in * SY), int(r_in * SX)
    yy, xx = np.mgrid[0:H, 0:W]
    return (xx - cx) ** 2 + (yy - cy) ** 2 <= r ** 2

# 圆心：左圆 (0.55+1.175, 1.55+1.175) = (1.725, 2.725)，半径 1.175
L = circle_mask(1.725, 2.725, 1.175)
R = circle_mask(1.725 + 1.3, 2.725, 1.175)   # 右圆圆心 3.025
inter = L & R
only_l = L & ~R
only_r = R & ~L

def avg_color(mask):
    if mask.sum() == 0:
        return None
    return tuple(int(v) for v in img[mask].mean(axis=0))

print("左圆独有区平均色:", avg_color(only_l), "预期:", BLUE_LIGHT)
print("右圆独有区平均色:", avg_color(only_r), "预期:", GOLD_LIGHT)
print("交集区平均色:", avg_color(inter), "(蓝金叠加)")
print()
print("左圆独有区像素数:", only_l.sum(), "右圆独有区:", only_r.sum(), "交集区:", inter.sum())

# 边框线检查：纯色边框像素数
def count_near(mask, target, tol=60):
    sub = img[mask]
    d = np.abs(sub - np.array(target)).max(axis=1)
    return int((d <= tol).sum())

print()
print("左圆内纯蓝(边框+文字)像素:", count_near(only_l, BLUE))
print("右圆内纯金(边框+文字)像素:", count_near(only_r, GOLD))
print("交集内深橄榄(蓝金叠加文字)像素:", count_near(inter, (198, 192, 61), tol=70))
