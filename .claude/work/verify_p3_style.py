# -*- coding: utf-8 -*-
"""验证：第3页右侧两个表格区域出现黑/灰边框线，且文字未溢出表格范围。"""
import sys, os
sys.stdout.reconfigure(encoding='utf-8')
import numpy as np
from PIL import Image

img = np.array(Image.open(r".claude\work\p3_check\slide3.png").convert("RGB")).astype(int)
H, W, _ = img.shape
SX, SY = W / 13.333, H / 7.5

def near(c, t, tol=50):
    return all(abs(int(c[i]) - t[i]) <= tol for i in range(3))

BLACK = (10, 10, 10)
GRAY = (214, 214, 214)

# 表格1区域: (6.90,4.43)-(12.96,5.57)
for name, (x0, y0, x1, y1) in [
    ("表格1(五类定位原因)", (6.90, 4.43, 12.96, 5.57)),
    ("表格2(诊断报告)", (6.90, 6.08, 12.96, 6.84)),
]:
    px0, py0, px1, py1 = int(x0*SX), int(y0*SY), int(x1*SX), int(y1*SY)
    region = img[py0:py1, px0:px1]
    # 横向扫描每行，统计黑线/灰线像素
    row_dark = []
    for yy in range(region.shape[0]):
        cnt_black = sum(near(p, BLACK) for p in region[yy])
        cnt_gray = sum(near(p, GRAY) for p in region[yy])
        row_dark.append((cnt_black, cnt_gray))
    # 找出黑色线行（连续>100黑色像素）
    black_rows = [i for i, (b, g) in enumerate(row_dark) if b > 100]
    gray_rows = [i for i, (b, g) in enumerate(row_dark) if g > 100 and b < 60]
    print(f"{name}: 黑线行={black_rows} 灰线行={gray_rows[:10]}")
    total_black = sum(b for b, g in row_dark)
    print(f"  总黑色像素={total_black}")
