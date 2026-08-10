# -*- coding: utf-8 -*-
"""渲染 V1.7 第2页为PNG。"""
import sys, os
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client

PPT = os.path.abspath("银行流水分享_V1.7.pptx")
OUT = os.path.abspath(".claude/work/p2_check")
os.makedirs(OUT, exist_ok=True)

app = win32com.client.DispatchEx("PowerPoint.Application")
app.Visible = True
try:
    prs = app.Presentations.Open(PPT, WithWindow=False)
    prs.Slides(2).Export(os.path.join(OUT, "slide2.png"), "PNG", 1280, 720)
    print("rendered")
finally:
    try:
        prs.Close()
    except Exception:
        pass
    app.Quit()
