# -*- coding: utf-8 -*-
"""用 PowerPoint COM 渲染 V1.5 第4页为 PNG，用于像素验证韦恩图。"""
import sys, os, time
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client

PPT = os.path.abspath("银行流水分享_V1.5.pptx")
OUT = os.path.abspath(".claude/work/venn_check")
os.makedirs(OUT, exist_ok=True)

app = win32com.client.DispatchEx("PowerPoint.Application")
app.Visible = True
try:
    prs = app.Presentations.Open(PPT, WithWindow=False)
    prs.Slides(4).Export(os.path.join(OUT, "slide4.png"), "PNG", 1280, 720)
    print("rendered")
finally:
    try:
        prs.Close()
    except Exception:
        pass
    app.Quit()
