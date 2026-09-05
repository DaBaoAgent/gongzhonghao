# -*- coding: utf-8 -*-
"""公众号推文 docx 生成（gongzhonghao 技能）。

用法：编辑脚本头部 CONTENT / IMAGES 配置后运行：
    python make_docx.py [输出路径.docx]

排版规范：微软雅黑全文；主标题16pt深蓝加粗居中；小节标题14pt红加粗；
正文11pt 1.5倍行距；图注9pt灰居中；配图宽15cm居中。
"""
import sys
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.shared import Cm, Pt, RGBColor

sys.stdout.reconfigure(encoding="utf-8")

# ===== 在这里配置内容 =====
TITLE = "直击福祉博览会｜福宏康复五大明星产品亮相北京，全碳电动轮椅成焦点"
SUBTITLE = "2026中国国际福祉博览会 · 北京国家会议中心（9.3-9.5）"
# 正文块列表：("h"=小节标题 / "p"=正文 / "img"=图片 / "cap"=图注, 文本或图片路径)
BLOCKS = [
    ("p", "9月3日至5日，2026中国国际福祉博览会在北京国家会议中心举行，深耕康复辅具近20年的昆山福宏康复携五大类明星产品亮相。"),
    ("h", "◆ 展位直击：蓝白科技风，人气持续爆棚"),
    ("p", "福宏展位蓝白配色、发光拱门与科技网格背板十分醒目，背景板“全球客户分布”地图见证产品远销欧美亚澳。"),
    ("img", r"D:\BaiduSyncdisk\5 @文档\@公众号推文\9.3-5号北京展会\微信图片_20260905205615_446_30.jpg"),
    ("cap", "蓝白展位人气火爆，全球客户分布地图见证“苏州智造”实力"),
    # ... 按需增删
]
DEFAULT_OUT = r"D:\BaiduSyncdisk\5 @文档\@公众号推文\output.docx"
# =========================

BLUE = RGBColor(0x1F, 0x4E, 0x79)
RED = RGBColor(0xC0, 0x00, 0x00)
GRAY = RGBColor(0x80, 0x80, 0x80)


def add_para(doc, text, size, color=None, bold=False, align=None, line15=False):
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    r = p.add_run(text)
    r.font.name = "微软雅黑"
    r.font.size = Pt(size)
    r.bold = bold
    if color is not None:
        r.font.color.rgb = color
    if line15:
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
        p.paragraph_format.space_after = Pt(6)
    return p


def main():
    doc = Document()
    add_para(doc, TITLE, 16, BLUE, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    add_para(doc, SUBTITLE, 11, GRAY, align=WD_ALIGN_PARAGRAPH.CENTER)
    doc.add_paragraph()
    for kind, val in BLOCKS:
        if kind == "h":
            add_para(doc, val, 14, RED, bold=True)
        elif kind == "p":
            add_para(doc, val, 11, line15=True)
        elif kind == "img":
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.add_run().add_picture(val, width=Cm(15))
        elif kind == "cap":
            add_para(doc, val, 9, GRAY, align=WD_ALIGN_PARAGRAPH.CENTER)
    out = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(DEFAULT_OUT)
    out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out))
    print("saved:", out)


if __name__ == "__main__":
    main()
