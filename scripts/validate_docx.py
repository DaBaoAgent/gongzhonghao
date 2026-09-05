# -*- coding: utf-8 -*-
"""公众号推文 docx 校验（gongzhonghao 技能）。

用法：python validate_docx.py <docx>
输出 JSON：字数/图片数/图注/标题长度/字体抽查/违禁词命中。全绿 valid=true。
"""
import json
import re
import sys

from docx import Document

sys.stdout.reconfigure(encoding="utf-8")

BANNED = [
    "【", "】", "镜头切到", "画面来到", "划重点", "说白了", "接下来",
    "总的来说", "值得注意的是", "不可否认的是", "真正重要的是",
    "本质上", "底层逻辑", "惊艳", "震撼全场", "燃爆", "刷屏",
]


def is_img(p):
    return "graphic" in p._p.xml or "pic:" in p._p.xml


def main():
    path = sys.argv[1]
    doc = Document(path)
    paras = doc.paragraphs
    text_paras = [p for p in paras if not is_img(p) and p.text.strip()]
    img_paras = [p for p in paras if is_img(p)]
    body = "\n".join(p.text for p in text_paras)
    chars = len(re.sub(r"\s", "", body))

    title = text_paras[0].text.strip() if text_paras else ""
    # 图注检查：每个图段下一非空段
    caps_missing = []
    for i, p in enumerate(paras):
        if not is_img(p):
            continue
        nxt = next((q.text.strip() for q in paras[i + 1:] if q.text.strip()), "")
        if not nxt or len(nxt) > 40 or len(nxt) < 5:
            caps_missing.append(i)

    # 字体抽查：前3个非空段的run字体
    fonts = set()
    for p in text_paras[:3]:
        for r in p.runs:
            if r.font.name:
                fonts.add(r.font.name)

    banned_hits = [b for b in BANNED if b in body]

    problems = []
    if not (0 < len(title) <= 35):
        problems.append(f"标题长度{len(title)}需1-30字")
    if len(img_paras) < 1:
        problems.append("无配图")
    if caps_missing:
        problems.append(f"图段{caps_missing}下方缺图注/图注异常")
    if not (500 <= chars <= 1500):
        problems.append(f"字数{chars}超出500-1500")
    if fonts and not fonts <= {"微软雅黑"}:
        problems.append(f"字体非微软雅黑: {fonts}")
    if banned_hits:
        problems.append(f"违禁词: {banned_hits}")

    result = {
        "file": path,
        "chars": chars,
        "images": len(img_paras),
        "title_len": len(title),
        "fonts": sorted(fonts),
        "banned": banned_hits,
        "caption_issues": caps_missing,
        "valid": not problems,
        "problems": problems,
    }
    print(json.dumps(result, ensure_ascii=False, indent=1))
    sys.exit(0 if result["valid"] else 1)


if __name__ == "__main__":
    main()
