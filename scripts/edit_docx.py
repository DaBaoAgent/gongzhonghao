# -*- coding: utf-8 -*-
"""公众号推文 docx 段落级改稿（gongzhonghao 技能）。

用法：
    python edit_docx.py --list  <docx>                # 列出段落（序号+[IMG]+文字预览）
    python edit_docx.py --set "13=新文字" --set "15=..." <docx>   # 改指定段（图段自动跳过）

只改文字不动图片段；保留 runs[0] 格式，其余 run 清空。
"""
import argparse
import sys

from docx import Document

sys.stdout.reconfigure(encoding="utf-8")


def is_img(p):
    return "graphic" in p._p.xml or "pic:" in p._p.xml


def set_text(doc, idx, text):
    paras = doc.paragraphs
    if idx < 0 or idx >= len(paras):
        raise SystemExit(f"段号越界: {idx} (共{len(paras)}段)")
    p = paras[idx]
    if is_img(p):
        raise SystemExit(f"段 {idx} 是图片段，跳过")
    if p.runs:
        p.runs[0].text = text
        for r in p.runs[1:]:
            r.text = ""
    else:
        p.add_run(text)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("docx")
    ap.add_argument("--list", action="store_true")
    ap.add_argument("--set", action="append", default=[], metavar="N=text")
    args = ap.parse_args()

    doc = Document(args.docx)
    if args.list:
        for i, p in enumerate(doc.paragraphs):
            tag = "[IMG]" if is_img(p) else "     "
            t = p.text.strip()
            if t or tag == "[IMG]":
                print(f"{i:>3} {tag} {t[:60]}")
        return
    total_chars = 0
    for kv in args.set:
        idx, _, text = kv.partition("=")
        set_text(doc, int(idx.strip()), text)
        total_chars += len(text)
    doc.save(args.docx)
    print(f"OK: {len(args.set)}段已改, 新增文字{total_chars}字")


if __name__ == "__main__":
    main()
