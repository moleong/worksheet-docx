#!/usr/bin/env python3
"""merge_docx.py — 把多個同風格的分章 docx 合併成單冊（純搬運，不重產內容）

用法：
    python merge_docx.py 輸出.docx 章1.docx 章2.docx ... \
        --header "泛珠補充A1" [--title "01 斜拋運動" --title "02 …" …]

行為（對應 SKILL.md「合併成冊」需求）：
- 第一章為基底（保留其頁面設定 sectPr）；之後每章前插入分頁符，再深拷貝 body
- 圖片關係自動重映射：逐一複製來源 media part 到基底，替換拷貝 XML 中的
  r:embed / r:id（依內容去重，同圖不複製）
- 每章首個有文字的段落（標題）文字替換為對應 --title 值，保留段落設定與
  字級（rPr 取原首個 run）。--title 省略時維持原标题
- 新增頁首：--header 文字置左、10 pt、中文新細明體/英文 Times New Roman，
  全冊每頁自動出現（单一 section，不做 different-first-page）
- 只保留基底的 w:sectPr（各章頁面設定必須一致，由同一 builder 產生）

前提：所有分章檔由 build_docx.py 同參數產生（字體、頁面、樣式一致）。
"""
import argparse
import io

from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.enum.text import WD_BREAK

_R_ATTRS = (qn("r:embed"), qn("r:id"), qn("r:link"))


def _replace_title(doc, new_title):
    """把文件首個有文字的段落的文字換成 new_title，保留段落設定與首 run 的 rPr。"""
    for p in doc.paragraphs:
        if p.text.strip():
            runs = p.runs
            if runs:
                runs[0].text = new_title
                for r in runs[1:]:
                    r._element.getparent().remove(r._element)
            else:
                r = p.add_run(new_title)
                r.font.size = Pt(18)
                r.font.name = "Times New Roman"
                r._element.rPr.rFonts.set(qn("w:eastAsia"), "新細明體")
            return True
    return False


def _copy_chapter(base, src):
    """把 src 的 body 內容（除去末尾 sectPr）接到 base 末尾，重映射圖片 rId。"""
    base_part, src_part = base.part, src.part
    body = base.element.body
    sectPr = body.find(qn("w:sectPr"))

    # 圖片關係重映射：依 blob 內容在基底取得（或建）對應的 rId
    rid_map = {}
    for rel in src_part.rels.values():
        if rel.is_external or "/image" not in rel.reltype:
            continue
        rid_map[rel.rId] = base_part.get_or_add_image(io.BytesIO(rel.target_part.blob))[0]

    for el in src_part.element.body:
        if el.tag == qn("w:sectPr"):
            continue
        clone = __import__("copy").deepcopy(el)
        for node in clone.iter():
            for attr in _R_ATTRS:
                v = node.get(attr)
                if v in rid_map:
                    node.set(attr, rid_map[v])
        sectPr.addprevious(clone)


def _add_page_break(doc):
    p = doc.add_paragraph()
    p.add_run().add_break(WD_BREAK.PAGE)


def _set_header(doc, text):
    hdr = doc.sections[0].header
    hdr.is_linked_to_previous = False
    p = hdr.paragraphs[0]
    for r in list(p.runs):
        r._element.getparent().remove(r._element)
    r = p.add_run(text)
    r.font.size = Pt(10)
    r.font.name = "Times New Roman"
    r._element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), "新細明體")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("output")
    ap.add_argument("chapters", nargs="+")
    ap.add_argument("--header", required=True, help="頁首文字，如「泛珠補充A1」")
    ap.add_argument("--title", action="append", default=[],
                    help="章節簡化標題，依序列出（可多次）；省略時不改標題")
    args = ap.parse_args()

    base = Document(args.chapters[0])
    if args.title:
        assert len(args.title) == len(args.chapters), \
            f"--title 數量（{len(args.title)}）需與章數（{len(args.chapters)}）相同"
        _replace_title(base, args.title[0])

    for i, path in enumerate(args.chapters[1:], start=1):
        src = Document(path)
        if args.title:
            _replace_title(src, args.title[i])
        _add_page_break(base)
        _copy_chapter(base, src)

    _set_header(base, args.header)
    base.save(args.output)
    print(f"merged {len(args.chapters)} chapters -> {args.output}")


if __name__ == "__main__":
    main()
