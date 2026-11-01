#!/usr/bin/env python3
"""diagnose_pdf.py — 診斷 PDF 結構：文字層 vs 截圖型（worksheet-docx skill）

用法：
    python diagnose_pdf.py 輸入.pdf

輸出：每頁文字層字元數、內嵌圖片清單（像素尺寸＋頁面位置），
並給出整體判定：TEXT_LAYER / IMAGE_ONLY / MIXED。
依賴：PyMuPDF（pip install pymupdf）
"""
import sys

try:
    import fitz  # PyMuPDF
except ImportError:
    sys.exit("需要 PyMuPDF：pip install pymupdf")


def main(path):
    doc = fitz.open(path)
    total_chars = 0
    total_images = 0
    print(f"檔案：{path}")
    print(f"頁數：{doc.page_count}　頁面尺寸：{doc[0].rect.width:.0f}x{doc[0].rect.height:.0f} pt")
    print("-" * 64)
    for i, page in enumerate(doc):
        text = page.get_text("text").strip()
        total_chars += len(text)
        infos = page.get_image_info(xrefs=True)
        total_images += len(infos)
        kind = "文字" if len(text) > 50 else "幾無文字"
        print(f"第 {i+1} 頁：{kind}（{len(text)} 字元），內嵌圖 {len(infos)} 張")
        for inf in infos:
            bbox = inf["bbox"]
            print(f"    圖 xref={inf.get('xref', '?')} "
                  f"{inf['width']}x{inf['height']} px "
                  f"位置 y={bbox[1]:.0f}–{bbox[3]:.0f}（愈小愈靠頁首）")
    print("-" * 64)
    if total_images == 0:
        verdict = "TEXT_LAYER（無內嵌圖，直接抽取文字即可）"
    elif total_chars < 200 * doc.page_count:
        verdict = ("IMAGE_ONLY（正文為截圖——'抽出文字少'是正常現象，"
                   "走 worksheet-docx 完整流程：直讀圖片轉錄）")
    else:
        verdict = "MIXED（文字層與截圖並存，兩邊都要處理）"
    print(f"判定：{verdict}")
    print(f"合計：{total_chars} 字元、{total_images} 張內嵌圖")


if __name__ == "__main__":
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    main(sys.argv[1])
