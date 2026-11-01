#!/usr/bin/env python3
"""verify_docx.py — 複驗成品 docx（worksheet-docx skill）

用法：
    python verify_docx.py 成品.docx [--export-pdf] [--render] [--outdir 目錄]

流程（SKILL.md 階段 4）：
1. 結構檢查：圖片數、字體、斜體 run、段落間距
2. --export-pdf：依 OS 分支匯出 PDF
     - macOS  → Pages AppleScript（本機 Word AppleScript 已知故障，不用 Word）
     - Windows → Word COM（PowerShell）
3. --render：PyMuPDF 渲染逐頁 PNG 到 <outdir>/pages/，供 agent 逐頁目視
   （必須用圖片檢視工具逐頁看過，才算完成複驗）
"""
import argparse
import os
import platform
import re
import subprocess
import sys
import zipfile


def structural_check(path):
    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml").decode("utf-8")
        media = [n for n in z.namelist() if n.startswith("word/media/")]
    print(f"結構：內嵌圖 {len(media)} 張、斜體 run {xml.count('<w:i/>')} 個、"
          f"下標 {xml.count('subscript')}、上標 {xml.count('superscript')}")
    fonts = sorted(set(re.findall(r'w:(?:ascii|eastAsia)="([^"]+)"', xml)))
    print(f"字體：{fonts}")
    nonzero = re.findall(r'w:(?:before|after)="(?!0")[^"]+"', xml)
    print(("OK  " if not nonzero else "WARN") +
          f" 段落間距：{'全部為 0' if not nonzero else nonzero}")


def export_pdf_mac(docx_path, pdf_path):
    script = f'''
tell application "Pages"
    activate
    open (POSIX file "{docx_path}")
    delay 8
    set theDoc to front document
    export theDoc to (POSIX file "{pdf_path}") as PDF
    close theDoc saving no
end tell'''
    subprocess.run(["osascript", "-e", script], check=True)


def export_pdf_windows(docx_path, pdf_path):
    docx_path = docx_path.replace("'", "''")
    pdf_path = pdf_path.replace("'", "''")
    ps = f"""
$ErrorActionPreference = 'Stop'
$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {{
    $doc = $word.Documents.Open('{docx_path}', $false, $true)
    $doc.SaveAs([ref] '{pdf_path}', [ref] 17)
    $doc.Close($false)
}} finally {{
    $word.Quit()
}}
"""
    subprocess.run(["powershell", "-NoProfile", "-Command", ps], check=True)


def render_pages(pdf_path, outdir, dpi=100):
    import fitz  # PyMuPDF
    os.makedirs(outdir, exist_ok=True)
    doc = fitz.open(pdf_path)
    paths = []
    for i, page in enumerate(doc):
        pix = page.get_pixmap(dpi=dpi)
        p = os.path.join(outdir, f"page-{i+1}.png")
        pix.save(p)
        paths.append(p)
    print(f"已渲染 {len(paths)} 頁到 {outdir}/（dpi={dpi}）")
    for p in paths:
        print("  " + p)
    return paths


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("docx")
    ap.add_argument("--export-pdf", action="store_true")
    ap.add_argument("--render", action="store_true")
    ap.add_argument("--outdir", default=None)
    args = ap.parse_args()

    docx_path = os.path.abspath(args.docx)
    outdir = args.outdir or os.path.splitext(docx_path)[0] + "_verify"
    pdf_path = os.path.splitext(docx_path)[0] + "_verify.pdf"

    structural_check(docx_path)

    if args.export_pdf:
        system = platform.system()
        try:
            if system == "Darwin":
                print("macOS：經 Pages 匯出 PDF…")
                export_pdf_mac(docx_path, pdf_path)
            elif system == "Windows":
                print("Windows：經 Word COM 匯出 PDF…")
                export_pdf_windows(docx_path, pdf_path)
            else:
                sys.exit(f"不支援的平台 {system}：請手動開啟 Word 匯出 PDF")
            print(f"PDF：{pdf_path}")
        except Exception as e:
            sys.exit(f"匯出失敗（{e}）。降級方案：結構檢查＋請用戶自行開啟目視。")

    if args.render:
        if not os.path.exists(pdf_path):
            sys.exit(f"找不到 {pdf_path}，先加 --export-pdf")
        render_pages(pdf_path, os.path.join(outdir, "pages"))
        print("下一步：逐頁目視渲染圖，確認題目齊全、圖文相隨、斜體/上下標正確。")


if __name__ == "__main__":
    main()
