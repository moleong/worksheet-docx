# worksheet-docx

把 **PDF 或圖片形式的物理練習卷**（截圖題、掃描題）轉製成**可編輯、排版規範的 Word 文檔**。
適用於 Kimi Code 等支援 SKILL.md 的 agent；核心價值是物理排版規範（變量斜體、單位正體）、
術語標準推導、插圖精確裁切，以及跨平台（macOS / Windows）的複驗流程。

## 觸發關鍵字

練習題、測驗、工作紙、講義轉Word、pdf轉Word、把 PDF 做成可編輯 Word、掃描題目錄入 Word

## 安裝

```bash
# 1. 安裝 skill（Kimi Code；整個 git 目錄即 skill 目錄）
git clone https://github.com/moleong/worksheet-docx.git ~/.kimi-code/skills/worksheet-docx

# 2. Python 依賴
pip install python-docx opencc-python-reimplemented pymupdf pillow

# 3. 平台額外依賴
#    macOS：brew install poppler（pdfimages/pdftotext 診斷用；渲染由 PyMuPDF 負責）
#    Windows：安裝 Microsoft Word（複驗匯出 PDF 用）
```

## 目錄結構

```
worksheet-docx/
├── SKILL.md              # 工作流程（五階段）、開工前確認清單、避雷清單
└── scripts/
    ├── diagnose_pdf.py   # 階段 0：診斷 PDF（文字層 vs 截圖型）
    ├── crop_figures.py   # 階段 2：插圖裁切（單張或 boxes.json 批次）
    ├── build_docx.py     # 階段 3：docx 構建函式庫（含四項自動檢查）
    ├── verify_docx.py    # 階段 4：複驗（OS 分支匯出 PDF＋逐頁渲染）
    └── example_build.py  # 內容填寫範本（每次任務複製改造）
```

## 快速開始（agent 視角）

1. `python scripts/diagnose_pdf.py 輸入.pdf` — 判定輸入類型
2. 依 SKILL.md「階段 1」與用戶確認術語/字體/簡繁（先 grep 用戶現有講義推導術語）
3. 直讀圖片轉錄題目，`crop_figures.py` 裁插圖並**逐張目視複驗**
4. 複製 `example_build.py` 填入內容，執行（內建 self_check）
5. `python scripts/verify_docx.py 成品.docx --export-pdf --render`，**逐頁目視**渲染圖
6. 原位部署

## 排版預設（可於 build_docx.py 修改）

| 項目 | 預設 |
|---|---|
| 中文字體 | 新細明體（需目標機安裝；未安裝時系統自動替代顯示） |
| 英文/數字 | Times New Roman |
| 頁面 | A4，邊界 2 cm，內文 11 pt |
| 行距 / 段落間距 | 1.0 / 0（題間一空行分隔） |
| 變量 | 斜體（x, t, F, k, m, R, L, T, s, g, h, α、點位 A/B/O/C） |
| 單位/函數 | 正體（cm, m, s, Hz, N, r/min、sin、Δ、選項前綴 A. B. C. D.） |
| 分頁 | 每題設「與下段同頁」，圖文永不分散 |

## 已知限制

- 圖片內的點陣文字（圖題、坐標軸）無法改字體或轉繁體
- 手寫/低解析度掃描需人工校對數字與符號
- 複雜公式以線性格式（3L/4、√2、10⁻²）呈現

## License

MIT（教學用途自由使用）
