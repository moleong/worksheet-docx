---
name: worksheet-docx
description: >
  把 PDF 或圖片形式的物理練習卷（截圖題、掃描題）轉製成可編輯的 Word 文檔。
  工作流程：診斷輸入（文字層 vs 截圖）→ 多模態直讀轉錄 → 插圖裁切嵌入 →
  python-docx 構建（物理排版規範：變量斜體、單位正體）→ 跨平台複驗（macOS 經
  Pages、Windows 經 Word 匯出 PDF 逐頁目視）。Triggers when user asks to
  練習題轉Word、測驗卷Word版、工作紙、講義轉Word、pdf轉Word、把這份PDF做成
  可編輯Word、掃描題目錄入Word; or gives a PDF/images of physics questions and
  wants an editable worksheet.
when_to_use:
  - "把這份練習題 PDF 做成 Word"
  - "測驗卷做成可編輯的 Word 文檔"
  - "講義轉Word"
  - "工作紙錄入 Word"
  - "pdf轉Word（題目是截圖/掃描）"
  - "把圖片上的題目做成 Word 練習卷"
---

# Worksheet → Word Skill（物理練習卷專用）

把 **PDF 或圖片形式的物理題目**（常見為截圖貼入的 Word 轉出 PDF，無文字層）
製作成**可編輯、排版規範的 Word 文檔**。核心價值：物理排版規範（變量斜體/單位正體）、
術語標準推導、插圖精確裁切、跨平台複驗。

## 適用與不適用

- ✅ 物理練習題、測驗卷、工作紙、講義（含圖的題目截圖、掃描檔）
- ❌ 一般 prose 文件（無物理符號排版需求）——本 skill 的斜體/術語規範是為物理內容設計
- ❌ 純文字層 PDF 的完美複刻（那應直接複製貼上，不必走本流程）

## 環境準備

```bash
pip install python-docx opencc-python-reimplemented pymupdf
```

- **macOS**：另需 poppler（`brew install poppler`；提供 pdfimages/pdftotext/pdfinfo）
- **Windows**：需已安裝 Microsoft Word（複驗用）；繪圖不需 poppler（PyMuPDF 取代）

## 工作流程（五階段）

### 階段 0｜診斷輸入 — 不得假設

```bash
python scripts/diagnose_pdf.py 輸入.pdf
```

- 抽出文字少 ≠ 內容少：**很多 Word 轉出的 PDF 題目全是內嵌截圖**（本 skill 存在的原因）
- 看報告判斷：`TEXT_LAYER`（直接複製文字）／`IMAGE_ONLY`（走本流程）／`MIXED`
- 輸入若是零散圖片檔（非 PDF），跳过本階段直接進階段 2

### 階段 1｜開工前確認清單（歧義必須先討論，一次問完）

1. **術語標準**（最重要）：先 grep 用戶工作區現有的講義/練習檔案推導慣用語，別憑感覺：
   ```bash
   grep -o -E "(週期|周期|圖象|圖像|回復力|回覆力|為|爲|係數|系數|向量|矢量)" 工作區/*.html | sort | uniq -c
   ```
   把結果寫入 `build_docx.py` 的 `TERM_OVERRIDES`。**已知 OpenCC 誤轉**：
   `回复力→回覆力`（應為回復力）、`周期` 在部分詞組漏轉、`為→爲`（港式應為「為」）——
   覆寫清單必含這三項的修正。
2. 簡轉繁？（預設：跟隨用戶語料標準，用 OpenCC s2t＋覆寫）
3. 字體確認（預設：中文 新細明體、英文 Times New Roman；行距 1.0；段落間距 0；題間一空行）
4. 圖內的簡體標註處理（預設：**保留原樣**——點陣圖改字得不償失）
5. 交付位置、檔名（預設：工作目錄根、原名加 `_Word版`）

### 階段 2｜轉錄與裁圖

**轉錄**：用多模態直讀圖片內容（本機通常無中文 OCR，模型直讀更準）。忠實保留原文用字
（簡體題目維持簡體，轉換在構建期統一做）。留意題目順序＝原檔視覺順序；
用 `pdftotext -bbox` 時**必須記錄每個字屬於第幾頁**（坑 #3）。

**裁圖**（題目附圖時）：
```bash
python scripts/crop_figures.py 頁面圖.png --box 左,上,右,下 --out 圖名
# 或批次：boxes.json = {"圖名": {"src": "來源.png", "box": [左,上,右,下]}, ...}
python scripts/crop_figures.py --spec boxes.json
```
- 先用 `pdfimages -png 檔.pdf 前綴` 抽出內嵌圖（macOS），或 `pymupdf`/`pdftoppm` 渲染整頁
- **每張裁圖必須目視複驗**：左緣常切到相鄰文字（把 box 左界逐步右移重裁），
  右緣常切斷曲線/標籤（右界外擴），需 2–3 輪迭代屬正常

**內容標記語言**（寫在內容字串裡）：
- `**粗體**`（題號標籤）、`{i}x{/i}` 斜體（變量、點位標籤）、`{sub}0{/sub}` 下標、`{sup}−2{/sup}` 上標
- `{p}…{/p}` 散文段落（英文/葡文整句）——豁免斜體歸類檢查；其中變量仍用 `{i}` 標記
- `{eq}…{/eq}` Word 方程式物件（OMML，行內）；`{eqd}…{/eqd}` 置中展示方程式。
  語法用 Word 線性輸入式：`^` 上標、`_` 下標、`/` 分式、`()` 分組，
  如 `{eq}GMm/R^2=m(2π/T_2)^2h{/eq}`。數學區內單字母自動斜體（OMML 行為）。
  雙擊 Word 中的方程式可進入方程式編輯器修改
- 考卷專用：`wb.page_break()`（強制分頁）、`wb.answer_box(高度cm)`（作答方框）
- 斜體規範（ISO 80000）：物理量符號斜體（x, t, F, f, k, m, M, R, L, T, s, g, h, α、點位 A/B/O/C、線段 AB、下標變量 x_A）
- 正體：單位（cm, m, s, Hz, N, r/min）、數字、函數名（sin/cos）、Δ（運算符）、選項前綴「A. B. C. D.」、下標數字（A₀ 的 0）

### 階段 3｜構建 docx

複製 `scripts/example_build.py` 為工作腳本，填入內容（參考其中的範例結構）。
`build_docx.py` 提供的預設已驗證：

- A4、邊界 2 cm、11 pt；**行距 1.0、段落前後間距 0**；題間以一個空段落分隔
- 每題設「與下段同頁」（keep_with_next）→ **圖文永不分散**（代價：Word 編輯視圖開¶時
  左邊界出現黑方塊，那只是提示符，不會列印）
- 雙字體掛載：單一字元串自動中文用新細明體、英文/數字用 Times New Roman
- **構建後自動四項檢查**（在腳本輸出中必須全 OK）：
  1. 無殘留簡體字 2. 每個拉丁/希臘字母都有斜體/正體歸類 3. 字體已寫入 4. 無非零段落間距

### 階段 4｜複驗（依 OS 分支，自動偵測）

```bash
python scripts/verify_docx.py 成品.docx --export-pdf --render
```

- **macOS**：Pages AppleScript 匯出 PDF → PyMuPDF 渲染逐頁 PNG
  （⚠️ 部分 Mac 的 Word AppleScript `save as` 故障回 -1708，**不要嘗試 Word 路徑**，直接用 Pages；
  若該機無 Pages，降級為結構檢查＋請用戶自行開啟目視）
- **Windows**：Word COM（PowerShell）匯出 PDF → PyMuPDF 渲染逐頁 PNG
- **必須逐頁目視渲染圖**（用 ReadMediaFile 工具）：題目齊全、圖在正確題目旁、斜體/上下標正確、
  無亂碼、無圖文分頁
- 複驗後原位部署（預設同名替換）

## 避雷清單（已實證的坑）

1. **`textutil -convert docx` 不嵌入圖片**（macOS 內建工具會靜默丟圖）——務必用 python-docx
2. **OpenCC 誤轉**：回复力→回覆力（應回復力）、固有周期等詞組漏轉、為→爲——覆寫清單必查
3. **`pdftotext -bbox` 不分頁碼**：grep 全部文字時會丟失頁歸屬，判斷位置前必須逐 `<page>` 元素追蹤
4. **裁圖邊界**：截圖題的文字與圖緊鄰，box 差 10–20px 就會切到字；逐張目視是唯一可靠驗證
5. **Word 黑方塊 ≠ 符號**：是「與下段同頁」的 ¶ 提示（⌘⇧8 關閉），不影響列印
6. **新細明體需目標機安裝**：未安裝的 Mac 會以系統宋體替代顯示；docx 內字型名稱正確，
   Windows/學校電腦正常。不要為此改用其他字體
7. **Pages 與 Word 分頁略有差異**：keep_with_next 的語意兩邊都遵循，圖文不分開；
   但分頁位置可能差一兩行，屬正常

## 已知限制

- 圖片內的點陣文字（圖題標註、坐標軸）無法改字體/轉繁體
- 手寫、極低解析度掃描的辨識準確度受限，轉錄後需人類校對數字與符號
- 複雜公式（分式、積分）以線性格式（3L/4、√2、10⁻²）呈現，非 OMML 方程式
