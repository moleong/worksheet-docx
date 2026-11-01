#!/usr/bin/env python3
"""example_build.py — 練習卷構建範本（worksheet-docx skill）

每次任務：複製本檔為工作腳本，替換內容區段，執行。
流程對應 SKILL.md：
  階段 0  diagnose_pdf.py 診斷輸入
  階段 1  確認術語標準（grep 語料 -> 修改 build_docx.py 的 TERM_OVERRIDES）
  階段 2  直讀圖片轉錄下列字串；crop_figures.py 裁圖並逐張目視複驗
  階段 3  執行本腳本（內含 self_check 四項自動檢查）
  階段 4  verify_docx.py 成品.docx --export-pdf --render，逐頁目視
  階段 5  原位部署
"""
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from build_docx import WorksheetBuilder   # noqa: E402

OUT = "成品.docx"

wb = WorksheetBuilder(crops_dir="crops")   # crops/ = crop_figures.py 的輸出目錄

# ---- 標題（note 可省略）----
wb.title("高一物理  第九章  機械振動")

# ---- 選擇題範例：stem 題幹、opts 選項、figs 插圖 ----
# 標記：{i}斜體變量{/i}、{sub}下標{/sub}、{sup}上標{/sup}
# 斜體：x t F f k m M R L T s g h α、點位 A B O C、線段 AB、下標變量
# 正體：單位 cm m s Hz N r/min、數字、sin、Δ、選項前綴 A. B. C. D.
wb.question(
    "例1",
    "如圖所示，將彈簧振子從平衡位置拉下一段距離 Δ{i}x{/i}，釋放後振子在 "
    "{i}A{/i}、{i}B{/i} 間振動，且 {i}AB{/i} = 20 cm，則下列說法中正確的是（  ）.",
    opts=[
        "A. 振子在振動過程中受到重力、彈簧的彈力、回復力三個力作用",
        "B. 振子做簡諧運動的振幅 {i}A{/i}{sub}0{/sub} = 20 cm",
        "C. 振子做簡諧運動的頻率 {i}f{/i} = 5 Hz",
        "D. 振子在 1.25 s 運動的路程 {i}s{/i} = 1.25 m",
    ],
    figs=[dict(file="fig_example.png", w=4.0)],   # w = 顯示寬度（cm）
)

# ---- 計算題範例：subs 小問；無圖題目省略 figs ----
wb.question(
    "5.",
    "如圖甲，{i}O{/i} 點為單擺的固定懸點。圖乙表示拉力 {i}F{/i} 隨時間 {i}t{/i} "
    "變化的曲線，{i}g{/i} 取 10 m/s{sup}2{/sup}。",
    subs=[
        "(1) 求單擺的振動週期和擺長。",
        "(2) 求擺球的質量。",
        "(3) 求擺球運動過程中的最大速度。",
    ],
    figs=[dict(file="fig_example_jia.png", w=3.7, cap="甲"),
          dict(file="fig_example_yi.png", w=8.7)],   # cap = 圖註（可省略）
)

# ---- 無圖、無選項的補充題 ----
wb.question(
    "補充：",
    "一邊長為 10 cm 的木塊，放在平靜的湖面上，{i}g{/i} = 10 m/s{sup}2{/sup}，求：",
    subs=[
        "(1) 平衡靜止時，下底面離水面的深度 {i}h{/i}{sub}0{/sub} = ?",
        "(2) 此時木塊受到的回復力 {i}F{/i} = ?",
    ],
)

# ---- 構建 + 四項自動檢查（全 OK 才進入複驗）----
wb.save(OUT)
ok = wb.self_check(OUT)
sys.exit(0 if ok else 1)
