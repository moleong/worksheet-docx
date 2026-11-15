#!/usr/bin/env python3
"""test_math_omml.py — _MathOMML 解析器回歸測試（worksheet-docx skill）

用法：
    python scripts/test_math_omml.py

每次修改 build_docx.py 的數學語法後必跑。斷言每條線性標記經
linear_math_to_omml -> omml_to_linear 往返後得到預期結構。
案例來源：已實證的寫法陷阱（A1/A2/校際 2022）＋ 全部已交付文件的方程式。

⚠️ 預期字串是「結構線性化」格式，不是原始標記：
   分式 -> (num)/(den)，上下標 -> x^(sup)、x_(sub)，根號 -> √(x)，
   下標內容會多一層括號（delim）：ε_0 -> ε_(0)
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from build_docx import linear_math_to_omml, omml_to_linear

# (原始標記, 預期線性化結構)
CASES = [
    # ---- 已交付文件的全部方程式（行為不得改變） ----
    ("V=Q/(4πε_0 r)", "V=(Q)/((4πε_(0)r))"),
    ("V_a=q/(πε_0 r)", "V_(a)=(q)/((πε_(0)r))"),
    ("0.1894ρa^2/ε_0", "(0.1894ρa^(2))/(ε_(0))"),
    ("0.126σa/ε_0", "(0.126σa)/(ε_(0))"),
    ("R=A+BP", "R=A+BP"),
    ("f=0.47hv/(a^2 (1-ρ^2)^(1/2))",
     "f=(0.47hv)/((a^(2)(1-ρ^(2))^(((1)/(2)))))"),
    ("δR/R", "(δR)/(R)"),
    ("V_x=2I_x+8I_x^2", "V_(x)=2I_(x)+8I_(x)^(2)"),
    # SKILL.md 文件範例
    ("GMm/R^2=m(2π/T_2)^2h", "(GMm)/(R^(2))=m((2π)/(T_(2)))^(2)h"),
    ("M = 4π^2h^3/GT_2^2", "M=(4π^(2)h^(3))/(GT_(2)^(2))"),
    # ---- 2022 校際卷第 3 題事故案例（治本目標） ----
    # 舊行為：(1)/((4πε_0)Q)/(r)（Q 被吞進分母）；
    # 新行為：分子 = 已累積乘積 = (1/(4πε₀))·Q，即 (((1/(4πε₀))·Q)/r)
    ("1/(4πε_0) Q/r", "((1)/((4πε_(0)))Q)/(r)"),
    # 等價寫法（前綴 V= 不影響結構）
    ("V = 1/(4πε_0) Q/r", "V=((1)/((4πε_(0)))Q)/(r)"),
    # 分組後接因子再接分式：括號保住，不再被吞
    ("(1/2) gt^2/h", "(((1)/(2))gt^(2))/(h)"),
    # ---- 已實證陷阱 ----
    ("v_(0y)", "v_((0y))"),                     # 多字元下標必須加括號
    ("x_(CM)", "x_((CM))"),
    ("√(2gR)", "√((2gR))"),                    # 根號多項式加括號
    ("√x^2", "√(x^(2))"),                      # 根號吃掉自己的上下標
    ("1/2 gt^2", "(1)/(2gt^(2))"),             # Word 慣例：裸分母向右延伸
    ("a/b/c", "((a)/(b))/(c)"),                # 左結合
    ("10^(-2)", "10^((-2))"),                  # 負指數：推薦寫法（帶括號）
    ("F=ma", "F=ma"),
    # ---- A1 ch10 / PA-H101 實戰案例 ----
    # 括號分式後接括號乘因子：因子是乘數不是分母（原書排版並排確認）
    ("U(r_1)-U(r) = GMm/(r·r_1) (r-r_1) = (-GMm/r_1) - (-GMm/r)",
     "U(r_(1))-U(r)=(GMm)/((r·r_(1)))(r-r_(1))=(-(GMm)/(r_(1)))-(-(GMm)/(r))"),
    # U+2212 減號（PA-H101 曾因此被靜默丟字）：
    ("(4000−323−1000)/55.6 = 48.1", "((4000−323−1000))/(55.6)=48.1"),
    ("x−y", "x−y"),
    # ---- 運算元符號吸收（10^-2 事故案例：舊行為 10^(-)·2，2 掉回基線）----
    ("10^-2", "10^(-2)"),
    ("x_-1", "x_(-1)"),
    ("x^−1", "x^(−1)"),                       # U+2212 變體
    ("√-2", "√(-2)"),
    ("e^-x", "e^(-x)"),
    ("a^+b", "a^(+b)"),
    ("-GMm/r^2", "-(GMm)/(r^(2))"),          # 前置一元符號行為不變（run 串接）
    # ---- 裝飾函數 m:acc（Pages + LibreOffice 渲染 spike 驗證）----
    ("vec(v)", "vec((v))"),
    ("vec(AB)", "vec((AB))"),
    ("hat(F)", "hat((F))"),
    ("dot(a)", "dot((a))"),
    ("bar(x)", "bar((x))"),
    ("vec(v)_(0)", "vec((v))_((0))"),         # 裝飾後接後綴
    ("2vec(AB)·vec(CD)", "2vec((AB))·vec((CD))"),
]

# 運算元守衛：這些寫法必須印 WARNING（結構垃圾，禁止靜默通過）
# （∝ 與字母同類 token，走字面值路徑不產生結構垃圾，不在此列——見 SKILL.md 已知限制）
WARN_CASES = ["10^=2", "x^/y", "a^×b", "x^_y"]


def run():
    import contextlib
    import io

    fails = 0
    for src, expected in CASES:
        got = omml_to_linear(linear_math_to_omml(src))
        if got != expected:
            fails += 1
            print(f"FAIL  {src!r}\n  預期 {expected}\n  實得 {got}")
        else:
            print(f"ok    {src!r} -> {got}")

    for src in WARN_CASES:
        err = io.StringIO()
        with contextlib.redirect_stderr(err):
            omml_to_linear(linear_math_to_omml(src))
        if "WARNING" in err.getvalue():
            print(f"ok    {src!r} 觸發守衛 WARNING")
        else:
            fails += 1
            print(f"FAIL  {src!r} 未觸發守衛 WARNING（結構垃圾靜默通過）")

    total = len(CASES) + len(WARN_CASES)
    print(f"\n{total - fails}/{total} PASS")
    return fails


if __name__ == "__main__":
    sys.exit(1 if run() else 0)
