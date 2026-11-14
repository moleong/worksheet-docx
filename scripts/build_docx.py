#!/usr/bin/env python3
"""build_docx.py — 物理練習卷 Word 構建函式庫（worksheet-docx skill）

用法：複製 example_build.py，填入題目內容後執行。
本模組提供：預設排版、標記語言解析、簡轉繁、四項自動檢查。

內容標記語言（寫在題目字串中）：
    **粗體**          題號標籤
    {i}x{/i}          斜體（變量、點位標籤 A/B/O、線段 AB、下標變量）
    {sub}0{/sub}      下標（下標數字保持正體，如 A₀ 寫 {i}A{/i}{sub}0{/sub}）
    {sup}−2{/sup}     上標（如 10⁻² 寫 10{sup}−2{/sup}）
    可巢狀：{i}x{/i}{sub}{i}A{/i}{/sub} → x_A（x 與 A 皆斜體）

斜體規範（ISO 80000）：物理量符號斜體；單位、數字、sin/cos、Δ、
選項前綴「A. B. C. D.」保持正體。
"""
import re
import zipfile

from lxml import etree
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

# ---------------- 預設值（依用戶慣用語修改 TERM_OVERRIDES） ----------------
DEFAULTS = dict(
    cn_font="新細明體",       # 中文字體（需目標機安裝）
    en_font="Times New Roman",  # 英文/數字字體
    font_size=11,           # 內文字級（pt）
    title_size=18,
    caption_size=10,
    margin_cm=2.0,
)

# OpenCC s2t 之後的慣用語覆寫——開工前 grep 用戶語料推導（見 SKILL.md 階段 1）
TERM_OVERRIDES = [
    ("周期", "週期"),      # OpenCC 在部分詞組（如 固有周期）漏轉
    ("回覆力", "回復力"),   # OpenCC 誤轉，物理術語為「復」
    ("爲", "為"),          # 港澳慣用「為」，OpenCC 預設轉「爲」
]

# 正體白名單：這些拉丁/希臘 token 允許不斜體（單位、函數、運算符）
UPRIGHT_TOKENS = {
    "sin", "cos", "tan", "ln", "log",
    "cm", "mm", "m", "km", "s", "ms", "min", "h",
    "Hz", "kHz", "N", "kg", "g", "J", "W", "Pa",
    "V", "A", "Ω", "r", "mol", "K", "T", "Δ", "π",
}

# s2t 檢查的已知誤報：這些字在繁體亦正確（OpenCC 仍視為簡體映射）
S2T_FALSE_POSITIVES = {
    "皂",   # 肥皂（泡）：繁體正字
    "干",   # 若干：繁體正字（OpenCC 視為「乾/幹」簡化映射）
    "游",   # 上游/游泳：繁體正字（OpenCC 視為「遊」簡化映射）
}

try:
    from opencc import OpenCC
    _s2t = OpenCC("s2t")
except ImportError:
    _s2t = None


def t2tw(text):
    """簡體轉繁體＋用戶慣用語覆寫。不需要轉換時原樣回傳。"""
    if _s2t is None:
        return text
    t = _s2t.convert(text)
    for src, dst in TERM_OVERRIDES:
        t = t.replace(src, dst)
    return t


# ---------------- 標記語言解析 ----------------
_TAG = re.compile(r"(\*\*|\{/?i\}|\{/?sub\}|\{/?sup\}|\{/?p\}|\{/?eq\}|\{/?eqd\})")


# ---------------- 線性數學語法 -> OMML（Word 方程式） ----------------
# {eq}GMm/R^2 = m(2π/T_2)^2 h{/eq}  →  行內 Word 方程式物件
# {eqd}M = 4π^2h^3/GT_2^2{/eqd}     →  置中展示方程式
# 語法：^ 上標、_ 下標、/ 分式、() 分組、字母/數字/運算符直寫
#
# 已實證的寫法陷阱（A1/A2 兩輪對照實驗，違者渲染必錯）：
#   1. 多字元下標必須加括號：v_(0y)、x_(CM)、F_(外1)
#      （裸寫 v_0y 只會把 0 放下標，y 留在基線）
#   2. 分式後面還有因子時必須分組：(1/2) gt^2
#      （裸寫 1/2 gt^2 會整串被吞進分母）
#   3. 根號：√ 是前綴運算子，作用範圍是「下一個 atom 及其上下標」：
#      √(2gR)、√x^2（= √(x^2)）；根號內多項式一律加括號
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"


def _m(tag):
    return f"{{{M_NS}}}{tag}"


class _MathOMML:
    """小型的線性數學語法 -> OMML 產生器。"""

    TOK = re.compile(
        r"[A-Za-zΑ-Ωα-ωℓ⊥∞′″…⋮→←∝一-鿿；：:≈><≥≤\u0300-\u036f]+|\d+(?:\.\d+)?|[()^_/=+\-·×±∓√]")

    def __init__(self, text):
        # ASCII 撇號在數學語境一律視為撇號（prime）：v' -> v′（U+2032，已在收錄範圍）
        text = text.replace("'", "′")
        self.toks = self.TOK.findall(text)
        self.pos = 0
        # 防呆：token 之外的字元會被靜默丟棄（A2 整理時發現 ±∓ 被吃掉），
        # 必須大聲警告而不是默默產生缺字的方程式
        covered = set()
        for m in self.TOK.finditer(text):
            covered.update(range(m.start(), m.end()))
        dropped = [ch for i, ch in enumerate(text)
                   if i not in covered and not ch.isspace()]
        if dropped:
            import sys
            print(f"WARNING 方程式有未識別字元被丟棄：{dropped} ← {text!r}",
                  file=sys.stderr)

    def peek(self):
        return self.toks[self.pos] if self.pos < len(self.toks) else None

    def next(self):
        t = self.peek()
        self.pos += 1
        return t

    # ---- 語法（優先級：根號 > 上下標 > 分式 > 乘除並置 > 關係/加減運算符）----
    # expr    := product (('='|'+'|'-'|'−'|'±'|'∓'|'·'|'×'|'∝') product)*
    # product := term ('/' term)*        分式左結合；並置因子相乘
    # term    := atom (^_ 後綴)*
    # atom    := √ atom (^_ 後綴)* | 字|數|(expr)
    _BINOPS = ("=", "+", "-", "−", "±", "∓", "·", "×", "∝")

    def parse(self):
        return self._expr()

    def _expr(self):
        nodes = self._product()
        while self.peek() in self._BINOPS:
            nodes.append(self._run(self.next()))
            nodes.extend(self._product())
        return nodes

    def _product(self):
        # 並置因子相乘；'/' 分式：左、右各取一個完整乘積，左結合
        left = self._factors()
        while True:
            t = self.peek()
            if t == "/":
                self.next()
                left = [self._frac(left, self._factors())]
            elif t is not None and t not in (")",) + self._BINOPS:
                left.append(self._term())
            else:
                break
        return left

    def _factors(self):
        nodes = []
        while True:
            t = self.peek()
            if t is None or t in (")", "/") + self._BINOPS:
                break
            nodes.append(self._term())
        return nodes

    def _term(self):
        node = self._atom()
        while self.peek() in ("^", "_"):
            if self.next() == "^":
                node = self._ssup(node, self._atom())
            else:
                sub = self._atom()
                if self.peek() == "^":           # _ 後接 ^ → 上下標
                    self.next()
                    node = self._ssubsup(node, sub, self._atom())
                else:
                    node = self._ssub(node, sub)
        return node

    def _atom(self):
        t = self.peek()
        if t == "√":
            self.next()
            inner = self._atom()
            # 根號內容的上下標歸根號管：√x^2 = √(x^2)
            while self.peek() in ("^", "_"):
                if self.next() == "^":
                    inner = self._ssup(inner, self._atom())
                else:
                    sub = self._atom()
                    if self.peek() == "^":
                        self.next()
                        inner = self._ssubsup(inner, sub, self._atom())
                    else:
                        inner = self._ssub(inner, sub)
            return self._rad(inner)
        if t == "(":
            self.next()
            inner = self._expr()
            if self.peek() == ")":
                self.next()
            return self._delim(inner)
        self.next()
        return self._run(t)

    # ---- OMML 節點工廠（回傳 lxml 元素）----
    @staticmethod
    def _el(tag, children=None):
        e = etree.Element(_m(tag))
        for c in (children or []):
            e.append(c)
        return e

    @classmethod
    def _run(cls, text):
        r = etree.Element(_m("r"))
        t = etree.SubElement(r, _m("t"))
        t.text = text
        return r

    @classmethod
    def _wrap(cls, tag, nodes):
        e = cls._el(tag)
        for n in nodes:
            e.append(n)
        return e

    @classmethod
    def _frac(cls, num, den):
        f = cls._el("f")
        cls._el("fPr")  # 無屬性即預設分式
        num_e = cls._wrap("num", num)
        den_e = cls._wrap("den", den)
        f.append(num_e)
        f.append(den_e)
        return f

    @classmethod
    def _ssup(cls, base, sup):
        e = cls._el("sSup")
        e.append(cls._wrap("e", base if isinstance(base, list) else [base]))
        e.append(cls._wrap("sup", sup if isinstance(sup, list) else [sup]))
        return e

    @classmethod
    def _ssub(cls, base, sub):
        e = cls._el("sSub")
        e.append(cls._wrap("e", base if isinstance(base, list) else [base]))
        e.append(cls._wrap("sub", sub if isinstance(sub, list) else [sub]))
        return e

    @classmethod
    def _ssubsup(cls, base, sub, sup):
        e = cls._el("sSubSup")
        e.append(cls._wrap("e", base if isinstance(base, list) else [base]))
        e.append(cls._wrap("sub", sub if isinstance(sub, list) else [sub]))
        e.append(cls._wrap("sup", sup if isinstance(sup, list) else [sup]))
        return e

    @classmethod
    def _delim(cls, inner):
        # 明確寫出 begChr/endChr：省略 dPr 雖合規，但 LibreOffice 會渲染成 |x|
        e = cls._el("d")
        dpr = cls._el("dPr")
        beg = cls._el("begChr")
        beg.set(_m("val"), "(")
        end = cls._el("endChr")
        end.set(_m("val"), ")")
        dpr.append(beg)
        dpr.append(end)
        e.append(dpr)
        e.append(cls._wrap("e", inner))
        return e

    @classmethod
    def _rad(cls, inner):
        # 平方根：Word 的標準寫法是 radPr/degHide + 空 m:deg + m:e。
        # 省略 m:deg 雖合規，但 LibreOffice 的 OMML 匯入會錯位（內容變方塊）
        r = cls._el("rad")
        rpr = cls._el("radPr")
        dh = cls._el("degHide")
        dh.set(_m("val"), "1")
        rpr.append(dh)
        r.append(rpr)
        r.append(cls._el("deg"))
        r.append(cls._wrap("e", inner if isinstance(inner, list) else [inner]))
        return r


def linear_math_to_omml(text, display=False):
    """線性數學字串 -> m:oMath（display=True 時包 m:oMathPara 置中）。"""
    nodes = _MathOMML(text).parse()
    omath = etree.Element(_m("oMath"))
    for n in nodes:
        omath.append(n)
    if not display:
        return omath
    para = etree.Element(_m("oMathPara"))
    para.append(omath)
    return para


def omml_to_linear(omath):
    """OMML 結構樹 -> 線性字串（linear_math_to_omml 的逆向，結構複驗用）。

    逐條比對 lin(linear_math_to_omml("v_(0y)")) 與預期字串，
    可抓到下標/分式分組等結構錯誤。
    ⚠️ 只驗證「結構正確」，不保證 Word 實際渲染正確——Quick Look 不渲染
    OMML，最終品質閘門永遠是匯出 PDF 逐頁目視（A1 的教訓：結構全對、
    渲染系統性出錯，只靠用戶匯出的 PDF 才抓到）。
    """
    def lin(e):
        if e is None:
            return ""
        tag = etree.QName(e).localname
        if tag == "r":
            return e.findtext(_m("t")) or ""
        if tag == "f":
            return f"({lin(e.find(_m('num')))})/({lin(e.find(_m('den')))})"
        if tag == "sSub":
            return f"{lin(e.find(_m('e')))}_({lin(e.find(_m('sub')))})"
        if tag == "sSup":
            return f"{lin(e.find(_m('e')))}^({lin(e.find(_m('sup')))})"
        if tag == "sSubSup":
            return (f"{lin(e.find(_m('e')))}_({lin(e.find(_m('sub')))})"
                    f"^({lin(e.find(_m('sup')))})")
        if tag == "d":
            return f"({lin(e.find(_m('e')))})"
        if tag == "rad":
            return f"√({lin(e.find(_m('e')))})"
        # num/den/e/sub/sup/oMath/oMathPara 及其他：串接子節點
        return "".join(lin(c) for c in e)

    return lin(omath)


def list_docx_equations(docx_path):
    """列出成品 docx 中所有 OMML 方程式的線性形式，供逐條與原題核對。

    用法：build 腳本在 self_check 之後呼叫，或互動式檢查已交付的 docx。
    同時是「同型錯掃描」的入口：發現一條渲染錯，先懷疑解析規則有洞
    （見 _MathOMML 開頭的寫法陷阱），再從本清單全文掃同型寫法——
    錯誤是系統性的，不是隨機的。
    """
    with zipfile.ZipFile(docx_path) as z:
        root = etree.fromstring(z.read("word/document.xml"))
    lines = [omml_to_linear(om) for om in root.iter(_m("oMath"))]
    for i, s in enumerate(lines, 1):
        print(f"{i:3d}. {s}")
    print(f"共 {len(lines)} 條方程式")
    return lines


def parse_markup(text):
    """解析標記語言 -> [(chunk, state)]，state = {b, i, sub, sup, prose, eq}"""
    out = []
    stack = [dict(b=False, i=False, sub=False, sup=False, prose=False)]
    for part in _TAG.split(text):
        if not part:
            continue
        if part == "**":
            top = stack[-1]
            stack.append(dict(top, b=not top["b"]))
        elif part in ("{/i}", "{/sub}", "{/sup}", "{/p}", "{/eq}", "{/eqd}"):
            if len(stack) > 1:
                stack.pop()
        elif part == "{i}":
            stack.append(dict(stack[-1], i=True))
        elif part == "{sub}":
            stack.append(dict(stack[-1], sub=True))
        elif part == "{sup}":
            stack.append(dict(stack[-1], sup=True))
        elif part == "{p}":
            stack.append(dict(stack[-1], prose=True))
        elif part == "{eq}":
            stack.append(dict(stack[-1], eq="inline"))
        elif part == "{eqd}":
            stack.append(dict(stack[-1], eq="display"))
        else:
            out.append((part, dict(stack[-1])))
    return out


# ---------------- 文件構建 ----------------
class WorksheetBuilder:
    """用法：
        wb = WorksheetBuilder()                  # 套用 DEFAULTS
        wb.title("高一物理 第一章 直線運動")
        wb.question("例1", "如圖所示，… {i}x{/i} = 5 {i}m{/i} …",
                    opts=["A. …", "B. …"], figs=[dict(file="fig_a.png", w=5.0)])
        wb.save("成品.docx")
        wb.self_check()                           # 四項自動檢查
    """

    def __init__(self, crops_dir="crops", upright_extra=(), **overrides):
        self.cfg = dict(DEFAULTS)
        self.cfg.update(overrides)
        self.upright_tokens = UPRIGHT_TOKENS | set(upright_extra)
        self.crops_dir = crops_dir
        self._source_strings = []          # 供斜體歸類檢查
        doc = Document()
        sec = doc.sections[0]
        sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)   # A4
        m = Cm(self.cfg["margin_cm"])
        sec.top_margin = sec.bottom_margin = m
        sec.left_margin = sec.right_margin = m
        normal = doc.styles["Normal"]
        normal.font.name = self.cfg["en_font"]
        normal.font.size = Pt(self.cfg["font_size"])
        normal.element.rPr.rFonts.set(qn("w:eastAsia"), self.cfg["cn_font"])
        self.doc = doc

    def _set_fonts(self, run):
        run.font.name = self.cfg["en_font"]
        rpr = run._element.get_or_add_rPr()
        rf = rpr.find(qn("w:rFonts"))
        if rf is None:
            rf = rpr.makeelement(qn("w:rFonts"), {})
            rpr.append(rf)
        rf.set(qn("w:eastAsia"), self.cfg["cn_font"])

    def _add_runs(self, p, text, size):
        has_content = lambda q: bool(q.runs) or bool(
            q._p.findall("{http://schemas.openxmlformats.org/officeDocument/2006/math}oMath"))
        for chunk, st in parse_markup(text):
            if st.get("eq") == "display":
                # 展示方程式：內容為 oMathPara，必須獨立成段（前置/後續文字各成一段）
                if has_content(p):
                    p = self._para(align="center")
                p._p.append(linear_math_to_omml(chunk, display=True))
                p = self._para()
                continue
            if st.get("eq"):
                p._p.append(linear_math_to_omml(chunk, display=False))
                continue
            chunk = t2tw(chunk)
            if not chunk:
                continue
            r = p.add_run(chunk)
            r.bold = st["b"] or None
            r.italic = st["i"] or None
            if st["sub"]:
                r.font.subscript = True
            if st["sup"]:
                r.font.superscript = True
            r.font.size = Pt(size)
            self._set_fonts(r)
        return p

    def _para(self, text="", size=None, align=None, indent_cm=None, keep_next=False):
        p = self.doc.add_paragraph()
        pf = p.paragraph_format
        pf.space_before = Pt(0)             # 段落間距 0
        pf.space_after = Pt(0)
        pf.line_spacing = 1.0               # 行距 1.0
        pf.keep_with_next = keep_next       # 圖文不分散
        if align == "center":
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if indent_cm:
            pf.left_indent = Cm(indent_cm)
        if text:
            p = self._add_runs(p, text, size or self.cfg["font_size"])
        return p

    # ---- 對外 API ----
    def title(self, text, note=None):
        self._para(text, size=self.cfg["title_size"], align="center")
        if note:
            self._para(note, size=self.cfg["caption_size"], align="center")
        self._para()                        # 標題後一空行

    def blank(self):
        self._para()

    def picture(self, fname, width_cm, caption=None):
        p = self._para(align="center", keep_next=True)
        p.add_run().add_picture(
            f"{self.crops_dir}/{fname}", width=Cm(width_cm))
        if caption:
            # 圖註不設 keep_with_next：否則圖+註會把後續段落整串拖走，
            # 留下整頁空白（A2-02 實證）。圖片本身 keep_next 已足以
            # 保證「圖在、註在」（註緊跟圖之後）
            self._para(caption, size=self.cfg["caption_size"],
                       align="center")

    def question(self, label, stem, opts=None, subs=None, figs=None):
        """figs: [dict(file="xxx.png", w=寬cm, cap="甲" 或 None)]"""
        self._source_strings.append(stem)
        self._para(f"**{label}**  {stem}", keep_next=True)
        for o in (opts or []):
            self._source_strings.append(o)
            self._para(o, indent_cm=0.8, keep_next=True)
        for s in (subs or []):
            self._source_strings.append(s)
            self._para(s, indent_cm=0.5, keep_next=True)
        for f in (figs or []):
            self.picture(f["file"], f["w"], f.get("cap"))
        self._para()                        # 題間一空行

    def page_break(self):
        """分頁（考卷每題一頁等場景）。"""
        from docx.enum.text import WD_BREAK
        p = self._para()
        p.add_run().add_break(WD_BREAK.PAGE)

    def answer_box(self, height_cm=10, width_cm=None):
        """作答方框：單格有框表格（考卷用）。height_cm 為框高度。"""
        from docx.enum.table import WD_ROW_HEIGHT_RULE
        from docx.oxml import OxmlElement
        width = Cm(width_cm or 17.0)
        table = self.doc.add_table(rows=1, cols=1)
        table.autofit = False
        table.columns[0].width = width
        cell = table.cell(0, 0)
        cell.width = width
        row = table.rows[0]
        row.height = Cm(height_cm)
        row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
        tblPr = table._tbl.tblPr
        borders = OxmlElement("w:tblBorders")
        for edge in ("top", "left", "bottom", "right"):
            el = OxmlElement(f"w:{edge}")
            el.set(qn("w:val"), "single")
            el.set(qn("w:sz"), "8")
            el.set(qn("w:color"), "000000")
            borders.append(el)
        tblPr.append(borders)
        self._para()   # 框後一空行

    def save(self, path):
        self.doc.save(path)
        print(f"saved {path}")
        return path

    # ---------------- 自動檢查（必須全數 OK 才可交付） ----------------
    def self_check(self, docx_path=None):
        full = "".join(p.text for p in self.doc.paragraphs)
        ok = True

        # 1. 無殘留簡體字
        if _s2t is not None:
            leftover = sorted({ch for ch in full if _s2t.convert(ch) != ch}
                              - S2T_FALSE_POSITIVES)
            print(("OK  " if not leftover else "FAIL") +
                  f" 簡體殘留：{leftover if leftover else '無'}")
            ok &= not leftover

            # 2. 每個拉丁/希臘字母都有斜體/正體歸類（{p} 散文段落豁免檢查；
            #    散文中的變量仍應用 {i} 標記，prose 只豁免未標記即違規的判定）
            bad = []
            for s in self._source_strings:
                for chunk, st in parse_markup(s):
                    if st.get("prose") or st.get("eq"):
                        continue   # 散文與方程式（OMML 自動斜體）豁免
                    c = t2tw(chunk)
                    for m in re.finditer(r"[A-Za-zΑ-Ωα-ω]+", c):
                        tok = m.group(0)
                        if st["i"] or tok in self.upright_tokens:
                            continue
                        rest = c[m.start():]
                        if re.match(r"^[A-D]\.", rest):
                            continue   # 選項前綴 A.
                        if c[m.start() - 1:m.start()] == "(":
                            continue   # 小題標記 (a) (b)
                        bad.append((tok, c[:40]))
        print(("OK  " if not bad else "FAIL") +
              f" 斜體歸類：{'無未歸類字母' if not bad else bad}")
        ok &= not bad

        # 3. 字體寫入 document.xml
        if docx_path:
            with zipfile.ZipFile(docx_path) as z:
                xmls = z.read("word/document.xml").decode("utf-8")
            for font in (self.cfg["en_font"], self.cfg["cn_font"]):
                hit = font in xmls
                print(("OK  " if hit else "FAIL") + f" 字體 {font} 已寫入")
                ok &= hit
            n_it = xmls.count("<w:i/>")
            print(f"     斜體 run 數：{n_it}")

        # 4. 無非零段落間距
        if docx_path:
            nonzero = re.findall(r'w:(?:before|after)="(?!0")[^"]+"', xmls)
            print(("OK  " if not nonzero else "FAIL") +
                  f" 段落間距：{'全部為 0' if not nonzero else nonzero}")
            ok &= not nonzero

        print("self_check:", "ALL PASS" if ok else "HAS FAILURES")
        return ok
