# -*- coding: utf-8 -*-
r"""
词表构建脚本（单一数据源 -> LaTeX + DOCX + PDF）

数据源：
  词表/DXX_*.py   或   词表/DXX_*.json      （每天一个文件，人工编辑只改这里）
  step4_wordlist.json                      （编号->星级，自动读取，勿手改）

输出：
  词表/tex/00_前言.tex            总表前言与表头
  词表/tex/DXX.tex                每天一张表（\input 片段）
  词表/tex/总表.tex               总表：\input 各天 + 非30天词汇
  词表/build/总表.pdf             xelatex 编译结果
  词表/docx/DXX_词汇表.docx       每天一个 Word 文件
  词表/docx/总表.docx             合并的 Word 文件

标记约定（数据源里写这些）：
  ⚠ 或 !!     无法联想的熟词生义  -> LaTeX 渲染为 \EX，DOCX 渲染为 ⚠
  {b}…{/b} 或 **…**   可联想的引申义 -> 黑体/加粗
用法：
  python build.py            # 全部构建
  python build.py D01 D02    # 只构建指定天
  python build.py --no-pdf   # 不编译 PDF
"""
import sys, os, re, json, glob, subprocess, importlib.util
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

sys.stdout.reconfigure(encoding="utf-8")
ROOT   = os.path.dirname(os.path.abspath(__file__))
WS     = os.path.dirname(ROOT)
TEXDIR = os.path.join(ROOT, "tex")
BUILD  = os.path.join(ROOT, "build")
DOCX   = os.path.join(ROOT, "docx")
for d in (TEXDIR, BUILD, DOCX): os.makedirs(d, exist_ok=True)

CN, LAT = "Microsoft YaHei", "Times New Roman"   # CN 必须有粗体字重；LAT 需要 IPA 音标字形
EX_SYM  = "！"          # 全角感叹号 U+FF01：生义标记（❗ ⚠ ▲ △ 统一归到它）

# ---------------- 星级 / 词形权威表 ----------------
def load_authority():
    p = os.path.join(ROOT, "step4_wordlist.json")
    if not os.path.exists(p): return {}, {}
    data = json.load(open(p, encoding="utf-8"))
    star, canon = {}, {}
    for t in data:
        star[t["id"]] = t["tier_use"]
        canon[t["id"]] = t["final"]
    return star, canon

STAR, CANON = load_authority()

# ---------------- 数据源读取 ----------------
def norm_pos(pos, first_def):
    """统一词性符号：n. C -> n. [C]；v. T -> vt.；v. I -> vi.；v. T,I -> vt.&vi. 等"""
    pos, defs = (pos or "").strip(), list(first_def)
    if not pos and defs:
        m = re.match(r"^((?:n\.|v\.|vt\.|vi\.|adj\.|adv\.|conj\.|prep\.|pron\.|det\.|num\.|int\.)"
                     r"\s*(?:[CU](?:\s*,\s*[CU])?|T(?:\s*,\s*I)?|I(?:\s*,\s*T)?)?)\s+(.*)$", defs[0])
        if m: pos, defs[0] = m.group(1).strip(), m.group(2).strip()
    pos = pos.replace("&", ",").replace(" ", "")
    m = re.match(r"^(n|v|vt|vi|adj|adv|conj|prep|pron|det|num|int)\.?(.*)$", pos)
    if not m: return pos, defs
    base, rest = m.group(1), m.group(2).replace("[", "").replace("]", "").replace(".", "")
    if base == "n":
        parts = [x for x in rest.split(",") if x in ("C", "U")]
        if parts: return "n. [" + "]&[".join(parts) + "]", defs
        return "n.", defs
    if base in ("v", "vt", "vi"):
        d = set(x for x in rest.replace("T", "T,").replace("I", "I,").split(",") if x in ("T", "I"))
        if base in ("vt", "v"): d.add("T")
        if base == "vi": d = {"I"}
        if d == {"T"}: return "vt.", defs
        if d == {"I"}: return "vi.", defs
        if d == {"T", "I"}: return "vt.&vi.", defs
        return "v.", defs
    return base + ".", defs

def split_defs(defs):
    """把义项列表归一成 [(pos, [义项...]), ...]"""
    groups = []
    for d in defs:
        if isinstance(d, (tuple, list)) and len(d) == 2 and isinstance(d[1], (list, tuple)):
            pos, items = d[0], list(d[1])
        else:
            s = str(d)
            m = re.match(r"^((?:n\.|vt\.|vi\.|adj\.|adv\.|conj\.|prep\.|pron\.|det\.|num\.|int\.|linking v\.)"
                         r"(?:\s*\[[CU]\]|\s*&)?)\s*(.*)$", s)
            if not m: pos, items = "", [s]
            else:     pos, items = m.group(1).strip(), [m.group(2).strip()]
        pos, items = norm_pos(pos, items)
        groups.append((pos, items))
    return groups

def normalize_mark(t):
    """统一生义标记为全角！(U+FF01)，统一粗体标记为 **…**"""
    t = str(t)
    for old in ("❗", "⚠", "▲", "△", "!!", "‼"):
        t = t.replace(old, EX_SYM)
    t = t.replace(EX_SYM + EX_SYM, EX_SYM)
    t = re.sub(r"\{b\}(.*?)\{/b\}", r"**\1**", t)
    return t.strip()

def load_day(path):
    """骨架文件只声明编号；词形取自权威表，释义取自 defs/"""
    name = os.path.basename(path)
    if path.endswith(".json"):
        raw = json.load(open(path, encoding="utf-8"))
        ids = [int(e[0]) for e in raw]
    else:
        spec = importlib.util.spec_from_file_location("mod_" + os.path.splitext(name)[0], path)
        mod = importlib.util.module_from_spec(spec); spec.loader.exec_module(mod)
        raw = getattr(mod, "ITEMS", None)
        if raw is None:
            for k in dir(mod):
                if k.startswith("D") and isinstance(getattr(mod, k), list):
                    raw = getattr(mod, k); break
        ids = [int(e[0]) for e in raw]
    items = []
    for eid in sorted(set(ids)):
        word = CANON.get(eid, "")
        if not word:            # 权威表里没有的编号直接跳过
            continue
        items.append((eid, word, [("", ["（待补释义）"])], ""))
    return apply_defs(items)

def list_sources():
    out = {}
    pats = (glob.glob(os.path.join(ROOT, "D[0-9][0-9]_*.py")) + glob.glob(os.path.join(ROOT, "Day[0-9][0-9].py"))
            + glob.glob(os.path.join(ROOT, "D[0-9][0-9]_*.json")) + glob.glob(os.path.join(ROOT, "Day[0-9][0-9].json")))
    for p in sorted(set(pats)):
        m = re.match(r"^(?:D|Day)(\d{2})", os.path.basename(p))
        if m: out["D" + m.group(1)] = p
    return out

# ---------------- 释义文件（defs/*.txt） ----------------
# 格式：编号|词性+义项|变化|备注
#   词性+义项：多组用 / 分隔；同组内义项用 ； 分隔
#   变化：留空则由 irregular.py 自动补不规则形式
#   备注：[课标无记录] 等
POS_TOKEN = re.compile(r"^(linking v\.|modal v\.|aux\. v\.|(?:vt\.|vi\.)&(?:vt\.|vi\.)|n\.\s*\[[CU]\]&\[[CU]\]|"
                       r"n\.\s*\[[CU]\]|n\.|vt\.|vi\.|v\.|adj\.|adv\.|conj\.|prep\.|pron\.|det\.|num\.|int\.)\s*(.*)$")

def parse_posdefs(text):
    """'n. [C] 政府 / n. [U] **管理、统治**' -> [(pos,[defs]), ...]"""
    out = []
    for chunk in text.split("/"):
        chunk = chunk.strip()
        if not chunk: continue
        m = POS_TOKEN.match(chunk)
        if m:
            pos = re.sub(r"\s+", " ", m.group(1)).strip().replace("vi.&vt.", "vt.&vi.").replace("aux. v.", "modal v.")
            pos = pos.replace("]", "]").replace("[ ", "[")
        else:
            pos, chunk = "", chunk
            m2 = POS_TOKEN.match(chunk)
            if m2: pos, chunk = m2.group(1).strip(), m2.group(2)
        rest = (m.group(2) if m else chunk).strip()
        out.append((pos, split_by_semicolon(rest)))
    return out

def split_by_semicolon(text):
    """按全角分号切分义项，但不在 **粗体** 标记内部切分
    （否则 '**（衣服）穿上；继续**' 会被切成两半，标记失效）"""
    out, buf, in_bold = [], [], False
    i = 0
    while i < len(text):
        if text.startswith("**", i):
            in_bold = not in_bold
            buf.append("**"); i += 2; continue
        ch = text[i]
        if ch == "；" and not in_bold:
            out.append("".join(buf)); buf = []
        else:
            buf.append(ch)
        i += 1
    out.append("".join(buf))
    return [x.strip() for x in out if x.strip()]

def load_defs():
    """读 defs/*.txt，返回 {编号: (posdefs, change, note)}"""
    d = {}
    folder = os.path.join(ROOT, "defs")
    if not os.path.isdir(folder): return d
    for fn in sorted(glob.glob(os.path.join(folder, "*.txt"))):
        for ln in open(fn, encoding="utf-8", errors="replace").read().split("\n"):
            ln = ln.strip()
            if not ln or ln.startswith("#"): continue
            parts = ln.split("|")
            if len(parts) < 2: continue
            try: eid = int(parts[0].strip())
            except ValueError: continue
            posdefs = parse_posdefs(parts[1].strip())
            chg = normalize_mark(parts[2]) if len(parts) > 2 and parts[2].strip() else ""
            note = parts[3].strip() if len(parts) > 3 else ""
            d[eid] = (posdefs, chg, note)
    return d

# ---------------- 课标无记录判定（以权威表为准，覆盖子代理的判断） ----------------
KCMISS = set()
def load_kcmiss():
    p = os.path.join(ROOT, "step4_wordlist.json")
    if not os.path.exists(p): return set()
    return {t["id"] for t in json.load(open(p, encoding="utf-8")) if t.get("tier_src") != "课标"}
KCMISS = load_kcmiss()

DEFS = None
def apply_defs(items):
    """用释义文件覆盖/补齐 ITEMS；词形与星级始终以权威表为准"""
    global DEFS
    if DEFS is None: DEFS = load_defs()
    out = []
    for eid, word, defs, chg in items:
        if eid not in CANON:          # 过滤权威表里不存在的编号
            continue
        if eid in DEFS:
            pd, ch, note = DEFS[eid]
            chg = ch or auto_change(CANON.get(eid, word))
        else:
            pd, note = defs, ""
            chg = chg or auto_change(CANON.get(eid, word))
        # 「课标无记录」一律以权威表为准，并且只出现一次
        chg = re.sub(r"\[课标无记录\]", "", chg).strip("； ;")
        note = "[课标无记录]" if eid in KCMISS else ""
        out.append((eid, word, pd, chg, note))
    out.sort(key=lambda x: x[0])          # 全文统一按编号升序
    return out

def auto_change(word):
    """由不规则引擎补"变化"列"""
    try:
        import irregular
        return "；".join(irregular.forms(word))
    except Exception:
        return ""

# ---------------- 文本渲染 ----------------
def tex_escape(s):
    for a, b in [("\\", r"\textbackslash{}"), ("&", r"\&"), ("%", r"\%"), ("$", r"\$"),
                 ("#", r"\#"), ("_", r"\_"), ("{", r"\{"), ("}", r"\}"),
                 ("~", r"\textasciitilde{}"), ("^", r"\textasciicircum{}")]:
        s = s.replace(a, b)
    return s

def to_tex(s):
    r"""**粗体** -> \B{…}；生义标记(！/!!) -> \EX；其余转义"""
    parts = re.split(r"(\*\*.*?\*\*)", s)
    out = []
    for p in parts:
        if p.startswith("**") and p.endswith("**"):
            out.append(r"\B{" + tex_escape(p[2:-2]) + "}")
        else:
            out.append(tex_escape(p))
    return "".join(out).replace("!!", r"\EX{}").replace(EX_SYM, r"\EX{}")

def to_docx_runs(p, s, size, gray=None):
    """**粗体** -> 加粗 run；生义标记 -> 全角！"""
    for part in re.split(r"(\*\*.*?\*\*)", s):
        if not part: continue
        if part.startswith("**") and part.endswith("**"):
            r = p.add_run(part[2:-2]); _font(r, size, True)
        else:
            for seg in re.split(r"(!!|" + re.escape(EX_SYM) + r")", part):
                if not seg: continue
                r = p.add_run(EX_SYM if seg in ("!!", EX_SYM) else seg); _font(r, size, False)

def _font(run, size=None, bold=None):
    run.font.name = LAT
    run._element.rPr.rFonts.set(qn("w:eastAsia"), CN)
    if size is not None: run.font.size = Pt(size)
    if bold is not None: run.bold = bold

def star_mark(eid):
    return {2: "**", 1: "*"}.get(STAR.get(eid, 0), "")

# ---------------- LaTeX ----------------
PREAMBLE = r"""% 本文件由 build.py 自动生成。手改会在下次构建时被覆盖。
% 要改内容请编辑 词表/DXX_*.py 或 .json（单一数据源）。
\documentclass[10pt]{ctexart}
\usepackage[a4paper,top=1.3cm,bottom=1.1cm,left=1.2cm,right=1.2cm]{geometry}
\usepackage{fontspec}
\usepackage{xeCJK}
\usepackage{array}
\usepackage{longtable}
\usepackage{booktabs}
\usepackage[table]{xcolor}
\usepackage{fancyhdr}
\setCJKmainfont{Microsoft YaHei}
\setmainfont{Times New Roman}
\setlength{\parindent}{0pt}
\renewcommand{\arraystretch}{1.12}
% 行间细实线，防止读串行（不用隔行灰底：\rowcolor 会覆盖 \hline）
\arrayrulecolor{black}
\setlength{\arrayrulewidth}{0.3pt}
\newcolumntype{N}{>{\raggedleft\arraybackslash}p{0.85cm}}
\newcolumntype{W}{>{\raggedright\arraybackslash}p{2.85cm}}
\newcolumntype{D}{>{\raggedright\arraybackslash}p{9.5cm}}
\newcolumntype{C}{>{\raggedright\arraybackslash}p{3.4cm}}
\newcommand{\B}[1]{\textbf{#1}}
% 生义标记：全角感叹号
\newcommand{\EX}{\textbf{！}}
\newcommand{\hd}[1]{\textbf{#1}}
\newcommand{\zebra}{\rowcolor{gray!12}}
\pagestyle{fancy}\fancyhf{}
\fancyfoot[C]{\small 第 \thepage\ 页}
\renewcommand{\headrulewidth}{0pt}
\begin{document}
"""

TABLE_HEAD = (r"\begin{longtable}{N W D C}" "\n"
              r"\toprule" "\n"
              r"\hd{编号} & \hd{原词} & \hd{义项} & \hd{变化} \\" "\n"
              r"\midrule" "\n"
              r"\endfirsthead" "\n"
              r"\toprule" "\n"
              r"\hd{编号} & \hd{原词} & \hd{义项} & \hd{变化} \\" "\n"
              r"\midrule" "\n"
              r"\endhead" "\n")
STRAY_BOLD = []   # 记录残留 ** 的条目，构建结束时报出
TABLE_TAIL = "\\bottomrule\n\\end{longtable}\n"

def day_tex(day, title, items):
    L = [f"% ==== {title} ====\n"]
    L.append(r"\vspace{2mm}{\large\bfseries " + tex_escape(title) + r"}\par\vspace{1.5mm}")
    L.append(TABLE_HEAD)
    for i, (eid, word, defs, chg, note) in enumerate(items):
        w = tex_escape(word + star_mark(eid))
        if note: chg = (chg + " " if chg else "") + note
        d = r" \newline ".join(r"\hd{" + tex_escape(p) + r"}" + (" " + "；".join(to_tex(x) for x in xs) if p else " ".join(to_tex(x) for x in xs))
                               for p, xs in defs)
        c = to_tex(chg)
        # 自检：转换后不应残留 ** 标记（通常是粗体标记被分隔符切断）
        if "**" in d or "**" in c:
            STRAY_BOLD.append((eid, word))
        # PDF 版靠 \hline 细线分隔；不用 \rowcolor（会覆盖横线）
        L.append(f"{eid} & \\textbf{{{w}}} & {d} & {c} \\\\ \\hline")
    L.append(TABLE_TAIL)
    return "\n".join(L) + "\n"

def write_prelude():
    p = os.path.join(TEXDIR, "00_前言.tex")
    open(p, "w", encoding="utf-8").write(
        r"\begin{center}" "\n"
        r"{\LARGE\bfseries 考纲词汇详解总表}\\[2mm]" "\n"
        r"{\small 星级：\textbf{**} 最重要\quad\textbf{*} 次之\quad 无星号 最弱\quad|\quad"
        r"\EX{} 无法联想的熟词生义\quad|\quad \textbf{黑体} 可联想的引申义\quad|\quad"
        r"[课标无记录] 课标文件中查不到该词}" "\n"
        r"\end{center}" "\n"
        r"\vspace{2mm}" "\n")
    return p

def write_master(days, extra_day=None):
    p = os.path.join(TEXDIR, "总表.tex")
    L = [PREAMBLE, r"\input{00_前言}", "\n"]
    for d in days:
        L.append(r"\input{" + d + "}")
        L.append(r"\clearpage" if d != days[-1] else "")
    if extra_day:
        L.append(r"\clearpage")
        L.append(r"\input{" + extra_day + "}")
    L.append(r"\end{document}")
    open(p, "w", encoding="utf-8").write("\n".join(L))

# ---------------- DOCX ----------------
def build_docx(title, items, path, merged_note=""):
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    sec.top_margin, sec.bottom_margin = Cm(1.3), Cm(1.1)
    sec.left_margin, sec.right_margin = Cm(1.2), Cm(1.2)
    st = doc.styles["Normal"]; st.font.name = LAT; st.font.size = Pt(9)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), CN); st.paragraph_format.space_after = Pt(0)

    fp = sec.footer.paragraphs[0]; fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = fp.add_run(title + " · 第 "); _font(r, 8)
    fld = OxmlElement("w:fldSimple"); fld.set(qn("w:instr"), "PAGE"); fp._p.append(fld)
    r = fp.add_run(" 页"); _font(r, 8)

    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(1)
    r = p.add_run(title); _font(r, 14, True)
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(4)
    _font(p.add_run("词性：n. [C] 可数 / n. [U] 不可数 / vt. 及物 / vi. 不及物 / linking v. 连系动词　｜　"
                    + EX_SYM + " 无法联想的熟词生义　｜　"), 8)
    _font(p.add_run("黑体"), 8, True)
    _font(p.add_run(" 可联想的引申义　｜　[课标无记录] 课标文件中查不到该词"), 8)

    W = [0.85, 2.85, 9.5, 3.4]
    t = doc.add_table(rows=1, cols=4)
    t.style = "Table Grid"; t.alignment = WD_TABLE_ALIGNMENT.CENTER; t.autofit = False
    mar = OxmlElement("w:tblCellMar")
    for tag, val in (("top",12),("left",60),("bottom",12),("right",60)):
        e = OxmlElement(f"w:{tag}"); e.set(qn("w:w"), str(val)); e.set(qn("w:type"), "dxa"); mar.append(e)
    t._tbl.tblPr.append(mar)
    trPr = t.rows[0]._tr.get_or_add_trPr()
    th = OxmlElement("w:tblHeader"); th.set(qn("w:val"), "true"); trPr.append(th)
    for i, h in enumerate(["编号", "原词", "义项", "变化"]):
        cell = t.rows[0].cells[i]
        q = cell.paragraphs[0]; q.alignment = WD_ALIGN_PARAGRAPH.CENTER
        q.paragraph_format.space_before = Pt(0); q.paragraph_format.space_after = Pt(0)
        q.paragraph_format.line_spacing = Pt(11)
        _font(q.add_run(h), 9.5, True)
        sh = OxmlElement("w:shd"); sh.set(qn("w:val"),"clear"); sh.set(qn("w:fill"),"E6E6E6")
        cell._tc.get_or_add_tcPr().append(sh)

    for idx, (eid, word, defs, chg, note) in enumerate(items):
        if note: chg = (chg + " " if chg else "") + note
        row = t.add_row().cells
        if idx % 2 == 1:                       # 隔行浅灰，防止读串行
            for c in row:
                sh = OxmlElement("w:shd"); sh.set(qn("w:val"), "clear"); sh.set(qn("w:fill"), "F2F2F2")
                c._tc.get_or_add_tcPr().append(sh)
        q = row[0].paragraphs[0]; q.alignment = WD_ALIGN_PARAGRAPH.CENTER
        q.paragraph_format.space_before = Pt(0); q.paragraph_format.space_after = Pt(0); q.paragraph_format.line_spacing = Pt(12)
        _font(q.add_run(str(eid)), 8.5)
        q = row[1].paragraphs[0]; q.paragraph_format.space_before = Pt(0); q.paragraph_format.space_after = Pt(0); q.paragraph_format.line_spacing = Pt(12)
        _font(q.add_run(word + star_mark(eid)), 9.5, True)
        q = row[2].paragraphs[0]; q.paragraph_format.space_before = Pt(0); q.paragraph_format.space_after = Pt(0); q.paragraph_format.line_spacing = Pt(12)
        first = True
        for pos, xs in defs:
            if not first: q.add_run().add_break()
            first = False
            if pos: _font(q.add_run(pos + " "), 9.5)
            to_docx_runs(q, "；".join(xs), 9.5)
        q = row[3].paragraphs[0]; q.paragraph_format.space_before = Pt(0); q.paragraph_format.space_after = Pt(0); q.paragraph_format.line_spacing = Pt(12)
        if chg: to_docx_runs(q, chg, 8.5)
    for row in t.rows:
        for i, w in enumerate(W): row.cells[i].width = Cm(w)
    # grid
    tbl = t._tbl
    old = tbl.find(qn("w:tblGrid"))
    if old is not None: tbl.remove(old)
    grid = OxmlElement("w:tblGrid")
    for w in W:
        gc = OxmlElement("w:gridCol"); gc.set(qn("w:w"), str(int(round(w*567)))); grid.append(gc)
    tbl.tblPr.addnext(grid)
    tw = OxmlElement("w:tblW"); tw.set(qn("w:w"), str(int(round(sum(W)*567)))); tw.set(qn("w:type"),"dxa")
    tbl.tblPr.append(tw)
    ly = OxmlElement("w:tblLayout"); ly.set(qn("w:type"), "fixed"); tbl.tblPr.append(ly)
    doc.save(path)

# ---------------- 主流程 ----------------
def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    no_pdf = "--no-pdf" in sys.argv

    # 构建前提示：是否有条目的星级判定待重算（只提示，不中断）
    try:
        import tier_check
        tier_check.warn_if_needed()
    except Exception:
        pass

    srcs = list_sources()
    if not srcs:
        print("没有找到数据源文件（词表/DXX_*.py 或 .json）"); return
    keys = [k for k in sorted(srcs) if not args or k in args]
    days = []
    built = []
    for k in keys:
        items = load_day(srcs[k])
        id0, id1 = items[0][0], items[-1][0]
        title = "新增词汇" if k == "D31" else f"{k} · 考纲第 {id0}-{id1} 词"
        # tex 片段
        p = os.path.join(TEXDIR, k + ".tex")
        open(p, "w", encoding="utf-8").write(day_tex(k, title, items))
        # docx
        build_docx(title, items, os.path.join(DOCX, f"{k}_词汇表.docx"))
        days.append(k); built.append((k, title, len(items)))
        print(f"  {k}: {len(items)} 词 -> tex/{k}.tex, docx/{k}_词汇表.docx")

    write_prelude()
    write_master(days)
    print(f"  总表: tex/总表.tex （\\input {len(days)} 个片段）")

    # 合并 DOCX
    allitems = []
    for k in days: allitems += load_day(srcs[k])
    build_docx("考纲词汇详解总表（Day " + "–".join([days[0][1:], days[-1][1:]]) + "）",
               allitems, os.path.join(DOCX, "总表.docx"))
    print(f"  合并 DOCX: docx/总表.docx （{len(allitems)} 词）")

    if STRAY_BOLD:
        bar = "=" * 68
        print("\n" + bar)
        print(f"⚠️  检测到 {len(STRAY_BOLD)} 条存在「粗体标记未闭合」——通常是 ** 被分隔符切断")
        print(bar)
        for eid, w in STRAY_BOLD[:10]:
            print(f"   {eid:>5} {w}")
        if len(STRAY_BOLD) > 10: print(f"   ...另有 {len(STRAY_BOLD)-10} 条")
        print("\n   修正：检查 defs/ 里该条的第2段，确认 ** 成对且不被 ；或 / 分开")
        print(bar + "\n")

    if no_pdf: return
    print("  编译 PDF ...")
    r = subprocess.run([r"D:\texlive\2025\bin\windows\xelatex.exe",
                        "-interaction=nonstopmode", "-halt-on-error", "总表.tex"],
                       cwd=TEXDIR, capture_output=True, text=True, errors="replace")
    if r.returncode != 0:
        log = (r.stdout or "")[-2500:]
        print("  PDF 编译失败，日志尾部：\n" + log)
    else:
        out = os.path.join(TEXDIR, "总表.pdf")
        if os.path.exists(out):
            import shutil; shutil.copy2(out, os.path.join(BUILD, "总表.pdf"))
            print("  PDF: build/总表.pdf")

if __name__ == "__main__":
    main()
