# -*- coding: utf-8 -*-
r"""
词表构建脚本（单一数据源 -> LaTeX + DOCX + PDF）

数据源：
  词表/days/DayXX.py                       骨架：每天含哪些编号（自动生成，勿手改）
  step4_wordlist.json                      编号 -> 词形、星级（自动读取，勿手改）
  defs/*.txt                               释义、变化、备注

输出：
  词表/tex/00_前言.tex            总表前言与表头
  词表/tex/DayXX.tex              每天一张表（\input 片段）
  词表/tex/总表.tex               总表：\input 各天（按编号升序）
  词表/tex/总表_字母序.tex         同样内容，全局按字母表排序
  词表/docx/DayXX_词汇表.docx     每天一个 Word 文件
  词表/docx/总表.docx             合并的 Word（按编号升序）
  词表/docx/总表_字母序.docx       合并的 Word（全局按字母表排序）
  词表/build/总表.pdf             xelatex 编译结果
  词表/build/总表_字母序.pdf       字母序版本

标记约定（数据源里写这些）：
  ！ 或 !!     无法联想的熟词生义  -> LaTeX 渲染为 \EX
  {b}…{/b} 或 **…**   可联想的引申义 -> 黑体/加粗
用法：
  python build.py                # 全部构建
  python build.py D01 D02        # 只构建指定天（不会动总表）
  python build.py --no-pdf       # 不编译 PDF
  python build.py --alpha-only   # 只出字母序总表（核验词形、星级用）
"""
import sys, os, re, json, glob, subprocess, importlib.util
import ids_util
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

# ---------------- 字体配置（见 fonts.json） ----------------
# CN 必须有粗体字重；LAT 需要 IPA 音标字形。列表按序尝试，取第一个已安装的。
FONT_WANT = {"cjk": ["Microsoft YaHei", "SimSun"], "latin": ["Times New Roman", "DejaVu Serif"]}
try:
    with open(os.path.join(ROOT, "fonts.json"), encoding="utf-8") as _f:
        _cfg = json.load(_f)
    for _k in ("cjk", "latin"):
        if _cfg.get(_k): FONT_WANT[_k] = [str(x) for x in _cfg[_k]]
except Exception:
    pass                       # 没有 fonts.json 也能构建，用上面内置的默认值

CN, LAT = None, None           # 由 pick_fonts() 在构建时确定
INSTALLED = set()              # 本机字体名缓存
EX_SYM  = "！"          # 全角感叹号 U+FF01：生义标记（❗ ⚠ ▲ △ 统一归到它）

def _add_font_names(names, s):
    """把字体名规范化后加入集合。
    注册表里形如 'Microsoft YaHei & Microsoft YaHei UI (TrueType)'，
    需去掉 '(TrueType)' 后缀、按 ' & ' 拆开；字体文件主名（msyh 等）则跳过。"""
    s = re.sub(r"\s*\((?:TrueType|OpenType|All res|VGA res|Plotter|PostScript[^)]*)\)\s*$", "", str(s))
    for part in re.split(r"\s*&\s*", s):
        part = part.strip()
        if part and not re.fullmatch(r"[0-9a-fA-F]{4,}|[A-Za-z]{0,3}\d+[A-Za-z]?\d*", part):
            names.add(re.sub(r"\s+", " ", part))

def installed_fonts():
    """列出本机已安装的字体名（Windows 读注册表，类 Unix 用 fc-list）"""
    names = set()
    try:
        if os.name == "nt":
            import winreg
            for hive, key in ((winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"),
                              (winreg.HKEY_CURRENT_USER,  r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts")):
                try:
                    with winreg.OpenKey(hive, key) as k:
                        n = winreg.QueryInfoKey(k)[1]
                        for i in range(n):
                            v = winreg.EnumValue(k, i)[0]      # 名称在 [0]，[1] 是文件名
                            if isinstance(v, str): _add_font_names(names, v)
                except OSError:
                    pass
            # 注册表读不到时，退回扫描字体目录（文件名不一定等于族名，仅作兜底）
            if not names:
                for d in (os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts"),
                          os.path.join(os.environ.get("LOCALAPPDATA", ""), r"Microsoft\Windows\Fonts")):
                    if os.path.isdir(d):
                        for fn in os.listdir(d):
                            _add_font_names(names, os.path.splitext(fn)[0])
        else:
            r = subprocess.run(["fc-list", "--format", "%{family}\n"], capture_output=True, text=True, errors="replace")
            for ln in (r.stdout or "").splitlines():
                for part in ln.split(","): names.add(part.strip())
    except Exception:
        pass
    return names

# DOCX 里写的字体名只是「建议」——Word 打开时会用本机同名字体，找不到就自动替换。
# 所以这里固定用最通用的写法，不跟随 fonts.json。
DOCX_CN  = "Noto Sans CJK SC" if os.name != "nt" else "Microsoft YaHei"
DOCX_LAT = "Times New Roman"

def pick_fonts():
    """从候选列表里挑出本机可用的中文字体与拉丁字体；挑不到就返回 None（跳过 PDF）"""
    global CN, LAT, INSTALLED
    INSTALLED = have = installed_fonts()
    def pick(cands):
        for want in cands:
            if want in have: return want           # 族名精确命中
        for want in cands:                          # 再放宽：去空格、前缀匹配
            w = want.replace(" ", "").lower()
            for hn in have:
                if hn.replace(" ", "").lower().startswith(w): return want
        for want in cands:                          # 最后：子串匹配
            for hn in have:
                if want.lower() in hn.lower(): return want
        return None
    CN, LAT = pick(FONT_WANT["cjk"]), pick(FONT_WANT["latin"])
    # 允许用环境变量临时顶替，便于在未改配置的机器上先编一遍
    if os.environ.get("FONT_CJK"): CN = os.environ["FONT_CJK"]
    if os.environ.get("FONT_LAT"): LAT = os.environ["FONT_LAT"]
    return CN, LAT

def find_texlive():
    """定位 xelatex：先看环境变量，再看 PATH，最后扫常见安装目录。找不到返回 None"""
    env = os.environ.get("XELATEX")
    if env and os.path.exists(env): return env
    import shutil as _sh
    p = _sh.which("xelatex")
    if p: return p
    pats = [r"D:\texlive\*\bin\windows\xelatex.exe", r"C:\texlive\*\bin\windows\xelatex.exe",
            "/usr/bin/xelatex", "/usr/local/bin/xelatex", "/Library/TeX/texbin/xelatex",
            "/opt/homebrew/bin/xelatex", "/usr/local/texlive/*/bin/*/xelatex",
            os.path.expanduser("~/Library/TinyTeX/bin/*/xelatex")]
    for pat in pats:
        for hit in sorted(glob.glob(pat) + glob.glob(os.path.expanduser(pat)), reverse=True):
            if os.path.exists(hit): return hit
    return None

XELATEX = None                 # 由 find_texlive() 在构建时填入

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
    else:
        spec = importlib.util.spec_from_file_location("mod_" + os.path.splitext(name)[0], path)
        mod = importlib.util.module_from_spec(spec); spec.loader.exec_module(mod)
        raw = getattr(mod, "ITEMS", None)
        if raw is None:
            for k in dir(mod):
                if k.startswith("D") and isinstance(getattr(mod, k), list):
                    raw = getattr(mod, k); break
    ids = [ids_util.parse(e[0]) for e in raw]
    items = []
    for eid in sorted({i for i in ids if i is not None}, key=ids_util.key):
        word = CANON.get(eid, "")
        if not word:            # 权威表里没有的编号直接跳过
            continue
        items.append((eid, word, [("", ["（待补释义）"])], ""))
    return apply_defs(items)

def list_sources():
    """骨架文件统一放在 days/ 下（也兼容根目录下的旧式 DXX_*.py）"""
    out = {}
    pats = []
    for base in (os.path.join(ROOT, "days"), ROOT):
        for pat in ("Day[0-9][0-9].py", "Day[0-9][0-9].json",
                    "D[0-9][0-9]_*.py", "D[0-9][0-9]_*.json"):
            pats += glob.glob(os.path.join(base, pat))
    for p in sorted(set(pats)):
        m = re.match(r"^(?:D|Day)(\d{2})", os.path.basename(p))
        if m: out["D" + m.group(1)] = p
    return out

# ---------------- 释义文件（defs/*.txt） ----------------
# 格式：编号|词性+义项|变化或备注|        （恰好 4 段、3 个竖线）
#   词性+义项：多组用 / 分隔；同组内义项用 ； 分隔
#   第3段：不规则变化，或搭配/易混/读音等备注。留空则由 irregular.py 自动补不规则形式
#   第4段：一律留空（旧文件里写在第4段的备注会被自动并入第3段）
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
    """读 defs/*.txt，返回 {编号: (posdefs, change, note)}
    第4段（末段）已废弃：若仍有内容，自动并入第3段，保证两种写法都能正常渲染。"""
    d = {}
    folder = os.path.join(ROOT, "defs")
    if not os.path.isdir(folder): return d
    for fn in sorted(glob.glob(os.path.join(folder, "*.txt"))):
        for ln in open(fn, encoding="utf-8", errors="replace").read().split("\n"):
            ln = ln.strip()
            if not ln or ln.startswith("#"): continue
            parts = ln.split("|")
            if len(parts) < 2: continue
            eid = ids_util.parse(parts[0].strip())
            if eid is None: continue
            posdefs = parse_posdefs(parts[1].strip())
            chg = normalize_mark(parts[2]) if len(parts) > 2 and parts[2].strip() else ""
            tail = parts[3].strip() if len(parts) > 3 else ""
            if tail:                     # 兼容旧写法：末段有内容则并到「变化」之后
                chg = (chg + "；" if chg else "") + tail
            d[eid] = (posdefs, chg, "")
    return d

# ---------------- 排序键 ----------------
def alpha_key(word):
    """字母序键：忽略大小写与空格/连字符，非字母字符不参与比较。
    因此 'by accident'、'byaccident'、'By accident' 会排在一起。"""
    s = re.sub(r"[^0-9a-z]+", "", str(word).lower())
    return (s, str(word).lower())

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
        out.append((eid, word, pd, chg, note))
    out.sort(key=lambda x: ids_util.key(x[0]))   # 编号升序，子编号紧随主编号之后
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
    run.font.name = DOCX_LAT
    run._element.rPr.rFonts.set(qn("w:eastAsia"), DOCX_CN)
    if size is not None: run.font.size = Pt(size)
    if bold is not None: run.bold = bold

def star_mark(eid):
    return {2: "**", 1: "*"}.get(STAR.get(eid, 0), "")

# ---------------- 字体自检（挑不到字体时给出可操作的提示） ----------------
def warn_missing_glyphs():
    tips = []
    if not CN:
        tips.append(("中文字体", FONT_WANT["cjk"],
                     "没有粗体字重的中文字体会让 \\textbf{中文} 编译失败"))
    if not LAT:
        tips.append(("拉丁与音标字体", FONT_WANT["latin"],
                     "缺 IPA 字形时音标（ə ʊ ɪ ɔː）会渲染成空白"))
    if not tips: return
    print("\n" + "=" * 68)
    print("  缺少可用字体 —— 无法保证 PDF 排版正确")
    print("=" * 68)
    for kind, cands, why in tips:
        print(f"  {kind}：候选 {', '.join(cands)} 全部未找到")
        print(f"      ({why})")
    have = sorted(INSTALLED or installed_fonts())
    if have:
        print(f"\n  本机共识别到 {len(have)} 个字体。请从中挑一个，填进 fonts.json 的对应列表。")
        shown = " / ".join(have[:40])
        print("      " + shown + (" …" if len(have) > 40 else ""))
        if len(have) > 40:
            print("      （另有更多；完整列表可在 PowerShell 里看："
                  "Get-ItemProperty 'HKLM:/SOFTWARE/Microsoft/Windows NT/CurrentVersion/Fonts'）")
        print("      挑中文字体时注意：必须有粗体字重（黑体、宋体、雅黑、Noto Sans CJK 均可）。")
    else:
        print("\n  （未能读取本机字体列表；可直接在 fonts.json 里试填字体名）")
    print("\n  临时替代（不改 fonts.json 也能编出 PDF，但排版未按本表调优）：")
    print("      set XELATEX=...        # 指定 xelatex 路径")
    print("      set FONT_CJK=SimSun    # 临时指定中文字体")
    print("      set FONT_LAT=DejaVu Serif   # 临时指定拉丁字体")
    print("\n  也可以先不生成 PDF：python build.py --no-pdf")
    print("  生成的 docx/ 不受字体影响，Word 里用什么字体由打开它的电脑决定。")
    print("=" * 68 + "\n")

# ---------------- LaTeX ----------------
def preamble():
    cn = CN or (FONT_WANT["cjk"][0] if FONT_WANT["cjk"] else "SimSun")
    lat = LAT or (FONT_WANT["latin"][0] if FONT_WANT["latin"] else "DejaVu Serif")
    return r"""% 本文件由 build.py 自动生成。手改会在下次构建时被覆盖。
% 要改内容请编辑 词表/defs/*.txt、词形表.txt 或 DayXX.py（单一数据源）。
% 字体在 fonts.json 里配置；此处按该文件挑出的本机可用字体写入：
%   中文 = """ + cn + "\n" + r"""%   拉丁 = """ + lat + "\n" + r"""\documentclass[10pt]{ctexart}
\usepackage[a4paper,top=1.3cm,bottom=1.1cm,left=1.2cm,right=1.2cm]{geometry}
\usepackage{fontspec}
\usepackage{xeCJK}
\usepackage{array}
\usepackage{longtable}
\usepackage{booktabs}
\usepackage[table]{xcolor}
\usepackage{fancyhdr}
\setCJKmainfont{""" + cn + r"""}
\setCJKmonofont{""" + cn + r"""}
\setmainfont{""" + lat + r"""}
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

def day_tex(day, title, items, sort=None):
    """sort=None 时按传入顺序（编号升序）；sort='alpha' 时按字母表排序"""
    if sort == "alpha":
        items = sorted(items, key=lambda x: alpha_key(x[1]))
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

def write_prelude(path=None, alpha=False):
    p = path or os.path.join(TEXDIR, "00_前言.tex")
    head = (r"{\LARGE\bfseries 考纲词汇详解总表}" if not alpha
            else r"{\LARGE\bfseries 考纲词汇详解总表（字母序）}")
    open(p, "w", encoding="utf-8", newline="\n").write(
        r"\begin{center}" "\n"
        + head + "\\\\[2mm]\n"
        r"{\small 星级：\textbf{**} 最重要\quad\textbf{*} 次之\quad 无星号 最弱\quad|\quad"
        r"\EX{} 无法联想的熟词生义\quad|\quad \textbf{黑体} 可联想的引申义}" "\n"
        r"\end{center}" "\n"
        r"\vspace{2mm}" "\n")
    return p

def write_master(days, path=None, alpha=False, prelude="00_前言"):
    """总表：只引用各天的片段，内容不重复维护"""
    p = path or os.path.join(TEXDIR, "总表.tex")
    L = [preamble(), r"\input{" + prelude + "}", "\n"]
    for d in days:
        L.append(r"\input{" + d + "}")
        L.append(r"\clearpage" if d != days[-1] else "")
    L.append(r"\end{document}")
    open(p, "w", encoding="utf-8", newline="\n").write("\n".join(L))
    return p

# ---------------- DOCX ----------------
def build_docx(title, items, path, sort=None, merged_note=""):
    """sort=None 按传入顺序；sort='alpha' 按字母表排序"""
    if sort == "alpha":
        items = sorted(items, key=lambda x: alpha_key(x[1]))
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    sec.top_margin, sec.bottom_margin = Cm(1.3), Cm(1.1)
    sec.left_margin, sec.right_margin = Cm(1.2), Cm(1.2)
    st = doc.styles["Normal"]; st.font.name = DOCX_LAT; st.font.size = Pt(9)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), DOCX_CN); st.paragraph_format.space_after = Pt(0)

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
    _font(p.add_run(" 可联想的引申义"), 8)

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

# ---------------- PDF 编译（可选步骤，失败不影响 tex/ 与 docx/ 的产出） ----------------
def compile_pdf(tex_name, out_name):
    """在 tex/ 目录编译 tex_name，成功后把 PDF 复制到 build/out_name"""
    global XELATEX
    if XELATEX is None:
        XELATEX = find_texlive()
    if not XELATEX:
        return None
    print(f"  编译 {tex_name} ... （{XELATEX}）")
    try:
        r = subprocess.run([XELATEX, "-interaction=nonstopmode", "-halt-on-error", tex_name],
                           cwd=TEXDIR, capture_output=True, text=True, errors="replace",
                           timeout=1800)
    except Exception as exc:
        print(f"  PDF 编译未能执行：{exc}")
        return None
    out = os.path.join(TEXDIR, os.path.splitext(tex_name)[0] + ".pdf")
    if os.path.exists(out) and (r.returncode == 0 or "Output written" in (r.stdout or "")):
        import shutil; shutil.copy2(out, os.path.join(BUILD, out_name))
        print(f"  PDF: build/{out_name}")
        return os.path.join(BUILD, out_name)
    print(f"  {tex_name} 编译失败，日志尾部：\n" + (r.stdout or "")[-2000:])
    print("  （tex/ 与 docx/ 已正常生成，可先使用；字体问题请检查 fonts.json）")
    return None

def no_tex_hint():
    print("\n" + "=" * 68)
    print("  未找到 xelatex —— 已跳过 PDF 生成（其他产物不受影响）")
    print("=" * 68)
    print("  · Word 文件已生成：docx/ 目录，可直接使用")
    print("  · LaTeX 源文件已生成：tex/ 目录，装了 TeX Live 的机器可直接编译")
    print("  · 通常不必自己装 TeX Live（完整安装约需数 GB）：")
    print("      成品 PDF 随版本发布提供，见仓库 Releases 页面")
    print("  · 若确实要自行编译，装 TeX Live 后重跑 python build.py 即可；")
    print("      也可用环境变量指定路径：XELATEX=<xelatex 的完整路径>")
    print("=" * 68 + "\n")

# ---------------- 主流程 ----------------
def main():
    argv   = sys.argv[1:]
    no_pdf = "--no-pdf" in argv
    only_alpha = "--alpha-only" in argv
    no_alpha   = "--no-alpha" in argv
    args = [a for a in argv if not a.startswith("--")]

    # 构建前提示：是否有条目的星级判定待重算（只提示，不中断）
    try:
        import tier_check
        tier_check.warn_if_needed()
    except Exception:
        pass

    srcs = list_sources()
    if not srcs:
        print("没有找到骨架文件（应为 词表/days/Day01.py … Day31.py）"); return

    # 先定字体：tex/ 源文件里要写入具体字体名（供日后编译或他人接手）
    pick_fonts()
    if CN or LAT:
        print(f"  字体：中文 {CN or '（未找到，见 fonts.json）'} ｜ 拉丁 {LAT or '（未找到，见 fonts.json）'}")

    keys  = [k for k in sorted(srcs) if not args or k in args]
    days, allitems = [], []
    for k in keys:
        items = load_day(srcs[k])
        id0, id1 = items[0][0], items[-1][0]
        title = "新增词汇" if k == "D31" else f"{k} · 考纲第 {id0}-{id1} 词"
        allitems += items
        open(os.path.join(TEXDIR, k + ".tex"), "w", encoding="utf-8", newline="\n") \
            .write(day_tex(k, title, items))
        build_docx(title, items, os.path.join(DOCX, f"{k}_词汇表.docx"))
        days.append(k)
        print(f"  {k}: {len(items)} 词 -> tex/{k}.tex, docx/{k}_词汇表.docx")

    # 总表只在「全量构建」时生成。若只是重建某几天（python build.py D04），
    # 只更新那几天的片段，否则会把总表覆盖成残缺版本。
    partial = bool(args) and not only_alpha
    if partial:
        print(f"\n  只重建了 {', '.join(days)}；tex/总表.tex 与 docx/总表.docx 未改动。")
        print("  要更新总表（合并全部 31 天），请运行：python build.py")
    else:
        span = "Day " + "–".join([days[0][1:], days[-1][1:]])
        if not only_alpha:
            write_prelude()
            write_master(days)
            print(f"  总表: tex/总表.tex （\\input {len(days)} 个片段，按编号升序）")
            build_docx("考纲词汇详解总表（" + span + "）", allitems,
                       os.path.join(DOCX, "总表.docx"))
            print(f"  合并 DOCX: docx/总表.docx （{len(allitems)} 词）")
        if not no_alpha:
            # 字母序版本：同样是 \input 引用，只是各片段内部改成按字母表排序
            write_prelude(os.path.join(TEXDIR, "00_前言_字母序.tex"), alpha=True)
            for k in days:
                items = load_day(srcs[k])
                id0, id1 = items[0][0], items[-1][0]
                title = "新增词汇" if k == "D31" else f"{k} · 考纲第 {id0}-{id1} 词"
                open(os.path.join(TEXDIR, "A_" + k + ".tex"), "w", encoding="utf-8", newline="\n") \
                    .write(day_tex(k, title, items, sort="alpha"))
            write_master(["A_" + k for k in days],
                         path=os.path.join(TEXDIR, "总表_字母序.tex"),
                         alpha=True, prelude="00_前言_字母序")
            print(f"  字母序总表: tex/总表_字母序.tex （{len(allitems)} 词，全局 A→Z）")
            build_docx("考纲词汇详解总表 · 字母序（" + span + "）", allitems,
                       os.path.join(DOCX, "总表_字母序.docx"), sort="alpha")
            print(f"  合并 DOCX: docx/总表_字母序.docx （{len(allitems)} 词）")

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

    if no_pdf:
        print("\n  （--no-pdf：已跳过 PDF 编译。tex/ 与 docx/ 均已生成。）")
        return

    if partial:
        print("\n  （只重建了部分天，未编译 PDF。要出成品请运行：python build.py）")
        return

    global XELATEX
    XELATEX = find_texlive()
    if not XELATEX:
        no_tex_hint()
        return
    warn_missing_glyphs()
    if not only_alpha:
        compile_pdf("总表.tex", "总表.pdf")
    if not no_alpha:
        compile_pdf("总表_字母序.tex", "总表_字母序.pdf")

if __name__ == "__main__":
    main()
