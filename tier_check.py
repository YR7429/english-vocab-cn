# -*- coding: utf-8 -*-
"""
星级判定核查

用途：检查权威表里每一条的「星级来源」是否与课标文件一致。
典型失配：损坏词被还原成真实词形后（SOITOw -> sorrow），
          该条的星级判定仍是旧的「原文星号」，于是成品里的星号会是错的。

被 build.py（构建时提示）与 recompute_tier.py（人工重算）共用，保证判定口径一致。
"""
import os, re, sys, json

ROOT  = os.path.dirname(os.path.abspath(__file__))
JSONP = os.path.join(ROOT, "step4_wordlist.json")
KCP   = os.path.join(ROOT, "考纲词汇含分级.txt")

PLACEHOLDER = ("待补", "损坏", "无法辨认")

def variants(w):
    """把 'favour (favor)'、'actor/actress'、'PE (=physical education)'、
    'tofu (bean curd)' 等拆成课标文件中可能出现的写法。
    同时归一化弯引号与连字符。

    括号里的内容也算一种写法：课标的体例是「主词（另一种拼写）」，
    而本项目的词形有时直接用了括号里的那个（例如 bean curd）。"""
    w = str(w).lower().strip()
    w = w.replace("\u2019", "'").replace("\u2018", "'")   # ’ ‘ -> '
    w = w.replace("\u2010", "-").replace("\u2011", "-").replace("\u2013", "-")
    out = {w}
    out.add(re.sub(r"\(.*?\)", "", w).strip())
    out.add(re.sub(r"=.*$", "", w).strip())
    for p in w.replace("=", "/").split("/"):
        p = p.strip()
        if p:
            out.add(p); out.add(p.replace(" ", ""))
    for m in re.finditer(r"\(([^)]*)\)", w):
        inner = m.group(1).strip()
        if inner and "=" not in inner:
            out.add(inner); out.add(inner.replace(" ", ""))
    out.add(w.replace(" ", ""))
    return {x for x in out if x}

def load_kc(path=None):
    """课标索引：写法 -> (原词, 星级)"""
    path = path or KCP
    kc = {}
    if not os.path.exists(path):
        return kc
    for i, l in enumerate(open(path, encoding="utf-8", errors="replace").read().split("\n"), 1):
        s = l.strip()
        if not s or i <= 16: continue
        if re.match(r"^[A-Z]$", s) or "年版" in s or s.startswith("("): continue
        if not re.match(r"^[A-Za-z]", s): continue
        core = re.sub(r"\*+$", "", s).strip()
        if not core: continue
        tier = s.count("*")
        for v in variants(core):
            kc.setdefault(v, (core, tier))
    return kc

def scan(data=None, kc=None):
    """返回 (待修正列表, 占位条目列表)
    待修正项：(编号, 词, 旧tier_kc, 旧tier_src, 新tier 或 None)"""
    if data is None:
        data = json.load(open(JSONP, encoding="utf-8"))
    if kc is None:
        kc = load_kc()
    changes, placeholders = [], []
    for t in data:
        w = t.get("final") or t.get("word") or ""
        if any(p in w for p in PLACEHOLDER):
            placeholders.append((t["id"], w)); continue
        hit = None
        for v in variants(w):
            if v in kc:
                hit = kc[v]; break
        if hit:
            tier = hit[1]
            if t.get("tier_kc") != tier or t.get("tier_src") != "课标":
                changes.append((t["id"], w, t.get("tier_kc"), t.get("tier_src"), tier))
        else:
            if t.get("tier_src") == "课标":
                changes.append((t["id"], w, t.get("tier_kc"), "课标", None))
    return changes, placeholders

def apply_changes(data, changes):
    """按 scan 的结果就地修改 data（不写文件）"""
    idx = {t["id"]: t for t in data}
    for eid, w, old_kc, old_src, new_tier in changes:
        t = idx.get(eid)
        if not t: continue
        if new_tier is None:
            t["tier_kc"] = None
            t["tier_src"] = "原文星号"
            t["tier_use"] = t.get("star_raw", 0)
        else:
            t["tier_kc"] = t["tier_use"] = new_tier
            t["tier_src"] = "课标"
    return data

def warn_if_needed(max_show=5):
    """供 build.py 调用：有待修正项时打印醒目提示，返回条数"""
    try:
        changes, _ = scan()
    except Exception as e:
        return 0
    if not changes:
        return 0
    bar = "=" * 68
    print("\n" + bar)
    print(f"⚠️  检测到 {len(changes)} 条「星级判定待重算」——通常是还原了损坏词却没重算星级")
    print(bar)
    for eid, w, old_kc, old_src, new_tier in changes[:max_show]:
        new = f"课标 {new_tier} 星" if new_tier is not None else "沿用原文星号（课标查不到）"
        print(f"   {eid:>5} {w:<24} 旧: {old_src}  ->  应为: {new}")
    if len(changes) > max_show:
        print(f"   ...另有 {len(changes) - max_show} 条")
    print("\n   本次构建仍会照旧渲染（成品里的星号可能不对）。")
    print("   修正请运行：  python recompute_tier.py --apply")
    print(bar + "\n")
    return len(changes)

if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    ch, ph = scan()
    print(f"待修正 {len(ch)} 处｜占位条目 {len(ph)} 处")
    for c in ch:
        print("   ", c)
    for p in ph:
        print("    [占位]", p)
