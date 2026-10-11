# -*- coding: utf-8 -*-
"""
词义错位自检

原理：正确写的释义里，常出现「与英文词同源的中文线索」。
反向做：如果一条的释义明显在讲另一个词（例如 its 的释义是 "运转；工作"），
就说明错位。这里用一组人工整理的「词 -> 该词释义必须包含的关键字」来抽查高危词。
纯自动判定不可靠，故本工具只做**已知高危词**的核对，并输出可疑清单供人复核。
"""
import json, re, os, sys

sys.stdout.reconfigure(encoding="utf-8")
ROOT = os.path.dirname(os.path.abspath(__file__))
JSONP = os.path.join(ROOT, "step4_wordlist.json")

# 抽查表：(编号, 词, 释义中应出现的字样之一)
PROBES = [
    (1601, "its",   ["它的"]),
    (1602, "found", ["创立", "创办"]),
    (1603, "mode",  ["方式", "模式"]),
    (1618, "cat",   ["猫"]),
    (1619, "package", ["包裹", "包装"]),
    (1625, "motive", ["动机"]),
    (1636, "planet", ["行星"]),
    (1637, "weapon", ["武器"]),
    (1646, "item",  ["物品", "项目"]),
    (1651, "violin", ["小提琴"]),
    (1657, "basin", ["盆"]),
    (1662, "army",  ["军队"]),
    (1683, "ninety", ["九十"]),
    (1685, None,    ["想要", "需要", "缺乏"]),
    (1694, "ankle", ["踝"]),
    (1695, "hen",   ["母鸡"]),
    (1700, "friendship", ["友谊"]),
]

def load_defs():
    d = {}
    folder = os.path.join(ROOT, "defs")
    for fn in sorted(os.listdir(folder)):
        if not fn.endswith(".txt"): continue
        for ln in open(os.path.join(folder, fn), encoding="utf-8"):
            m = re.match(r"^(\d+)\|(.*?)\|(.*?)\|(.*)$", ln.rstrip("\n"))
            if m:
                d.setdefault(int(m.group(1)), m.group(2))
    return d

def main():
    data = json.load(open(JSONP, encoding="utf-8"))
    W = {t["id"]: (t["final"] or "") for t in data}
    D = load_defs()
    bad = []
    for eid, word, keys in PROBES:
        text = D.get(eid, "")
        if not any(k in text for k in keys):
            bad.append((eid, W.get(eid, word or "?"), text[:50], keys))
    print(f"抽查 {len(PROBES)} 处，可疑 {len(bad)} 处")
    for eid, w, text, keys in bad:
        print(f"   {eid} 词={w!r}  释义={text!r}  期望包含 {keys}")
    return len(bad)

if __name__ == "__main__":
    main()
