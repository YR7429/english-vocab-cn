# -*- coding: utf-8 -*-
"""
校验 defs/*.txt 与权威词表的一致性
用法：python check_defs.py
检查：
  1. 编号是否都在权威表里
  2. 词性符号是否规范
  3. 第2段是否为空
  4. 覆盖到哪一天、还缺哪些编号
"""
import os, re, sys, json, glob
sys.stdout.reconfigure(encoding="utf-8")
ROOT = r"D:\DSH_01\词表"
WS   = r"D:\DSH_01"

AUTH = json.load(open(os.path.join(ROOT, "step4_wordlist.json"), encoding="utf-8"))
DAY_OF = {t["id"]: t["day"] for t in AUTH}
WORD_OF = {t["id"]: t["final"] for t in AUTH}

POS_OK = re.compile(r"^(linking v\.|modal v\.|n\. \[[CU]\]&\[CU\]|n\. \[[CU]\]|n\.|"
                    r"vt\.&vi\.|vt\.|vi\.|v\.|adj\.|adv\.|conj\.|prep\.|pron\.|det\.|num\.|int\.)$")

def load():
    rows = {}
    for fn in sorted(glob.glob(os.path.join(ROOT, "defs", "*.txt"))):
        for ln in open(fn, encoding="utf-8", errors="replace").read().split("\n"):
            ln = ln.strip()
            if not ln or ln.startswith("#"): continue
            p = ln.split("|")
            if not p[0].strip().isdigit(): continue
            rows.setdefault(int(p[0]), []).append((os.path.basename(fn), p))
    return rows

rows = load()
print(f"defs 覆盖编号数: {len(rows)}")

bad_pos, empty, unknown, dup = [], [], [], []
for eid, lst in rows.items():
    if len(lst) > 1: dup.append(eid)
    if eid not in WORD_OF: unknown.append(eid)
    fn, p = lst[0]
    if len(p) < 2 or not p[1].strip(): empty.append((eid, WORD_OF.get(eid))); continue
    for chunk in p[1].split("/"):
        chunk = chunk.strip()
        if not chunk: continue
        m = re.match(r"^([^ ]+(?: [a-z]\.)?)\s+(.*)$", chunk)
        pos = m.group(1) if m else ""
        if not POS_OK.match(pos): bad_pos.append((eid, WORD_OF.get(eid), pos, chunk[:40]))

print(f"  重复编号: {len(dup)} {dup[:10]}")
print(f"  权威表外的编号: {len(unknown)} {unknown[:10]}")
print(f"  第2段为空: {len(empty)} {empty[:10]}")
print(f"  词性不规范: {len(bad_pos)}")
for x in bad_pos[:25]: print("     ", x)

# 覆盖进度
have = set(rows)
print("\n各日覆盖情况：")
for d in sorted(set(DAY_OF.values()), key=lambda z: (z != "新增", int(z.replace("Day","")) if z != "新增" else 0)):
    ids = [i for i in WORD_OF if DAY_OF[i] == d]
    n = sum(1 for i in ids if i in have)
    flag = "✓" if n == len(ids) else ("·" if n else " ")
    print(f"  {flag} {d:8s} {n:3d}/{len(ids):3d}")
