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
ROOT = os.path.dirname(os.path.abspath(__file__))   # 本脚本所在目录，换机器无需修改

AUTH = json.load(open(os.path.join(ROOT, "step4_wordlist.json"), encoding="utf-8"))
DAY_OF = {t["id"]: t["day"] for t in AUTH}
WORD_OF = {t["id"]: t["final"] for t in AUTH}

# 允许的词性符号。与 build.py 的 POS_TOKEN / 修订指南第四节 保持一致：
#   vi.&vt. 与 vt.&vi. 等价；aux. v. 与 modal v. 等价（build.py 会归一化）
POS_OK = re.compile(r"^(linking v\.|modal v\.|aux\. v\.|"
                    r"n\. \[[CU]\]&\[CU\]|n\. \[[CU]\]|n\.|"
                    r"vt\.&vi\.|vi\.&vt\.|vt\.|vi\.|v\.|"
                    r"adj\.|adv\.|conj\.|prep\.|pron\.|det\.|num\.|int\.)$")

def load():
    """返回 {编号: [(文件名, 行号, 分段), ...]}；同时记录分段数异常的行"""
    rows, badfmt = {}, []
    for fn in sorted(glob.glob(os.path.join(ROOT, "defs", "*.txt"))):
        base = os.path.basename(fn)
        for i, ln in enumerate(open(fn, encoding="utf-8", errors="replace"), 1):
            ln = ln.rstrip("\n").strip()
            if not ln or ln.startswith("#"): continue
            p = ln.split("|")
            if not p[0].strip().isdigit(): continue
            # 每行应恰好 3 个竖线 -> 4 段；多了少了都会让解析出错位
            if len(p) != 4: badfmt.append((base, i, ln.count("|"), ln[:60]))
            rows.setdefault(int(p[0]), []).append((base, i, p))
    return rows, badfmt

rows, badfmt = load()
print(f"defs 覆盖编号数: {len(rows)}")

bad_pos, empty, unknown, dup = [], [], [], []
for eid, lst in rows.items():
    if len(lst) > 1: dup.append(eid)
    if eid not in WORD_OF: unknown.append(eid)
    fn, lineno, p = lst[0]
    if len(p) < 2 or not p[1].strip(): empty.append((eid, WORD_OF.get(eid), f"{fn}:{lineno}")); continue
    for chunk in p[1].split("/"):
        chunk = chunk.strip()
        if not chunk: continue
        m = re.match(r"^([^ ]+(?: [a-z]\.)?)\s+(.*)$", chunk)
        pos = m.group(1) if m else ""
        if not POS_OK.match(pos): bad_pos.append((eid, WORD_OF.get(eid), pos, chunk[:40]))

print(f"  重复编号: {len(dup)} {dup[:10]}")
print(f"  权威表外的编号: {len(unknown)} {unknown[:10]}")
print(f"  第2段为空: {len(empty)} {empty[:10]}")
print(f"  竖线数不是 3 的行: {len(badfmt)}")
for fn, i, n, seg in badfmt[:25]: print(f"      {fn}:{i}  竖线 {n} 个 | {seg}")
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
