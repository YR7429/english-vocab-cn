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
import ids_util
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

# 释义还没写的条目，第2段统一写这个占位符（不算错，单独统计）
PENDING = "（待补）"

def load():
    """返回 {编号: [(文件名, 行号, 分段), ...]}；同时记录分段数异常的行"""
    rows, badfmt = {}, []
    for fn in sorted(glob.glob(os.path.join(ROOT, "defs", "*.txt"))):
        base = os.path.basename(fn)
        for i, ln in enumerate(open(fn, encoding="utf-8", errors="replace"), 1):
            ln = ln.rstrip("\n").strip()
            if not ln or ln.startswith("#"): continue
            p = ln.split("|")
            eid = ids_util.parse(p[0].strip())
            if eid is None: continue
            # 每行应恰好 3 个竖线 -> 4 段；多了少了都会让解析出错位
            if len(p) != 4: badfmt.append((base, i, ln.count("|"), ln[:60]))
            rows.setdefault(eid, []).append((base, i, p))
    return rows, badfmt

rows, badfmt = load()
print(f"defs 覆盖编号数: {len(rows)}")

bad_pos, empty, unknown, dup, pending = [], [], [], [], []
for eid, lst in rows.items():
    if len(lst) > 1: dup.append(eid)
    if eid not in WORD_OF: unknown.append(eid)
    fn, lineno, p = lst[0]
    if len(p) < 2 or not p[1].strip(): empty.append((eid, WORD_OF.get(eid), f"{fn}:{lineno}")); continue
    if p[1].strip() == PENDING:           # 释义待补：不算错，单独计数
        pending.append((eid, WORD_OF.get(eid), f"{fn}:{lineno}")); continue
    for chunk in p[1].split("/"):
        chunk = chunk.strip()
        if not chunk: continue
        m = re.match(r"^([^ ]+(?: [a-z]\.)?)\s+(.*)$", chunk)
        pos = m.group(1) if m else ""
        if not POS_OK.match(pos): bad_pos.append((eid, WORD_OF.get(eid), pos, chunk[:40]))

print(f"  重复编号: {len(dup)} {dup[:10]}")
print(f"  权威表外的编号: {len(unknown)} {unknown[:10]}")
print(f"  第2段为空: {len(empty)} {empty[:10]}")
print(f"  释义待补（{PENDING}）: {len(pending)}")
for x in pending[:25]: print("     ", x)
print(f"  竖线数不是 3 的行: {len(badfmt)}")
for fn, i, n, seg in badfmt[:25]: print(f"      {fn}:{i}  竖线 {n} 个 | {seg}")
print(f"  词性不规范: {len(bad_pos)}")
for x in bad_pos[:25]: print("     ", x)

# ---------------- 编号体检 ----------------
# 编号是全表的主键：重号会让构建静默丢一条，空号说明还原时漏了词。
from collections import Counter

ALL_IDS = [t["id"] for t in AUTH]
CNT = Counter(str(i) for i in ALL_IDS)
DUP_AUTH = sorted([k for k, v in CNT.items() if v > 1], key=ids_util.key)
PLAIN = {i for i in ALL_IDS if isinstance(i, int)}
SUBS = sorted([i for i in ALL_IDS if ids_util.is_sub(i)], key=ids_util.key)
PARENTS = {int(str(s).split("-")[0]) for s in SUBS}
TOP = max(PLAIN) if PLAIN else 0
# 空号：1..TOP 里既没有主编号、也没有子编号接替它
GAPS = [i for i in range(1, TOP + 1) if i not in PLAIN and i not in PARENTS]
BAD_SUBS = []
for p in sorted(PARENTS):
    got = sorted(int(str(s).split("-")[1]) for s in SUBS if int(str(s).split("-")[0]) == p)
    if got != list(range(1, len(got) + 1)):
        BAD_SUBS.append((p, got))
VIEW_IDS = set()
_vp = os.path.join(ROOT, "词形表.txt")
if os.path.exists(_vp):
    for _ln in open(_vp, encoding="utf-8"):
        _v = ids_util.parse(_ln.split("\t")[0].strip())
        if _v is not None: VIEW_IDS.add(_v)
VIEW_MISSING = [i for i in ALL_IDS if i not in VIEW_IDS]

print("\n编号体检：")
print(f"  权威表条目 {len(ALL_IDS)} ｜ 唯一编号 {len(CNT)} ｜ 词形表 {len(VIEW_IDS)} 行")
print(f"  重号: {len(DUP_AUTH)} {DUP_AUTH[:10]}")
print(f"  空号（1–{TOP}）: {len(GAPS)} {GAPS[:10]}")
print(f"  被子编号替代的主编号: {len(PARENTS)} {sorted(PARENTS)}")
print(f"  子编号: {len(SUBS)} {SUBS}")
if BAD_SUBS:    print(f"  子编号未从 1 起连续: {BAD_SUBS}")
if VIEW_MISSING: print(f"  词形表缺号: {len(VIEW_MISSING)} {VIEW_MISSING[:10]}")

# 覆盖进度
have = set(rows)
print("\n各日覆盖情况：")
for d in sorted(set(DAY_OF.values()), key=lambda z: (z != "新增", int(z.replace("Day","")) if z != "新增" else 0)):
    ids = [i for i in WORD_OF if DAY_OF[i] == d]
    n = sum(1 for i in ids if i in have)
    flag = "✓" if n == len(ids) else ("·" if n else " ")
    print(f"  {flag} {d:8s} {n:3d}/{len(ids):3d}")
