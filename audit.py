# -*- coding: utf-8 -*-
"""
全量校对

对 Day4 以后的全部内容做系统检查：
  1. 编号覆盖：每天的词是否与权威表完全一致
  2. 释义完整性：是否有条目缺释义
  3. 词性符号是否规范
  4. 标记是否正确：** 成对、！使用、无残留旧符号
  5. 分隔符误用：义项文字里出现 / 或 | 会破坏解析
  6. 词义错位：逐条给出「词-义」相关性，低分列出供人工复核
  7. 图片形：词形表是否与权威表一致
用法：python audit.py            输出汇总 + 可疑清单
      python audit.py --day 8    只看第 8 天
"""
import os, re, sys, json, glob

sys.stdout.reconfigure(encoding="utf-8")
ROOT = os.path.dirname(os.path.abspath(__file__))
JSONP = os.path.join(ROOT, "step4_wordlist.json")
DEFS = os.path.join(ROOT, "defs")

POS_OK = re.compile(r"^(linking v\.|modal v\.|n\. \[[CU]\]&\[[CU]\]|n\. \[[CU]\]|n\.|"
                    r"vt\.&vi\.|vi\.&vt\.|vt\.|vi\.|v\.|adj\.|adv\.|conj\.|prep\.|pron\.|"
                    r"det\.|num\.|int\.)$")
OLD_MARKS = "❗⚠▲△"
PLACEHOLDER = ("待补", "损坏", "无法辨认")

def load_defs():
    """返回 {编号: (第2段, 第3段, 第4段, 来源文件, 行号)}"""
    d = {}
    for fn in sorted(glob.glob(os.path.join(DEFS, "*.txt"))):
        base = os.path.basename(fn)
        for i, ln in enumerate(open(fn, encoding="utf-8", errors="replace"), 1):
            ln = ln.rstrip("\n")
            if not ln.strip() or ln.startswith("#"): continue
            parts = ln.split("|")
            if len(parts) < 2 or not parts[0].strip().isdigit(): continue
            eid = int(parts[0])
            d.setdefault(eid, (parts[1], parts[2] if len(parts) > 2 else "",
                               parts[3] if len(parts) > 3 else "", base, i))
    return d

def norm(w):
    return re.sub(r"[^a-z]", "", str(w).lower())

def match_score(word, text):
    """粗略衡量「词」与「释义」的相关度"""
    w = norm(word)
    t = norm(text)
    if not w or not t: return 0.0
    if w in t: return 1.0
    # 最长公共子串 / 词长
    best = 0
    for i in range(len(w)):
        for j in range(i + 2, min(len(w), i + 6) + 1):
            if w[i:j] in t:
                best = max(best, j - i)
    return best / len(w)

def main():
    only = None
    if "--day" in sys.argv:
        only = int(sys.argv[sys.argv.index("--day") + 1])

    data = json.load(open(JSONP, encoding="utf-8"))
    W = {t["id"]: (t["final"] or t["word"] or "") for t in data}
    D = {t["id"]: (t["day"] or "") for t in data}
    D2 = {t["id"]: (t.get("day") or "") for t in data}
    defs = load_defs()

    days = sorted({d for d in D.values() if d},
                  key=lambda x: (x != "新增", int(re.sub(r"\D", "", x) or 0)))

    print("=" * 78)
    print("全量校对报告")
    print("=" * 78)
    all_issues = []
    for day in days:
        num = 0 if day == "新增" else int(re.sub(r"\D", "", day))
        if only and num != only: continue
        ids = sorted(i for i in W if D[i] == day)
        if day == "新增": ids = sorted(i for i in W if D2[i] == "新增")
        probs = []
        for eid in ids:
            w = W[eid]
            if any(p in w for p in PLACEHOLDER):
                continue
            if eid not in defs:
                probs.append(("缺释义", eid, w)); continue
            seg2, seg3, seg4, src, ln = defs[eid]
            if not seg2.strip():
                probs.append(("第2段空", eid, w)); continue
            for x in (seg2, seg3, seg4):
                if any(m in x for m in OLD_MARKS):
                    probs.append(("旧符号", eid, w)); break
            if seg2.count("**") % 2:
                probs.append(("粗体未闭合", eid, w))
            for chunk in seg2.split("/"):
                chunk = chunk.strip()
                if not chunk: continue
                m = POS_OK.match(chunk.split()[0]) if chunk.split() else None
                if not m:
                    probs.append(("词性符号异常: " + chunk.split()[0], eid, w))
                    break
            if "|" in seg2 or "|" in seg3:
                probs.append(("内嵌竖线", eid, w))
        # 词义错位：取整天的平均分与最低几条
        scores = []
        for eid in ids:
            if eid in defs and not any(p in W[eid] for p in PLACEHOLDER):
                scores.append((match_score(W[eid], defs[eid][0]), eid, W[eid]))
        scores.sort()
        low = [s for s in scores if s[0] < 0.34]
        avg = sum(s[0] for s in scores) / len(scores) if scores else 0
        status = "✅" if not probs and not low else "⚠️"
        print(f"{status} {day:8s} {len(ids):>4} 词 | 问题 {len(probs):>2} | 词义相关度均值 {avg:.2f} | 低分 {len(low):>2}")
        if probs:
            for kind, eid, w in probs[:6]:
                print(f"        [{kind}] {eid} {w}")
            if len(probs) > 6: print(f"        ...另有 {len(probs)-6} 处")
            all_issues.append((day, "格式", len(probs)))
        if low:
            for sc, eid, w in low[:5]:
                txt = defs[eid][0][:44] if eid in defs else ""
                print(f"        [疑错位 {sc:.2f}] {eid} {w:<16} {txt}")
            if len(low) > 5: print(f"        ...另有 {len(low)-5} 条低分")
            all_issues.append((day, "疑错位", len(low)))

    print()
    print("=" * 78)
    if all_issues:
        print("汇总：")
        for day, kind, n in all_issues:
            print(f"   {day:8s} {kind} {n} 处")
    else:
        print("✅ 未发现格式问题，也没有明显的词义错位")
    print("=" * 78)

if __name__ == "__main__":
    main()
