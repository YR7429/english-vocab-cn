# -*- coding: utf-8 -*-
"""
全量校对

对全部内容做系统检查：
  1. 编号覆盖：每天的词是否与权威表完全一致
  2. 释义完整性：是否有条目缺释义、是否仍为待补
  3. 词性符号是否规范
  4. 标记是否正确：** 成对、！使用、无残留旧符号
  5. 分隔符误用：义项文字里出现 | 会破坏解析
  6. 释义雷同：两条不同编号的释义完全相同 —— 这是「词义错位」最可靠的信号
     （原始资料是六列网格，行错位时会让相邻条目的释义互相串位，串位后必然雷同）
  7. 释义无中文：一条纯英文/空白的释义几乎肯定是错的

说明：曾用「英文词与中文释义的字符串重合度」判错位，但中文释义里本来就不含
英文字母，该指标恒为 0，等于每天报上百条假警。已废弃，改用上面的雷同检测。

用法：python audit.py            输出汇总 + 可疑清单
      python audit.py --day 8    只看第 8 天
"""
import os, re, sys, json, glob
from collections import defaultdict

import ids_util

sys.stdout.reconfigure(encoding="utf-8")
ROOT = os.path.dirname(os.path.abspath(__file__))
JSONP = os.path.join(ROOT, "step4_wordlist.json")
DEFS = os.path.join(ROOT, "defs")

# 允许的词性符号（与 build.py 的 POS_TOKEN、check_defs.py 保持一致）
POS_OK = re.compile(r"^(linking v\.|modal v\.|aux\. v\.|"
                    r"n\. \[[CU]\]&\[[CU]\]|n\. \[[CU]\]|n\.|"
                    r"vt\.&vi\.|vi\.&vt\.|vt\.|vi\.|v\.|"
                    r"adj\.|adv\.|conj\.|prep\.|pron\.|det\.|num\.|int\.)$")
OLD_MARKS = "❗⚠▲△"
PLACEHOLDER = ("待补", "损坏", "无法辨认")
PENDING = "（待补）"            # 释义尚未填写的占位符
CJK = re.compile(r"[\u4e00-\u9fff]")


def load_defs():
    """返回 {编号: (第2段, 第3段, 第4段, 来源文件, 行号)}"""
    d = {}
    for fn in sorted(glob.glob(os.path.join(DEFS, "*.txt"))):
        base = os.path.basename(fn)
        for i, ln in enumerate(open(fn, encoding="utf-8", errors="replace"), 1):
            ln = ln.rstrip("\n")
            if not ln.strip() or ln.startswith("#"): continue
            parts = ln.split("|")
            if len(parts) < 2: continue
            eid = ids_util.parse(parts[0].strip())
            if eid is None: continue
            d.setdefault(eid, (parts[1], parts[2] if len(parts) > 2 else "",
                               parts[3] if len(parts) > 3 else "", base, i))
    return d


def pos_of(chunk):
    """从 'n. [C] 政府' 里取出 'n. [C]'；取不到返回空串"""
    m = re.match(r"^([^ ]+(?: [a-z]\.)?)\s+(.*)$", chunk)
    return m.group(1) if m else ""


def main():
    only = None
    if "--day" in sys.argv:
        i = sys.argv.index("--day")
        if i + 1 < len(sys.argv):
            only = int(sys.argv[i + 1])

    data = json.load(open(JSONP, encoding="utf-8"))
    W = {t["id"]: (t.get("final") or t.get("word") or "") for t in data}
    D = {t["id"]: (t.get("day") or "") for t in data}
    defs = load_defs()

    # 释义雷同：同一段释义出现在两个以上不同编号上
    groups = defaultdict(list)
    for eid, rec in defs.items():
        txt = rec[0].strip()
        if txt and txt != PENDING:
            groups[txt].append(eid)
    twins = {}
    for txt, ks in groups.items():
        if len(ks) > 1:
            ks = sorted(ks, key=ids_util.key)
            for eid in ks:
                twins[eid] = [x for x in ks if x != eid]

    days = sorted({d for d in D.values() if d},
                  key=lambda x: (x != "新增", int(re.sub(r"\D", "", x) or 0)))

    print("=" * 78)
    print("全量校对报告")
    print("=" * 78)
    total = {"词条": 0, "格式": 0, "待补": 0, "雷同": 0, "无中文": 0}
    for day in days:
        num = 0 if day == "新增" else int(re.sub(r"\D", "", day))
        if only is not None and num != only: continue
        ids = sorted((i for i in W if D[i] == day), key=ids_util.key)
        probs, pend, dup_ids, nocjk = [], [], [], []
        for eid in ids:
            w = W[eid]
            if any(p in w for p in PLACEHOLDER):
                continue
            if eid not in defs:
                probs.append(("缺释义", eid, w)); continue
            seg2, seg3, seg4, src, ln = defs[eid]
            if not seg2.strip():
                probs.append(("第2段空", eid, w)); continue
            if seg2.strip() == PENDING:
                pend.append((eid, w)); continue
            for x in (seg2, seg3, seg4):
                if any(m in x for m in OLD_MARKS):
                    probs.append(("旧符号", eid, w)); break
            if seg2.count("**") % 2:
                probs.append(("粗体未闭合", eid, w))
            for chunk in seg2.split("/"):
                chunk = chunk.strip()
                if not chunk: continue
                pos = pos_of(chunk)
                if not POS_OK.match(pos):
                    probs.append(("词性符号异常: " + (pos or chunk[:8]), eid, w))
                    break
            if "|" in seg2 or "|" in seg3:
                probs.append(("内嵌竖线", eid, w))
            if not CJK.search(seg2):
                nocjk.append((eid, w, seg2[:36]))
            if eid in twins:
                dup_ids.append((eid, w, twins[eid], seg2[:36]))

        total["词条"] += len(ids)
        total["格式"] += len(probs)
        total["待补"] += len(pend)
        total["雷同"] += len(dup_ids)
        total["无中文"] += len(nocjk)
        ok = not (probs or dup_ids or nocjk)
        print(f"{'OK ' if ok else '疑难'} {day:8s} {len(ids):>4} 词 | 格式 {len(probs):>2} | "
              f"待补 {len(pend):>2} | 释义雷同 {len(dup_ids):>2}"
              + (f" | 无中文 {len(nocjk)}" if nocjk else ""))
        for kind, eid, w in probs[:6]:
            print(f"        [{kind}] {eid} {w}")
        if len(probs) > 6: print(f"        ...另有 {len(probs)-6} 处")
        for eid, w, other, txt in dup_ids[:8]:
            print(f"        [释义雷同] {eid} {w:<16} 与 {', '.join(map(str, other))} 完全相同 | {txt}")
        if len(dup_ids) > 8: print(f"        ...另有 {len(dup_ids)-8} 条雷同")
        for eid, w, txt in nocjk[:5]:
            print(f"        [释义无中文] {eid} {w:<16} {txt}")
        for eid, w in pend[:5]:
            print(f"        [待补释义] {eid} {w}")
        if len(pend) > 5: print(f"        ...另有 {len(pend)-5} 条待补")

    print()
    print("=" * 78)
    print(f"合计：{total['词条']} 词 | 格式问题 {total['格式']} 处 | 待补释义 {total['待补']} 条 | "
          f"释义雷同 {total['雷同']} 条 | 释义无中文 {total['无中文']} 条")
    if total["雷同"]:
        print()
        print("「释义雷同」最值得逐条核对：原始资料是六列网格，行错位会让相邻条目")
        print("的释义互相串位，串位后两条释义必然完全相同。但同义词（如 two 与 second")
        print("的中文都可能写作「第二」）也会命中，需人工判断。")
    print("=" * 78)


if __name__ == "__main__":
    main()
