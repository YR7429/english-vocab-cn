# -*- coding: utf-8 -*-
"""
权威表工具

设计原则：
  step4_wordlist.json   —— 机器可读的完整数据，只由本工具写入，请勿手工编辑
  词表/词形表.txt        —— 人工编辑用的视图，只用于改「词形」

三种用法：
  python authority.py fmt       把 json 重新排版（缩进、一行一条），方便查看
  python authority.py view      由 json 生成 词形表.txt
  python authority.py sync      从 词形表.txt 读回改动，更新 json，并重建骨架
  python authority.py check     只校验 词形表.txt 是否与 json 一致
"""
import os, re, sys, json, hashlib, shutil

sys.stdout.reconfigure(encoding="utf-8")
TABLE = os.path.dirname(os.path.abspath(__file__))
WS    = TABLE
JSONP = os.path.join(WS, "step4_wordlist.json")
VIEWP = os.path.join(TABLE, "词形表.txt")
HASHP = os.path.join(TABLE, ".词形表.hash")

# 词条排列顺序：Day1..Day30, 然后新增
def sort_key(t):
    d = t.get("day", "")
    n = 999 if d == "新增" else int(re.sub(r"\D", "", d) or 0)
    return (0 if d != "新增" else 1, n, t["id"])

def load_json():
    return json.load(open(JSONP, encoding="utf-8"))

def save_json(data):
    """缩进排版，一行一条，便于查看与 diff"""
    data = sorted(data, key=sort_key)
    # 每个对象内部保持字段顺序：核心字段在前
    order = ["id", "word", "final", "day", "tier_use", "tier_kc", "tier_src",
             "star_raw", "line", "note", "conf", "needs_review"]
    out = []
    for t in data:
        o = {}
        for k in order:
            if k in t: o[k] = t[k]
        for k in t:
            if k not in o: o[k] = t[k]
        out.append(o)
    with open(JSONP, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=1)
        f.write("\n")

def view_rows(data):
    """生成可编辑视图的行：编号 | 词形 | 星级 | 是否改动"""
    rows = []
    for t in sorted(data, key=sort_key):
        w = t.get("final") or t.get("word") or ""
        tier = t.get("tier_use")
        tier = "" if tier is None else str(tier)
        rows.append(f'{t["id"]}\t{w}\t{tier}')
    return rows

def write_view(data):
    body = "\n".join(view_rows(data)) + "\n"
    header = (
        "# 词形表（可编辑）\n"
        "#\n"
        "# 这是唯一可以手工编辑的权威数据文件。格式：编号 <TAB> 词形 <TAB> 星级\n"
        "#   - 编号：不要改\n"
        "#   - 词形：可以改（改错别字、补损坏条目、统一拼写）\n"
        "#   - 星级：不要改（由课标核对结果决定）\n"
        "#\n"
        "# 改完保存，然后运行：  python authority.py sync\n"
        "#\n"
    )
    open(VIEWP, "w", encoding="utf-8").write(header + body)
    h = hashlib.sha256((header + body).encode("utf-8")).hexdigest()
    open(HASHP, "w", encoding="utf-8").write(h)
    return h, len(view_rows(data))

def read_view():
    """读回视图，返回 {编号: 词形}；忽略 # 注释行"""
    if not os.path.exists(VIEWP):
        return None
    rows = {}
    for ln in open(VIEWP, encoding="utf-8"):
        s = ln.rstrip("\n")
        if not s.strip() or s.lstrip().startswith("#"): continue
        parts = s.split("\t")
        if len(parts) < 2: continue
        try: eid = int(parts[0].strip())
        except ValueError: continue
        rows[eid] = parts[1].strip()
    return rows

def current_hash():
    return hashlib.sha256(open(VIEWP, encoding="utf-8").read().encode("utf-8")).hexdigest()

def cmd_fmt():
    data = load_json()
    save_json(data)
    print(f"已重新排版：{JSONP}（{len(data)} 条，缩进 1）")

def cmd_view():
    data = load_json()
    h, n = write_view(data)
    print(f"已生成 {VIEWP}（{n} 条）；哈希已记录")

def cmd_check():
    data = load_json()
    saved = open(HASHP, encoding="utf-8").read().strip() if os.path.exists(HASHP) else ""
    cur = current_hash()
    rows = read_view() or {}
    ids = {t["id"]: (t.get("final") or t.get("word")) for t in data}
    diff = [(i, ids[i], rows.get(i)) for i in ids if i in rows and rows[i] != ids[i]]
    missing = [i for i in ids if i not in rows]

    if diff:
        print(f"提示：词形表.txt 有 {len(diff)} 处改动尚未写入权威表。")
        print("      如果这些改动是你的，请运行 python authority.py sync 写回权威表。")
    elif saved and saved != cur:
        # 数据一致但哈希不同：通常是换行符（LF/CRLF）转换或编辑器重存造成的
        print("词形表.txt 与权威表内容一致，无需同步。")
        print("（内部记录的时间戳不同，通常只是换行符转换，不影响构建结果）")
    else:
        print("词形表.txt 与上次生成时一致。")

    print(f"视图条目 {len(rows)} 条，权威表 {len(ids)} 条，词形差异 {len(diff)} 处")
    for d in diff[:20]:
        print("   ", d)
    if missing:
        print(f"   ！视图缺少 {len(missing)} 个编号（可运行 python authority.py view 重建视图）")
        print("     ", missing[:20])

def rebuild_skeletons(data):
    AUTH = {t["id"]: (t.get("final") or t.get("word")) for t in data}
    DAY  = {t["id"]: t.get("day") for t in data}
    Q3 = '"' * 3
    n = 0
    for num in range(1, 32):
        sel = sorted(i for i, d in DAY.items() if (d == "新增" and num == 31) or d == f"Day{num}")
        if not sel: continue
        title = "新增词汇" if num == 31 else f"Day {num} · 考纲第 {sel[0]}-{sel[-1]} 词"
        L = ["# -*- coding: utf-8 -*-", Q3, title,
             "本文件只是骨架：词形、编号、星级取自权威表 step4_wordlist.json；",
             "释义、变化、备注全部来自 defs/ 目录下对应的释义文件。", "",
             "要修订内容，请编辑 defs/ 里的释义文件，不要改本文件。", Q3, "ITEMS = ["]
        for i in sel:
            L.append(f"({i}, {AUTH[i]!r}),")
        L.append("]")
        open(os.path.join(TABLE, f"Day{num:02d}.py"), "w", encoding="utf-8").write("\n".join(L) + "\n")
        n += 1
    return n

def cmd_sync():
    rows = read_view()
    if rows is None:
        print("找不到词形表.txt，请先运行 python authority.py view"); return
    data = load_json()
    fixed, skipped = [], []
    for t in data:
        eid = t["id"]
        if eid not in rows: continue
        new = rows[eid]
        old = t.get("final") or t.get("word") or ""
        if new == old: continue
        if not new.strip():                       # 不允许清空
            skipped.append((eid, old, "词形为空，已忽略")); continue
        t["word"] = new; t["final"] = new
        fixed.append((eid, old, new))
    save_json(data)
    h, n = write_view(data)
    made = rebuild_skeletons(data)
    print(f"已同步 {len(fixed)} 处词形改动，重建骨架 {made} 个文件")
    for f in fixed[:40]:
        print(f"   {f[0]}: {f[1]}  ->  {f[2]}")
    for s in skipped:
        print(f"   [忽略] {s[0]}: {s[2]}")

if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else "check"
    {"fmt": cmd_fmt, "view": cmd_view, "sync": cmd_sync, "check": cmd_check}.get(cmd, cmd_check)()
