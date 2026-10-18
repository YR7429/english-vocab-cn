# -*- coding: utf-8 -*-
r"""
校对台 —— 带界面的校对工具

只用 Python 自带的 tkinter，不需要安装任何第三方库。

两个页面：
  ① 考纲词汇表校对
       核对《考纲词汇含分级.txt》的词形与星级。左边列出课标每一条，右边能看到
       它在我方词表里对应哪一条。双击可改词形与星级，改完立刻写回该文件。
       适合做「词形 + 星级 + 顺序」的校对。

  ② 字母序校对
       把我方 3186 条按字母序排开，逐条核对词性与星级；右侧可以直接把释义、
       变化、备注写进 defs/ 对应文件。适合「一边核对一边查字典补释义」。

快捷键
  ① 考纲词汇表校对（焦点在列表上时）
       Enter    标记已核对，并自动跳到下一条
       空格     标记 / 取消标记（不跳转）
       Delete   删除当前行（弹确认窗）
       Insert   在当前行后插入一条
       F2       编辑当前行的词形与星级（Enter 保存、Esc 取消）
       Ctrl+F   跳到搜索框；Esc 清空筛选；F5 重新载入
  ② 字母序校对（焦点在列表上时）
       Enter    跳到右侧释义框，光标落在末尾
       Ctrl+S   保存；Ctrl+Enter 保存并下一条
       Ctrl+F   跳到搜索框；Esc 回到列表；F5 重新载入

用法：
  python 校对台.py

进度自动存在 校对进度.json（不入库）。改完仍要跑：
  python check_defs.py
  python build.py

安全措施：
  · 第一次写某个文件前，会在同目录留一份 .bak 备份
  · 写释义时校验「该行开头的编号」是否与当前条目一致，不一致就拒绝写入
  · 只替换单行，不动文件其他内容，换行符统一 LF

自检（不开窗口，检查数据能正常载入）：
  python 校对台.py --selftest
"""
import os, re, sys, json, shutil, datetime

import tkinter as tk
from tkinter import ttk, messagebox

sys.stdout.reconfigure(encoding="utf-8")
ROOT = os.path.dirname(os.path.abspath(__file__))
KCP = os.path.join(ROOT, "考纲词汇含分级.txt")
JSONP = os.path.join(ROOT, "step4_wordlist.json")
DEFSD = os.path.join(ROOT, "defs")
PROGP = os.path.join(ROOT, "校对进度.json")
PENDING = "（待补）"

sys.path.insert(0, ROOT)


# ---------------- 数据层 ----------------
def variants(w):
    """复用 tier_check 的写法归一化；它缺失时退回简单小写"""
    try:
        import tier_check
        return tier_check.variants(w)
    except Exception:
        return {str(w).lower().strip()}


def load_kc():
    """读课标文件，返回 [{line, word, stars, raw}]"""
    rows = []
    for i, ln in enumerate(open(KCP, encoding="utf-8", errors="replace").read().split("\n"), 1):
        s = ln.strip()
        if not s or i <= 16: continue
        if re.match(r"^[A-Z]$", s) or "年版" in s or s.startswith("("): continue
        if not re.match(r"^[A-Za-z]", s): continue
        core = re.sub(r"\*+$", "", s).strip()
        if not core: continue
        rows.append({"line": i, "word": core, "stars": s.count("*"), "raw": ln})
    return rows


def load_defs_index():
    """编号 -> {file, lineno, pos, chg, note}"""
    out = {}
    if not os.path.isdir(DEFSD): return out
    for fn in sorted(os.listdir(DEFSD)):
        if not fn.endswith(".txt"): continue
        p = os.path.join(DEFSD, fn)
        for i, ln in enumerate(open(p, encoding="utf-8", errors="replace"), 1):
            if ln.startswith("#") or "|" not in ln: continue
            parts = ln.rstrip("\n").split("|")
            if len(parts) != 4: continue
            out.setdefault(parts[0].strip(),
                           {"file": fn, "lineno": i, "pos": parts[1], "chg": parts[2], "note": parts[3]})
    return out


def load_entries():
    data = json.load(open(JSONP, encoding="utf-8"))
    defs = load_defs_index()
    out = []
    for t in data:
        eid = str(t["id"])
        d = defs.get(eid) or {"file": "", "lineno": 0, "pos": PENDING, "chg": "", "note": ""}
        out.append({
            "id": eid, "word": t.get("final") or t.get("word") or "",
            "day": t.get("day") or "", "star": t.get("tier_use"),
            "src": t.get("tier_src") or "", **d,
        })
    return out


def alpha_key(e):
    s = re.sub(r"[^0-9a-z]+", "", e["word"].lower())
    return (s, e["word"].lower())


def core_variants(w):
    """我方一侧的匹配键：variants() 去掉「虚词级的括号内容」。
    否则 'according (to)' 会把课标的 'to' 误判为已收录；
    而 'television (TV)' 的 'tv' 要保留，课标正是用 'TV (=television)' 收录的。
    注意：只删「来自括号」的虚词，词条本身就是虚词（to、and、of…）时照常保留。"""
    STOP = {"to", "be", "a", "an", "the", "of", "in", "on", "at", "by", "for", "and", "or"}
    w0 = str(w).lower().strip()
    paren = set()
    for m in re.finditer(r"\(([^)]*)\)", w0):
        inner = m.group(1).strip()
        if inner and "=" not in inner:
            paren.add(inner); paren.add(inner.replace(" ", ""))
    return {v for v in variants(w) if not (v in paren and v in STOP)}


def kc_coverage(entries):
    """课标写法 -> 我方编号。
    我方一侧只登记主写法；课标一侧允许括号内容（如 TV (=television)、
    tofu (bean curd)），这样两边都能对上。"""
    idx = {}
    for e in entries:
        for v in core_variants(e["word"]):
            idx.setdefault(v, e["id"])
    return idx


def _pkey(r):
    """进度键：用词形而不是行号，这样插入/删除行不会让已核对标记错位。
    同形重复（China/china）会共用同一个键，两者一起被标记。"""
    k = re.sub(r"[^a-z]", "", str(r["word"]).lower())
    return k or f"#{r['line']}"


def load_prog():
    if os.path.exists(PROGP):
        try:
            p = json.load(open(PROGP, encoding="utf-8"))
            p.setdefault("考纲", {}); p.setdefault("词表", {})
            return p
        except Exception:
            pass
    return {"考纲": {}, "词表": {}}


def save_prog(p):
    with open(PROGP, "w", encoding="utf-8", newline="\n") as f:
        json.dump(p, f, ensure_ascii=False, indent=1)


_BACKED = set()
def backup_once(path):
    """第一次写某文件前留一份备份"""
    if path in _BACKED: return
    if os.path.exists(path):
        stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        shutil.copy2(path, path + f".{stamp}.bak")
    _BACKED.add(path)


def write_line(path, lineno, newline):
    """只替换第 lineno 行（1 起算）"""
    backup_once(path)
    lines = open(path, encoding="utf-8", errors="replace").read().split("\n")
    i = lineno - 1
    if i < 0 or i >= len(lines):
        raise RuntimeError(f"{os.path.basename(path)} 第 {lineno} 行不存在，可能已被改动，请重开校对台")
    lines[i] = newline
    with open(path, "w", encoding="utf-8", newline="\n") as f:
        f.write("\n".join(lines))


# ---------------- 界面 ----------------
class App(tk.Tk):
    def __init__(self, selftest=False):
        super().__init__()
        self.title("词表校对台")
        self.geometry("1280x820")
        self.selftest = selftest
        self.prog = load_prog()
        self.entries = load_entries()
        self.kc = load_kc()
        self.cov = kc_coverage(self.entries)

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True)
        self.tab_kc = KcTab(nb, self)
        self.tab_al = AlphaTab(nb, self)
        nb.add(self.tab_kc, text="① 考纲词汇表校对")
        nb.add(self.tab_al, text="② 字母序校对（可写释义）")

        self.status = tk.StringVar()
        bar = ttk.Frame(self); bar.pack(fill="x")
        ttk.Label(bar, textvariable=self.status, anchor="w").pack(fill="x", padx=8, pady=3)
        self.refresh_status()

    def refresh_status(self):
        a = sum(1 for _, v in self.prog["考纲"].items() if v)
        b = sum(1 for _, v in self.prog["词表"].items() if v)
        self.status.set(f"考纲 {a}/{len(self.kc)} 已核对　｜　词表 {b}/{len(self.entries)} 已核对　｜　"
                        f"进度文件：校对进度.json")

    def reload_data(self):
        """增删课标行、加入新词之后，把两页的数据都重新读一遍"""
        self.entries = load_entries()
        self.kc = load_kc()
        self.cov = kc_coverage(self.entries)
        self.tab_kc.rows = list(self.kc)
        self.tab_kc.fill()
        self.tab_al.reload()
        self.refresh_status()


class KcTab(ttk.Frame):
    def __init__(self, master, app):
        super().__init__(master)
        self.app = app
        self.rows = list(app.kc)
        self._cache = {}          # iid -> {"miss": bool, "dup": str}

        top = ttk.Frame(self); top.pack(fill="x", padx=6, pady=5)
        ttk.Label(top, text="搜索").pack(side="left")
        self.q = tk.StringVar()
        e = ttk.Entry(top, textvariable=self.q, width=24); e.pack(side="left", padx=4)
        e.bind("<Return>", lambda _: self.fill())
        e.bind("<Escape>", lambda _: (self.tree.focus_set(), "break")[1])
        self.q_entry = e
        ttk.Button(top, text="筛选", command=self.fill).pack(side="left", padx=2)
        ttk.Button(top, text="清空", command=self.clear_q).pack(side="left", padx=2)
        self.only_miss = tk.BooleanVar(value=False)
        ttk.Checkbutton(top, text="只看我方没有的", variable=self.only_miss,
                        command=self.fill).pack(side="left", padx=8)
        self.only_todo = tk.BooleanVar(value=False)
        ttk.Checkbutton(top, text="只看未核对", variable=self.only_todo,
                        command=self.fill).pack(side="left", padx=8)
        ttk.Button(top, text="标记已核对", command=self.mark).pack(side="left", padx=8)
        ttk.Button(top, text="取消标记", command=self.unmark).pack(side="left", padx=2)

        self.stat = tk.StringVar()
        ttk.Label(top, textvariable=self.stat).pack(side="left", padx=12)

        top2 = ttk.Frame(self); top2.pack(fill="x", padx=6, pady=(0, 5))
        ttk.Button(top2, text="在此行后插入一条", command=self.insert_row).pack(side="left", padx=2)
        ttk.Button(top2, text="删除选中行", command=self.delete_row).pack(side="left", padx=2)
        ttk.Button(top2, text="加入我方词表", command=self.add_to_table).pack(side="left", padx=12)
        ttk.Label(top2, text="增删行会立刻写回 txt 并重读；行号整体前移/后移属正常",
                  foreground="#666").pack(side="left")

        ttk.Label(self, foreground="#777",
                  text="快捷键：Enter 标记并跳下一条 ｜ 空格 标记/取消 ｜ Delete 删除 ｜ "
                       "Insert 插入 ｜ F2 编辑 ｜ Ctrl+F 搜索 ｜ Esc 清筛选 ｜ F5 重载").pack(anchor="w", padx=8)

        cols = ("line", "word", "stars", "ours", "flag")
        self.tree = ttk.Treeview(self, columns=cols, show="headings")
        for c, t, w in (("line", "行号", 60), ("word", "课标词形", 240),
                        ("stars", "星级", 50), ("ours", "我方编号/词形", 200),
                        ("flag", "状态", 140)):
            self.tree.heading(c, text=t)
            self.tree.column(c, width=w, anchor="w")
        vs = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vs.set)
        self.tree.pack(side="left", fill="both", expand=True, padx=(6, 0), pady=(0, 6))
        vs.pack(side="left", fill="y", pady=(0, 6))
        self.tree.bind("<Double-1>", self.edit)
        self.tree.bind("<Return>", lambda e: self._k(self.mark_next))
        self.tree.bind("<space>", lambda e: self._k(self.toggle_mark))
        self.tree.bind("<Delete>", lambda e: self._k(self.delete_row))
        self.tree.bind("<Insert>", lambda e: self._k(self.insert_row))
        self.tree.bind("<F2>", lambda e: self._k(self.edit_selected))
        self.tree.bind("<Control-f>", lambda e: self._k(self.focus_search))
        self.tree.bind("<Escape>", lambda e: self._k(self.clear_all_filters))
        self.tree.bind("<F5>", lambda e: self._k(self.app.reload_data))
        self.fill()

    # ---- 键盘与选中辅助 ----
    def _k(self, fn):
        """执行动作并返回 'break'，阻止 Treeview 的默认按键行为"""
        fn()
        return "break"

    def _row(self, iid):
        return next((r for r in self.rows if str(r["line"]) == iid), None)

    def _cur_iid(self):
        s = self.tree.selection()
        return s[0] if s else None

    def select_next(self):
        kids = list(self.tree.get_children())
        if not kids:
            return
        cur = self._cur_iid()
        i = kids.index(cur) + 1 if cur in kids else 0
        if i >= len(kids):
            i = len(kids) - 1
        self.tree.selection_set(kids[i]); self.tree.see(kids[i]); self.tree.focus(kids[i])

    def _pkey_for(self, iid):
        r = self._row(iid)
        return _pkey(r) if r else iid

    def _flag_text(self, iid):
        c = self._cache.get(iid, {})
        parts = []
        if self.app.prog["考纲"].get(self._pkey_for(iid)):
            parts.append("✓ 已核对")
        if c.get("miss"):
            parts.append("我方无")
        if c.get("dup"):
            parts.append(c["dup"])
        return "；".join(parts)

    def _refresh_flag(self, iid):
        try:
            self.tree.set(iid, "flag", self._flag_text(iid))
        except Exception:
            pass

    def mark_next(self):
        """Enter：标为已核对，并跳到下一条"""
        iid = self._cur_iid()
        kids = list(self.tree.get_children())
        if not iid:
            if not kids:
                return
            iid = kids[0]
            self.tree.selection_set(iid)
        self.app.prog["考纲"][self._pkey_for(iid)] = True
        save_prog(self.app.prog); self.app.refresh_status()
        if self.only_todo.get():
            idx = kids.index(iid) if iid in kids else 0
            self.tree.delete(iid)
            kids = list(self.tree.get_children())
            if kids:
                nxt = kids[min(idx, len(kids) - 1)]
                self.tree.selection_set(nxt); self.tree.see(nxt); self.tree.focus(nxt)
        else:
            self._refresh_flag(iid)
            self.select_next()

    def toggle_mark(self):
        """空格：已核对则取消，否则标记；不跳转"""
        iid = self._cur_iid()
        if not iid:
            return
        key = self._pkey_for(iid)
        if self.app.prog["考纲"].get(key):
            self.app.prog["考纲"].pop(key, None)
        else:
            self.app.prog["考纲"][key] = True
        save_prog(self.app.prog); self.app.refresh_status()
        if self.only_todo.get():
            self.fill(); return
        self._refresh_flag(iid)

    def edit_selected(self):
        """F2：打开编辑框"""
        if self._cur_iid():
            self.edit(None)

    def focus_search(self):
        self.q_entry.focus_force(); self.q_entry.select_range(0, "end")

    def clear_all_filters(self):
        self.q.set(""); self.only_miss.set(False); self.only_todo.set(False)
        self.fill(); self.tree.focus_set()

    def clear_q(self):
        self.q.set(""); self.fill()

    def fill(self):
        q = self.q.get().strip().lower()
        self.tree.delete(*self.tree.get_children())
        n_miss = n_todo = n_dup = 0
        seen = {}
        for r in self.rows:
            iid = str(r["line"])
            key = _pkey(r)
            done = self.app.prog["考纲"].get(key, False)
            hit = None
            for v in variants(r["word"]):
                if v in self.app.cov:
                    hit = self.app.cov[v]; break
            # 同形重复：同一个写法在课标文件里出现了第二次
            norm = re.sub(r"[^a-z]", "", r["word"].lower())
            dup = norm in seen and norm not in ("", "an")
            seen.setdefault(norm, r["line"])
            if not hit: n_miss += 1
            if not done: n_todo += 1
            if dup: n_dup += 1
            if q and q not in r["word"].lower(): continue
            if self.only_miss.get() and hit: continue
            if self.only_todo.get() and done: continue
            self._cache[iid] = {"miss": not hit,
                                "dup": f"重复(另见第{seen[norm]}行)" if dup else ""}
            self.tree.insert("", "end", iid=iid,
                             values=(r["line"], r["word"], str(r["stars"]) if r["stars"] else "0",
                                     hit or "—", self._flag_text(iid)))
        self.stat.set(f"共 {len(self.rows)} 条｜我方无 {n_miss}｜同形重复 {n_dup}｜未核对 {n_todo}")

    def reload(self):
        self.rows = list(self.app.kc)
        self.fill()

    # ---- 结构性操作：插入 / 删除 / 加入我方词表 ----
    def insert_row(self):
        r = self._sel()
        lines = open(KCP, encoding="utf-8", errors="replace").read().split("\n")
        pos = r["line"] if r else len(lines) - 1      # 插在选中行之后；未选中则插到末尾
        pos = max(1, min(pos, len(lines)))
        lines.insert(pos, "placeholder")
        backup_once(KCP)
        with open(KCP, "w", encoding="utf-8", newline="\n") as f:
            f.write("\n".join(lines))
        self.app.reload_data()
        messagebox.showinfo("已插入",
                            f"已在第 {pos} 行插入占位词 placeholder。\n"
                            "双击它改成真实词形（漏项补这里）。")

    def delete_row(self):
        r = self._sel()
        if not r: return
        if not messagebox.askyesno("确认删除",
                                   f"从考纲文件里删掉第 {r['line']} 行「{r['word']}」？\n\n"
                                   "删除后，后面的行号会整体前移一位；已核对标记按行号记录，也会跟着错位，"
                                   "建议少量多次。"):
            return
        lines = open(KCP, encoding="utf-8", errors="replace").read().split("\n")
        i = r["line"] - 1
        if lines[i].strip() != r["raw"].strip():
            messagebox.showerror("行已变动", "这一行内容与载入时不同，请重新打开校对台再试"); return
        del lines[i]
        backup_once(KCP)
        with open(KCP, "w", encoding="utf-8", newline="\n") as f:
            f.write("\n".join(lines))
        self.app.reload_data()
        messagebox.showinfo("已删除", f"已删除「{r['word']}」。")

    def add_to_table(self):
        r = self._sel()
        if not r: return
        for v in variants(r["word"]):
            if v in self.app.cov:
                messagebox.showinfo("已经在词表里",
                                    f"「{r['word']}」已经对应我方编号 {self.app.cov[v]}，不需要再加。")
                return
        if not messagebox.askyesno("加入我方词表",
                                   f"把「{r['word']}」（课标 {r['stars']} 星）加入我方词表？\n\n"
                                   "会做四件事：\n"
                                   "  1. 分配一个新编号（接在 3185 之后）\n"
                                   "  2. 写进 step4_wordlist.json，天字段设为「新增」\n"
                                   "  3. 在 defs/D31_新增.txt 里加一行，释义先填（待补）\n"
                                   "  4. 重建 days/ 骨架\n\n"
                                   "加完请到第二页填释义，再跑 check_defs.py 与 build.py。"):
            return
        import authority as A
        data = A.load_json()
        low = r["word"].strip().lower()
        for t in data:
            if (t.get("final") or "").strip().lower() == low:
                messagebox.showerror("重复", f"词表里已经有「{t.get('final')}」（编号 {t['id']}），不能重复加入")
                return
        ints = [t["id"] for t in data if isinstance(t["id"], int)]
        new_id = max(max(ints) if ints else 0, 3185) + 1
        used = {str(t["id"]) for t in data}
        while str(new_id) in used: new_id += 1
        data.append({
            "id": new_id, "word": r["word"], "final": r["word"], "day": "新增",
            "tier_use": r["stars"], "tier_kc": r["stars"], "tier_src": "课标",
            "star_raw": r["stars"], "line": r["line"],
            "note": "补录：课标有、原词表缺", "how": "补录：课标有、原词表缺",
            "conf": "确定", "needs_review": False,
        })
        backup_once(JSONP)
        A.save_json(data)
        A.write_view(data)
        A.rebuild_skeletons(data)
        dp = os.path.join(DEFSD, "D31_新增.txt")
        backup_once(dp)
        txt = open(dp, encoding="utf-8", errors="replace").read() if os.path.exists(dp) else ""
        if not txt.endswith("\n") and txt: txt += "\n"
        with open(dp, "w", encoding="utf-8", newline="\n") as f:
            f.write(txt + f"{new_id}|（待补）||\n")
        self.app.reload_data()
        messagebox.showinfo("已加入",
                            f"「{r['word']}」已加入，新编号 {new_id}，归入「新增」。\n"
                            "下一步：第二页填释义 → python check_defs.py → python build.py")

    def _sel(self):
        s = self.tree.selection()
        if not s:
            messagebox.showinfo("提示", "先选一条"); return None
        return next(r for r in self.rows if str(r["line"]) == s[0])

    def mark(self):
        s = self.tree.selection()
        if not s: return
        for i in s:
            self.app.prog["考纲"][self._pkey_for(i)] = True
        save_prog(self.app.prog); self.app.refresh_status()
        if self.only_todo.get():
            self.fill()
        else:
            for i in s: self._refresh_flag(i)

    def unmark(self):
        s = self.tree.selection()
        for i in s:
            self.app.prog["考纲"].pop(self._pkey_for(i), None)
        save_prog(self.app.prog); self.app.refresh_status()
        for i in s: self._refresh_flag(i)

    def edit(self, event):
        r = self._sel()
        if not r: return
        dlg = tk.Toplevel(self); dlg.title(f"修改第 {r['line']} 行"); dlg.transient(self); dlg.grab_set()
        ttk.Label(dlg, text="词形").grid(row=0, column=0, padx=8, pady=6, sticky="e")
        w = tk.StringVar(value=r["word"])
        we = ttk.Entry(dlg, textvariable=w, width=40)
        we.grid(row=0, column=1, padx=8, pady=6)
        we.focus_set(); we.select_range(0, "end")
        ttk.Label(dlg, text="星级").grid(row=1, column=0, padx=8, pady=6, sticky="e")
        st = tk.IntVar(value=r["stars"])
        f = ttk.Frame(dlg); f.grid(row=1, column=1, sticky="w", padx=8)
        for v, t in ((0, "0（义务教育）"), (1, "*（必修）"), (2, "**（选择性必修）")):
            ttk.Radiobutton(f, text=t, value=v, variable=st).pack(side="left", padx=4)

        def ok():
            word, stars = w.get().strip(), int(st.get())
            if not word:
                messagebox.showwarning("提示", "词形不能为空"); return
            if not re.fullmatch(r"[A-Za-z][A-Za-z0-9 .,'’()/&=\-]*", word):
                if not messagebox.askyesno("确认", f"词形 {word!r} 含不常见字符，仍要保存？"): return
            write_line(KCP, r["line"], word + "*" * stars)
            r["word"], r["stars"] = word, stars
            dlg.destroy(); self.fill()

        ttk.Button(dlg, text="保存", command=ok).grid(row=2, column=1, sticky="e", padx=8, pady=10)
        ttk.Button(dlg, text="取消", command=dlg.destroy).grid(row=2, column=1, sticky="w", padx=8, pady=10)
        dlg.bind("<Return>", lambda _: ok())
        dlg.bind("<Escape>", lambda _: dlg.destroy())


class AlphaTab(ttk.Frame):
    def __init__(self, master, app):
        super().__init__(master)
        self.app = app
        self.rows = sorted(app.entries, key=alpha_key)
        self.by_id = {e["id"]: e for e in app.entries}

        top = ttk.Frame(self); top.pack(fill="x", padx=6, pady=5)
        ttk.Label(top, text="搜索").pack(side="left")
        self.q = tk.StringVar()
        e = ttk.Entry(top, textvariable=self.q, width=24); e.pack(side="left", padx=4)
        e.bind("<Return>", lambda _: self.fill())
        e.bind("<Escape>", lambda _: (self.tree.focus_set(), "break")[1])
        self.q_entry = e
        ttk.Button(top, text="筛选", command=self.fill).pack(side="left", padx=2)
        ttk.Button(top, text="清空", command=self.clear_q).pack(side="left", padx=2)
        ttk.Label(top, text="显示").pack(side="left", padx=(10, 2))
        self.mode = tk.StringVar(value="全部")
        cb = ttk.Combobox(top, textvariable=self.mode, width=12, state="readonly",
                          values=("全部", "未核对", "课标未命中", "释义待补"))
        cb.pack(side="left"); cb.bind("<<ComboboxSelected>>", lambda _: self.fill())
        ttk.Button(top, text="下一条未核对", command=self.next_todo).pack(side="left", padx=8)

        self.stat = tk.StringVar()
        ttk.Label(top, textvariable=self.stat).pack(side="left", padx=12)

        body = ttk.Panedwindow(self, orient="horizontal")
        body.pack(fill="both", expand=True, padx=6, pady=(0, 6))

        left = ttk.Frame(body)
        cols = ("id", "word", "star", "pos")
        self.tree = ttk.Treeview(left, columns=cols, show="headings")
        for c, t, w in (("id", "编号", 70), ("word", "词形", 160),
                        ("star", "星", 40), ("pos", "释义（摘要）", 300)):
            self.tree.heading(c, text=t)
            self.tree.column(c, width=w, anchor="w")
        vs = ttk.Scrollbar(left, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vs.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vs.pack(side="left", fill="y")
        self.tree.bind("<<TreeviewSelect>>", self.on_select)
        self.tree.bind("<Return>", lambda e: self._k(self.focus_editor))
        self.tree.bind("<Control-s>", lambda e: self._k(self.save))
        self.tree.bind("<Control-Return>", lambda e: self._k(lambda: self.save(next_one=True)))
        self.tree.bind("<Control-f>", lambda e: self._k(self.focus_search))
        self.tree.bind("<Escape>", lambda e: self._k(self.clear_q))
        self.tree.bind("<F5>", lambda e: self._k(self.app.reload_data))
        body.add(left, weight=3)

        right = ttk.Frame(body)
        body.add(right, weight=2)

        self.hdr = tk.StringVar(value="（未选择）")
        ttk.Label(right, textvariable=self.hdr, font=("", 11, "bold")).pack(anchor="w", padx=8, pady=(6, 2))

        ttk.Label(right, text="词性 + 义项（第 2 段）").pack(anchor="w", padx=8)
        self.pos = tk.Text(right, height=9, wrap="word", font=("Consolas", 10))
        self.pos.pack(fill="x", padx=8, pady=2)
        for w in (self.pos,):
            w.bind("<Control-s>", lambda e: self._k(self.save))
            w.bind("<Control-Return>", lambda e: self._k(lambda: self.save(next_one=True)))
            w.bind("<Escape>", lambda e: self._k(self.tree.focus_set))

        ttk.Label(right, text="变化（第 3 段，可留空）").pack(anchor="w", padx=8)
        self.chg = tk.StringVar(); ttk.Entry(right, textvariable=self.chg).pack(fill="x", padx=8, pady=2)

        ttk.Label(right, text="备注（与变化同栏，可留空）").pack(anchor="w", padx=8)
        self.note = tk.StringVar(); ttk.Entry(right, textvariable=self.note).pack(fill="x", padx=8, pady=2)

        btns = ttk.Frame(right); btns.pack(fill="x", padx=8, pady=8)
        ttk.Button(btns, text="保存", command=self.save).pack(side="left")
        ttk.Button(btns, text="保存并下一条", command=lambda: self.save(next_one=True)).pack(side="left", padx=6)
        ttk.Button(btns, text="仅标为已核对", command=self.mark).pack(side="left")

        ttk.Label(right, foreground="#777", wraplength=420,
                  text="快捷键：Enter 跳到释义框 ｜ Ctrl+S 保存 ｜ Ctrl+Enter 保存并下一条 ｜ "
                       "Ctrl+F 搜索 ｜ Esc 回到列表 ｜ F5 重载").pack(anchor="w", padx=8, pady=(2, 0))

        self.hint = tk.StringVar(value="提示：双击列表或选中后编辑，保存会直接写回 defs/ 文件。")
        ttk.Label(right, textvariable=self.hint, wraplength=420, foreground="#555").pack(anchor="w", padx=8, pady=(6, 0))

        self.fill()

    def clear_q(self):
        self.q.set(""); self.fill()

    # ---- 键盘辅助 ----
    def _k(self, fn):
        fn()
        return "break"

    def focus_editor(self):
        """Enter：跳到释义框，光标落在末尾"""
        self.pos.focus_force()
        self.pos.mark_set("insert", "end-1c")
        self.pos.see("insert")

    def focus_search(self):
        self.q_entry.focus_force(); self.q_entry.select_range(0, "end")

    def reload(self):
        """数据被外部改动（新增词、改了词形）后重新载入"""
        self.rows = sorted(self.app.entries, key=alpha_key)
        self.by_id = {e["id"]: e for e in self.app.entries}
        self.fill()

    def fill(self):
        q = self.q.get().strip().lower()
        mode = self.mode.get()
        self.tree.delete(*self.tree.get_children())
        shown = todo = 0
        for e in self.rows:
            done = self.app.prog["词表"].get(e["id"], False)
            if mode == "未核对" and done: continue
            if mode == "课标未命中" and e["src"] == "课标": continue
            if mode == "释义待补" and e["pos"].strip() != PENDING: continue
            if q and q not in e["word"].lower() and q not in e["id"]: continue
            shown += 1
            if not done: todo += 1
            summary = e["pos"].replace("\n", " ")[:80]
            self.tree.insert("", "end", iid=e["id"],
                             values=(e["id"], e["word"], "*" * (e["star"] or 0) or "—", summary))
        self.stat.set(f"显示 {shown} 条｜其中未核对 {todo}")

    def _cur(self):
        s = self.tree.selection()
        return self.by_id.get(s[0]) if s else None

    def on_select(self, _=None):
        e = self._cur()
        if not e: return
        done = self.app.prog["词表"].get(e["id"], False)
        self.hdr.set(f"{e['id']}　{e['word']}　{e['day']}　{'已核对' if done else '未核对'}")
        self.pos.delete("1.0", "end"); self.pos.insert("1.0", e["pos"])
        self.chg.set(e["chg"]); self.note.set(e["note"])

    def next_todo(self):
        kids = self.tree.get_children()
        if not kids: return
        cur = self.tree.selection()
        start = kids.index(cur[0]) + 1 if cur else 0
        for k in kids[start:] + kids[:start]:
            if not self.app.prog["词表"].get(k, False):
                self.tree.selection_set(k); self.tree.see(k)
                self.on_select(); return
        messagebox.showinfo("提示", "当前列表里没有未核对的了")

    def mark(self):
        e = self._cur()
        if not e: return
        self.app.prog["词表"][e["id"]] = True
        save_prog(self.app.prog); self.app.refresh_status(); self.fill()

    def save(self, next_one=False):
        e = self._cur()
        if not e:
            messagebox.showinfo("提示", "先选一条"); return
        pos = self.pos.get("1.0", "end").strip().replace("\n", " ")
        chg = self.chg.get().strip()
        note = self.note.get().strip()
        if "|" in pos or "|" in chg or "|" in note:
            messagebox.showerror("不能保存", "释义/变化/备注里不能有竖线 |（它会把一行切成两段）"); return
        if not pos:
            messagebox.showerror("不能保存", "词性+义项不能为空；确实没写就填 （待补）"); return
        if not e["file"]:
            messagebox.showerror("不能保存", f"defs/ 里找不到 {e['id']} 这一行，请先运行 python check_defs.py 查看")
            return
        try:
            write_line(os.path.join(DEFSD, e["file"]), e["lineno"], f'{e["id"]}|{pos}|{chg}|{note}')
        except Exception as ex:
            messagebox.showerror("保存失败", str(ex)); return
        e["pos"], e["chg"], e["note"] = pos, chg, note
        self.app.prog["词表"][e["id"]] = True
        save_prog(self.app.prog); self.app.refresh_status(); self.fill()
        self.tree.selection_set(e["id"])
        self.hint.set(f"已写入 defs/{e['file']} 第 {e['lineno']} 行。记得跑 check_defs.py 与 build.py。")
        if next_one: self.next_todo()


def selftest():
    app = App(selftest=True)
    app.withdraw()
    app.update()
    print(f"考纲条目 {len(app.kc)}｜词表条目 {len(app.entries)}｜已覆盖映射 {len(app.cov)}")
    app.tab_kc.fill(); app.tab_al.fill()
    app.update()
    print("两个页面均能填充，自检通过（未写任何文件）")
    app.destroy()


if __name__ == "__main__":
    if "--selftest" in sys.argv:
        selftest()
    else:
        App().mainloop()
