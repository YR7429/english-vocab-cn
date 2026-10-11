# -*- coding: utf-8 -*-
"""
Day 骨架文件读取器
骨架格式：ITEMS = [(编号, '原词'), ...]   或   [(编号, '原词', ...), ...]（兼容旧格式）
词形与编号一律以权威表 step4_wordlist.json 为准，骨架只用来确认"这一天包含哪些编号"。
"""
import os, json, importlib.util

WS = r"D:\DSH_01"

def load_authority():
    p = os.path.join(ROOT, "step4_wordlist.json")
    if not os.path.exists(p): return {}, {}, {}
    data = json.load(open(p, encoding="utf-8"))
    star  = {t["id"]: t["tier_use"] for t in data}
    canon = {t["id"]: t["final"]    for t in data}
    day   = {t["id"]: t["day"]      for t in data}
    miss  = {t["id"] for t in data if t.get("tier_src") != "课标"}
    return star, canon, day, miss

def load_skeleton(path):
    """返回该文件声明的编号列表（升序）"""
    spec = importlib.util.spec_from_file_location(
        "skel_" + os.path.splitext(os.path.basename(path))[0], path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    ids = [int(e[0]) for e in getattr(mod, "ITEMS", [])]
    return sorted(set(ids))
