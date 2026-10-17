# -*- coding: utf-8 -*-
"""
Day 骨架文件读取器

骨架格式：ITEMS = [(编号, '原词'), ...]   或   [(编号, '原词', ...), ...]（兼容旧格式）
词形与编号一律以权威表 step4_wordlist.json 为准，骨架只用来确认"这一天包含哪些编号"。

注意：build.py 有自己的骨架读取逻辑（带缓存与释义合并），本模块供外部脚本单独调用。
"""
import os, json, importlib.util

import ids_util

ROOT = os.path.dirname(os.path.abspath(__file__))
DAYS = os.path.join(ROOT, "days")


def load_authority():
    """返回 (星级, 词形, 所在天, 课标未命中集合)；无权威表时返回四个空字典"""
    p = os.path.join(ROOT, "step4_wordlist.json")
    if not os.path.exists(p):
        return {}, {}, {}, set()
    data = json.load(open(p, encoding="utf-8"))
    star = {t["id"]: t.get("tier_use") for t in data}
    canon = {t["id"]: (t.get("final") or t.get("word") or "") for t in data}
    day = {t["id"]: t.get("day") for t in data}
    miss = {t["id"] for t in data if t.get("tier_src") != "课标"}
    return star, canon, day, miss


def load_skeleton(path):
    """返回该文件声明的编号列表（升序，子编号紧随主编号之后）"""
    if not os.path.isabs(path):
        path = os.path.join(DAYS, path)
    spec = importlib.util.spec_from_file_location(
        "skel_" + os.path.splitext(os.path.basename(path))[0], path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    got = [ids_util.parse(e[0]) for e in getattr(mod, "ITEMS", [])]
    return sorted({i for i in got if i is not None}, key=ids_util.key)
