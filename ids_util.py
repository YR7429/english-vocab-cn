# -*- coding: utf-8 -*-
"""
编号工具

本项目绝大多数条目用普通整数编号（1–3185，对应原始资料）。
个别条目因一个词形承载了过多互不相干的义项，拆成两条，用子编号表示：

    1840-1    bind 的过去式、过去分词
    1840-2    形容词/名词义（一定的、前往、界限）

子编号与普通编号地位相同，都是稳定标识。JSON 里普通编号存为整数，
子编号存为字符串，读取时统一交给本模块解析，避免各处各写一套。

三个函数：
    parse(x)   把视图/释义文件里读到的字符串转成编号；无法识别返回 None
    key(x)     排序键：主编号升序，子编号紧跟主编号（1840 < 1840-1 < 1841）
    text(x)    输出用字符串
"""
import re

_SUB = re.compile(r"^(\d+)(?:-(\d+))?$")


def parse(x):
    """'993' -> 993（int）；'1840-1' -> '1840-1'（str）；不合法 -> None"""
    if isinstance(x, int):
        return x
    s = str(x).strip()
    m = _SUB.match(s)
    if not m:
        return None
    return int(s) if m.group(2) is None else s


def key(x):
    """排序键。子编号排在主编号之后：1639, 1640, 1640-1, 1640-2, 1641"""
    s = str(x).strip()
    m = _SUB.match(s)
    if not m:
        return (10 ** 9, 0)
    return (int(m.group(1)), int(m.group(2) or 0))


def text(x):
    return str(x)


def is_sub(x):
    return "-" in str(x)
