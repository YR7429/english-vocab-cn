# -*- coding: utf-8 -*-
"""
星级重算工具

背景：损坏词被人工还原成真实词形后（例如 SOITOw -> sorrow / pieidds -> pie），
      该条的星级判定仍是旧的「原文星号」，于是成品里的星号会是错的。
      本工具重新对照课标文件，修正这类条目的星级。

用法：
  python recompute_tier.py            只报告，不改任何文件
  python recompute_tier.py --apply    确认无误后写回权威表

判定逻辑与 build.py 的构建前提示共用 tier_check.py，口径一致。
"""
import os, sys, json

sys.stdout.reconfigure(encoding="utf-8")
ROOT  = os.path.dirname(os.path.abspath(__file__))
JSONP = os.path.join(ROOT, "step4_wordlist.json")

sys.path.insert(0, ROOT)
import tier_check

def main():
    apply = "--apply" in sys.argv
    kc = tier_check.load_kc()
    data = json.load(open(JSONP, encoding="utf-8"))
    changes, placeholders = tier_check.scan(data, kc)

    print(f"权威表 {len(data)} 条｜课标索引 {len(kc)} 键")
    print(f"需要修正的星级判定：{len(changes)} 处")
    for eid, w, old_kc, old_src, new_tier in changes[:60]:
        new = f"课标 {new_tier} 星" if new_tier is not None else "沿用原文星号（课标查不到）"
        print(f"   {eid:>5} {w:<24} 旧: tier_kc={old_kc} src={old_src}  ->  {new}")
    if len(changes) > 60:
        print(f"   ...另有 {len(changes) - 60} 处")

    print(f"\n占位条目（跳过，不参与重算）：{len(placeholders)}")
    for p in placeholders:
        print(f"   {p[0]:>5} {p[1]}")

    if apply and changes:
        tier_check.apply_changes(data, changes)
        with open(JSONP, "w", encoding="utf-8", newline="\n") as f:
            json.dump(data, f, ensure_ascii=False, indent=1)
            f.write("\n")
        bar = "=" * 68
        print("\n" + bar)
        print(f"✅ 已写回权威表（{len(changes)} 处已修正）")
        print(bar)
        print("接着依次运行：")
        print("   python authority.py view     # 刷新词形表")
        print("   python build.py              # 重新生成 PDF / Word")
    elif changes:
        bar = "=" * 68
        print("\n" + bar)
        print(f"⚠️  发现 {len(changes)} 处需要修正，但【尚未写回】")
        print(bar)
        print("上面只是报告，文件没有被改动，成品里仍会显示旧标记。")
        print("确认无误后，运行这一条才会真正生效：")
        print("   python recompute_tier.py --apply")
    else:
        print("\n✅ 全部一致，无需修正。")

if __name__ == "__main__":
    main()
