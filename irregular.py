# -*- coding: utf-8 -*-
"""
不规则变化引擎
用途：自动生成词表"变化"列的内容。
原则：只输出「不能靠规则推出」的形式；规则变化返回空字符串。
数据来源：人工整理的不规则形式表。
"""

PLURAL = {
    "tooth": "teeth", "foot": "feet", "mouse": "mice", "goose": "geese",
    "man": "men", "woman": "women", "policeman": "policemen", "policewoman": "policewomen",
    "child": "children", "person": "people", "basis": "bases", "crisis": "crises",
    "analysis": "analyses", "phenomenon": "phenomena", "criterion": "criteria",
    "medium": "media", "datum": "data", "bacterium": "bacteria",
    "sheep": "sheep", "deer": "deer", "fish": "fish", "means": "means",
    "series": "series", "species": "species", "Chinese": "Chinese", "Japanese": "Japanese",
    "knife": "knives", "leaf": "leaves", "life": "lives", "wife": "wives",
    "wolf": "wolves", "shelf": "shelves", "thief": "thieves", "half": "halves",
    "loaf": "loaves", "self": "selves", "scarf": "scarves",
    "potato": "potatoes", "tomato": "tomatoes", "hero": "heroes", "echo": "echoes",
    "photo": "photos", "piano": "pianos", "radio": "radios", "zoo": "zoos",
    "mosquito": "mosquitoes", "kilo": "kilos", "bamboo": "bamboos",
}

PLURAL_TWO = {
    "handkerchief": "handkerchiefs / handkerchieves",
    "mosquito": "mosquitoes / mosquitos",
    "staff": "staffs（指多批人员）",
    "index": "indexes / indices",
    "appendix": "appendixes / appendices",
}

VERB = {
    "be": ("was/were", "been", "being"), "have": ("had", "had", "having"),
    "do": ("did", "done", "doing"), "go": ("went", "gone", "going"),
    "say": ("said", "said", "saying"), "get": ("got", "got/gotten", "getting"),
    "make": ("made", "made", "making"), "know": ("knew", "known", "knowing"),
    "think": ("thought", "thought", "thinking"), "take": ("took", "taken", "taking"),
    "see": ("saw", "seen", "seeing"), "come": ("came", "come", "coming"),
    "give": ("gave", "given", "giving"), "find": ("found", "found", "finding"),
    "tell": ("told", "told", "telling"), "become": ("became", "become", "becoming"),
    "leave": ("left", "left", "leaving"), "feel": ("felt", "felt", "feeling"),
    "put": ("put", "put", "putting"), "bring": ("brought", "brought", "bringing"),
    "begin": ("began", "begun", "beginning"), "keep": ("kept", "kept", "keeping"),
    "hold": ("held", "held", "holding"), "write": ("wrote", "written", "writing"),
    "stand": ("stood", "stood", "standing"), "hear": ("heard", "heard", "hearing"),
    "let": ("let", "let", "letting"), "mean": ("meant", "meant", "meaning"),
    "set": ("set", "set", "setting"), "meet": ("met", "met", "meeting"),
    "run": ("ran", "run", "running"), "pay": ("paid", "paid", "paying"),
    "sit": ("sat", "sat", "sitting"), "speak": ("spoke", "spoken", "speaking"),
    "lie": ("lay", "lain", "lying"), "lead": ("led", "led", "leading"),
    "read": ("read", "read", "reading"), "grow": ("grew", "grown", "growing"),
    "lose": ("lost", "lost", "losing"), "fall": ("fell", "fallen", "falling"),
    "send": ("sent", "sent", "sending"), "build": ("built", "built", "building"),
    "understand": ("understood", "understood", "understanding"),
    "draw": ("drew", "drawn", "drawing"), "break": ("broke", "broken", "breaking"),
    "spend": ("spent", "spent", "spending"), "cut": ("cut", "cut", "cutting"),
    "rise": ("rose", "risen", "rising"), "drive": ("drove", "driven", "driving"),
    "buy": ("bought", "bought", "buying"), "wear": ("wore", "worn", "wearing"),
    "choose": ("chose", "chosen", "choosing"), "seek": ("sought", "sought", "seeking"),
    "throw": ("threw", "thrown", "throwing"), "catch": ("caught", "caught", "catching"),
    "deal": ("dealt", "dealt", "dealing"), "win": ("won", "won", "winning"),
    "forget": ("forgot", "forgotten", "forgetting"), "lay": ("laid", "laid", "laying"),
    "sell": ("sold", "sold", "selling"), "fight": ("fought", "fought", "fighting"),
    "bear": ("bore", "borne", "bearing"), "beat": ("beat", "beaten", "beating"),
    "hide": ("hid", "hidden", "hiding"), "shine": ("shone", "shone", "shining"),
    "strike": ("struck", "struck", "striking"), "stick": ("stuck", "stuck", "sticking"),
    "hang": ("hung", "hung", "hanging"), "shake": ("shook", "shaken", "shaking"),
    "steal": ("stole", "stolen", "stealing"), "swim": ("swam", "swum", "swimming"),
    "sing": ("sang", "sung", "singing"), "ring": ("rang", "rung", "ringing"),
    "drink": ("drank", "drunk", "drinking"), "ride": ("rode", "ridden", "riding"),
    "fly": ("flew", "flown", "flying"), "blow": ("blew", "blown", "blowing"),
    "teach": ("taught", "taught", "teaching"),
    "freeze": ("froze", "frozen", "freezing"), "steal2": ("stole", "stolen", "stealing"),
    "wake": ("woke", "woken", "waking"), "sink": ("sank", "sunk", "sinking"),
    "arise": ("arose", "arisen", "arising"), "awake": ("awoke", "awoken", "awaking"),
    "quit": ("quit", "quit", "quitting"), "shut": ("shut", "shut", "shutting"),
    "spread": ("spread", "spread", "spreading"), "hurt": ("hurt", "hurt", "hurting"),
    "cost": ("cost", "cost", "costing"), "hit": ("hit", "hit", "hitting"),
    "burst": ("burst", "burst", "bursting"), "feed": ("fed", "fed", "feeding"),
    "flee": ("fled", "fled", "fleeing"), "weep": ("wept", "wept", "weeping"),
    "sweep": ("swept", "swept", "sweeping"), "bend": ("bent", "bent", "bending"),
    "lend": ("lent", "lent", "lending"), "spell": ("spelt/spelled", "spelt/spelled", "spelling"),
    "burn": ("burnt/burned", "burnt/burned", "burning"),
    "learn": ("learnt/learned", "learnt/learned", "learning"),
    "smell": ("smelt/smelled", "smelt/smelled", "smelling"),
    "dream": ("dreamt/dreamed", "dreamt/dreamed", "dreaming"),
    "lie2": ("lied", "lied", "lying"),
    "die": ("died", "died", "dying"), "tie": ("tied", "tied", "tying"),
    "panic": ("panicked", "panicked", "panicking"),
    "picnic": ("picnicked", "picnicked", "picnicking"),
    "jog": ("jogged", "jogged", "jogging"),
    "plan": ("planned", "planned", "planning"),
    "stop": ("stopped", "stopped", "stopping"),
    "prefer": ("preferred", "preferred", "preferring"),
    "occur": ("occurred", "occurred", "occurring"),
    "refer": ("referred", "referred", "referring"),
    "admit": ("admitted", "admitted", "admitting"),
    "permit": ("permitted", "permitted", "permitting"),
    "regret": ("regretted", "regretted", "regretting"),
    "forget2": ("forgot", "forgotten", "forgetting"),
    "begin2": ("began", "begun", "beginning"),
    "travel": ("travelled/traveled", "travelled/traveled", "travelling/traveling"),
    "cancel": ("cancelled/canceled", "cancelled/canceled", "cancelling/canceling"),
    "quarrel": ("quarrelled/quarreled", "quarrelled/quarreled", "quarrelling/quarreling"),
}

ADJ = {
    "good": "better / best", "bad": "worse / worst", "far": "farther / further；farthest / furthest",
    "little": "less / least", "many": "more / most", "much": "more / most",
    "old": "older / elder；oldest / eldest",
    "late": "later / latter；latest / last",
}

ADV = {"well": "better / best", "badly": "worse / worst", "far2": "farther / further"}

# 需要特别提示的读音/同形异音
PHON = {
    "record": "名词读 /ˈrekɔːd/；动词读 /rɪˈkɔːd/",
    "present": "名词/形容词读 /ˈpreznt/；动词读 /prɪˈzent/",
    "object": "名词读 /ˈɒbdʒɪkt/；动词读 /əbˈdʒekt/",
    "close": "形容词/副词读 /kləʊs/；动词读 /kləʊz/",
    "separate": "动词读 /ˈsepəreɪt/；形容词读 /ˈseprət/",
    "use": "名词读 /juːs/；动词读 /juːz/",
    "excuse": "名词读 /ɪkˈskjuːs/；动词读 /ɪkˈskjuːz/",
    "increase": "名词读 /ˈɪŋkriːs/；动词读 /ɪnˈkriːs/",
    "decrease": "名词读 /ˈdiːkriːs/；动词读 /dɪˈkriːs/",
    "refuse": "名词读 /ˈrefjuːs/；动词读 /rɪˈfjuːz/",
    "wound": "读 /wuːnd/（与 wind 的过去式 /waʊnd/ 不同）",
    "desert": "名词（沙漠）读 /ˈdezət/；动词（抛弃）读 /dɪˈzɜːt/",
    "lead": "动词（带领）读 /liːd/；名词（铅）读 /led/",
    "live": "动词读 /lɪv/；形容词（现场的）读 /laɪv/",
    "tear": "名词（眼泪）读 /tɪə/；动词（撕）读 /teə/",
    "wind": "名词（风）读 /wɪnd/；动词（缠绕）读 /waɪnd/",
    "produce": "动词读 /prəˈdjuːs/；名词读 /ˈprɒdjuːs/",
    "content": "名词（内容）读 /ˈkɒntent/；形容词（满意的）读 /kənˈtent/",
    "conduct": "名词读 /ˈkɒndʌkt/；动词读 /kənˈdʌkt/",
}

def _clean(w):
    """去掉括号注、斜杠变体、末尾小数点，取主拼写用于查表"""
    w = w.split("(")[0].split("=")[0].strip().rstrip(".")
    return w.strip()

def forms(word):
    """返回该词需要提示的不规则变化列表（可能为空）"""
    out = []
    for seg in word.split("/"):
        w = _clean(seg)
        if not w: continue
        low = w.lower()
        if w in PLURAL_TWO: out.append("复数 " + PLURAL_TWO[w])
        elif low in PLURAL and PLURAL[low] != low:
            out.append("复数 " + PLURAL[low])
        if low in VERB:
            p, pp, ing = VERB[low]
            out.append(f"过去式 {p}；过去分词 {pp}")
        if low in ADJ:
            out.append("比较级 / 最高级 " + ADJ[low])
        if low in ADV:
            out.append("比较级 / 最高级 " + ADV[low])
        if low in PHON:
            out.append(PHON[low])
    # 去重、保序
    seen, uniq = set(), []
    for x in out:
        if x not in seen:
            seen.add(x); uniq.append(x)
    return uniq

if __name__ == "__main__":
    import sys
    sys.stdout.reconfigure(encoding="utf-8")
    for w in ["tooth", "teach", "good", "record", "handkerchief", "mouse", "book", "play", "desert", "leaf"]:
        print(f"{w:16s} -> {'；'.join(forms(w)) or '（规则变化，留空）'}")
