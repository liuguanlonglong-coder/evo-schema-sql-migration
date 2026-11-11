# -*- coding: utf-8 -*-
"""分析单位的三层分解（第五轮评审第 1、2 项：样本量到底支持什么推论）。

同一批冻结结果可以在四个不同分析单位上检验，结论强度差别很大：

  条目级（item）        —— 610 条端到端条目；本章正文的配对检验用这一层
  基础题级（question）  —— 同一基础题在多档重复出现时按其基础题聚类
  变换事实级（op）      —— 一道 schema 变换（如"把 A、B 两列合成 C"）是真正独立的演化实例
  数据库级（db）        —— 外部效度的单元，但合并/拆分各只有 5 个库

本脚本把同一组对照在四层上并列，报每层的单元数、差值、精确符号检验 p 与按单元重采样的
bootstrap 区间，并显式给出该层的最小可达 p（2/2^n）。评审的核心质疑是「题目级显著性
不能代替按数据库或变换事实留出的检验」——这里用数字把差距摆出来。

输出 results/op_level.json
"""
import collections
import json
import math
import os
import random

BASE = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE)

VERSION = "op-level-v1"
BOOT = 5000
SEED = 7

FILES = ["e1_dev_full.jsonl", "e1_diff_3lvl.jsonl", "e1_c3m_paired.jsonl",
         "e1_expand_med.jsonl", "e1_expand_hard.jsonl", "e1_ds_c3f.jsonl",
         "e1_ms_ds.jsonl", "e1_ms_ds_p.jsonl", "e1_ms_ds_p_split.jsonl",
         "e1_ds_one.jsonl", "e1_ds_one_b1.jsonl", "e1_ds_one_b2.jsonl",
         "e3_rw_main.jsonl"]

MAN = {}
for _l in open("data/manifest_v3.jsonl", encoding="utf-8"):
    _r = json.loads(_l)
    MAN[_r["id"]] = _r

BY = collections.defaultdict(dict)
for f in FILES:
    p = os.path.join("results", f)
    if not os.path.exists(p):
        continue
    for l in open(p, encoding="utf-8"):
        r = json.loads(l)
        if r.get("cond"):
            BY[r["id"]][r["cond"]] = r["outcome"]

IDS_XM = [l.strip() for l in open("data/ids_xm_full.txt", encoding="utf-8") if l.strip()]
IDS_MS = [l.strip() for l in open("data/ids_ms.txt", encoding="utf-8") if l.strip()]

GROUPS = {
    "easy": ("rename", IDS_XM),
    "medium": ("rename_medium", IDS_XM),
    "hard": ("rename_hard", IDS_XM),
    "merge": ("merge_col", IDS_MS),
    "split": ("split_col", IDS_MS),
}
CONTRASTS = [("merge", "c3mp", "c3m", "合并：段位 vs 列级"),
             ("merge", "c3mp", "c_rw", "合并：段位 vs 直接改写"),
             ("merge", "c3m", "c2", "合并：列级 vs 无谱系"),
             ("split", "c3mp", "c_rw", "拆分：段位 vs 直接改写"),
             ("split", "c3m", "c2", "拆分：列级 vs 无谱系"),
             ("easy", "c3m", "c_rw", "改名 easy"),
             ("medium", "c3m", "c_rw", "改名 medium"),
             ("hard", "c3m", "c_rw", "改名 hard")]


def ids_of(g):
    evo, pool = GROUPS[g]
    return [i for i in pool if MAN[i]["evo_type"] == evo]


def opsig(i):
    return tuple(tuple(x) for x in MAN[i].get("ops") or [])


def oks(i, c):
    return 1 if BY[i].get(c) == "EX" else 0


def sign_p(pos, neg):
    """精确双侧符号检验。"""
    n = pos + neg
    if n == 0:
        return 1.0
    k = min(pos, neg)
    return min(1.0, 2.0 * sum(math.comb(n, j) for j in range(k + 1)) / (2 ** n))


def boot_ci(units, reps=BOOT, seed=SEED):
    """按单元重采样的百分位 bootstrap 区间（百分点）；单元内条目一起进出。"""
    rnd = random.Random(seed)
    n = len(units)
    if n < 3:
        return None
    vals = []
    for _ in range(reps):
        num = den = 0.0
        for _ in range(n):
            u = units[rnd.randrange(n)]
            num += u["sa"] - u["sb"]
            den += u["n"]
        if den:
            vals.append(num / den * 100)
    vals.sort()
    return [round(vals[int(0.025 * len(vals))], 1),
            round(vals[min(len(vals) - 1, int(0.975 * len(vals)))], 1)]


def cluster_test(units):
    """簇级：等权簇均值差、符号检验、簇 bootstrap 区间。"""
    n = len(units)
    if n == 0:
        return {}
    pos = sum(1 for u in units if u["sa"] / u["n"] > u["sb"] / u["n"])
    neg = sum(1 for u in units if u["sa"] / u["n"] < u["sb"] / u["n"])
    return {"n_units": n, "n_positive": pos, "n_negative": neg,
            "n_tie": n - pos - neg,
            "mean_cluster_diff_pt": round(
                sum((u["sa"] - u["sb"]) / u["n"] for u in units) / n * 100, 1),
            "p_sign": round(sign_p(pos, neg), 6),
            "p_min": round(2 / 2 ** n, 6),
            "ci_boot": boot_ci(units)}


out = {"version": VERSION, "bootstrap": BOOT, "seed": SEED,
       "levels": {
           "item": "条目级：一条演化样本记一次（正文配对检验用这一层）",
           "question": "基础题级：同一基础题在多档重复时聚为一簇",
           "op_atom": "原子 op 级：一个列级变更动作为一簇（合并档 10 个、拆分档 5 个）",
           "op_event": "变更事件级：同库同表的一次列重整（合并档 5 个，因其 10 个原子 op 两两属同一次合并）",
           "db": "数据库级：一个库为一簇（外部效度的单元）"},
       "unit_note": "合并档的 10 个原子 op 两两来自同一次列合并（一个合并列的两个来源），"
                    "因此独立变更事件只有 5 个；把它当作 10 个独立事实会高估功效。",
       "contrasts": []}

print("=" * 100)
print("同一组对照在四个分析单位上的强度（Δ 单位：百分点）")
print("=" * 100)

for g, a, b, label in CONTRASTS:
    ids = [i for i in ids_of(g) if a in BY[i] and b in BY[i]]
    if not ids:
        continue
    n_it = len(ids)
    ao = sum(1 for i in ids if oks(i, a) and not oks(i, b))
    bo = sum(1 for i in ids if oks(i, b) and not oks(i, a))
    rec = {"group": g, "a": a, "b": b, "label": label, "n_items": n_it,
           "levels": {}}
    rec["levels"]["item"] = {
        "n_units": n_it, "a_only": ao, "b_only": bo,
        "diff_pt": round((sum(oks(i, a) for i in ids) - sum(oks(i, b) for i in ids)) / n_it * 100, 1),
        "p": sign_p(ao, bo),
        "p_min": round(2 / 2 ** (ao + bo), 8) if (ao + bo) else 1.0,
        "note": "p_min 为当前不一致对数下的最小可达双侧 p"}

    def pack(keyfn):
        """keyfn 返回该条目所属的键列表（条目可同时属于多个键，如一次合并的两个来源列）。"""
        d = collections.defaultdict(list)
        for i in ids:
            for k in keyfn(i):
                d[k].append(i)
        return [{"key": k, "n": len(v),
                 "sa": sum(oks(i, a) for i in v), "sb": sum(oks(i, b) for i in v)}
                for k, v in d.items()]

    def atoms(i):
        return [(MAN[i]["db_id"], tuple(x)) for x in MAN[i].get("ops") or []]

    def events(i):
        """一次 schema 变更事件 = 同库同表的一次列重整。

        合并档里 10 个原子 op 两两属于同一次合并（同一个合并列的两个来源列，
        表达式互为 SUBSTR 的前后两段），按表聚簇后每库剩 1 次事件；
        拆分档本就每库 1 次。"""
        return sorted({(MAN[i]["db_id"], str(x[1]))
                       for x in MAN[i].get("ops") or []})

    q = cluster_test(pack(lambda i: [MAN[i]["question_id"]]))
    q["n_multi_item_clusters"] = sum(
        1 for _k, v in collections.Counter(MAN[i]["question_id"] for i in ids).items() if v > 1)

    au = pack(atoms)
    oa = cluster_test(au)
    oa["n_units"] = len(au)

    eu = pack(events)
    oe = cluster_test(eu)
    oe["n_units"] = len(eu)
    oe["items_per_event_min"] = min(u["n"] for u in eu)
    oe["items_per_event_max"] = max(u["n"] for u in eu)

    dbu = pack(lambda i: [MAN[i]["db_id"]])
    d = cluster_test(dbu)
    d["per_db"] = {u["key"]: round((u["sa"] - u["sb"]) / u["n"] * 100, 1) for u in dbu}
    rec["levels"].update({"question": q, "op_atom": oa, "op_event": oe, "db": d})
    out["contrasts"].append(rec)

    L = rec["levels"]
    print(f"\n【{label}】  n_items={n_it}")
    print(f"  条目级    n={n_it:4d}  Δ={L['item']['diff_pt']:+6.1f}  "
          f"p={L['item']['p']:.2e}  不一致 {ao}:{bo}")
    print(f"  基础题级  n={q['n_units']:4d}  Δ={q['mean_cluster_diff_pt']:+6.1f}  "
          f"p={q['p_sign']:.4f}  {q['n_positive']}正/{q['n_negative']}负/{q['n_tie']}平  "
          f"CI={q['ci_boot']}")
    print(f"  原子 op   n={oa['n_units']:4d}  Δ={oa['mean_cluster_diff_pt']:+6.1f}  "
          f"p={oa['p_sign']:.4f}  {oa['n_positive']}正/{oa['n_negative']}负/{oa['n_tie']}平  "
          f"CI={oa['ci_boot']}  最小可达 p={oa['p_min']:.4f}")
    print(f"  变更事件  n={oe['n_units']:4d}  Δ={oe['mean_cluster_diff_pt']:+6.1f}  "
          f"p={oe['p_sign']:.4f}  {oe['n_positive']}正/{oe['n_negative']}负/{oe['n_tie']}平  "
          f"CI={oe['ci_boot']}  最小可达 p={oe['p_min']:.4f}  每事件 {oe['items_per_event_min']}-{oe['items_per_event_max']} 条")
    print(f"  数据库级  n={d['n_units']:4d}  Δ={d['mean_cluster_diff_pt']:+6.1f}  "
          f"p={d['p_sign']:.4f}  {d['n_positive']}正/{d['n_negative']}负/{d['n_tie']}平  "
          f"CI={d['ci_boot']}")

# ================================================================ 跨档池化
# 单档各自只有 5 个变更事件，精确符号检验的最小可达 p = 0.0625，无法在 0.05 水平判定。
# 合并档与拆分档施加的是同一批 5 个库上的两类相反编码改变（多列→一列、一列→多列），
# 若统一到同一对照口径（C3mp − c_rw），可把 10 次不同的列重整事件放在一起做符号检验。
# 必须同时报库级口径：库仍只有 5 个，事件级 10 个不等于库级 10 个。

POOLED = [("段位 vs 直接改写（两类编码改变池化）",
           [("merge", "c3mp", "c_rw"), ("split", "c3mp", "c_rw")])]


def event_units(g, a, b):
    ids = [i for i in ids_of(g) if a in BY[i] and b in BY[i]]
    d = collections.defaultdict(list)
    for i in ids:
        for k in sorted({(MAN[i]["db_id"], str(x[1])) for x in MAN[i].get("ops") or []}):
            d[k].append(i)
    return [{"key": k, "group": g, "n": len(v),
             "sa": sum(oks(i, a) for i in v), "sb": sum(oks(i, b) for i in v)}
            for k, v in d.items()]


out["pooled"] = []
print("\n" + "=" * 100)
print("跨档池化：统一对照口径后的事件级与库级（不是 10 个库，仍是 5 个库）")
print("=" * 100)

for label, spec in POOLED:
    evs = []
    for g, a, b in spec:
        evs += event_units(g, a, b)
    n_ev = len(evs)
    pos = sum(1 for e in evs if e["sa"] / e["n"] > e["sb"] / e["n"])
    neg = sum(1 for e in evs if e["sa"] / e["n"] < e["sb"] / e["n"])
    tie = n_ev - pos - neg
    # 库级：每库内的事件等权平均，再对库做符号检验
    per = collections.defaultdict(list)
    for e in evs:
        per[e["key"][0]].append((e["sa"] - e["sb"]) / e["n"])
    db_vals = {k: sum(v) / len(v) for k, v in per.items()}
    dpos = sum(1 for v in db_vals.values() if v > 0)
    dneg = sum(1 for v in db_vals.values() if v < 0)
    n_db = len(db_vals)
    rec = {
        "label": label, "spec": [f"{g}:{a}-{b}" for g, a, b in spec],
        "n_events": n_ev, "n_db": n_db,
        "event_level": {
            "n_positive": pos, "n_negative": neg, "n_tie": tie,
            "mean_diff_pt": round(sum((e["sa"] - e["sb"]) / e["n"] for e in evs) / n_ev * 100, 1),
            "p_sign": round(sign_p(pos, neg), 6),
            "p_min": round(2 / 2 ** n_ev, 6),
            "events": [{"db": e["key"][0], "table": e["key"][1], "group": e["group"],
                        "n": e["n"], "diff_pt": round((e["sa"] - e["sb"]) / e["n"] * 100, 1)}
                       for e in evs]},
        "db_level": {
            "n_positive": dpos, "n_negative": dneg, "n_tie": n_db - dpos - dneg,
            "p_sign": round(sign_p(dpos, dneg), 6),
            "p_min": round(2 / 2 ** n_db, 6),
            "per_db_pt": {k: round(v * 100, 1) for k, v in sorted(db_vals.items())}},
        "reading": "事件级 n=10、库级 n=5；两类编码改变在同批 5 个库上的 10 次不同列重整，"
                   "因此事件级显著不等于跨库泛化，外部效度仍受 5 个库限制。",
    }
    out["pooled"].append(rec)
    print(f"\n【{label}】")
    print(f"  事件级  n={n_ev:2d}  Δ={rec['event_level']['mean_diff_pt']:+6.1f}  "
          f"p={rec['event_level']['p_sign']:.4f}  "
          f"{pos}正/{neg}负/{tie}平  最小可达 p={rec['event_level']['p_min']:.4f}")
    print(f"  库级    n={n_db:2d}  p={rec['db_level']['p_sign']:.4f}  "
          f"{dpos}正/{dneg}负/{n_db-dpos-dneg}平  最小可达 p={rec['db_level']['p_min']:.4f}")
    print("  逐库:", rec["db_level"]["per_db_pt"])

json.dump(out, open("results/op_level.json", "w", encoding="utf-8"),
          ensure_ascii=False, indent=1)
print("\n-> results/op_level.json")
