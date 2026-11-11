# -*- coding: utf-8 -*-
"""基础题级聚类配对检验（回应「伪重复」意见）。

问题：同一道基础题可派生出多条演化条目，逐条 McNemar 会把同一题的多条记录
当成独立样本，低估不确定性。

做法：
  1. 以「基础题」= (db_id, question_id) 为聚类单位；
  2. 每条对比在聚类内取平均正确率（同一基础题派生条目数不同也不影响权重）；
  3. 主检验：聚类级符号检验 —— 统计「A 优于 B」与「B 优于 A」的基础题数，
     用精确二项检验（两侧）；
  4. 区间：对基础题做 5,000 次 cluster bootstrap，报告差值 95% 百分位区间；
  5. 同时给出逐条 McNemar 的 p 作对照，说明两者差异。

输出 results/stats_cluster.json
"""
import collections
import json
import math
import os
import random
import sys

BASE = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE)
R = "results"

MAN = {}
for _l in open("data/manifest_v3.jsonl", encoding="utf-8"):
    _r = json.loads(_l)
    MAN[_r["id"]] = _r

FROZEN_REN = {x.strip() for x in open("data/ids_xm_full.txt", encoding="utf-8").read().split()
              if x.strip()}

# 只取 DeepSeek-V3.2 的记录（聚类检验在主模型上做；其他模型的同题结论见 §4.10）
def is_v32(rec):
    m = str(rec.get("model") or "")
    return ("V3.2" in m) or (m == "")


def load_all():
    by = collections.defaultdict(dict)
    for f in sorted(os.listdir(R)):
        if not (f.startswith("e1_") or f.startswith("e3_")):
            continue
        p = os.path.join(R, f)
        try:
            fh = open(p, encoding="utf-8")
        except Exception:
            continue
        for l in fh:
            l = l.strip()
            if not l:
                continue
            try:
                rec = json.loads(l)
            except Exception:
                continue
            if not is_v32(rec):
                continue
            by[rec["id"]][rec["cond"]] = rec["outcome"]
    return by


BY = load_all()


def cluster_of(i):
    r = MAN.get(i, {})
    return (r.get("db_id"), str(r.get("question_id")))


def group_ids(evos, frozen=None):
    ids = [i for i in BY if i in MAN and MAN[i]["evo_type"] in evos]
    if frozen is not None:
        ids = [i for i in ids if i in frozen]
    return sorted(ids)


def paired_items(ids, a, b):
    out = []
    for i in ids:
        v = BY.get(i, {})
        x, y = v.get(a), v.get(b)
        if x is None or y is None:
            continue
        out.append((i, x == "EX", y == "EX"))
    return out


def mcnemar_exact(b, c):
    n = b + c
    if n == 0:
        return 1.0
    k = min(b, c)
    s = sum(math.comb(n, i) for i in range(0, k + 1)) / (2 ** n)
    return min(1.0, 2 * s)


def bootstrap_diff(clusters, a_scores, b_scores, B=5000, seed=7):
    keys = sorted(clusters)
    if not keys:
        return (0.0, 0.0)
    rnd = random.Random(seed)
    diffs = []
    n = len(keys)
    for _ in range(B):
        pick = [keys[rnd.randrange(n)] for _ in range(n)]
        sa = sum(a_scores[k] for k in pick)
        sb = sum(b_scores[k] for k in pick)
        diffs.append((sa - sb) / n * 100)
    diffs.sort()
    lo = diffs[int(0.025 * B)]
    hi = diffs[int(0.975 * B) - 1]
    return (lo, hi)


def analyse(name, ids, a, b):
    pr = paired_items(ids, a, b)
    if not pr:
        return None
    # 聚类内平均
    agg = collections.defaultdict(lambda: [[], []])
    for i, x, y in pr:
        k = cluster_of(i)
        agg[k][0].append(1.0 if x else 0.0)
        agg[k][1].append(1.0 if y else 0.0)
    a_scores, b_scores = {}, {}
    for k, (xs, ys) in agg.items():
        a_scores[k] = sum(xs) / len(xs)
        b_scores[k] = sum(ys) / len(ys)
    better = sum(1 for k in agg if a_scores[k] > b_scores[k] + 1e-12)
    worse = sum(1 for k in agg if b_scores[k] > a_scores[k] + 1e-12)
    tie = len(agg) - better - worse
    p_cluster = mcnemar_exact(better, worse)
    lo, hi = bootstrap_diff(list(agg), a_scores, b_scores)
    point = (sum(a_scores.values()) - sum(b_scores.values())) / len(agg) * 100
    # 逐条对照
    a_only = sum(1 for _, x, y in pr if x and not y)
    b_only = sum(1 for _, x, y in pr if y and not x)
    p_item = mcnemar_exact(a_only, b_only)
    rate_a = sum(1 for _, x, _ in pr if x) / len(pr) * 100
    rate_b = sum(1 for _, x, y in pr if y) / len(pr) * 100
    return {
        "name": name, "a": a, "b": b, "n_items": len(pr), "n_clusters": len(agg),
        "rate_a": rate_a, "rate_b": rate_b, "diff_item_pt": rate_a - rate_b,
        "a_only": a_only, "b_only": b_only, "p_item": p_item,
        "better": better, "worse": worse, "tie": tie,
        "diff_cluster_pt": point, "ci_cluster": [lo, hi], "p_cluster": p_cluster,
    }


REN = group_ids({"rename", "rename_medium", "rename_hard"}, FROZEN_REN)
REN_HARD = group_ids({"rename_hard"}, FROZEN_REN)
MERGE = group_ids({"merge_col"})
SPLIT = group_ids({"split_col"})

TESTS = [
    ("改名档：C3m − C2（谱系 vs 旧 SQL）", REN, "c3m", "c2"),
    ("改名 hard：C3m − c_rw（显式谱系 vs 自行改写）", REN_HARD, "c3m", "c_rw"),
    ("合并档：C3mp − C2（段位谱系 vs 旧 SQL）", MERGE, "c3mp", "c2"),
    ("合并档：C3mp − C3m（段位 vs 列级）", MERGE, "c3mp", "c3m"),
    ("拆分档：C3mp − C2", SPLIT, "c3mp", "c2"),
]

res = {"tests": [], "note": "聚类单位 = (db_id, question_id)；主检验为聚类级符号检验，区间为聚类 bootstrap"}
for name, ids, a, b in TESTS:
    r = analyse(name, ids, a, b)
    if r:
        res["tests"].append(r)

json.dump(res, open(os.path.join(R, "stats_cluster.json"), "w", encoding="utf-8"),
          ensure_ascii=False, indent=1)

print(f"{'对比':46s} {'条':>5s} {'簇':>4s} {'Δ条pt':>7s} {'Δ簇pt':>7s} {'簇CI':>18s} "
      f"{'p簇':>9s} {'p条':>9s}")
for t in res["tests"]:
    print(f"{t['name'][:44]:46s} {t['n_items']:5d} {t['n_clusters']:4d} "
          f"{t['diff_item_pt']:+7.1f} {t['diff_cluster_pt']:+7.1f} "
          f"[{t['ci_cluster'][0]:+6.1f},{t['ci_cluster'][1]:+6.1f}] "
          f"{t['p_cluster']:9.3g} {t['p_item']:9.3g}")
print("\nSAVED results/stats_cluster.json")
