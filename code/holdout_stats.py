# -*- coding: utf-8 -*-
"""留出库确认实验的统计：各条件 EX 率、配对 McNemar、逐库与事件级方向。

口径与第 7 章主实验一致：
  · 分母剔除 LLM_ERROR（接口失败不计入任何条件）
  · 配对检验用精确双侧 McNemar（条件二项，p = 0.5）
  · 事件级 = 同库同表的一次列重整，用精确符号检验

输入 results/e1_holdout_<db>.jsonl
输出 results/holdout_stats.json
"""
import collections
import glob
import json
import math
import os

BASE = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE)

VERSION = "holdout-stats-v1"
CONDS = ["c2", "c3m", "c3mp"]

rows = []
for p in sorted(glob.glob("results/e1_holdout_*.jsonl")):
    for l in open(p, encoding="utf-8"):
        rows.append(json.loads(l))
if not rows:
    raise SystemExit("没有留出实验结果")

MAN = {}
for l in open("data/manifest_holdout_sample.jsonl", encoding="utf-8"):
    r = json.loads(l)
    MAN[r["id"]] = r


def mcnemar(b, c):
    n = b + c
    if n == 0:
        return 1.0
    k = min(b, c)
    return min(1.0, 2.0 * sum(math.comb(n, j) for j in range(k + 1)) / (2 ** n))


def sign_p(pos, neg):
    return mcnemar(pos, neg)


# 每个条件下每条记录是否 EX
ok = collections.defaultdict(dict)      # (evo, id) -> cond -> 1/0
err = collections.defaultdict(set)      # (evo, id) -> {cond} 接口失败的
meta = {}
for r in rows:
    k = (r["evo_type"], r["id"])
    meta[k] = {"db": r["db_id"], "qid": r["question_id"]}
    if r["outcome"] == "LLM_ERROR":
        err[k].add(r["cond"])
        continue
    ok[k][r["cond"]] = 1 if r["outcome"] == "EX" else 0

EVOS = sorted({k[0] for k in ok})


def rate(evo, cond):
    keys = [k for k in ok if k[0] == evo and cond in ok[k] and cond not in err[k]]
    if not keys:
        return None
    n = len(keys)
    e = sum(ok[k][cond] for k in keys)
    return {"n": n, "ex": e, "rate": round(e / n * 100, 1)}


def paired(evo, a, b):
    keys = [k for k in ok if k[0] == evo and a in ok[k] and b in ok[k]
            and a not in err[k] and b not in err[k]]
    ao = sum(1 for k in keys if ok[k][a] and not ok[k][b])
    bo = sum(1 for k in keys if ok[k][b] and not ok[k][a])
    both = sum(1 for k in keys if ok[k][a] and ok[k][b])
    neither = len(keys) - ao - bo - both
    return {"n": len(keys), "a_only": ao, "b_only": bo, "both": both,
            "neither": neither,
            "delta_pt": round((sum(ok[k][a] for k in keys) - sum(ok[k][b] for k in keys))
                              / len(keys) * 100, 1) if keys else None,
            "p": round(mcnemar(ao, bo), 6) if keys else None}


out = {"version": VERSION, "model": rows[0].get("model"), "n_records": len(rows),
       "evos": {}, "paired": {}, "per_db": {}, "event_level": {}}

for evo in EVOS:
    out["evos"][evo] = {c: rate(evo, c) for c in CONDS}
    for a, b in (("c3mp", "c2"), ("c3mp", "c3m"), ("c3m", "c2")):
        if rate(evo, a) and rate(evo, b):
            out["paired"][f"{evo}|{a}-{b}"] = paired(evo, a, b)
    # 逐库
    dbs = sorted({meta[k]["db"] for k in ok if k[0] == evo})
    for db in dbs:
        keys = [k for k in ok if k[0] == evo and meta[k]["db"] == db]
        d = {}
        for c in CONDS:
            kk = [k for k in keys if c in ok[k] and c not in err[k]]
            if kk:
                d[c] = {"n": len(kk), "rate": round(sum(ok[k][c] for k in kk) / len(kk) * 100, 1)}
        sel = [k for k in keys if "c3mp" in ok[k] and "c3m" in ok[k]]
        if sel:
            d["c3mp_minus_c3m_pt"] = round(
                (sum(ok[k]["c3mp"] for k in sel) - sum(ok[k]["c3m"] for k in sel)) / len(sel) * 100, 1)
        out["per_db"][f"{evo}|{db}"] = d
    # 事件级：同库同表的一次列重整
    ev = collections.defaultdict(list)
    for k in ok:
        if k[0] != evo:
            continue
        for op in MAN[k[1]].get("ops") or []:
            if isinstance(op, list) and len(op) >= 3:
                ev[(meta[k]["db"], str(op[1]))].append(k)
    cells = []
    for key, ks in ev.items():
        sel = [k for k in ks if "c3mp" in ok[k] and "c3m" in ok[k]]
        if not sel:
            continue
        cells.append({"db": key[0], "table": key[1], "n": len(sel),
                      "c3mp_rate": round(sum(ok[k]["c3mp"] for k in sel) / len(sel) * 100, 1),
                      "c3m_rate": round(sum(ok[k]["c3m"] for k in sel) / len(sel) * 100, 1)})
    pos = sum(1 for c in cells if c["c3mp_rate"] > c["c3m_rate"])
    neg = sum(1 for c in cells if c["c3mp_rate"] < c["c3m_rate"])
    out["event_level"][evo] = {"n_cells": len(cells), "positive": pos, "negative": neg,
                               "tie": len(cells) - pos - neg,
                               "p": round(sign_p(pos, neg), 6), "cells": cells}

json.dump(out, open("results/holdout_stats.json", "w", encoding="utf-8"),
          ensure_ascii=False, indent=1)

print(f"模型 {out['model']}，记录 {len(rows)} 条")
for evo in EVOS:
    print(f"\n【{evo}】")
    for c in CONDS:
        r = out["evos"][evo].get(c)
        if r:
            print(f"  {c:5s} n={r['n']:3d} EX={r['ex']:3d} ({r['rate']:.1f}%)")
    for k, v in out["paired"].items():
        if k.startswith(evo + "|"):
            print(f"  {k.split('|')[1]:12s} Δ={v['delta_pt']:+.1f}pt  "
                  f"{v['a_only']}:{v['b_only']}  p={v['p']}")
    e = out["event_level"][evo]
    print(f"  事件级 n={e['n_cells']}  {e['positive']}正/{e['negative']}负/{e['tie']}平  p={e['p']}")
print("\n-> results/holdout_stats.json")
