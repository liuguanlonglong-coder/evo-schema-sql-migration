# -*- coding: utf-8 -*-
"""合并 evoschema_eval.py 的分片输出，重算总体与分层汇总。"""
import collections
import glob
import json
import os

BASE = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE)

parts = sorted(glob.glob("results/evoschema_eval.shard*.json"))
if not parts:
    raise SystemExit("没有分片文件")

rows, model, src, n_raw = [], None, None, 0
for p in parts:
    d = json.load(open(p, encoding="utf-8"))
    rows += d["rows"]
    model = model or d.get("model")
    src = src or d.get("source")
    n_raw = max(n_raw, d.get("n_raw") or 0)

conds = sorted({c for r in rows for c in r["conds"]})
summary = {}
for cond in conds:
    vals = [r["conds"][cond] for r in rows
            if cond in r["conds"] and "error" not in r["conds"][cond]]
    f1s = [v["f1"] for v in vals if v["f1"] is not None]
    summary[cond] = {
        "n": len(vals),
        "mean_f1": round(sum(f1s) / len(f1s), 4) if f1s else None,
        "exact_cols_rate": round(sum(1 for v in vals if v["exact_cols"]) / len(vals), 4) if vals else None,
        "uses_deleted_old_col_rate": round(
            sum(1 for v in vals if v["uses_deleted_old_col"]) / len(vals), 4) if vals else None,
        "n_cached": sum(1 for v in vals if v["cached"]),
        "n_error": sum(1 for r in rows if cond in r["conds"] and "error" in r["conds"][cond]),
    }

by_kind = {}
for cond in conds:
    d = collections.defaultdict(list)
    for r in rows:
        v = r["conds"].get(cond) or {}
        if "error" in v or v.get("f1") is None:
            continue
        d[r["recovered_kind"]].append(v)
    by_kind[cond] = {k: {"n": len(v),
                         "mean_f1": round(sum(x["f1"] for x in v) / len(v), 4),
                         "uses_deleted_old_col_rate": round(
                             sum(1 for x in v if x["uses_deleted_old_col"]) / len(v), 4)}
                     for k, v in d.items()}

# 配对：同一条样本上 c2 与 c3m 的列 F1 差
paired = {"n": 0, "c3m_better": 0, "c2_better": 0, "tie": 0, "delta_pt": None}
diffs = []
for r in rows:
    a, b = r["conds"].get("c2"), r["conds"].get("c3m")
    if not a or not b or a.get("f1") is None or b.get("f1") is None:
        continue
    paired["n"] += 1
    d = b["f1"] - a["f1"]
    diffs.append(d)
    if d > 1e-9:
        paired["c3m_better"] += 1
    elif d < -1e-9:
        paired["c2_better"] += 1
    else:
        paired["tie"] += 1
if diffs:
    paired["delta_pt"] = round(sum(diffs) / len(diffs) * 100, 1)

out = {"version": "evoschema-eval-v1", "source": src, "model": model,
       "n_raw": n_raw, "n_unique": len(rows), "shards": len(parts),
       "note": "EvoSchema 只提供 DDL 与 gold SQL，未提供演化后的库文件，"
               "故不评执行等价（EX），只评列引用 F1；gold 仅作列集合参照。",
       "summary": summary, "by_kind": by_kind, "paired_c2_vs_c3m": paired,
       "rows": rows}
json.dump(out, open("results/evoschema_eval.json", "w", encoding="utf-8"),
          ensure_ascii=False, indent=1)

print(f"合并 {len(parts)} 个分片，{len(rows)} 条样本")
for c, s in summary.items():
    print(f"  {c:4s} n={s['n']:3d} 列F1={s['mean_f1']}  全对={s['exact_cols_rate']}  "
          f"用已删列={s['uses_deleted_old_col_rate']}")
print("配对 c2 vs c3m:", paired)
for c, d in by_kind.items():
    print(f"  {c}: " + "  ".join(f"{k}(n={v['n']}) F1={v['mean_f1']}"
                                 for k, v in d.items()))
