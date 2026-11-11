# -*- coding: utf-8 -*-
"""m1_holdout_eval.py - 留出库上 M1 段位表达式的直接评估（P1 第 1 项）。

审稿 P1 要求「在独立留出库上单独评估 M1 段位抽取质量，而不仅是 SQL 执行结果」。
此前留出库只报 C1/C2/C3m 的执行结果一致率；本脚本把 M1 推断出的段位表达式
与 gold drift_ops[].new 逐字符比对。

独立性口径（沿用第 13/14 轮）：条目 481 -> 唯一 op 16 -> 独立事件 8 -> 库 8。
不得把 481 条当成 481 个独立样本。
留出库来自 BIRD train 划分（主实验为 dev 划分），因此阈值校准与数据划分两重独立。

零 API 调用；复用 m1_part.infer_db / to_expr。
输出 results/m1_holdout_eval.json
"""
import collections
import json
import math
import os
import re

BASE = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE)

OLD = "data/bird_train/train/train_databases"
EVO = "data/evolved_merge_holdout"
os.environ["OLD_DBDIR"] = OLD
os.environ["EVO_DIR"] = EVO

import m1_part as mp

VERSION = "m1-holdout-eval-v1"
GOLD = "data/evolution_merge_holdout.jsonl"
OUT = "results/m1_holdout_eval.json"


def norm(s):
    return re.sub(r"\s+", "", str(s or "")).lower()


def wilson(k, n, z=1.96):
    if n == 0:
        return [0.0, 0.0]
    p = k / n
    d = 1 + z * z / n
    c = (p + z * z / (2 * n)) / d
    h = z * math.sqrt(p * (1 - p) / n + z * z / (4 * n * n)) / d
    return [round(max(0.0, c - h) * 100, 1), round(min(1.0, c + h) * 100, 1)]


def attr_of(expr):
    m = re.search(r"attr_\d+", str(expr), re.I)
    return m.group(0).lower() if m else ""


rows = [json.loads(l) for l in open(GOLD, encoding="utf-8")]
ops = collections.Counter()
for r in rows:
    for o in r["drift_ops"]:
        if o["op"] != "merge_columns":
            continue
        ops[(r["db_id"], o["table"], o["old"], o["new"])] += 1

dbs = sorted({k[0] for k in ops})
infer = {db: (mp.infer_db(db) or {}) for db in dbs}

detail = []
for (db, tab, old, new), cnt in sorted(ops.items()):
    found = None
    for t, d in infer[db].items():
        if t.lower() != tab.lower():
            continue
        for mc, meta in d.items():
            parts = {k.lower(): v for k, v in meta["parts"].items()}
            if old.lower() in parts:
                p = parts[old.lower()]
                expr = mp.to_expr(mc, p["idx"], p["nseg"], meta["sep"])
                found = {"attr": mc, "idx": p["idx"], "nseg": p["nseg"],
                         "score": p["score"], "expr": expr}
                break
        if found:
            break
    if found is None:
        status = "no_infer"
    elif norm(found["expr"]) == norm(new):
        status = "matched"
    else:
        status = "expr_mismatch"
    detail.append({"db": db, "table": tab, "old": old, "gold_expr": new,
                   "items": cnt,
                   "event": "%s|%s|%s" % (db, tab.lower(), attr_of(new)),
                   "status": status, "inferred": found})


def group_stat(keyfn, label):
    g = collections.defaultdict(list)
    for x in detail:
        g[keyfn(x)].append(x)
    out = []
    for k, xs in sorted(g.items(), key=lambda kv: str(kv[0])):
        hit = sum(1 for x in xs if x["status"] == "matched")
        out.append({"key": str(k), "n": len(xs), "matched": hit,
                    "rate_pct": round(hit / len(xs) * 100, 1)})
    n = len(out)
    full = sum(1 for x in out if x["matched"] == x["n"])
    return {"unit": label, "n_units": n, "fully_recovered": full,
            "rate_pct": round(full / n * 100, 1) if n else 0.0,
            "wilson95_pct": wilson(full, n), "detail": out}


hit_op = sum(1 for x in detail if x["status"] == "matched")
op_stat = {"unit": "atomic_op", "n_units": len(detail), "matched": hit_op,
           "rate_pct": round(hit_op / len(detail) * 100, 1),
           "wilson95_pct": wilson(hit_op, len(detail)),
           "by_status": dict(collections.Counter(x["status"] for x in detail))}

ev_stat = group_stat(lambda x: x["event"], "change_event")
db_stat = group_stat(lambda x: x["db"], "database")

report = {
    "version": VERSION,
    "source": {"old_dbdir": OLD, "evo_dir": EVO,
               "note": "BIRD train split; main experiments use dev split. "
                       "M1 thresholds were calibrated on dev, so split and calibration are independent."},
    "units": {"items": len(rows), "unique_ops": len(ops),
              "events": ev_stat["n_units"], "dbs": len(dbs)},
    "op_level": op_stat,
    "event_level": ev_stat,
    "db_level": db_stat,
    "failures": [x for x in detail if x["status"] != "matched"],
    "method_note": "M1 part branch only; no shape fingerprint, no B4. "
                   "Expression comparison is character-level after whitespace/case normalization. "
                   "Only unique ops are independent; per-item counts are not used as denominators.",
}
json.dump(report, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("saved", OUT)
print("items %d -> unique_op %d -> event %d -> db %d" %
      (report["units"]["items"], report["units"]["unique_ops"],
       report["units"]["events"], report["units"]["dbs"]))
print("op   : %d/%d = %.1f%%  wilson %s" %
      (hit_op, len(detail), op_stat["rate_pct"], op_stat["wilson95_pct"]))
print("event: %d/%d = %.1f%%" %
      (ev_stat["fully_recovered"], ev_stat["n_units"], ev_stat["rate_pct"]))
print("db   : %d/%d = %.1f%%" %
      (db_stat["fully_recovered"], db_stat["n_units"], db_stat["rate_pct"]))
print("by_status:", op_stat["by_status"])
for x in report["failures"][:10]:
    print("  FAIL", x["db"], x["old"], x["status"],
          (x["inferred"] or {}).get("expr"), "| gold:", x["gold_expr"])
