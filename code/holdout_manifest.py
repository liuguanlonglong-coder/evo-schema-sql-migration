# -*- coding: utf-8 -*-
"""把留出库（BIRD train 划分）的演化结果转成与 manifest_v3 同构的清单。

留出库确认实验的目的是回答「在完全未参与开发的数据库上，段位谱系的优势
是否仍然存在」。因此这份清单与 manifest_v3 的区别只有两处：
  1. split 一律记为 "holdout"（不参与任何训练/阈值调整）；
  2. evo_type 带 _holdout 后缀，使下游脚本就读到独立的演化库目录。

字段与 make_manifest_v3.py 完全一致，便于 run_c1c2c3.py / m1_lineage.py
在不改代码的情况下直接消费。
"""
import argparse
import collections
import json
import os

BASE = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE)

SLICES = [
    ("merge_col_holdout", "evolution_merge_holdout.jsonl"),
    ("split_col_holdout", "evolution_split_holdout.jsonl"),
    ("merge_col3_holdout", "evolution_merge3_holdout.jsonl"),
]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="data/manifest_holdout.jsonl")
    ap.add_argument("--evo", default="", help="逗号分隔，只纳入指定 evo_type")
    a = ap.parse_args()
    keep = {x.strip() for x in a.evo.split(",") if x.strip()} or None

    records, per_slice, skipped = [], {}, []
    for evo, fn in SLICES:
        if keep and evo not in keep:
            continue
        p = os.path.join("data", fn)
        if not os.path.exists(p):
            skipped.append(fn)
            continue
        rows = [json.loads(l) for l in open(p, encoding="utf-8")]
        per_slice[evo] = len(rows)
        for r in rows:
            records.append({
                "id": f"{evo}:{r['db_id']}:{r['question_id']}",
                "evo_type": evo,
                "level": "-",
                "expected_action": r["expected_action"],
                "split": "holdout",
                "db_id": r["db_id"],
                "question_id": r["question_id"],
                "difficulty": r.get("difficulty"),
                "n_ops": len(r["drift_ops"]),
                "ops": [(o["op"], o["table"], o["old"], o["new"])
                        for o in r["drift_ops"]],
                "old_sql": r["old_sql"],
                "new_sql": r["new_sql"],
                "question": r["question"],
                "check_pass": bool(r.get("check", {}).get("pass", False)),
            })

    if not records:
        raise SystemExit(f"没有可用输入；缺失 {skipped}")

    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    with open(a.out, "w", encoding="utf-8") as f:
        for r in records:
            f.write(json.dumps(r, ensure_ascii=False) + "\n")

    by_evo = collections.Counter(r["evo_type"] for r in records)
    by_db = collections.Counter(r["db_id"] for r in records)
    by_ok = collections.Counter((r["evo_type"], r["check_pass"]) for r in records)
    stats = {"out": a.out, "records": len(records), "per_slice": per_slice,
             "missing": skipped, "by_evo": dict(by_evo), "by_db": dict(by_db),
             "pass": {f"{k[0]}|{k[1]}": v for k, v in by_ok.items()},
             "n_db": len(by_db)}
    json.dump(stats, open("data/holdout_manifest_stats.json", "w", encoding="utf-8"),
              ensure_ascii=False, indent=1)

    print(f"写出 {a.out}：{len(records)} 条，{len(by_db)} 个库")
    for evo, n in by_evo.items():
        p = by_ok[(evo, True)]
        print(f"  {evo:22s} {n:4d} 条  自检通过 {p:4d}")
    print("  库分布:", dict(by_db))


if __name__ == "__main__":
    main()
