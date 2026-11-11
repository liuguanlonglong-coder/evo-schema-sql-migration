# -*- coding: utf-8 -*-
"""recompute_ext_summary.py - 重算扩充集端到端汇总（第 22 轮修复后）。

run_ext_eval.py 的汇总段落只在同一次进程里见到全部条件时才会算配对；
第 22 轮只重跑了 c3mp（c1/c2/c3m 不受谱系注入缺陷影响），
故这里按同一套定义从逐条 jsonl 重算，不改变任何判定。

定义与 run_ext_eval.py 保持一致：
  条件成功率 = EX / 该条件全部条目
  事件级配对 = 同一 (事件, 题) 下两条件判定不同才计入；p 为精确双侧符号检验
"""
import collections
import json
import math
import os

BASE = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE)

SRC = "results/ext_eval_merge_col.jsonl"
OUT = "results/ext_eval_merge_col.json"
VERSION = "ext-eval-v2"
CONDS = ["c1", "c2", "c3m", "c3mp"]


def main():
    rows = [json.loads(l) for l in open(SRC, encoding="utf-8") if l.strip()]
    man = {json.loads(l)["id"] for l in
           open("data/manifest_ext_merge_col.jsonl", encoding="utf-8") if l.strip()}
    summ = {}
    for c in CONDS:
        xs = [r for r in rows if r["cond"] == c]
        cnt = collections.Counter(r["outcome"] for r in xs)
        summ[c] = {"n": len(xs), "EX": cnt["EX"],
                   "rate": round(100 * cnt["EX"] / max(1, len(xs)), 1),
                   "by_outcome": dict(cnt),
                   "events": len({r["tag"] for r in xs})}

    paired = {}
    for hi, lo in (("c3mp", "c3m"), ("c3m", "c2"), ("c2", "c1")):
        A = {(r["tag"], r["id"]): r["outcome"] for r in rows if r["cond"] == hi}
        B = {(r["tag"], r["id"]): r["outcome"] for r in rows if r["cond"] == lo}
        per = collections.defaultdict(lambda: [0, 0])
        for k in set(A) & set(B):
            a_ok, b_ok = A[k] == "EX", B[k] == "EX"
            if a_ok and not b_ok:
                per[k[0]][0] += 1
            elif b_ok and not a_ok:
                per[k[0]][1] += 1
        pos = sum(1 for v in per.values() if v[0] > v[1])
        neg = sum(1 for v in per.values() if v[1] > v[0])
        tie = len(per) - pos - neg
        n_ = pos + neg
        p = (min(1.0, 2.0 * sum(math.comb(n_, j) for j in range(min(pos, neg) + 1))
                 / 2 ** n_) if n_ else 1.0)
        paired[f"{hi}-{lo}"] = {"pos_events": pos, "neg_events": neg, "tie": tie,
                                "n_events": len(per), "p": round(p, 8)}

    js = {"version": VERSION, "task": "merge_col", "n_items": len(man),
          "conditions": summ, "paired_by_event": paired,
          "provenance": {
              "note": "第 22 轮修复 m1_ext.py 的 EVO_DIR 注入缺陷后只重跑 c3mp；"
                      "c1/c2/c3m 不读段位谱系，不受该缺陷影响，沿用第 20 轮产物。",
              "c3mp_rerun": "results/ext_eval_merge_col.jsonl（479 条，2026-11-11）",
              "superseded": "results/ext_eval_merge_col_v1_wrongtable.json（44.9%，"
                            "其中 18/28 事件注入了另一张表的段位表达式）"}}
    json.dump(js, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    for c, v in summ.items():
        print(f"{c:5s} {v['EX']:3d}/{v['n']:<3d} {v['rate']:5.1f}%  {v['by_outcome']}")
    print("事件级配对:", json.dumps(paired, ensure_ascii=False))
    print("写入", OUT)


if __name__ == "__main__":
    main()
