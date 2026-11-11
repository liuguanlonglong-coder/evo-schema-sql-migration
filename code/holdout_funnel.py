# -*- coding: utf-8 -*-
"""留出库筛选漏斗（回答第七轮评审第二节末段）。

评审要求：把筛选规则与**所有未入选原因**报清楚，不得把这 7 个库说成
BIRD train 的代表性样本。本脚本按阶段复原漏斗，每一步都记录剔除原因与库名。

阶段
  S0  BIRD train 全部库
  S1  剔除单库文件 >= 110 MB（DROP COLUMN 需重写整表，构造阶段会跳过）
  S2  剔除 train 划分内题数 < 20（不足以做配对检验）
  S3  剔除与 dev 11 库重叠（留出的定义；train 与 dev 本无交集，实测 0）
  S4  施加演化后实际产出条目的库（无可用列对或无统一分隔符的库在此出局）
  S5  自检通过（旧 SQL 在旧库 与 迁移 SQL 在新库 的结果多重集一致）
  S6  分层抽样：每（档, 库）按 id 的 md5 取前 15

输出 results/holdout_funnel.json
"""
import collections
import json
import os

BASE = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE)

VERSION = "holdout-funnel-v1"
BIG_MB = 110
MIN_Q = 20

CAND = json.load(open("data/holdout_db_candidates.json", encoding="utf-8"))
ROWS = CAND["dbs"]

MAN = [json.loads(l) for l in open("data/manifest_holdout.jsonl", encoding="utf-8")]
SMP = [json.loads(l) for l in open("data/manifest_holdout_sample.jsonl", encoding="utf-8")]

stages = []
names = set()

s0 = len(ROWS)
stages.append({"stage": "S0", "label": "BIRD train 全部库", "n_db": s0,
               "n_items": CAND["n_items"], "removed": [], "reason": "起点"})

big = sorted(r["db"] for r in ROWS if (r.get("mb") or 0) >= BIG_MB)
stages.append({"stage": "S1", "label": "单库文件 < %d MB" % BIG_MB, "n_db": s0 - len(big),
               "removed": big,
               "reason": "DROP COLUMN 会重写整表，超大库在构造阶段被跳过"})

rest = [r for r in ROWS if (r.get("mb") or 0) < BIG_MB]
few = sorted(r["db"] for r in rest if (r.get("n_questions") or 0) < MIN_Q)
stages.append({"stage": "S2", "label": "该库题数 >= %d" % MIN_Q,
               "n_db": s0 - len(big) - len(few), "removed": few,
               "reason": "题数过少，配对检验没有意义"})

rest2 = [r for r in rest if (r.get("n_questions") or 0) >= MIN_Q]
dev = sorted(r["db"] for r in rest2 if r.get("in_dev"))
stages.append({"stage": "S3", "label": "与 dev 11 库不重叠", "n_db": len(rest2) - len(dev),
               "removed": dev, "reason": "「留出」的定义；train 与 dev 本无交集"})

eligible = {r["db"] for r in rest2 if not r.get("in_dev")}
evolved = {r["db_id"] for r in MAN}
noevo = sorted(eligible - evolved)
stages.append({"stage": "S4", "label": "演化可施加（有可用列对 / 统一分隔符）",
               "n_db": len(evolved), "n_items": len(MAN), "removed": noevo,
               "reason": "合并需同表两列均为文本且各 >= 5 个不同值；"
                         "拆分需某文本列的**全部**非空值都含同一分隔符"})

by_db_ok = collections.defaultdict(lambda: [0, 0])
for r in MAN:
    by_db_ok[r["db_id"]][0 if r.get("check_pass") else 1] += 1
fail_db = sorted(d for d, (a, b) in by_db_ok.items() if a == 0)
part_fail = sorted(d for d, (a, b) in by_db_ok.items() if b and a)
stages.append({"stage": "S5", "label": "自检通过（旧结果 == 迁移后结果）",
               "n_db": len([d for d in evolved if d not in fail_db]),
               "n_items": sum(1 for r in MAN if r.get("check_pass")),
               "removed": fail_db, "partial": part_fail,
               "per_db": {d: {"pass": a, "fail": b} for d, (a, b) in sorted(by_db_ok.items())},
               "reason": "整库 0 条通过者出局；部分失败者保留，只抽通过的条目"})

stages.append({"stage": "S6", "label": "每（档, 库）按 id md5 取前 15",
               "n_db": len({r["db_id"] for r in SMP}), "n_items": len(SMP),
               "removed": [], "reason": "分层等量抽样，避免依赖文件顺序"})

cells = collections.Counter((r["evo_type"], r["db_id"]) for r in SMP)
out = {"version": VERSION, "big_mb": BIG_MB, "min_q": MIN_Q,
       "stages": stages,
       "final_dbs": sorted({r["db_id"] for r in SMP}),
       "final_cells": {f"{k[0]}|{k[1]}": v for k, v in sorted(cells.items())},
       "n_final": len(SMP),
       "by_evo": dict(collections.Counter(r["evo_type"] for r in SMP))}
json.dump(out, open("results/holdout_funnel.json", "w", encoding="utf-8"),
          ensure_ascii=False, indent=1)

for s in stages:
    print("%s %-34s 库 %3d  条 %5s  剔除 %d"
          % (s["stage"], s["label"], s["n_db"], s.get("n_items", ""), len(s["removed"])))
print("\n最终 %d 库 %d 条" % (len(out["final_dbs"]), out["n_final"]))
print("按档:", out["by_evo"])
print("-> results/holdout_funnel.json")
