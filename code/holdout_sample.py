# -*- coding: utf-8 -*-
"""对留出库清单做分层抽样：每个（演化类型, 库）取固定条数，保证可复现。

留出实验的目的是「方向是否在未见库上重现」，不是精确估计效应量，
因此每层取等量样本即可；用 id 的 md5 排序，避免依赖文件顺序。

输出 data/manifest_holdout_sample.jsonl 与 data/ids_holdout.txt
"""
import collections
import hashlib
import json
import os

BASE = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE)

PER_CELL = 15

rows = [json.loads(l) for l in open("data/manifest_holdout.jsonl", encoding="utf-8")]
rows = [r for r in rows if r.get("check_pass")]
print("通过自检:", len(rows))

by = collections.defaultdict(list)
for r in rows:
    by[(r["evo_type"], r["db_id"])].append(r)

picked = []
for k in sorted(by):
    v = sorted(by[k], key=lambda r: hashlib.md5(r["id"].encode()).hexdigest())
    picked += v[:PER_CELL]
    print(f"  {k[0]:20s} {k[1]:16s} 全 {len(v):3d} -> 取 {min(PER_CELL, len(v)):2d}")

picked.sort(key=lambda r: r["id"])
with open("data/manifest_holdout_sample.jsonl", "w", encoding="utf-8") as f:
    for r in picked:
        f.write(json.dumps(r, ensure_ascii=False) + "\n")
with open("data/ids_holdout.txt", "w", encoding="utf-8") as f:
    for r in picked:
        f.write(r["id"] + "\n")

cnt = collections.Counter((r["evo_type"], r["db_id"]) for r in picked)
stats = {"n": len(picked), "per_cell": PER_CELL,
         "by_evo": dict(collections.Counter(r["evo_type"] for r in picked)),
         "by_db": dict(collections.Counter(r["db_id"] for r in picked)),
         "cells": {f"{k[0]}|{k[1]}": v for k, v in sorted(cnt.items())}}
json.dump(stats, open("data/holdout_sample_stats.json", "w", encoding="utf-8"),
          ensure_ascii=False, indent=1)
print(f"\n抽样 {len(picked)} 条，{len(set(r['db_id'] for r in picked))} 个库")
print("按档:", stats["by_evo"])
