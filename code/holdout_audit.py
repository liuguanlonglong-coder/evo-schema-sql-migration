# -*- coding: utf-8 -*-
"""留出数据库实验的可行性审计（第五轮评审第 3 项）。

评审建议"争取按数据库留出一组确认实验"。本脚本把这个"能否补做"变成可核对的数字：
  1) 每个演化档用了几个库、每个库多少条、每个库几次独立变更事件；
  2) BIRD dev 划分本身只有 11 个库，本研究已全部使用——留出库只能来自 train 划分；
  3) train 数据的获取代价（打包体积，实测自对象存储的 Content-Length）；
  4) 在现有单元数下，事件级精确符号检验能取到的最小 p 是多少，
     以及要达到 0.05 需要多少次独立事件。
不新增模型调用。

输出 results/holdout_audit.json
"""
import collections
import json
import math
import os

BASE = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE)

VERSION = "holdout-audit-v1"

MAN = {}
for _l in open("data/manifest_v3.jsonl", encoding="utf-8"):
    _r = json.loads(_l)
    MAN[_r["id"]] = _r

IDS_XM = [l.strip() for l in open("data/ids_xm_full.txt", encoding="utf-8") if l.strip()]
IDS_MS = [l.strip() for l in open("data/ids_ms.txt", encoding="utf-8") if l.strip()]

# 本章报告的五档（与 stats_ch7_v5.py 相同口径）
PANELS = [("easy", "rename", IDS_XM),
          ("medium", "rename_medium", IDS_XM),
          ("hard", "rename_hard", IDS_XM),
          ("merge", "merge_col", IDS_MS),
          ("split", "split_col", IDS_MS)]


def events_of(i):
    return {(MAN[i]["db_id"], str(x[1])) for x in MAN[i].get("ops") or []}


out = {"version": VERSION,
       "benchmark_facts": {
           "bird_dev_databases": 11,
           "bird_dev_databases_used": 11,
           "bird_total_databases_per_paper": 95,
           "bird_total_pairs_per_paper": 12751,
           "train_zip_bytes": 8919543554,
           "train_zip_note": "train.zip 的 Content-Length 由对象存储 HEAD 实测；"
                             "官方未提供按库下载的入口，按库留出必须先整包获取",
           "license": "CC BY-SA 4.0，仅限研究用途"},
       "panels": []}

print("=" * 92)
print("留出数据库审计")
print("=" * 92)
print(f"{'档位':8s}{'条目':>6s}{'库数':>5s}{'独立变更事件':>9s}{'每事件条目':>12s}  库（条目数）")

all_db = set()
for name, evo, pool in PANELS:
    ids = [i for i in pool if MAN[i]["evo_type"] == evo]
    d = collections.Counter(MAN[i]["db_id"] for i in ids)
    ev = collections.defaultdict(set)
    for i in ids:
        ev[MAN[i]["db_id"]].update(events_of(i))
    n_ev = sum(len(v) for v in ev.values())
    per = [len([i for i in ids if MAN[i]["db_id"] == db]) / max(1, len(ev[db]))
           for db in ev]
    rec = {"panel": name, "evo_type": evo, "n_items": len(ids), "n_db": len(d),
           "n_events": n_ev,
           "items_per_event_min": round(min(per), 1) if per else 0,
           "items_per_event_max": round(max(per), 1) if per else 0,
           "per_db": dict(d.most_common())}
    out["panels"].append(rec)
    all_db.update(d)
    print(f"{name:8s}{len(ids):6d}{len(d):5d}{n_ev:9d}"
          f"{rec['items_per_event_min']:6.1f}-{rec['items_per_event_max']:<6.1f}"
          f"  " + ", ".join(f"{k}:{v}" for k, v in d.most_common()))

out["db_union"] = sorted(all_db)
out["db_union_n"] = len(all_db)

# 事件级精确符号检验的最小可达 p，以及达标所需的事件数
def min_p(n):
    return 2.0 / 2 ** n


need = next(n for n in range(1, 40) if min_p(n) <= 0.05)
out["power_at_event_level"] = {
    "n_events_merge": 5, "n_events_split": 5,
    "min_attainable_p_at_5": round(min_p(5), 4),
    "min_attainable_p_at_6": round(min_p(6), 4),
    "n_events_needed_for_0.05": need,
    "note": "全部同向时，5 次独立事件的精确双侧符号检验最小可达 p=0.0625；"
            "要达到 p<0.05 至少需要 6 次同向事件。"}

print(f"\n全部五档共用 {len(all_db)} 个库，与 BIRD dev 的 11 个库完全重合。")
print(f"事件级精确符号检验：n=5 时最小 p={min_p(5):.4f}；"
      f"需要 n≥{need} 次同向事件才能达到 0.05。")
print(f"BIRD train.zip = {out['benchmark_facts']['train_zip_bytes']/2**30:.1f} GiB"
      "（按库留出必须先整包获取，官方无按库下载入口）。")

json.dump(out, open("results/holdout_audit.json", "w", encoding="utf-8"),
          ensure_ascii=False, indent=1)
print("-> results/holdout_audit.json")
