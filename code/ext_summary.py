# -*- coding: utf-8 -*-
"""扩充事件集的统计汇总（供正文表格直接引用）。

只读三份构造产物与三份抽取产物，不重算任何构造：
  results/evolution_ext_<task>.json   每事件明细（表、列、题数、通过数）
  data/manifest_ext_<task>.jsonl      逐题清单
  data/lineage_infer_ext/*             M1 抽取结果

输出 results/ext_events_summary.json

口径约定（写进正文时必须同时出现）
  events            构造出来的事件数
  events_effective  至少 1 道题通过四步验证的事件数（0 题的片段无法评估）
  events_ge3        至少 3 道题通过的事件数（事件级检验可用的事件）
  tables / dbs      事件落在多少张表 / 多少个库上
  separators        用到的分隔符字面量种类（拆分档）
"""
import collections
import json
import os

BASE = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE)

TASKS = ["merge_col", "split_col", "value_shift"]
out = {"version": "ext-events-summary-v1", "by_task": {}}

for t in TASKS:
    p = f"results/evolution_ext_{t}.json"
    if not os.path.exists(p):
        continue
    st = json.load(open(p, encoding="utf-8"))
    pe = st["per_event"]
    rows = [json.loads(l) for l in
            open(f"data/manifest_ext_{t}.jsonl", encoding="utf-8")]
    eff = [e for e in pe if e["n_pass"] > 0]
    ge3 = [e for e in pe if e["n_pass"] >= 3]

    d = {
        "events": len(pe),
        "events_effective": len(eff),
        "events_ge3": len(ge3),
        "dbs": len({e["db"] for e in pe}),
        "tables": len({(e["db"], e["table"]) for e in pe}),
        "items": len(rows),
        "items_pass": sum(1 for r in rows if r["check_pass"]),
        "pass_rate": round(sum(1 for r in rows if r["check_pass"])
                           / max(1, len(rows)) * 100, 1),
        "items_per_event": round(len(rows) / max(1, len(pe)), 1),
        "items_per_effective_event": round(len(rows) / max(1, len(eff)), 1),
        "per_db_events": dict(collections.Counter(e["db"] for e in pe)),
        "max_event_items": max((e["n_affected"] for e in pe), default=0),
        "max_event_share_pct": round(max((e["n_affected"] for e in pe), default=0)
                                     / max(1, len(rows)) * 100, 1),
    }
    if t == "split_col":
        seps = collections.Counter(r.get("sep") for r in rows if r.get("sep"))
        d["separators"] = {k: v for k, v in seps.items()}
        d["separator_kinds"] = len(seps)
    # 每事件的列组合（说明「不同字段组合」到底换了哪些列）
    d["column_pairs"] = [
        {"db": e["db"], "table": e["table"], "cols": e["cols"],
         "n_items": e["n_affected"], "n_pass": e["n_pass"]}
        for e in sorted(pe, key=lambda x: (-x["n_pass"], x["db"]))]
    out["by_task"][t] = d

# 与主实验的对照：主实验各档事件数（audit_independence v2）
adi = json.load(open("results/audit_independence.json", encoding="utf-8"))
ps = adi["A_units"]["paper_set"]
seg = adi["A_units"]["segment_set"]
out["main_experiment"] = {
    "paper_set": ps, "segment_set": seg,
    "note": ("主实验 610 条 / 90 op / 61 事件 / 11 库；其中段位相关（合并+拆分）"
             "334 条 / 15 op / 10 事件 / 5 库。扩充事件集只用于事件级复核，"
             "不与主实验条目合并计算，也不作为主实验的分母。"),
}

json.dump(out, open("results/ext_events_summary.json", "w", encoding="utf-8"),
          ensure_ascii=False, indent=1)

print(f"{'task':12s} {'事件':>4s} {'有效':>4s} {'≥3题':>5s} {'库':>3s} {'表':>3s} "
      f"{'题':>5s} {'通过':>5s} {'通过率':>6s} {'最大事件占比':>10s}")
for t, d in out["by_task"].items():
    print(f"{t:12s} {d['events']:4d} {d['events_effective']:4d} {d['events_ge3']:5d} "
          f"{d['dbs']:3d} {d['tables']:3d} {d['items']:5d} {d['items_pass']:5d} "
          f"{d['pass_rate']:5.1f}% {d['max_event_share_pct']:9.1f}%")
    if "separator_kinds" in d:
        print("   分隔符:", d["separators"])
