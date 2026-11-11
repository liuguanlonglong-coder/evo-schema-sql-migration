# -*- coding: utf-8 -*-
"""EvoSchema 外部对照的独立性重算（回答第七轮评审第三节）。

评审指出：合并档 44 条里 41 条来自同一个 student_club 的
first_name + last_name → full_name 映射，把这些题目当独立样本做符号检验
会高估独立证据。因此需要按「独立映射 / 事件 / 库」重新报告。

单元定义
  item   （库, 问题）去重后的 81 条
  event  （库, 被演化的目标列集合）——一次列重整，是真正的独立单元
  db     8 个库

另外把两处计数差异讲清楚：
  · 3 条 recovered_kind = unknown（DDL 列集合差无法恢复映射）
    仍在 item 级 n=81 里，但在按变异类型分层的表里被剔除，故该表 n=78
  · column_changes 为空 38 条，映射由 DDL 差恢复

输出 results/evoschema_events.json
"""
import collections
import json
import math
import os

BASE = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE)

VERSION = "evoschema-events-v1"
EV = json.load(open("results/evoschema_eval.json", encoding="utf-8"))
ROWS = EV["rows"]
CONDS = ["c1", "c2", "c3m"]


def sign_p(pos, neg):
    n = pos + neg
    if n == 0:
        return 1.0
    k = min(pos, neg)
    return min(1.0, 2.0 * sum(math.comb(n, j) for j in range(k + 1)) / (2 ** n))


def exact(r, c):
    return bool((r["conds"].get(c) or {}).get("exact_cols"))


def f1(r, c):
    return float((r["conds"].get(c) or {}).get("f1") or 0.0)


# ---------------------------------------------------------------- 单元划分
def event_key(r):
    ch = r.get("changes") or []
    olds, news = [], []
    for c in ch:
        olds += c.get("old") or []
        news += c.get("new") or []
    if not olds and not news:
        # column_changes 为空时由 DDL 差恢复，targets 记在 gold_cols 之外
        return (r["db"], "recovered", tuple(sorted(r.get("gold_cols") or []))[:1])
    return (r["db"], tuple(sorted(olds)), tuple(sorted(news)))


EVENT = collections.defaultdict(list)
for r in ROWS:
    EVENT[event_key(r)].append(r)

DB = collections.defaultdict(list)
for r in ROWS:
    DB[r["db"]].append(r)


def rates(group, cond):
    xs = [r for r in group if r["conds"].get(cond)]
    if not xs:
        return None
    return {"n": len(xs),
            "exact": sum(exact(r, cond) for r in xs),
            "rate": round(sum(exact(r, cond) for r in xs) / len(xs) * 100, 1),
            "f1": round(sum(f1(r, cond) for r in xs) / len(xs), 4)}


out = {"version": VERSION, "n_items": len(ROWS),
       "unknown_kind": sum(1 for r in ROWS if r.get("recovered_kind") == "unknown"),
       "by_recovered_kind": dict(collections.Counter(r.get("recovered_kind") for r in ROWS)),
       "units": {}, "event_table": [], "db_table": [], "paired_by_unit": {}}

for name, groups in (("event", EVENT), ("db", DB)):
    out["units"][name] = len(groups)

# 逐事件明细
for key, g in sorted(EVENT.items(), key=lambda kv: -len(kv[1])):
    d = {"db": key[0], "n_items": len(g),
         "kind": collections.Counter(r.get("kind") for r in g).most_common(1)[0][0]}
    for c in CONDS:
        d[c] = rates(g, c)
    if d.get("c2") and d.get("c3m"):
        d["c3m_minus_c2_pt"] = round(d["c3m"]["rate"] - d["c2"]["rate"], 1)
    out["event_table"].append(d)

# 逐库明细
for db, g in sorted(DB.items(), key=lambda kv: -len(kv[1])):
    d = {"db": db, "n_items": len(g)}
    for c in CONDS:
        d[c] = rates(g, c)
    if d.get("c2") and d.get("c3m"):
        d["c3m_minus_c2_pt"] = round(d["c3m"]["rate"] - d["c2"]["rate"], 1)
    out["db_table"].append(d)

# 单元级符号检验：c3m vs c2
for name, groups in (("event", EVENT), ("db", DB)):
    pos = neg = tie = 0
    for key, g in groups.items():
        a = rates(g, "c3m")
        b = rates(g, "c2")
        if not a or not b:
            continue
        if a["rate"] > b["rate"]:
            pos += 1
        elif a["rate"] < b["rate"]:
            neg += 1
        else:
            tie += 1
    out["paired_by_unit"][name] = {"pos": pos, "neg": neg, "tie": tie,
                                   "n": pos + neg + tie,
                                   "p": round(sign_p(pos, neg), 6)}

# 分类型的事件单元（评审要求：合并与拆分各自有几个独立映射）
KIND_EVENT = collections.defaultdict(list)
for key, g in EVENT.items():
    k = collections.Counter(r.get("recovered_kind") for r in g).most_common(1)[0][0]
    KIND_EVENT[k].append((key, g))

out["by_kind_unit"] = {}
for k, groups in KIND_EVENT.items():
    pos = neg = tie = 0
    for _key, g in groups:
        a, b = rates(g, "c3m"), rates(g, "c2")
        if not a or not b:
            continue
        if a["rate"] > b["rate"]:
            pos += 1
        elif a["rate"] < b["rate"]:
            neg += 1
        else:
            tie += 1
    out["by_kind_unit"][k] = {
        "events": len(groups), "items": sum(len(g) for _k, g in groups),
        "max_event_items": max(len(g) for _k, g in groups),
        "pos": pos, "neg": neg, "tie": tie, "p": round(sign_p(pos, neg), 6)}

# 集中度：最大单一映射占该类型的比例（评审点名 41/44 来自同一映射）
for k, v in out["by_kind_unit"].items():
    v["max_event_share_pct"] = round(v["max_event_items"] / v["items"] * 100, 1) if v["items"] else None

# item 级（原口径）保留对照
for a, b in (("c3m", "c2"), ("c1", "c2")):
    ao = sum(1 for r in ROWS if exact(r, a) and not exact(r, b))
    bo = sum(1 for r in ROWS if exact(r, b) and not exact(r, a))
    out["paired_by_unit"]["item|%s-%s" % (a, b)] = {
        "pos": ao, "neg": bo, "tie": len(ROWS) - ao - bo,
        "n": len(ROWS), "p": round(sign_p(ao, bo), 6)}

json.dump(out, open("results/evoschema_events.json", "w", encoding="utf-8"),
          ensure_ascii=False, indent=1)

print("items %d | events %d | dbs %d | unknown %d"
      % (out["n_items"], out["units"]["event"], out["units"]["db"], out["unknown_kind"]))
print("recovered_kind:", out["by_recovered_kind"])
print("\n逐事件（按条目数降序）")
for d in out["event_table"]:
    print("  %-22s %-6s n=%2d  c2 %5.1f%%  c3m %5.1f%%  Δ %+.1f"
          % (d["db"], d["kind"], d["n_items"],
             (d["c2"] or {}).get("rate", 0), (d["c3m"] or {}).get("rate", 0),
             d.get("c3m_minus_c2_pt", 0)))
print("\n单元级 c3m vs c2:", json.dumps(out["paired_by_unit"], ensure_ascii=False))
