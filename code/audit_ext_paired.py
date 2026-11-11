# -*- coding: utf-8 -*-
"""扩充事件集的逐事件配对复核（回答审稿意见 Q1/Q2）。

只读三份产物，不重算、不改数：
  公开材料包/outputs_ext.jsonl        四条件逐题输出（1916 = 4 × 479）
  data/manifest_ext_merge_col.jsonl   逐题清单（含 check_pass 四步验证）
  results/evolution_ext_merge_col.json 事件明细（28 事件）

输出 results/audit_ext_paired.json，并打印可直接抄进表注的口径。
"""
import collections
import json
import math
import os

BASE = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE)

CONDS = ["c1", "c2", "c3m", "c3mp"]
OUTCOMES = ["EX", "SILENT", "ABSTAIN", "ERROR"]


def binom_two_sided(k, n):
    """精确符号检验（双侧，p=0.5）。n 为有效对数（不含持平）。"""
    if n == 0:
        return 1.0
    k = min(k, n - k)
    tail = sum(math.comb(n, i) for i in range(0, k + 1)) / 2 ** n
    return min(1.0, 2 * tail)


# ---------- 1. 逐题输出 ----------
rows = [json.loads(l) for l in open("公开材料包/outputs_ext.jsonl", encoding="utf-8")]
by_cond = collections.defaultdict(dict)   # cond -> id -> outcome
for r in rows:
    by_cond[r["cond"]][r["id"]] = r["outcome"]

ids = sorted(by_cond["c3mp"])
assert all(len(by_cond[c]) == len(ids) for c in CONDS), \
    {c: len(by_cond[c]) for c in CONDS}

# ---------- 2. 逐题清单：check_pass + 事件归属 ----------
man = {}
for l in open("data/manifest_ext_merge_col.jsonl", encoding="utf-8"):
    m = json.loads(l)
    man[m["id"]] = {"pass": bool(m.get("check_pass")), "tag": m.get("tag") or m.get("shard")}
miss = [i for i in ids if i not in man]
print(f"清单命中 {len(man)} 条；输出 id 未命中 {len(miss)} 条")

# ---------- 3. 事件明细 ----------
st = json.load(open("results/evolution_ext_merge_col.json", encoding="utf-8"))
pe = st["per_event"]
tag_info = {e["tag"]: e for e in pe}

# 事件 → 题号（用输出行的 shard 字段做桥）
shard_of = {r["id"]: r.get("shard") for r in rows if r["cond"] == "c3mp"}
ev_items = collections.defaultdict(list)
for i in ids:
    ev_items[shard_of[i]].append(i)

print(f"\n构造事件 {len(pe)}；输出中出现的事件 {len(ev_items)}")
print("事件 | 题数 | 四步通过 | " + " ".join(f"{c:>5s}" for c in CONDS))
tab = []
for tag, e in sorted(tag_info.items(), key=lambda x: (-x[1]["n_affected"], x[1]["db"])):
    items = ev_items.get(tag, [])
    ex = {c: sum(1 for i in items if by_cond[c][i] == "EX") for c in CONDS}
    np_ = sum(1 for i in items if man.get(i, {}).get("pass"))
    tab.append({"tag": tag, "db": e["db"], "table": e["table"], "cols": e["cols"],
                "n": len(items), "n_pass": np_, **{f"ex_{c}": ex[c] for c in CONDS}})
    print(f"{e['db'][:18]:18s}/{e['table'][:12]:12s} {len(items):3d} {np_:3d}  "
          + " ".join(f"{ex[c]:5d}" for c in CONDS))

# ---------- 4. 事件级配对：三种口径 ----------
def paired(a, b, subset=None, rule="rate"):
    win = loss = tie = 0
    detail = []
    for t in tab:
        if subset is not None and t["n_pass"] == 0:
            continue
        na, nb = t[f"ex_{a}"], t[f"ex_{b}"]
        if rule == "rate":
            x, y = na, nb            # 同一事件内题数相同，直接比 EX 计数
        else:
            x, y = na, nb
        if x > y:
            win += 1
            detail.append((t["tag"], "+"))
        elif x < y:
            loss += 1
            detail.append((t["tag"], "-"))
        else:
            tie += 1
            detail.append((t["tag"], "="))
    n_cmp = win + loss
    return {"win": win, "loss": loss, "tie": tie, "n_events": len(detail),
            "n_cmp": n_cmp, "p": round(binom_two_sided(win, n_cmp), 8),
            "detail": detail}

res = {}
for a, b in [("c3mp", "c3m"), ("c3m", "c2"), ("c2", "c1"), ("c3mp", "c2")]:
    key = f"{a}-{b}"
    res[key] = {
        "all_events": paired(a, b),
        "effective_only": paired(a, b, subset="eff"),
        "ge3_only": {k: v for k, v in paired(a, b).items() if k != "detail"},
    }
    p = res[key]["all_events"]
    print(f"\n{key}: 事件 {p['n_events']}，优 {p['win']} 劣 {p['loss']} 平 {p['tie']}，"
          f"有效对 {p['n_cmp']}，p={p['p']:.3g}")

# 合格事件（≥3 题通过）子集
ge3 = [t for t in tab if t["n_pass"] >= 3]
for a, b in [("c3mp", "c3m"), ("c3m", "c2")]:
    w = sum(1 for t in ge3 if t[f"ex_{a}"] > t[f"ex_{b}"])
    l = sum(1 for t in ge3 if t[f"ex_{a}"] < t[f"ex_{b}"])
    print(f"合格事件(≥3题, n={len(ge3)}) {a}-{b}: {w}:{l}:{len(ge3)-w-l}，"
          f"p={binom_two_sided(w, w + l):.3g}")

# ---------- 4b. 逐库汇总（供表 S-38 插值） ----------
per_db = collections.defaultdict(lambda: {"events": 0, "events_with_items": 0,
                                          "events_pass": 0, "items": 0, "items_pass": 0})
for t in tab:
    d0 = per_db[t["db"]]
    d0["events"] += 1
    d0["items"] += t["n"]
    d0["items_pass"] += t["n_pass"]
    if t["n"]:
        d0["events_with_items"] += 1
    if t["n_pass"]:
        d0["events_pass"] += 1
print("\n逐库：")
for k, v in sorted(per_db.items(), key=lambda x: -x[1]["events"]):
    print(f"  {k:24s} 事件{v['events']:2d} 有题{v['events_with_items']:2d} "
          f"有效{v['events_pass']:2d} 题{v['items']:3d} 通过{v['items_pass']:3d}")

# ---------- 5. 分母敏感性：453 通过子集上的四条件正确率 ----------
sub = [i for i in ids if man.get(i, {}).get("pass")]
print(f"\n通过四步验证的题 {len(sub)}/{len(ids)} = {len(sub)/len(ids)*100:.1f}%")
sens = {}
for c in CONDS:
    k = sum(1 for i in sub if by_cond[c][i] == "EX")
    sens[c] = {"ex": k, "n": len(sub), "pct": round(k / len(sub) * 100, 1),
               "full_ex": sum(1 for i in ids if by_cond[c][i] == "EX"),
               "full_pct": round(sum(1 for i in ids if by_cond[c][i] == "EX") / len(ids) * 100, 1)}
    print(f"  {c:5s} 全 479: {sens[c]['full_ex']:3d} ({sens[c]['full_pct']:4.1f}%)   "
          f"453 子集: {k:3d} ({sens[c]['pct']:4.1f}%)")

json.dump({"version": "audit-ext-paired-v1",
           "denominator": {"items": len(ids), "items_pass": len(sub),
                           "pass_pct": round(len(sub) / len(ids) * 100, 1)},
           "events": {"constructed": len(pe), "with_items": len(ev_items),
                      "effective": sum(1 for t in tab if t["n_pass"] > 0),
                      "ge3": len(ge3)},
           "per_event": tab, "paired": res,
           "sensitivity_453": sens, "per_db": dict(per_db)},
          open("results/audit_ext_paired.json", "w", encoding="utf-8"),
          ensure_ascii=False, indent=1)
print("\n写出 results/audit_ext_paired.json")
