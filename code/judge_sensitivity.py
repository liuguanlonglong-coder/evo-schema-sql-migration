# -*- coding: utf-8 -*-
"""判定规则敏感性分析（第五轮评审第 4 项：量化判定器宽严的影响）。

现行判定器（run_c1c2c3.judge）的规则是：
    把结果行逐格转成字符串、NULL 当空串、**排序后**比较简单元组多重集。
三条规则都"宽严并存"：
  · 排序            → ORDER BY 写错的 SQL 可能被判对（偏宽）
  · 全转字符串      → 1 与 1.0、1.0 与 1.000001 被判错（偏严）
  · NULL ≡ 空串     → 本应不同的结果被判同（偏宽）

本脚本不新增模型调用：它从冻结的逐题输出里取每条预测 SQL，与 manifest 的
gold SQL 一起在演化库上重执行，然后按五种规则重判，量化每条规则对
EX 率的实际影响。基线规则必须复现出与冻结 outcome 相同的 EX 计数，
否则说明重执行环境与采集时不一致，脚本会显式报告偏差。

输出 results/judge_sensitivity.json
"""
import collections
import hashlib
import json
import math
import os
import sqlite3
import sys
import time

BASE = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE)

VERSION = "judge-sensitivity-v2"
QUERY_TIMEOUT = 4.0


def mcnemar_exact(b, c):
    """精确双侧 McNemar（条件二项，p = 0.5）。"""
    n = b + c
    if n == 0:
        return 1.0
    k = min(b, c)
    return min(1.0, 2.0 * sum(math.comb(n, j) for j in range(k + 1)) / (2 ** n))

EVOLVED_DIR = {
    "rename": "data/evolved",
    "rename_medium": "data/evolved_medium",
    "rename_hard": "data/evolved_hard",
    "merge_col": "data/evolved_merge",
    "split_col": "data/evolved_split",
}

FILES = ["e1_dev_full.jsonl", "e1_diff_3lvl.jsonl", "e1_c3m_paired.jsonl",
         "e1_expand_med.jsonl", "e1_expand_hard.jsonl", "e1_ds_c3f.jsonl",
         "e1_ms_ds.jsonl", "e1_ms_ds_p.jsonl", "e1_ms_ds_p_split.jsonl",
         "e1_ds_one.jsonl", "e1_ds_one_b1.jsonl", "e1_ds_one_b2.jsonl",
         "e3_rw_main.jsonl"]

MAN = {}
for _l in open("data/manifest_v3.jsonl", encoding="utf-8"):
    _r = json.loads(_l)
    MAN[_r["id"]] = _r

IDS_XM = [l.strip() for l in open("data/ids_xm_full.txt", encoding="utf-8") if l.strip()]
IDS_MS = [l.strip() for l in open("data/ids_ms.txt", encoding="utf-8") if l.strip()]

GROUPS = {
    "easy": ("rename", IDS_XM),
    "medium": ("rename_medium", IDS_XM),
    "hard": ("rename_hard", IDS_XM),
    "merge": ("merge_col", IDS_MS),
    "split": ("split_col", IDS_MS),
}
CONDS = ["c2", "c3m", "c3mp", "c_rw"]


def ids_of(name):
    evo, pool = GROUPS[name]
    return [i for i in pool if MAN[i]["evo_type"] == evo]


# ---------------------------------------------------------------- 执行与规范化
def db_path(evo_type, db_id):
    return os.path.join(BASE, EVOLVED_DIR[evo_type], f"{db_id}.sqlite")


_EXEC = {}


def run_sql(path, sql):
    """带超时的只读执行；同一 (path, sql) 只执行一次。"""
    key = (path, sql)
    if key in _EXEC:
        return _EXEC[key]
    if not sql or not sql.strip():
        _EXEC[key] = (None, "empty")
        return _EXEC[key]
    try:
        con = sqlite3.connect(f"file:{path}?mode=ro", uri=True)
    except Exception as e:                                   # pragma: no cover
        _EXEC[key] = (None, str(e)[:200])
        return _EXEC[key]
    t0 = time.time()
    con.set_progress_handler(lambda: 1 if time.time() - t0 > QUERY_TIMEOUT else 0, 2000)
    try:
        _EXEC[key] = (con.execute(sql).fetchall(), None)
    except Exception as e:
        _EXEC[key] = (None, str(e)[:200])
    finally:
        con.close()
    return _EXEC[key]


NUM_DP = 6


def cell(v, null_distinct=False, numeric=False):
    """单格的规范化。三种规则各开关一处。"""
    if v is None:
        return ("__NULL__",) if null_distinct else ("S", "")
    if numeric:
        try:
            return ("N", round(float(v), NUM_DP))
        except (TypeError, ValueError):
            pass
    return ("S", str(v))


def key_rows(rows, order_sensitive=False, numeric=False, null_distinct=False):
    ks = [tuple(cell(v, null_distinct, numeric) for v in r) for r in rows]
    return tuple(ks if order_sensitive else sorted(ks))


def same(gold, pred, order_sensitive=False, numeric=False, null_distinct=False):
    if gold is None or pred is None:
        return False
    return (key_rows(gold, order_sensitive, numeric, null_distinct)
            == key_rows(pred, order_sensitive, numeric, null_distinct))


RULES = {
    "r0_current": "排序多重集 + 全转字符串 + NULL≡空串（现行判定器）",
    "r1_order": "行序敏感（其余同 r0）",
    "r2_numeric": "数值容差 1e-6（其余同 r0）",
    "r3_null": "NULL 与空串区分（其余同 r0）",
    "r_all_strict": "行序敏感 + 数值容差 + NULL 区分（三者同时）",
}


def judge_under(gold, pred, rule):
    return same(gold,
                pred,
                order_sensitive=rule in ("r1_order", "r_all_strict"),
                numeric=rule in ("r2_numeric", "r_all_strict"),
                null_distinct=rule in ("r3_null", "r_all_strict"))


# ---------------------------------------------------------------- 装配冻结输出
SQL = collections.defaultdict(dict)
OUTCOME = collections.defaultdict(dict)
USED = []
for f in FILES:
    p = os.path.join("results", f)
    if not os.path.exists(p):
        continue
    USED.append(f)
    for l in open(p, encoding="utf-8"):
        r = json.loads(l)
        c = r.get("cond")
        if c is None:
            continue
        SQL[r["id"]][c] = r.get("sql") or ""
        OUTCOME[r["id"]][c] = r.get("outcome")

report = {"version": VERSION,
          "inputs": USED,
          "rules": RULES,
          "note": "不新增模型调用：从冻结输出取预测 SQL，重执行后按各规则重判；"
                  "基线规则 r0 应与冻结 outcome 的 EX 计数一致，偏差见 reproduce_check。",
          "groups": [],
          "gold_sql_shape": {},
          "examples": []}

print("=" * 78)
print("判定规则敏感性：同一批预测 SQL，五种规则下的 EX 计数")
print("=" * 78)

gold_cache = {}
order_by_share = {}
perm_items = []

for gname in ("easy", "medium", "hard", "merge", "split"):
    evo = GROUPS[gname][0]
    ids = ids_of(gname)
    # 先算 gold（并缓存），顺带统计 gold SQL 是否含 ORDER BY
    n_ob = 0
    for i in ids:
        gsql = MAN[i].get("new_sql") or ""
        if "order by" in gsql.lower():
            n_ob += 1
        path = db_path(evo, MAN[i]["db_id"])
        if i not in gold_cache:
            gold_cache[i] = run_sql(path, gsql)[0]
    order_by_share[gname] = {"n": len(ids), "n_gold_order_by": n_ob,
                             "share_pct": round(n_ob / len(ids) * 100, 1)}

    rec = {"group": gname, "n": len(ids), "conds": {}}
    for c in CONDS:
        have = [i for i in ids if c in SQL[i]]
        if not have:
            continue
        cnt = collections.Counter()
        detail = []
        n_err = 0
        for i in have:
            pred, err = run_sql(db_path(evo, MAN[i]["db_id"]), SQL[i][c])
            if pred is None:
                n_err += 1
                continue
            row = {r: judge_under(gold_cache[i], pred, r) for r in RULES}
            for r, ok in row.items():
                if ok:
                    cnt[r] += 1
            detail.append((i, row))
        n = len(have)
        entry = {"n": n, "n_exec_error": n_err}
        for r in RULES:
            entry[r] = cnt[r]
            entry[r + "_pct"] = round(cnt[r] / n * 100, 1)
        # 翻转计数
        entry["flip_order_lost"] = sum(1 for _, d in detail if d["r0_current"] and not d["r1_order"])
        entry["flip_order_gain"] = sum(1 for _, d in detail if not d["r0_current"] and d["r1_order"])
        entry["mcnemar_order_p"] = round(
            mcnemar_exact(entry["flip_order_lost"], entry["flip_order_gain"]), 6)
        entry["flip_numeric_gain"] = sum(1 for _, d in detail if not d["r0_current"] and d["r2_numeric"])
        entry["flip_null_lost"] = sum(1 for _, d in detail if d["r0_current"] and not d["r3_null"])
        # 含 ORDER BY 子集单独报告（评审第七轮第五节）：宽判只可能发生在这批条目上，
        # 因此需要给出该子集内的保序准确率，而不只是全体的差值。
        sub = [(i, d) for i, d in detail
               if "order by" in (MAN[i].get("new_sql") or "").lower()]
        entry["order_by_subset"] = {
            "n": len(sub),
            "ex_current": sum(1 for _, d in sub if d["r0_current"]),
            "ex_strict": sum(1 for _, d in sub if d["r1_order"]),
            "ex_current_pct": round(sum(1 for _, d in sub if d["r0_current"]) / len(sub) * 100, 1)
            if sub else None,
            "ex_strict_pct": round(sum(1 for _, d in sub if d["r1_order"]) / len(sub) * 100, 1)
            if sub else None,
            "lost": sum(1 for _, d in sub if d["r0_current"] and not d["r1_order"]),
        }
        # 冻结 outcome 复现校验
        frozen = sum(1 for i in have if OUTCOME[i].get(c) == "EX")
        entry["frozen_EX"] = frozen
        entry["reproduce_delta"] = cnt["r0_current"] - frozen
        rec["conds"][c] = entry
        for i, d in detail:
            if d["r0_current"] and not d["r1_order"] and len(perm_items) < 40:
                perm_items.append({"id": i, "group": gname, "cond": c,
                                   "gold_order_by": "order by" in (MAN[i].get("new_sql") or "").lower()})
    report["groups"].append(rec)

report["gold_sql_shape"] = order_by_share
report["examples"] = perm_items

# ---------------------------------------------------------------- 主结论的稳健性
# 关心的问题不是「某个格子的 EX 变成几点」，而是「换规则后主结论的差值变不变」。
CONTRASTS = [("merge", "c3mp", "c3m"), ("merge", "c3mp", "c_rw"),
             ("merge", "c3m", "c2"), ("split", "c3mp", "c_rw"),
             ("split", "c3m", "c2"), ("easy", "c3m", "c_rw"),
             ("medium", "c3m", "c_rw"), ("hard", "c3m", "c_rw")]
GD = {rec["group"]: rec for rec in report["groups"]}
rows = []
for g, a, b in CONTRASTS:
    if g not in GD or a not in GD[g]["conds"] or b not in GD[g]["conds"]:
        continue
    ea, eb = GD[g]["conds"][a], GD[g]["conds"][b]
    row = {"group": g, "a": a, "b": b, "n": ea["n"]}
    for r in RULES:
        row[r] = round((ea[r] - eb[r]) / ea["n"] * 100, 1)
    row["max_abs_change"] = round(max(abs(row[r] - row["r0_current"]) for r in RULES), 1)
    row["sign_stable"] = all((row[r] > 0) == (row["r0_current"] > 0) for r in RULES)
    rows.append(row)
report["contrasts"] = rows

# ---------------------------------------------------------------- 严格 EX 作为正式指标
# 现行判定排序后比较，ORDER BY 写错仍可能判对。上一版只把这一点写成敏感性脚注，
# 本版把「保序 EX」（行序敏感）作为正式指标并列报告：逐档×条件给出两套计数、
# 绝对差、受影响条数与同条件配对 McNemar。
strict = {"note": "保序 EX = 行序敏感比较下的执行一致率；与现行 EX 的差即行序不敏感"
                  "带来的宽判幅度。同条件配对 McNemar 的 b 为「现行对、保序错」，"
                  "c 为「现行错、保序对」，后者恒为 0（收严只会减少判对）。",
          "by_group": [], "total": {}}
T = {"n": 0, "r0": 0, "r1": 0, "lost": 0, "gain": 0}
for rec in report["groups"]:
    g = {"group": rec["group"], "n": rec["n"], "conds": []}
    G = {"n": 0, "r0": 0, "r1": 0, "lost": 0, "gain": 0}
    for c, e in rec["conds"].items():
        g["conds"].append({
            "cond": c, "n": e["n"],
            "ex_current": e["r0_current"], "ex_current_pct": e["r0_current_pct"],
            "ex_strict": e["r1_order"], "ex_strict_pct": e["r1_order_pct"],
            "diff_pt": round(e["r0_current_pct"] - e["r1_order_pct"], 1),
            "n_order_sensitive": e["flip_order_lost"],
            "mcnemar_order_p": e["mcnemar_order_p"],
            "order_by_subset": e["order_by_subset"]})
        for k, kk in (("n", "n"), ("r0", "r0_current"), ("r1", "r1_order"),
                      ("lost", "flip_order_lost"), ("gain", "flip_order_gain")):
            G[k] += e[kk]
            T[k] += e[kk]
    g["pooled"] = {"n": G["n"], "ex_current": G["r0"], "ex_strict": G["r1"],
                   "ex_current_pct": round(G["r0"] / G["n"] * 100, 1) if G["n"] else None,
                   "ex_strict_pct": round(G["r1"] / G["n"] * 100, 1) if G["n"] else None,
                   "n_order_sensitive": G["lost"],
                   "mcnemar_order_p": round(mcnemar_exact(G["lost"], G["gain"]), 6)}
    g["order_by_subset_pooled"] = {
        "n": sum(e["order_by_subset"]["n"] for e in rec["conds"].values()),
        "ex_current": sum(e["order_by_subset"]["ex_current"] for e in rec["conds"].values()),
        "ex_strict": sum(e["order_by_subset"]["ex_strict"] for e in rec["conds"].values()),
        "lost": sum(e["order_by_subset"]["lost"] for e in rec["conds"].values())}
    O = g["order_by_subset_pooled"]
    O["ex_current_pct"] = round(O["ex_current"] / O["n"] * 100, 1) if O["n"] else None
    O["ex_strict_pct"] = round(O["ex_strict"] / O["n"] * 100, 1) if O["n"] else None
    strict["by_group"].append(g)
strict["total"] = {"n": T["n"], "ex_current": T["r0"], "ex_strict": T["r1"],
                   "ex_current_pct": round(T["r0"] / T["n"] * 100, 1) if T["n"] else None,
                   "ex_strict_pct": round(T["r1"] / T["n"] * 100, 1) if T["n"] else None,
                   "n_order_sensitive": T["lost"],
                   "share_order_sensitive_pct": round(T["lost"] / T["n"] * 100, 2) if T["n"] else None,
                   "mcnemar_order_p": round(mcnemar_exact(T["lost"], T["gain"]), 6)}
report["strict_ex"] = strict

print("\n主结论在各判定规则下的差值（百分点）：")
print(f"{'对照':24s} {'r0现行':>7s} {'r1行序':>7s} {'r2数值':>7s} {'全严':>7s} {'最大变动':>8s} 符号不变")
for r in rows:
    print(f"{r['group'] + ':' + r['a'] + '−' + r['b']:24s} "
          f"{r['r0_current']:+7.1f} {r['r1_order']:+7.1f} {r['r2_numeric']:+7.1f} "
          f"{r['r_all_strict']:+7.1f} {r['max_abs_change']:8.1f}  {'是' if r['sign_stable'] else '否'}")

# ---------------------------------------------------------------- 打印
hdr = f"{'档位':7s} {'条件':6s} {'n':>4s} {'冻结EX':>6s} {'r0':>4s} {'r1行序':>6s} {'r2数值':>6s} {'r3NULL':>6s} {'全严':>5s}  翻转(序/NULL)"
print(hdr)
for rec in report["groups"]:
    for c, e in rec["conds"].items():
        print(f"{rec['group']:7s} {c:6s} {e['n']:4d} {e['frozen_EX']:6d} "
              f"{e['r0_current']:4d} {e['r1_order']:6d} {e['r2_numeric']:6d} "
              f"{e['r3_null']:6d} {e['r_all_strict']:5d}   "
              f"{e['flip_order_lost']}/{e['flip_null_lost']}"
              + (f"  Δ复现{e['reproduce_delta']:+d}" if e["reproduce_delta"] else ""))

print("\n各档 gold SQL 含 ORDER BY 的比例：")
for g, v in order_by_share.items():
    print(f"  {g:7s} {v['n_gold_order_by']:3d}/{v['n']:3d} = {v['share_pct']:.1f}%")

print("\n仅差行序却被判对的条目（前 12 条）：")
for x in perm_items[:12]:
    print("   ", x)

json.dump(report, open("results/judge_sensitivity.json", "w", encoding="utf-8"),
          ensure_ascii=False, indent=1)
print("\n-> results/judge_sensitivity.json")
