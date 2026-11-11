# -*- coding: utf-8 -*-
"""规则基线 vs LLM 的逐题配对（回答第七轮评审第四节）。

三件事：
1. 在同一批条目上，把 c3r_* 与 c2 / c3m / c3mp 逐题配对，给精确 McNemar。
2. 分解「证据粒度」与「执行器」两个因素：
     证据效应   = c3r_name  -> c3r_infer  （同一确定性执行器，只换证据）
     执行器效应 = c3r_infer -> c3mp        （同一份段位谱系，只换取数方式）
3. 报告规则基线的失败类型分布；STAR_PROJ 是标识符替换的固有上界。

输入 results/rule_rewrite_items.jsonl、results/e1_*.jsonl、data/ids_ms.txt
输出 results/rule_vs_llm.json
"""
import collections
import glob
import json
import math
import os

BASE = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE)

VERSION = "rule-vs-llm-v2"

# v1（COUNT(*) 误判为投影星号）已记录的四档 EX 率，用于量化修正幅度。
# 数值取自 v1 运行结果 results/rule_rewrite.json（v1 版，已被 v2 覆盖）。
V1_RECORDED = {"merge": {"c3r_gt": 97.4, "c3r_infer": 97.4, "c3r_name": 0.9},
               "split": {"c3r_gt": 95.2, "c3r_infer": 95.2, "c3r_name": 95.2},
               "merge_holdout": {"c3r_gt": 95.0, "c3r_infer": 79.2, "c3r_name": 15.8},
               "split_holdout": {"c3r_gt": 92.2, "c3r_infer": 92.2, "c3r_name": 92.2}}

DEV_FILES = ["e1_dev_full.jsonl", "e1_diff_3lvl.jsonl", "e1_c3m_paired.jsonl",
             "e1_expand_med.jsonl", "e1_expand_hard.jsonl", "e1_ds_c3f.jsonl",
             "e1_ms_ds.jsonl", "e1_ms_ds_p.jsonl", "e1_ms_ds_p_split.jsonl",
             "e1_ds_one.jsonl", "e1_ds_one_b1.jsonl", "e1_ds_one_b2.jsonl",
             "e3_rw_main.jsonl"]

LLM = collections.defaultdict(dict)
for f in DEV_FILES:
    p = os.path.join("results", f)
    if not os.path.exists(p):
        continue
    for l in open(p, encoding="utf-8"):
        r = json.loads(l)
        LLM[r["id"]][r["cond"]] = r["outcome"]
for p in sorted(glob.glob("results/e1_holdout_*.jsonl")):
    for l in open(p, encoding="utf-8"):
        r = json.loads(l)
        LLM[r["id"]][r["cond"]] = r["outcome"]

RULE = collections.defaultdict(dict)
STATUS = collections.defaultdict(lambda: collections.defaultdict(collections.Counter))
for l in open("results/rule_rewrite_items.jsonl", encoding="utf-8"):
    r = json.loads(l)
    RULE[r["id"]][r["cond"]] = (r["status"] == "EX")
    STATUS[r["evo_type"]][r["cond"]][r["status"]] += 1

RULE_CONDS = ["c3r_gt", "c3r_infer", "c3r_name"]
LLM_CONDS = ["c2", "c3m", "c3mp"]

GROUPS = {"merge": ("merge_col", "data/ids_ms.txt"),
          "split": ("split_col", "data/ids_ms.txt"),
          "merge_holdout": ("merge_col_holdout", None),
          "split_holdout": ("split_col_holdout", None)}


def mcnemar(b, c):
    n = b + c
    if n == 0:
        return 1.0
    k = min(b, c)
    return min(1.0, 2.0 * sum(math.comb(n, j) for j in range(k + 1)) / (2 ** n))


def rule_val(i, c):
    v = RULE.get(i, {}).get(c)
    return None if v is None else bool(v)


def llm_val(i, c):
    v = LLM.get(i, {}).get(c)
    if v is None or v == "LLM_ERROR":
        return None
    return v == "EX"


def paired(ids, sa, a, sb, b):
    sel = [i for i in ids
           if sa(i, a) is not None and sb(i, b) is not None]
    if not sel:
        return None
    ao = sum(1 for i in sel if sa(i, a) and not sb(i, b))
    bo = sum(1 for i in sel if sb(i, b) and not sa(i, a))
    ka = sum(1 for i in sel if sa(i, a))
    kb = sum(1 for i in sel if sb(i, b))
    return {"n": len(sel), "a_ex": ka, "b_ex": kb,
            "a_rate": round(ka / len(sel) * 100, 1),
            "b_rate": round(kb / len(sel) * 100, 1),
            "a_only": ao, "b_only": bo,
            "delta_pt": round((ka - kb) / len(sel) * 100, 1),
            "p": round(mcnemar(ao, bo), 6)}


out = {"version": VERSION, "groups": {}, "failure_profile": {},
       "decomposition": {}, "star_audit": {}, "v1_vs_v2": {}}

# 星号审计（v2 新增：区分投影星号与 COUNT(*)）
if os.path.exists("results/rule_rewrite.json"):
    _rr = json.load(open("results/rule_rewrite.json", encoding="utf-8"))
    out["star_audit"] = _rr.get("audit") or {}

for gname, (evo, idf) in GROUPS.items():
    if idf and os.path.exists(idf):
        ids = [l.strip() for l in open(idf, encoding="utf-8") if l.strip()]
        ids = [i for i in ids if i in RULE and i.startswith(evo + ":")]
    else:
        ids = sorted(i for i in RULE if i.startswith(evo + ":"))
    g = {"n": len(ids), "rates": {}, "paired": {}}
    for c in RULE_CONDS:
        k = sum(1 for i in ids if rule_val(i, c))
        n = sum(1 for i in ids if rule_val(i, c) is not None)
        g["rates"][c] = {"n": n, "ex": k,
                         "rate": round(k / n * 100, 1) if n else None}
    for c in LLM_CONDS:
        k = sum(1 for i in ids if llm_val(i, c))
        n = sum(1 for i in ids if llm_val(i, c) is not None)
        g["rates"][c] = {"n": n, "ex": k,
                         "rate": round(k / n * 100, 1) if n else None}
    # (a, a_reader, b, b_reader)
    PAIRS = [("c3r_gt", rule_val, "c3mp", llm_val),
             ("c3r_infer", rule_val, "c3mp", llm_val),
             ("c3r_infer", rule_val, "c3m", llm_val),
             ("c3r_name", rule_val, "c3m", llm_val),
             ("c3r_infer", rule_val, "c2", llm_val),
             ("c3r_gt", rule_val, "c3r_infer", rule_val),
             ("c3r_name", rule_val, "c3r_infer", rule_val),
             ("c3mp", llm_val, "c3m", llm_val)]
    for a, ra, b, rb in PAIRS:
        p = paired(ids, ra, a, rb, b)
        if p:
            g["paired"]["%s-%s" % (a, b)] = p
    out["groups"][gname] = g

# v1 -> v2 修正幅度（评审第八轮第 1 条的量化答复）
for gname, g in out["groups"].items():
    v1 = V1_RECORDED.get(gname)
    if not v1:
        continue
    row = {}
    for c in RULE_CONDS:
        new = (g["rates"].get(c) or {}).get("rate")
        if new is None:
            continue
        row[c] = {"v1": v1[c], "v2": new, "gain_pt": round(new - v1[c], 1)}
    out["v1_vs_v2"][gname] = row

# 分解：两个因素各自的百分点贡献
for gname, g in out["groups"].items():
    r = g["rates"]
    get = lambda c: (r.get(c) or {}).get("rate")
    if None in (get("c3r_name"), get("c3r_infer"), get("c3mp")):
        continue
    out["decomposition"][gname] = {
        "evidence_effect_pt": round(get("c3r_infer") - get("c3r_name"), 1),
        "executor_effect_pt": round(get("c3r_infer") - get("c3mp"), 1),
        "total_pt": round(get("c3r_infer") - get("c3r_name"), 1),
        "note": "证据效应=同一确定性执行器下 列级->段位；"
                "执行器效应=同一段位谱系下 确定性替换->LLM 生成",
        "scope_note": "两值均为当前构造与当前实现下的配对差值，不是普遍因果效应："
                      "gold new_sql 由同一替换规则机械生成，规则执行器的上界被构造抬高；"
                      "LLM 侧只测了一个模型、一套提示模板。"}

# 失败画像
for evo, per in STATUS.items():
    out["failure_profile"][evo] = {c: dict(v) for c, v in per.items()}

json.dump(out, open("results/rule_vs_llm.json", "w", encoding="utf-8"),
          ensure_ascii=False, indent=1)

print("%-16s %5s %8s %8s %8s %8s %8s %8s" %
      ("档", "n", "c3r_gt", "c3r_inf", "c3r_nam", "c2", "c3m", "c3mp"))
for gname, g in out["groups"].items():
    r = g["rates"]
    f = lambda c: ("%5.1f" % r[c]["rate"]) if (r.get(c) or {}).get("rate") is not None else "    -"
    print("%-16s %5d %8s %8s %8s %8s %8s %8s" %
          (gname, g["n"], f("c3r_gt"), f("c3r_infer"), f("c3r_name"),
           f("c2"), f("c3m"), f("c3mp")))
print("\n配对（精确 McNemar）")
for gname, g in out["groups"].items():
    for k, v in g["paired"].items():
        print("  %-16s %-22s Δ%+6.1fpt  %d:%d  p=%s" %
              (gname, k, v["delta_pt"], v["a_only"], v["b_only"], v["p"]))
print("\n分解")
for k, v in out["decomposition"].items():
    print("  %-16s 证据 %+.1fpt  执行器 %+.1fpt" %
          (k, v["evidence_effect_pt"], v["executor_effect_pt"]))
print("\nv1 -> v2 修正（COUNT(*) 误判修复 + 星号展开能力）")
for gname, row in out["v1_vs_v2"].items():
    cells = "  ".join("%s %.1f->%.1f(%+.1f)" % (c, v["v1"], v["v2"], v["gain_pt"])
                      for c, v in row.items())
    print("  %-16s %s" % (gname, cells))
sa = out.get("star_audit") or {}
if sa:
    print("\n星号审计：投影星号 %s 条 / COUNT(*)%s 条 / 展开命中 %s 次（模式 %s）"
          % (sa.get("projection_star_items"), sa.get("count_star_items"),
             sa.get("star_expanded"), sa.get("star_mode")))
print("\n失败画像（v2）")
for evo, per in out["failure_profile"].items():
    for c, cnt in per.items():
        bad = {k: v for k, v in cnt.items() if k not in ("EX",)}
        if bad:
            print("  %-20s %-10s %s" % (evo, c, bad))
print("\n-> results/rule_vs_llm.json")
