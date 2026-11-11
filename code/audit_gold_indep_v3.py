# -*- coding: utf-8 -*-
"""audit_gold_indep_v3.py - Gold SQL 的同源度量与信息补贴量化（回应 MO-2）。

审稿意见：若 Gold SQL 由预设映射与确定性替换规则构造、待评估方法又用到同一映射，
则高成功率可能部分反映「答案与规则同源」；C3mp 79.6% 与固定执行器 100% 都不能
直接解释为真实迁移准确率。要求提供独立人工核验子集或明确承认局限。

人工核验本轮无法伪造，故本脚本只做能被复核的三件事：

  A 同源核对：把构造期的 rewrite(old_sql, mapping) 重新执行一遍，与 manifest 里
    的 new_sql 逐字比较。若 100% 相同，说明 gold 完全由替换规则生成，
    「构造内上界」这个说法就有了可核对的依据，而不是自我评价。
  B 信息补贴量化：gold 中由替换插入的字符占比（中位数），即模型被「送」了多少。
  C 表达形式消融：c3mp（给可复制表达式）与 c3mp_nat（给等价的自然语言段位描述）
    的逐题配对，把「理解段位」与「照抄表达式」分开。这是本文唯一能把两者
    剥离的现成证据，n=230。

输出 results/gold_independence_v3.json
"""
import collections
import json
import math
import os
import re
import sys

BASE = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE)
sys.path.insert(0, BASE)

import build_evolution_v2 as V2  # noqa: E402

OUT = "results/gold_independence_v3.json"
VERSION = "gold-independence-v3"


def comb(n, k):
    return math.comb(n, k)


def mcnemar(b, c):
    n = b + c
    if n == 0:
        return 1.0
    k = min(b, c)
    return min(1.0, 2 * sum(comb(n, i) for i in range(k + 1)) / 2 ** n)


def load_manifest(evo, path="data/manifest_v3.jsonl"):
    out = []
    for line in open(path, encoding="utf-8"):
        r = json.loads(line)
        if r["evo_type"] == evo:
            out.append(r)
    return out


def rewrite_check(rows, evo):
    """逐条重放构造期的替换，与 manifest 的 new_sql 逐字比较。"""
    same = diff = err = 0
    shares = []
    examples = []
    for r in rows:
        ops = r["ops"]
        if isinstance(ops, str):
            ops = json.loads(ops)
        mapping = {}
        lit_map = {}
        for o in ops:
            if len(o) < 4:
                continue
            # o = [kind, table, old_col, new_expr]
            mapping[o[2]] = o[3]
        try:
            got = V2.rewrite(r["old_sql"], mapping, lit_map)
        except Exception:
            err += 1
            continue
        if got is None:
            err += 1
            continue
        if _norm(got) == _norm(r["new_sql"]):
            same += 1
        else:
            diff += 1
            if len(examples) < 3:
                examples.append({"id": r["id"], "rewrite": got[:160],
                                 "gold": r["new_sql"][:160]})
        # 信息补贴：gold 中替换插入的字符占比
        ins = sum(len(v) for v in mapping.values())
        if r["new_sql"]:
            shares.append(min(1.0, ins / max(1, len(r["new_sql"]))))
    shares.sort()
    med = shares[len(shares) // 2] if shares else None
    return {"n": len(rows), "identical": same, "different": diff,
            "rewrite_failed": err,
            "identical_pct": round(100 * same / max(1, len(rows)), 1),
            "inserted_char_share_median": round(100 * med, 1) if med else None,
            "examples": examples}


def _norm(s):
    return re.sub(r"\s+", " ", (s or "").strip()).lower()


def ablation():
    rows = []
    for i in range(6):
        f = f"results/e1_ab_nat_{i}.jsonl"
        if os.path.exists(f):
            rows += [json.loads(l) for l in open(f, encoding="utf-8") if l.strip()]
    by = collections.defaultdict(dict)
    for r in rows:
        by[r["id"]][r["cond"]] = r["outcome"]
    pair = [(v["c3mp"], v["c3mp_nat"]) for v in by.values()
            if "c3mp" in v and "c3mp_nat" in v]
    n = len(pair)
    k_lit = sum(1 for a, _b in pair if a == "EX")
    k_nat = sum(1 for a, b in pair if b == "EX")
    b = sum(1 for a, x in pair if a == "EX" and x != "EX")
    c = sum(1 for a, x in pair if a != "EX" and x == "EX")
    return {"n_pairs": n,
            "c3mp_literal": {"k": k_lit, "rate": round(100 * k_lit / max(1, n), 1)},
            "c3mp_nat": {"k": k_nat, "rate": round(100 * k_nat / max(1, n), 1)},
            "delta_pt": round(100 * (k_lit - k_nat) / max(1, n), 1),
            "only_literal": b, "only_natural": c,
            "mcnemar_p": round(mcnemar(b, c), 8),
            "note": "同一批题、同一模型、同一提示骨架，只把「可复制的段位表达式」换成"
                    "等价自然语言描述；因此差值度量的是表达形式，而非信息量。"}


def main():
    merge = load_manifest("merge_col")
    split = load_manifest("split_col")
    out = {"version": VERSION,
           "merge": rewrite_check(merge, "merge_col"),
           "split": rewrite_check(split, "split_col"),
           "ablation": ablation()}
    json.dump(out, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    for k in ("merge", "split"):
        v = out[k]
        print(f"{k}: {v['identical']}/{v['n']} 逐字相同（{v['identical_pct']}%），"
              f"不同 {v['different']}，替换失败 {v['rewrite_failed']}，"
              f"插入字符占比中位数 {v['inserted_char_share_median']}%")
    a = out["ablation"]
    print("消融:", json.dumps(a, ensure_ascii=False))
    print("写入", OUT)


if __name__ == "__main__":
    main()
