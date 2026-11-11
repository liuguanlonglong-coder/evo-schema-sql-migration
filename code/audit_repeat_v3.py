# -*- coding: utf-8 -*-
"""audit_repeat_v3.py - 重复判定的一致性（回应 MO-11）。

审稿意见：材料报告五轮重复实验最大偏差约 0.1 个百分点，但缺少逐轮逐题输出，
无法判断误差是「边缘率稳定、逐题翻转」还是「整体漂移」。

五轮的逐题输出确实没有归档，这一点不能编造。但仓库里存在另一类可用的重复判定：
同一模型、同一条件、同一道题在**不同批次文件**里被各评过一次（例如消融实验的
c3mp 对照组与主实验的 c3mp、扩样分片与合并后的全量文件）。这些重复判定可以用来
回答审稿人真正关心的问题——误差是集中在个别题上，还是整体漂移。

输出 results/repeat_consistency_v3.json
"""
import collections
import glob
import itertools
import json
import os

BASE = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE)

OUT = "results/repeat_consistency_v3.json"
VERSION = "repeat-consistency-v3"


def load():
    recs = []
    for f in sorted(glob.glob("results/*.jsonl")):
        name = os.path.basename(f)
        for line in open(f, encoding="utf-8"):
            line = line.strip()
            if not line:
                continue
            try:
                r = json.loads(line)
            except Exception:
                continue
            if not r.get("model") or not r.get("id") or not r.get("cond"):
                continue
            recs.append({"file": name, "model": r["model"], "cond": r["cond"],
                         "id": r["id"], "outcome": r.get("outcome")})
    return recs


def main():
    recs = load()
    # (model, cond, id) -> [(file, outcome)]
    by = collections.defaultdict(list)
    for r in recs:
        by[(r["model"], r["cond"], r["id"])].append((r["file"], r["outcome"]))
    dup = {k: v for k, v in by.items() if len({f for f, _ in v}) > 1}
    agree = sum(1 for v in dup.values() if len({o for _, o in v}) == 1)
    out = {"version": VERSION,
           "scope": {"rows_with_model_cond_id": len(recs),
                     "keys": len(by), "keys_evaluated_more_than_once": len(dup),
                     "keys_identical_outcome": agree,
                     "keys_flipped": len(dup) - agree,
                     "item_agreement_pct": round(100 * agree / max(1, len(dup)), 1)},
           "note": "同一 (模型, 条件, 题号) 在不同批次文件里各判一次。"
                   "这一致率度量的是判定稳定性，不是五轮重复实验的逐题输出。"}

    # 文件对级：哪些对之间重合最多、边缘率差多少
    pairs = collections.Counter()
    per = collections.defaultdict(lambda: {"n": 0, "same": 0, "a_ex": 0, "b_ex": 0})
    for k, v in dup.items():
        files = sorted({f for f, _ in v})
        outmap = {}
        for f, o in v:
            outmap.setdefault(f, o)
        for fa, fb in itertools.combinations(files, 2):
            key = (k[1], fa, fb)
            per[key]["n"] += 1
            oa, ob = outmap[fa], outmap[fb]
            a_ex, b_ex = oa == "EX", ob == "EX"
            per[key]["same"] += oa == ob
            per[key]["a_ex"] += a_ex
            per[key]["b_ex"] += b_ex
            pairs[(fa, fb)] += 1
    rows = []
    for (cond, fa, fb), d in per.items():
        if d["n"] < 20:
            continue
        rows.append({"cond": cond, "file_a": fa, "file_b": fb, "n": d["n"],
                     "agree_pct": round(100 * d["same"] / d["n"], 1),
                     "rate_a": round(100 * d["a_ex"] / d["n"], 1),
                     "rate_b": round(100 * d["b_ex"] / d["n"], 1),
                     "delta_pt": round(100 * (d["a_ex"] - d["b_ex"]) / d["n"], 1)})
    rows.sort(key=lambda x: (-x["n"]))
    out["file_pairs"] = rows

    # 只保留「同条件、同提示族」的重复判定做头条读数。
    # _oldsep 是另一种分隔符构造（不是重复），必须剔除，否则会出现 53.8pt 的假差值。
    strict = [r for r in rows
              if "oldsep" not in r["file_a"] and "oldsep" not in r["file_b"]]
    n_s = sum(r["n"] for r in strict)
    same_s = sum(round(r["agree_pct"] * r["n"] / 100) for r in strict)
    out["strict"] = {
        "n_pairs": len(strict), "n_judgements": n_s,
        "item_agreement_pct": round(100 * same_s / max(1, n_s), 1),
        "max_abs_delta_pt": max((abs(r["delta_pt"]) for r in strict), default=None),
        "median_abs_delta_pt": sorted(abs(r["delta_pt"]) for r in strict)[
            len(strict) // 2] if strict else None,
        "note": "剔除 oldsep 构造差异后，同条件跨批次重复判定的汇总。"
                "这是「五轮逐题输出未归档」下能给出的最强替代证据。"}

    # 消融对照臂 vs 主实验 c3mp：提示完全一致的 230 题，按折汇总
    foc = [r for r in rows if "ab_nat" in r["file_a"] and "ms_ds_p" in r["file_b"]]
    if foc:
        n_f = sum(r["n"] for r in foc)
        same_f = sum(round(r["agree_pct"] * r["n"] / 100) for r in foc)
        out["focus_ablation_vs_main"] = {
            "n": n_f, "item_agreement_pct": round(100 * same_f / n_f, 1),
            "per_fold_delta_pt": [r["delta_pt"] for r in foc],
            "max_abs_delta_pt": max(abs(r["delta_pt"]) for r in foc),
            "note": "消融实验的 c3mp 对照臂与主实验 c3mp 是同一提示、同一模型、"
                    "同一批题的两个批次，差异只能来自服务端不确定性。"}

    # 汇总：所有重复键上，两次判定的边缘率差绝对值的分布
    deltas = sorted(abs(r["delta_pt"]) for r in rows)
    out["pair_delta_pt"] = {
        "n_pairs": len(deltas),
        "median": deltas[len(deltas) // 2] if deltas else None,
        "max": max(deltas) if deltas else None,
        "pairs_within_0_2pt": sum(1 for d in deltas if d <= 0.2),
        "pairs_over_5pt": sum(1 for d in deltas if d > 5),
    }
    json.dump(out, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)

    s = out["scope"]
    print("重复键", s["keys_evaluated_more_than_once"], "逐题一致",
          s["item_agreement_pct"], "%（翻转", s["keys_flipped"], "）")
    print("文件对（n>=20）：", len(rows))
    for r in rows[:8]:
        print(f"  {r['cond']:5s} {r['file_a'][:24]:24s} vs {r['file_b'][:24]:24s} "
              f"n={r['n']:3d} 一致{r['agree_pct']:5.1f}% "
              f"{r['rate_a']:5.1f}->{r['rate_b']:5.1f} Δ{r['delta_pt']:5.1f}")
    print("边缘率差分布:", out["pair_delta_pt"])
    print("严格重复:", out["strict"])
    print("消融臂 vs 主实验:", out.get("focus_ablation_vs_main"))
    print("写入", OUT)


if __name__ == "__main__":
    main()
