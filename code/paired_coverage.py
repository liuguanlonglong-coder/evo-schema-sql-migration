# -*- coding: utf-8 -*-
"""配对区间覆盖率模拟（v5）

背景
----
v4 的覆盖率只报了五个情景，据此宣称方法在「整个范围」达到名义水平。四个情景
可以说明这些配置下的表现，不能证明整个参数空间；而且那五个情景的区间实现本身
是错的（对 b/n、c/n 各取 Wilson，见 paired_ci.py）。本文件做两件事：

1) 情景（scenario）：六个与实际档位一一对应的四格配置。真值直接取该档观测到的
   四格比例（p11, p10, p01, p00），再加一个 δ=0 的中等不一致率情形。
2) 网格（grid）：p10, p01 ∈ {0.005, 0.02, 0.05, 0.10, 0.20, 0.40}、
   p11 ∈ {0.10, 0.30, 0.50, 0.70}、n ∈ {96, 230}，凡 p11+p10+p01 ≤ 0.98 的组合
   全部纳入，报告最小值 / 均值 / 低于名义水平的比例。

结论口径：只对上述情景与网格覆盖到的 (p11, p10, p01, n) 成立，不外推。

方法
----
对每个真值参数，枚举全部四格表 (a,b,c,d)（多项式分布，概率低于 THRESH 的格子
截断并重新归一，记录被丢弃的质量），逐表计算四种区间，加权得到覆盖率：

    M10         Newcombe 方法 10（四格，本版主方法）
    v4_retired  v4 的错误实现（只用 b、c）
    conditional v3 的条件区间（只用 b、c，条件于 n_d）
    wald        McNemar 方差 Wald

覆盖率 = Σ w · 1{lo ≤ δ_true ≤ hi}，δ_true = p10 − p01。
名义水平 95%（alpha=0.05）。

用法
----
    python paired_coverage.py            # 情景 + 网格 → results/paired_ci_coverage.json
    python paired_coverage.py --quick    # 只跑情景与缩小网格（自检用）
"""

import itertools
import json
import math
import os
import sys
import time

import numpy as np

BASE = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE)

import importlib.util

_spec = importlib.util.spec_from_file_location(
    "paired_ci", os.path.join(BASE, "paired_ci.py"))
PCI = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(PCI)

THRESH = 1e-11
ALPHA = 0.05
OUT = "results/paired_ci_coverage.json"


def tables(n, p11, p10, p01, thresh=THRESH):
    """枚举四格表 (a,b,c,d) 及其概率。

    b ~ Bin(n, p10)；c | b ~ Bin(n-b, p01/(1-p10))；
    a | b,c ~ Bin(n-b-c, p11/(p11+p00))。
    返回 (list of (a,b,c,d,w), 被截断的质量)。
    """
    p00 = 1.0 - p11 - p10 - p01
    if p00 < -1e-12:
        return [], 1.0
    p00 = max(p00, 0.0)
    pc = p01 / (1.0 - p10) if (1.0 - p10) > 0 else 0.0
    denom = p11 + p00
    pa = p11 / denom if denom > 0 else 0.0
    keep = []
    mass = 0.0
    for b in range(n + 1):
        wb = PCI.binom_pmf(b, n, p10)
        if wb < thresh:
            continue
        cv = PCI.binom_pmf_vec(n - b, min(max(pc, 0.0), 1.0))
        for c in range(n - b + 1):
            wbc = wb * cv[c]
            if wbc < thresh:
                continue
            m = n - b - c
            av = PCI.binom_pmf_vec(m, min(max(pa, 0.0), 1.0))
            ws = wbc * av
            idx = np.nonzero(ws >= thresh)[0]
            for a in idx:
                keep.append((int(a), b, c, int(m - a), float(ws[a])))
                mass += float(ws[a])
    return keep, max(0.0, 1.0 - mass)


def coverage(n, p11, p10, p01, methods=("m10", "v4", "conditional", "wald")):
    """一个真值参数下的覆盖率与平均宽度。

    只有 M10 依赖全部四格；v4 / conditional / wald 只依赖 (b, c)，按 (b,c) 缓存。
    """
    T, lost = tables(n, p11, p10, p01)
    if not T:
        return None
    tot = sum(t[4] for t in T)
    delta = (p10 - p01) * 100.0
    res = {"n": n, "p11": p11, "p10": p10, "p01": p01,
           "delta_pt": round(delta, 2), "n_tables": len(T),
           "lost_mass": round(lost, 8)}
    cache = {}
    for m in methods:
        cover = 0.0
        width = 0.0
        for (a, b, c, d, w) in T:
            if m == "m10":
                lo, hi = PCI.ci_newcombe_pt(a, b, c, d, ALPHA)
            else:
                key = (m, b, c)
                got = cache.get(key)
                if got is None:
                    if m == "v4":
                        got = PCI.ci_discordant_v4(b, c, n, ALPHA)
                    elif m == "conditional":
                        got = PCI.ci_conditional(b, c, n, ALPHA)
                    else:
                        got = PCI.ci_wald(b, c, n, ALPHA)
                    cache[key] = got
                lo, hi = got
            if lo <= delta <= hi:
                cover += w
            width += w * (hi - lo)
        res[m] = {"coverage": round(cover / tot, 4),
                  "mean_width_pt": round(width / tot, 2)}
    return res


# 六个情景：真值取各档实际观测到的四格比例
SCEN = [
    ("改名 easy（c3m−c_rw）", 96, 87 / 96, 0 / 96, 4 / 96),
    ("改名 medium（c3m−c_rw）", 90, 82 / 90, 1 / 90, 2 / 90),
    ("改名 hard（c3m−c_rw）", 90, 68 / 90, 13 / 90, 3 / 90),
    ("合并（c3mp−c_rw）", 230, 25 / 230, 158 / 230, 3 / 230),
    ("拆分（c3mp−c_rw）", 104, 10 / 104, 71 / 104, 0 / 104),
    ("改名全档（c3m−c_rw）", 276, 237 / 276, 14 / 276, 9 / 276),
    ("δ=0、中等不一致率", 100, 0.80, 0.10, 0.10),
]

P10S = [0.005, 0.02, 0.05, 0.10, 0.20, 0.40]
P01S = [0.005, 0.02, 0.05, 0.10, 0.20, 0.40]
P11S = [0.10, 0.30, 0.50, 0.70]
NS = [96, 230]


def main(quick=False):
    t0 = time.time()
    out = {"version": "paired-coverage-v5",
           "alpha": ALPHA, "nominal": 1 - ALPHA,
           "prune_threshold": THRESH,
           "methods": {
               "m10": "Newcombe 方法 10（四格，本版主方法）",
               "v4": "v4 错误实现（只用 b、c）",
               "conditional": "v3 条件区间（条件于 n_d）",
               "wald": "McNemar 方差 Wald"},
           "delta_definition": "delta = p10 - p01（百分点）",
           "note": "覆盖率由四格多项式分布完全枚举加权得到；概率低于阈值的格子截断并"
                   "重新归一，lost_mass 记录丢弃质量。"}

    scen_rows = []
    print("=== 情景 ===")
    for label, n, p11, p10, p01 in SCEN:
        r = coverage(n, p11, p10, p01)
        r["label"] = label
        scen_rows.append(r)
        line = f"  {label:26s} n={n:3d} δ={r['delta_pt']:+6.1f} tab={r['n_tables']:5d} "
        for m in ("m10", "v4", "conditional", "wald"):
            line += f"{m}={r[m]['coverage']:.3f}/w{r[m]['mean_width_pt']:.1f}  "
        print(line, flush=True)
    out["scenarios"] = scen_rows

    p10s = [0.005, 0.02, 0.05, 0.10] if quick else P10S
    p01s = [0.005, 0.02, 0.05, 0.10] if quick else P01S
    p11s = [0.30, 0.70] if quick else P11S
    ns = [96] if quick else NS
    grid = []
    combos = [x for x in itertools.product(p11s, p10s, p01s, ns)
              if x[0] + x[1] + x[2] <= 0.98]
    print(f"\n=== 网格（{len(combos)} 点） ===")
    for i, (p11, p10, p01, n) in enumerate(combos):
        r = coverage(n, p11, p10, p01, methods=("m10", "v4", "wald"))
        if r is None:
            continue
        grid.append(r)
        if (i + 1) % 20 == 0:
            print(f"  {i+1}/{len(combos)}  {time.time()-t0:.0f}s", flush=True)
    out["grid"] = grid
    out["grid_spec"] = {"p11": p11s, "p10": p10s, "p01": p01s, "n": ns,
                        "n_points": len(grid)}

    summ = {}
    for m in ("m10", "v4", "wald"):
        cov = [g[m]["coverage"] for g in grid]
        wid = [g[m]["mean_width_pt"] for g in grid]
        summ[m] = {
            "min": round(min(cov), 4),
            "mean": round(sum(cov) / len(cov), 4),
            "max": round(max(cov), 4),
            "frac_below_nominal": round(sum(1 for x in cov if x < 1 - ALPHA) / len(cov), 4),
            "frac_below_0.93": round(sum(1 for x in cov if x < 0.93) / len(cov), 4),
            "frac_below_0.90": round(sum(1 for x in cov if x < 0.90) / len(cov), 4),
            "mean_width_pt": round(sum(wid) / len(wid), 2),
        }
    out["grid_summary"] = summ
    print("\n=== 网格汇总 ===")
    for m, s in summ.items():
        print(f"  {m:12s} min={s['min']:.3f} mean={s['mean']:.3f} "
              f"<95%:{s['frac_below_nominal']:.2f} <93%:{s['frac_below_0.93']:.2f} "
              f"<90%:{s['frac_below_0.90']:.2f} w={s['mean_width_pt']}")

    # 情景层面的头条数字，供正文直接引用
    sm = {r["label"]: r for r in scen_rows}
    out["headline"] = {
        "scenario_m10_min": round(min(r["m10"]["coverage"] for r in scen_rows), 4),
        "scenario_m10_max": round(max(r["m10"]["coverage"] for r in scen_rows), 4),
        "scenario_v4_min": round(min(r["v4"]["coverage"] for r in scen_rows), 4),
        "scenario_conditional_min": round(
            min(r["conditional"]["coverage"] for r in scen_rows), 4),
        "scenario_conditional_max": round(
            max(r["conditional"]["coverage"] for r in scen_rows), 4),
        "grid_m10_min": summ["m10"]["min"],
        "grid_v4_min": summ["v4"]["min"],
        "grid_m10_frac_below_nominal": summ["m10"]["frac_below_nominal"],
        "grid_v4_frac_below_nominal": summ["v4"]["frac_below_nominal"],
    }
    json.dump(out, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print(f"\nsaved {OUT}  ({time.time()-t0:.0f}s)")


if __name__ == "__main__":
    main(quick="--quick" in sys.argv)
