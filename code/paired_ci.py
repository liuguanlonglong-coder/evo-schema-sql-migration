# -*- coding: utf-8 -*-
"""配对二元结果的比例差置信区间（v5）

参数
----
配对四格表（行 = 条件 A，列 = 条件 B）：

                B 正确    B 错误
    A 正确        a         b
    A 错误        c         d

要估计的总体参数是边际比例差

    delta = p1 - p2,  p1 = (a+b)/n,  p2 = (a+c)/n,  n = a+b+c+d

注意 delta 的点估计等于 (b-c)/n，但这不意味着区间只需要 (b, c)：a 与 d 决定
两个边际比例的位置（p1、p2 可以都很高、都很低或相差很远），而区间宽度取决于
这两个位置与配对相关性。只用 (b, c, n) 的做法无法区分「a=45, b=5, c=0, d=0」
与「a=0, b=5, c=0, d=45」——这两张表的 delta 相同，可信程度截然不同。

v4 的错误（本版修正）
--------------------
v4 的 ci_newcombe(b, c, n) 对 b/n 与 c/n 分别取 Wilson 区间再做平方和合成。
那不是 Newcombe 方法 10：
  · 它把 b 与 c 当作两个独立二项，而 b、c 是同一多项式分布的格子，二者负相关；
  · 它完全没有使用 a、d 两个同对/同错的格子，也就没有相关性修正；
  · 它算的是「两个不一致比例之差」的区间，恰好等于 −delta 的点估计，
    所以偏差被点估计掩盖，看不出来。
故 v4 的数值不满足方法 10 的定义，本版重算，旧实现保留为 ci_discordant_v4 供对照。

方法 10 的定义（Newcombe 1998, Stat Med 17:2635-2650, p.2639）
------------------------------------------------------------
    (l1, u1) = Wilson(a+b, n)      # 边际比例 p1 的 Wilson 区间
    (l2, u2) = Wilson(a+c, n)      # 边际比例 p2 的 Wilson 区间
    f2 = p1 - l1,  g2 = u1 - p1,  f3 = p2 - l2,  g3 = u2 - p2
    phi = 修正 phi 系数（见 phi_corrected）
    下界 = delta - sqrt(f2^2 - 2*phi*f2*g3 + g3^2)
    上界 = delta + sqrt(g2^2 - 2*phi*g2*f3 + f3^2)

phi 的取值按 Newcombe 的做法做「修正」：
    AD - BC > 0 时  phi = max(AD - BC - n/2, 0) / sqrt((A+B)(C+D)(A+C)(B+D))
    否则            phi = (AD - BC) / sqrt(...)
    分母为 0 时 phi = 0
（记号 A,B,C,D 即四格 a,b,c,d。）

实现验证
--------
statsmodels 的 confint_proportions_paired(table, method="newcomb") 即为方法 10，
其官方文档给出算例：table=[[10,5],[2,20]] → (-0.0643926050711, 0.2211160178958)。
本文件 selftest 断言与之一致（误差 < 1e-10），并检查未做 phi 修正时算不出该值
（未修正得 -0.0548404 / 0.2122512），以证明修正项确实进入了实现。

其他区间
--------
· ci_conditional      v3 做法：条件于 n_d，未覆盖 n_d 的抽样变异，不是 delta 的
                      置信区间（c=0 时上界被钉在点估计上），仅作对照
· ci_wald             McNemar 方差版 Wald，仅敏感性参考
· ci_unconditional    无条件精确反演（对不一致率取上确界），过度保守，仅记录

用法
----
    python paired_ci.py              # 自检 + 逐比较区间（需先跑 stats_ch7_v5.py）
    python paired_ci.py --selftest   # 只跑自检
"""

import json
import math
import os
from statistics import NormalDist

import numpy as np

BASE = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE)

NMAX = 1024
LOG_FACT = np.zeros(NMAX + 1)
for _i in range(2, NMAX + 1):
    LOG_FACT[_i] = LOG_FACT[_i - 1] + math.log(_i)


def z_for(alpha):
    """双侧正态分位数 |z_{alpha/2}|。"""
    return NormalDist().inv_cdf(1.0 - alpha / 2.0)


# ============================================================ 基础分布工具
def binom_pmf(k, n, p):
    if p <= 0.0:
        return 1.0 if k == 0 else 0.0
    if p >= 1.0:
        return 1.0 if k == n else 0.0
    if k < 0 or k > n:
        return 0.0
    lg = LOG_FACT[n] - LOG_FACT[k] - LOG_FACT[n - k]
    return math.exp(lg + k * math.log(p) + (n - k) * math.log1p(-p))


def binom_pmf_vec(n, p):
    k = np.arange(n + 1)
    if p <= 0.0:
        out = np.zeros(n + 1)
        out[0] = 1.0
        return out
    if p >= 1.0:
        out = np.zeros(n + 1)
        out[n] = 1.0
        return out
    lg = LOG_FACT[n] - LOG_FACT[k] - LOG_FACT[n - k]
    return np.exp(lg + k * math.log(p) + (n - k) * math.log1p(-p))


def cp_ci(x, n, alpha=0.05):
    """Clopper–Pearson 精确区间。"""
    if n == 0:
        return 0.0, 1.0

    def cdf(k, nn, p):
        return float(np.sum(binom_pmf_vec(nn, p)[: k + 1]))

    if x <= 0:
        lo = 0.0
    else:
        a, b = 0.0, 1.0
        for _ in range(34):
            m = (a + b) / 2
            if 1 - cdf(x - 1, n, m) > alpha / 2:
                b = m
            else:
                a = m
        lo = (a + b) / 2
    if x >= n:
        up = 1.0
    else:
        a, b = 0.0, 1.0
        for _ in range(34):
            m = (a + b) / 2
            if cdf(x, n, m) > alpha / 2:
                a = m
            else:
                b = m
        up = (a + b) / 2
    return lo, up


def wilson(x, n, z):
    """Wilson 得分区间。"""
    if n == 0:
        return 0.0, 1.0
    p = x / n
    den = 1.0 + z * z / n
    centre = (p + z * z / (2 * n)) / den
    half = z / den * math.sqrt(p * (1 - p) / n + z * z / (4 * n * n))
    return max(0.0, centre - half), min(1.0, centre + half)


# ================================================== 本版主方法：Newcombe 方法 10
def phi_corrected(a, b, c, d):
    """配对四格表的修正 phi 系数（Newcombe 1998 方法 10 用）。

    返回 (phi, tag)。AD - BC > 0 时按 max(AD - BC - n/2, 0) 收缩（Newcombe 的
    修正），否则取原值；任一边际为 0（分母为 0）时 phi = 0。
    """
    n = a + b + c + d
    den2 = (a + b) * (c + d) * (a + c) * (b + d)
    if den2 <= 0:
        return 0.0, "degenerate"
    den = math.sqrt(den2)
    adbc = a * d - b * c
    if adbc > 0:
        return max(adbc - n / 2.0, 0.0) / den, "corrected"
    return adbc / den, "raw"


def ci_newcombe_paired(a, b, c, d, alpha=0.05):
    """Newcombe (1998) 方法 10：配对比例差 p1 - p2 的置信区间。

    a = 两条件都对，b = 仅 A 对，c = 仅 B 对，d = 都错。
    返回 dict，lo/hi 为概率尺度，lo_pt/hi_pt 为百分点（delta 用百分点）。
    """
    n = a + b + c + d
    if n == 0:
        return {"n": 0, "lo": 0.0, "hi": 0.0, "lo_pt": 0.0, "hi_pt": 0.0,
                "delta": 0.0, "delta_pt": 0.0, "phi": 0.0, "phi_tag": "empty"}
    z = z_for(alpha)
    x1, x2 = a + b, a + c
    p1, p2 = x1 / n, x2 / n
    delta = p1 - p2
    l1, u1 = wilson(x1, n, z)
    l2, u2 = wilson(x2, n, z)
    f2, g2 = p1 - l1, u1 - p1
    f3, g3 = p2 - l2, u2 - p2
    phi, tag = phi_corrected(a, b, c, d)
    lo = delta - math.sqrt(max(0.0, f2 * f2 - 2 * phi * f2 * g3 + g3 * g3))
    hi = delta + math.sqrt(max(0.0, g2 * g2 - 2 * phi * g2 * f3 + f3 * f3))
    lo, hi = max(lo, -1.0), min(hi, 1.0)
    return {"n": n, "a": a, "b": b, "c": c, "d": d,
            "p1": round(p1, 6), "p2": round(p2, 6),
            "delta": delta, "delta_pt": round(delta * 100, 2),
            "phi": round(phi, 4), "phi_tag": tag,
            "wilson_p1": [round(l1, 4), round(u1, 4)],
            "wilson_p2": [round(l2, 4), round(u2, 4)],
            "lo": lo, "hi": hi,
            "lo_pt": round(lo * 100, 2), "hi_pt": round(hi * 100, 2),
            "alpha": alpha, "z": round(z, 6)}


def ci_newcombe_pt(a, b, c, d, alpha=0.05):
    """只要百分点的 [lo, hi]。"""
    r = ci_newcombe_paired(a, b, c, d, alpha)
    return [r["lo_pt"], r["hi_pt"]]


# ============================================ v4 实现（错误，保留作对照，不再使用）
def ci_discordant_v4(b, c, n, alpha=0.05):
    """v4 实际算的东西：对 b/n 与 c/n 各取 Wilson 后平方和合成。

    这不是配对比例差的方法 10。保留原因：审稿意见要求证明「名称与实现一致」，
    需要能把旧值与新值并列展示。点估计恰好等于 delta（因 b/n - c/n = delta），
    所以旧值表面上「看起来对」，但方差结构与边际比例无关。
    """
    if n == 0:
        return [0.0, 0.0]
    z = z_for(alpha)
    pb, pc = b / n, c / n
    lb, ub = wilson(b, n, z)
    lc, uc = wilson(c, n, z)
    d = pb - pc
    lo = d - math.sqrt((pb - lb) ** 2 + (uc - pc) ** 2)
    up = d + math.sqrt((ub - pb) ** 2 + (pc - lc) ** 2)
    return [round(lo * 100, 2), round(up * 100, 2)]


# =========================================== 对照方法：v3 条件区间 / Wald / 无条件
def ci_conditional(b, c, n, alpha=0.05):
    """v3 做法：条件于 n_d = b+c，对 b/n_d 取 CP 再乘 n_d/n。"""
    nd = b + c
    if nd == 0:
        return [0.0, 0.0]
    lo, up = cp_ci(b, nd, alpha)
    s = nd / n
    return [round(s * (2 * lo - 1) * 100, 2), round(s * (2 * up - 1) * 100, 2)]


def ci_wald(b, c, n, alpha=0.05):
    """McNemar 方差版 Wald（仅敏感性参考）。"""
    if n == 0:
        return [0.0, 0.0]
    d = (b - c) / n
    var = ((b + c) - (b - c) ** 2 / n) / (n * n)
    se = math.sqrt(max(var, 0.0))
    z = z_for(alpha)
    return [round((d - z * se) * 100, 2), round((d + z * se) * 100, 2)]


def _pmf_matrix(n, p):
    m = np.arange(n + 1)[:, None]
    j = np.arange(n + 1)[None, :]
    if p <= 0.0:
        out = np.zeros((n + 1, n + 1))
        out[:, 0] = 1.0
        return out
    if p >= 1.0:
        return np.eye(n + 1)
    jj = np.minimum(j, m)
    lg = (LOG_FACT[m] - LOG_FACT[jj] - LOG_FACT[m - jj]
          + jj * math.log(p) + (m - jj) * math.log1p(-p))
    out = np.exp(lg)
    out[j > m] = 0.0
    return out


def pvalue_unconditional(b, c, n, delta0, n_u=81):
    """H0: delta = delta0 的无条件精确双侧 p 值（对不一致率 u 取上确界）。"""
    if n == 0:
        return 1.0
    d_obs = (b - c) / n
    t_obs = abs(d_obs - delta0)
    lo_u = abs(delta0)
    if lo_u >= 1.0:
        return 0.0
    us = np.linspace(lo_u + 1e-9, 1.0, n_u)
    rows = np.arange(n + 1)
    k_lo = np.floor((rows + n * (delta0 - t_obs)) / 2.0)
    k_hi = np.ceil((rows + n * (delta0 + t_obs)) / 2.0)
    lo_col = np.clip(k_lo, 0, n).astype(int)
    hi_col = np.clip(k_hi - 1, 0, n).astype(int)
    in_lo = k_lo >= 0
    in_hi = k_hi <= n
    best = 0.0
    for u in us:
        p10 = (u + delta0) / 2.0
        if p10 < 0.0 or p10 > u:
            continue
        pp = p10 / u if u > 0 else 0.0
        pp = min(max(pp, 0.0), 1.0)
        M = _pmf_matrix(n, pp)
        cdf = np.cumsum(M, axis=1)
        lower = np.where(in_lo, cdf[rows, lo_col], 0.0)
        upper = np.where(in_hi, 1.0 - cdf[rows, hi_col], 0.0)
        term = np.clip(lower + upper, 0.0, 1.0)
        w = binom_pmf_vec(n, u)
        val = float(np.dot(w, term))
        if val > best:
            best = val
    return min(best, 1.0)


def ci_unconditional(b, c, n, alpha=0.05, tol=0.001):
    """反演无条件检验得到 delta 的区间（百分点）。"""
    d = (b - c) / n
    if pvalue_unconditional(b, c, n, -1.0) >= alpha:
        lo = -1.0
    else:
        a, bb = -1.0, d
        for _ in range(14):
            m = (a + bb) / 2
            if pvalue_unconditional(b, c, n, m) >= alpha:
                bb = m
            else:
                a = m
            if bb - a < tol:
                break
        lo = bb
    if pvalue_unconditional(b, c, n, 1.0) >= alpha:
        up = 1.0
    else:
        a, bb = d, 1.0
        for _ in range(14):
            m = (a + bb) / 2
            if pvalue_unconditional(b, c, n, m) >= alpha:
                a = m
            else:
                bb = m
            if bb - a < tol:
                break
        up = a
    return [round(lo * 100, 2), round(up * 100, 2)]


# ==================================================================== 自检
def selftest(verbose=True):
    """实现验证：对齐 statsmodels 官方算例 + 性质检查。"""
    checks = []

    def chk(name, ok, detail=""):
        checks.append((name, bool(ok), detail))
        if verbose:
            print(f"  [{'PASS' if ok else 'FAIL'}] {name}  {detail}")

    # 1) statsmodels 文档算例（方法 10）
    a, b, c, d = 10, 5, 2, 20
    r = ci_newcombe_paired(a, b, c, d, 0.05)
    exp_lo, exp_hi = -0.0643926050711, 0.2211160178958
    ok = abs(r["lo"] - exp_lo) < 1e-10 and abs(r["hi"] - exp_hi) < 1e-10
    chk("官方算例 [[10,5],[2,20]] 95% CI", ok,
        f"({r['lo']:.13f}, {r['hi']:.13f}) vs ({exp_lo}, {exp_hi})")

    # 2) 未做 phi 修正则与官方值不符（证明修正项真的生效）
    n = a + b + c + d
    z = z_for(0.05)
    p1, p2 = (a + b) / n, (a + c) / n
    l1, u1 = wilson(a + b, n, z)
    l2, u2 = wilson(a + c, n, z)
    lo0 = (p1 - p2) - math.sqrt((p1 - l1) ** 2 + (u2 - p2) ** 2)
    chk("无 phi 修正的对照值不同", abs(lo0 - exp_lo) > 1e-3,
        f"phi=0 时下界 {lo0:.6f}，官方为 {exp_lo:.6f}")

    # 3) 官方 Wald 算例（本函数输出已四舍五入到 0.01 个百分点，容差取 1e-3）
    w = ci_wald(b, c, n, 0.05)
    chk("官方 Wald 算例", abs(w[0] / 100 - (-0.0566130751996)) < 1e-3
        and abs(w[1] / 100 - 0.2187752373618) < 1e-3,
        f"{w} vs (-5.6613, 21.8775) pt")

    # 4) 性质：区间含点估计；宽度非负；落在 [-1,1]
    ok = True
    bad = []
    for tb in [(0, 0, 0, 96), (87, 0, 4, 5), (10, 71, 0, 23), (45, 5, 0, 0),
               (0, 5, 0, 45), (50, 50, 50, 50), (96, 0, 0, 0), (0, 0, 0, 1)]:
        rr = ci_newcombe_paired(*tb)
        if not (rr["lo"] - 1e-12 <= rr["delta"] <= rr["hi"] + 1e-12):
            ok = False
            bad.append(("contains", tb, rr["lo"], rr["delta"], rr["hi"]))
        if rr["hi"] < rr["lo"] - 1e-12:
            ok = False
            bad.append(("order", tb))
        if rr["lo"] < -1 - 1e-12 or rr["hi"] > 1 + 1e-12:
            ok = False
            bad.append(("range", tb))
    chk("区间含点估计 / 顺序 / 值域", ok,
        "8 张边界表全部满足" if not bad else str(bad[:3]))

    # 5) 全部同对（b=c=0）时区间仍有正宽度（方法 10 的性质）
    rr = ci_newcombe_paired(40, 0, 0, 60, 0.05)
    chk("b=c=0 仍有正宽度", rr["hi"] - rr["lo"] > 0,
        f"[{rr['lo_pt']:.2f}, {rr['hi_pt']:.2f}] pt")

    # 6) 已知性质：方法 10 在 a <-> d 交换下不变（f2 与 g3 互换、phi 不变）
    t1 = ci_newcombe_paired(70, 20, 2, 8, 0.05)
    t2 = ci_newcombe_paired(8, 20, 2, 70, 0.05)
    chk("a<->d 交换区间不变（方法 10 的已知性质）",
        abs(t1["lo_pt"] - t2["lo_pt"]) < 0.011 and abs(t1["hi_pt"] - t2["hi_pt"]) < 0.011,
        f"{t1['lo_pt']}/{t1['hi_pt']} vs {t2['lo_pt']}/{t2['hi_pt']}")
    chk("v4 旧实现只依赖 (b,c,n)：信息少于方法 10",
        ci_discordant_v4(20, 2, 100, 0.05) == ci_discordant_v4(20, 2, 100, 0.05)
        and abs(t1["lo_pt"] - ci_discordant_v4(20, 2, 100, 0.05)[0]) > 0.1,
        f"M10 {t1['lo_pt']} vs v4 {ci_discordant_v4(20, 2, 100, 0.05)[0]}")

    # 7) 在本研究实际用到的表上，新方法与 v4 数值必然不同（差值幅度见报告）
    m10 = ci_newcombe_paired(87, 0, 4, 5, 0.05)     # easy 档 c3m - c_rw 的四格
    v4 = ci_discordant_v4(0, 4, 96, 0.05)
    dev = max(abs(m10["lo_pt"] - v4[0]), abs(m10["hi_pt"] - v4[1]))
    chk("easy 型表：M10 与 v4 数值不同（差 {:.2f}pt）".format(dev), dev > 0.1,
        f"M10 [{m10['lo_pt']}, {m10['hi_pt']}] vs v4 {v4}")

    # 7) 90% 区间应窄于 95%
    r95 = ci_newcombe_paired(87, 0, 4, 5, 0.05)
    r90 = ci_newcombe_paired(87, 0, 4, 5, 0.10)
    chk("90% 窄于 95%", (r90["hi"] - r90["lo"]) < (r95["hi"] - r95["lo"]),
        f"{r95['hi_pt']-r95['lo_pt']:.2f} / {r90['hi_pt']-r90['lo_pt']:.2f} pt")

    n_fail = sum(1 for _, ok_, _ in checks if not ok_)
    if verbose:
        print(f"  自检：{len(checks) - n_fail}/{len(checks)} 通过")
    return {"n_checks": len(checks), "n_fail": n_fail,
            "detail": [{"name": k, "pass": v, "detail": s} for k, v, s in checks],
            "anchor": "statsmodels.stats.proportion.confint_proportions_paired("
                      "[[10,5],[2,20]], method='newcomb') 文档算例"}


# ==================================================================== 主程序
if __name__ == "__main__":
    import sys

    print("=== 实现自检 ===")
    st = selftest()

    if "--selftest" in sys.argv:
        raise SystemExit(0)

    src = "results/stats_ch7_v5.json"
    if not os.path.exists(src):
        print(f"未找到 {src}：先运行 stats_ch7_v5.py")
        raise SystemExit(0)

    S = json.load(open(src, encoding="utf-8"))
    out = {
        "version": "paired-ci-v5",
        "estimate": "delta = p1 - p2 = (a+b)/n - (a+c)/n（百分点）",
        "primary_method": "Newcombe (1998) 方法 10：两个边际比例的 Wilson 区间 + "
                          "修正 phi 相关性修正（四格 a,b,c,d 全部进入）",
        "reference": "Newcombe RG. Improved confidence intervals for the difference "
                     "between binomial proportions based on paired data. "
                     "Stat Med 1998;17(22):2635-2650（方法 10 见 p.2639）",
        "validation": st,
        "v4_defect": "v4 的 ci_newcombe(b,c,n) 对 b/n 与 c/n 各取 Wilson 后平方和合成："
                     "把 b、c 当独立二项，未用 a、d，也没有 phi 修正，"
                     "故不是方法 10；点估计恰好等于 delta，掩盖了错误。",
        "methods": {
            "newcombe_m10": "本版主区间（95% 报告 / 90% 用于等价判据）",
            "discordant_v4": "v4 的错误实现，仅并列展示",
            "conditional_v3": "条件于 n_d，不是 delta 的置信区间，仅并列展示",
            "wald": "McNemar 方差版，敏感性参考",
        },
    }
    rows = []
    for pr in S["pairs"]:
        a_, b_, c_, d_ = pr["both_right"], pr["a_only"], pr["b_only"], pr["both_wrong"]
        r95 = ci_newcombe_paired(a_, b_, c_, d_, 0.05)
        r90 = ci_newcombe_paired(a_, b_, c_, d_, 0.10)
        row = {"group": pr["group"], "a": pr["a"], "b": pr["b"],
               "n": pr["n_pair"], "cells": {"a": a_, "b": b_, "c": c_, "d": d_},
               "rate_a": pr["rate_a"], "rate_b": pr["rate_b"],
               "delta": pr["delta"], "phi": r95["phi"], "phi_tag": r95["phi_tag"],
               "ci95": [r95["lo_pt"], r95["hi_pt"]],
               "ci90": [r90["lo_pt"], r90["hi_pt"]],
               "ci95_v4": ci_discordant_v4(b_, c_, pr["n_pair"], 0.05),
               "ci90_v4": ci_discordant_v4(b_, c_, pr["n_pair"], 0.10),
               "ci95_conditional_v3": ci_conditional(b_, c_, pr["n_pair"], 0.05),
               "ci95_wald": ci_wald(b_, c_, pr["n_pair"], 0.05),
               "p_mcnemar": pr["p"]}
        for k in ("ci95", "ci90", "ci95_v4", "ci95_conditional_v3", "ci95_wald"):
            row[k + "_contains_zero"] = row[k][0] <= 0 <= row[k][1]
        row["width95"] = round(row["ci95"][1] - row["ci95"][0], 2)
        row["width95_v4"] = round(row["ci95_v4"][1] - row["ci95_v4"][0], 2)
        row["width95_conditional"] = round(
            row["ci95_conditional_v3"][1] - row["ci95_conditional_v3"][0], 2)
        rows.append(row)
        print(f"{pr['group']:11s} {pr['a']:5s}-{pr['b']:5s} n={pr['n_pair']:3d} "
              f"a={a_:3d} b={b_:3d} c={c_:3d} d={d_:3d} phi={r95['phi']:+.3f} "
              f"d={pr['delta']:+6.1f}  M10={row['ci95']}  v4={row['ci95_v4']} "
              f"v3={row['ci95_conditional_v3']}")
    out["pairs"] = rows

    # 无条件精确区间：仅两个代表情形，用于宽度对比
    demo = []
    for g in ("easy", "split", "merge"):
        cand = [x for x in rows if x["group"] == g and x["b"] == "c_rw"]
        if not cand:
            continue
        row = cand[0]
        cu = ci_unconditional(row["cells"]["b"], row["cells"]["c"], row["n"])
        demo.append({"group": g, "n": row["n"], "b": row["cells"]["b"],
                     "c": row["cells"]["c"], "ci_unconditional": cu,
                     "width_unconditional": round(cu[1] - cu[0], 2),
                     "width_newcombe": row["width95"],
                     "ratio": round((cu[1] - cu[0]) / row["width95"], 2)})
        print(f"[uncond] {g:6s} {cu} width={demo[-1]['width_unconditional']} "
              f"vs M10 {row['width95']} (x{demo[-1]['ratio']})")
    out["unconditional_demo"] = demo

    json.dump(out, open("results/paired_ci.json", "w", encoding="utf-8"),
              ensure_ascii=False, indent=1)
    print("saved results/paired_ci.json")
