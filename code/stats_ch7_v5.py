# -*- coding: utf-8 -*-
"""第 7 章统计量（v5，取代 stats_ch7_v4.py）。

相对 v4 的改动（起因：第四轮评审指出配对区间的实现与方法名不符）
1) 配对差值的区间：v4 的 PCI.ci_newcombe(b, c, n) 对 b/n 与 c/n 各取 Wilson
   后平方和合成，只用了两个不一致格，把 b、c 当独立二项，没有 a、d，也没有
   phi 相关性修正——它不是 Newcombe 方法 10。v5 改为完整四格
   (a=都对, b=仅 A 对, c=仅 B 对, d=都错) 输入：边际比例 p_A=(a+b)/n、
   p_B=(a+c)/n 各取 Wilson，再用修正 phi 修正相关性（见 paired_ci.py，已对齐
   statsmodels 官方算例）。旧值保留为 ci_v4_retired 供对照。
2) 覆盖率证据改为「情景 + 参数网格」两段（results/paired_ci_coverage.json），
   结论限定在网格范围内，不再由五个情景推断全参数范围。
3) 复现链：meta 记录每个输入文件的 sha1、脚本版本与生成顺序；
   results/stats_cluster.json 的 bootstrap 次数与种子继续登记。

输出 results/stats_ch7_v5.json，供 make_chapter7_v5.py 使用。
"""
import collections
import hashlib
import importlib.util
import json
import math
import os

BASE = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE)

_spec = importlib.util.spec_from_file_location(
    "paired_ci", os.path.join(BASE, "paired_ci.py"))
PCI = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(PCI)

VERSION = "ch7-stats-v5"
Z95 = 1.959964
Z90 = 1.644854


def sha1_of(path):
    h = hashlib.sha1()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


MAN = {}
for l in open("data/manifest_v3.jsonl", encoding="utf-8"):
    r = json.loads(l)
    MAN[r["id"]] = r

BY = collections.defaultdict(dict)
FILES = ["e1_dev_full.jsonl", "e1_diff_3lvl.jsonl", "e1_c3m_paired.jsonl",
         "e1_expand_med.jsonl", "e1_expand_hard.jsonl", "e1_ds_c3f.jsonl",
         "e1_ms_ds.jsonl", "e1_ms_ds_p.jsonl", "e1_ms_ds_p_split.jsonl",
         "e1_ds_one.jsonl", "e1_ds_one_b1.jsonl", "e1_ds_one_b2.jsonl",
         "e3_rw_main.jsonl"]
USED = []
for f in FILES:
    p = os.path.join("results", f)
    if not os.path.exists(p):
        continue
    USED.append(f)
    for l in open(p, encoding="utf-8"):
        r = json.loads(l)
        BY[r["id"]][r["cond"]] = r["outcome"]

IDS_XM = [l.strip() for l in open("data/ids_xm_full.txt", encoding="utf-8") if l.strip()]
IDS_MS = [l.strip() for l in open("data/ids_ms.txt", encoding="utf-8") if l.strip()]

GROUPS = {
    "easy": ("rename", IDS_XM),
    "medium": ("rename_medium", IDS_XM),
    "hard": ("rename_hard", IDS_XM),
    "rename_all": (("rename", "rename_medium", "rename_hard"), IDS_XM),
    "merge": ("merge_col", IDS_MS),
    "split": ("split_col", IDS_MS),
}


def ids_of(name):
    evo, pool = GROUPS[name]
    evos = {evo} if isinstance(evo, str) else set(evo)
    return [i for i in pool if MAN[i]["evo_type"] in evos]


# ============================================================ 基础分布工具
def wilson(k, n, z=Z95):
    """单条件正确率的 Wilson 得分区间（百分点）。"""
    if n == 0:
        return [0.0, 0.0]
    p = k / n
    d = 1 + z * z / n
    c = p + z * z / (2 * n)
    h = z * math.sqrt(p * (1 - p) / n + z * z / (4 * n * n))
    return [round((c - h) / d * 100, 1), round((c + h) / d * 100, 1)]


def mcnemar_exact(a_only, b_only):
    """精确双侧 McNemar（条件二项，p=0.5）。"""
    n = a_only + b_only
    if n == 0:
        return 1.0
    k = min(a_only, b_only)
    return min(1.0, 2.0 * sum(math.comb(n, j) for j in range(k + 1)) / (2 ** n))


def binom_cdf(k, n, p):
    if p <= 0.0:
        return 1.0 if k >= 0 else 0.0
    if p >= 1.0:
        return 1.0 if k >= n else 0.0
    return sum(math.comb(n, j) * p ** j * (1 - p) ** (n - j) for j in range(0, k + 1))


def binom_sf(k, n, p):
    """P(X >= k)。"""
    return 1.0 - binom_cdf(k - 1, n, p)


def cp_ci(x, n, alpha=0.05):
    """Clopper–Pearson 精确区间：x ~ Bin(n, p)。"""
    if n == 0:
        return 0.0, 1.0
    if x == 0:
        lo = 0.0
    else:
        lo, hi = 0.0, 1.0
        for _ in range(200):
            mid = (lo + hi) / 2
            if binom_sf(x, n, mid) < alpha / 2:
                lo = mid
            else:
                hi = mid
        lo = (lo + hi) / 2
    if x == n:
        up = 1.0
    else:
        lo2, hi2 = 0.0, 1.0
        for _ in range(200):
            mid = (lo2 + hi2) / 2
            if binom_cdf(x, n, mid) > alpha / 2:
                lo2 = mid
            else:
                hi2 = mid
        up = (lo2 + hi2) / 2
    return lo, up


def paired_exact_ci(a_only, b_only, n, alpha=0.05):
    """配对差值（a − b）的精确条件区间（百分点）。

    条件于不一致对总数 n_d = a_only + b_only，
    θ = a_only / n_d 表示不一致对中偏向 a 的比例，
    差值 = (n_d / n) * (2θ − 1)。对 θ 取 Clopper–Pearson 区间再线性变换，
    该区间与精确 McNemar 检验同源：区间包含 0 ⟺ p ≥ alpha。
    """
    nd = a_only + b_only
    if nd == 0:
        return [0.0, 0.0], {"nd": 0, "theta_ci": [None, None]}
    lo, up = cp_ci(a_only, nd, alpha)
    scale = nd / n
    d_lo = scale * (2 * lo - 1) * 100
    d_up = scale * (2 * up - 1) * 100
    return ([round(d_lo, 1), round(d_up, 1)],
            {"nd": nd, "theta_ci": [round(lo, 4), round(up, 4)]})


def paired_wald_ci(a_only, b_only, n, alpha=0.05):
    """配对差值的 Wald 区间（百分点），仅作敏感性参考。"""
    nd = a_only + b_only
    if nd == 0:
        return [0.0, 0.0], 0.0
    var = (nd - (a_only - b_only) ** 2 / n) / (n * n)
    se = math.sqrt(max(var, 0.0))
    z = Z95 if abs(alpha - 0.05) < 1e-9 else Z90
    d = (a_only - b_only) / n
    return [round((d - z * se) * 100, 1), round((d + z * se) * 100, 1)], se


def tost_paired(a_only, b_only, n, margin_pt):
    """配对二元结果的 TOST 等价检验（Wald 版），margin 以百分点给出。

    H0_lower: δ ≤ −m  vs  H1: δ > −m
    H0_upper: δ ≥ +m  vs  H1: δ < +m
    TOST p = max(两者)。
    """
    nd = a_only + b_only
    if nd == 0:
        return {"margin_pt": margin_pt, "p_lower": None, "p_upper": None,
                "p_tost": None, "note": "无不一致对，无法实施检验"}
    var = (nd - (a_only - b_only) ** 2 / n) / (n * n)
    se = math.sqrt(max(var, 1e-18))
    d = (a_only - b_only) / n
    m = margin_pt / 100.0
    # H0_lower: δ ≤ −m  →  Z=(δ̂+m)/SE，右侧拒绝
    # H0_upper: δ ≥ +m  →  Z=(δ̂−m)/SE，左侧拒绝
    z_lower = (d + m) / se
    z_upper = (d - m) / se
    p_lower = 1 - 0.5 * (1 + math.erf(z_lower / math.sqrt(2)))
    p_upper = 0.5 * (1 + math.erf(z_upper / math.sqrt(2)))
    return {"margin_pt": margin_pt,
            "p_lower": round(p_lower, 5), "p_upper": round(p_upper, 5),
            "p_tost": round(max(p_lower, p_upper), 5),
            "pass": max(p_lower, p_upper) < 0.05}


def tost_exact(a_only, b_only, n, margin_pt):
    """等价性的精确条件判据：90% 精确条件区间是否整体落入 ±margin。

    警惕：条件于不一致对数 n_d，差值最大只能到 ±n_d/n。当 n_d/n ≤ margin 时，
    该判据必然通过，不构成等价证据，故以 degenerate 标记出来。
    """
    ci90, _ = paired_exact_ci(a_only, b_only, n, alpha=0.10)
    ok = (ci90[0] > -margin_pt) and (ci90[1] < margin_pt)
    max_attr = (a_only + b_only) / n * 100
    return {"margin_pt": margin_pt, "ci90_exact": ci90, "pass": ok,
            "max_attainable_pt": round(max_attr, 2),
            "degenerate": max_attr <= margin_pt}


# ============================================================ 主结果表
main = {}
for gname in GROUPS:
    ids = ids_of(gname)
    row = {"n": len(ids)}
    for cond in ("c1", "c2", "c3m", "c3mp", "c_rw"):
        have = [i for i in ids if cond in BY[i]]
        if len(have) < len(ids) * 0.5:
            continue
        k = sum(1 for i in have if BY[i][cond] == "EX")
        row[cond] = {"k": k, "n": len(have),
                     "rate": round(k / len(have) * 100, 1),
                     "ci": wilson(k, len(have))}
    main[gname] = row

# ============================================================ 配对表
PAIR_SPECS = [("easy", "c3m", "c_rw"), ("medium", "c3m", "c_rw"),
              ("hard", "c3m", "c_rw"), ("merge", "c3mp", "c_rw"),
              ("split", "c3mp", "c_rw"), ("rename_all", "c3m", "c_rw"),
              ("rename_all", "c3m", "c2"), ("merge", "c3mp", "c2"),
              ("split", "c3m", "c2")]

MARGINS = [3, 5, 10]

pairs = []
for gname, a, b in PAIR_SPECS:
    ids = [i for i in ids_of(gname) if a in BY[i] and b in BY[i]]
    n = len(ids)
    ka = sum(1 for i in ids if BY[i][a] == "EX")
    kb = sum(1 for i in ids if BY[i][b] == "EX")
    ao = sum(1 for i in ids if BY[i][a] == "EX" and BY[i][b] != "EX")
    bo = sum(1 for i in ids if BY[i][a] != "EX" and BY[i][b] == "EX")
    # 四格：a=都对，b=仅 A 对，c=仅 B 对，d=都错
    both_right = sum(1 for i in ids if BY[i][a] == "EX" and BY[i][b] == "EX")
    both_wrong = n - both_right - ao - bo
    assert both_right + ao == ka and both_right + bo == kb, (gname, a, b, ka, kb)
    assert both_right + ao + bo + both_wrong == n, (gname, a, b, n)
    assert both_wrong >= 0
    ci_exact, ci_meta = paired_exact_ci(ao, bo, n, alpha=0.05)
    ci_wald, se = paired_wald_ci(ao, bo, n)
    m10 = PCI.ci_newcombe_paired(both_right, ao, bo, both_wrong, alpha=0.05)
    m10_90 = PCI.ci_newcombe_paired(both_right, ao, bo, both_wrong, alpha=0.10)
    ci_nc = [m10["lo_pt"], m10["hi_pt"]]
    ci_nc90 = [m10_90["lo_pt"], m10_90["hi_pt"]]
    assert abs(m10["delta"] * 100 - (ka - kb) / n * 100) < 1e-9, (gname, a, b)
    ci_v4 = PCI.ci_discordant_v4(ao, bo, n, alpha=0.05)
    rec = {
        "group": gname, "a": a, "b": b, "n_pair": n,
        "k_a": ka, "k_b": kb,
        "rate_a": round(ka / n * 100, 1), "rate_b": round(kb / n * 100, 1),
        "a_only": ao, "b_only": bo,
        "both_right": both_right, "both_wrong": both_wrong,
        "delta": round((ka - kb) / n * 100, 1),
        "p": mcnemar_exact(ao, bo),
        "ci": ci_nc,
        "ci_method": "newcombe-1998-method10-four-cells",
        "ci90_newcombe": ci_nc90,
        "phi": m10["phi"], "phi_tag": m10["phi_tag"],
        "wilson_marginals": {"p_A": m10["wilson_p1"], "p_B": m10["wilson_p2"]},
        "ci_v4_retired": ci_v4,
        "ci_dev_v4_pt": round(max(abs(ci_nc[0] - ci_v4[0]),
                                  abs(ci_nc[1] - ci_v4[1])), 2),
        "ci_conditional_v3": ci_exact,
        "ci_wald_sensitivity": ci_wald,
        "n_discordant": ci_meta["nd"],
        "theta_ci": ci_meta["theta_ci"],
        "tost": [tost_paired(ao, bo, n, m) for m in MARGINS],
        "tost_exact": [tost_exact(ao, bo, n, m) for m in MARGINS],
        "tost_newcombe": [{"margin_pt": m, "ci90": ci_nc90,
                           "pass": (ci_nc90[0] > -m and ci_nc90[1] < m)}
                          for m in MARGINS],
    }
    pairs.append(rec)

# ============================================================ Holm：族固定为 5 项
HOLM_FAMILY = [("easy", "c3m", "c_rw"), ("medium", "c3m", "c_rw"),
               ("hard", "c3m", "c_rw"), ("merge", "c3mp", "c_rw"),
               ("split", "c3mp", "c_rw")]
IDX = {(p["group"], p["a"], p["b"]): p for p in pairs}
fam = [IDX[k] for k in HOLM_FAMILY]
for it in fam:
    it["in_holm_family"] = True
for p in pairs:
    if not p.get("in_holm_family"):
        p["in_holm_family"] = False
order = sorted(fam, key=lambda x: x["p"])
m_len = len(order)
for idx, it in enumerate(order):
    it["p_holm"] = min(1.0, it["p"] * (m_len - idx))
    it["holm_rank"] = idx + 1

# ============================================================ 两层评测范围
def uniq_base(ids):
    return len({(MAN[i]["db_id"], str(MAN[i]["question_id"])) for i in ids})


IDS_VS = [i for i in (l.strip() for l in open("data/ids_vshift.txt", encoding="utf-8"))
          if i and MAN[i]["evo_type"] == "value_shift"]
IDS_VN = [i for i in (l.strip() for l in open("data/ids_vs_xm.txt", encoding="utf-8"))
          if i and MAN[i]["evo_type"] == "value_shift_named"]
IDS_DP = [i for i in (l.strip() for l in open("data/ids_dp_xm.txt", encoding="utf-8"))
          if i and MAN[i]["evo_type"] == "dep_abstain"]

E2E = IDS_XM + IDS_MS
LINEAGE = E2E + IDS_VS + IDS_VN + IDS_DP * 2

scope = {
    "e2e": {"items": len(E2E), "base": uniq_base(E2E),
            "dbs": len({MAN[i]["db_id"] for i in E2E})},
    "lineage": {"items": len(LINEAGE),
                "base": uniq_base(E2E + IDS_VS + IDS_VN + IDS_DP),
                "dbs": len({MAN[i]["db_id"] for i in LINEAGE})},
    "tiers": [
        {"name": "改名 easy", "items": len(ids_of("easy")), "base": uniq_base(ids_of("easy")),
         "dbs": len({MAN[i]["db_id"] for i in ids_of("easy")}), "layer": "两"},
        {"name": "改名 medium", "items": len(ids_of("medium")), "base": uniq_base(ids_of("medium")),
         "dbs": len({MAN[i]["db_id"] for i in ids_of("medium")}), "layer": "两"},
        {"name": "改名 hard", "items": len(ids_of("hard")), "base": uniq_base(ids_of("hard")),
         "dbs": len({MAN[i]["db_id"] for i in ids_of("hard")}), "layer": "两"},
        {"name": "合并", "items": len(ids_of("merge")), "base": uniq_base(ids_of("merge")),
         "dbs": len({MAN[i]["db_id"] for i in ids_of("merge")}), "layer": "两"},
        {"name": "拆分", "items": len(ids_of("split")), "base": uniq_base(ids_of("split")),
         "dbs": len({MAN[i]["db_id"] for i in ids_of("split")}), "layer": "两"},
        {"name": "值改写", "items": len(IDS_VS), "base": uniq_base(IDS_VS),
         "dbs": len({MAN[i]["db_id"] for i in IDS_VS}), "layer": "抽取"},
        {"name": "值改写（语义名）", "items": len(IDS_VN), "base": uniq_base(IDS_VN),
         "dbs": len({MAN[i]["db_id"] for i in IDS_VN}), "layer": "抽取"},
        {"name": "依赖后继", "items": len(IDS_DP), "base": uniq_base(IDS_DP),
         "dbs": len({MAN[i]["db_id"] for i in IDS_DP}), "layer": "抽取"},
        {"name": "依赖弃答", "items": len(IDS_DP), "base": uniq_base(IDS_DP),
         "dbs": len({MAN[i]["db_id"] for i in IDS_DP}), "layer": "抽取"},
    ],
}

units = {}
for gname in ("merge", "split", "rename_all"):
    ids = ids_of(gname)
    ops = set()
    for i in ids:
        for op in MAN[i].get("ops") or []:
            ops.add((MAN[i]["db_id"], tuple(str(x) for x in op)))
    units[gname] = {"items": len(ids), "db": len({MAN[i]["db_id"] for i in ids}),
                    "op_facts": len(ops),
                    "items_per_op": round(len(ids) / len(ops), 1) if ops else None}

# ============================================================ 跨基座
XM = collections.defaultdict(lambda: collections.defaultdict(dict))
for f in os.listdir("results"):
    if not f.endswith(".jsonl"):
        continue
    for l in open(os.path.join("results", f), encoding="utf-8"):
        try:
            r = json.loads(l)
        except Exception:
            continue
        if r["id"] in set(IDS_XM) and r.get("cond") in ("c1", "c2", "c3f", "c3m"):
            XM[r.get("model")][r["cond"]][r["id"]] = r["outcome"]

xmodel = {}
for m_, cs in XM.items():
    if "c3m" not in cs:
        continue
    ids = sorted(cs["c3m"])
    row = {"model": m_, "n": len(ids)}
    for c in ("c1", "c2", "c3f", "c3m"):
        if c in cs:
            k = sum(1 for i in ids if cs[c].get(i) == "EX")
            row[c] = {"k": k, "n": len(ids), "rate": round(k / len(ids) * 100, 1)}
    if "c2" in row and "c3m" in row:
        row["delta"] = round(row["c3m"]["rate"] - row["c2"]["rate"], 1)
    xmodel[m_] = row


# ============================================================ 判定分布
def dist(gname, cond):
    ids = [i for i in ids_of(gname) if cond in BY[i]]
    c = collections.Counter(BY[i][cond] for i in ids)
    return {"n": len(ids), "EX": c["EX"], "SILENT": c["SILENT"],
            "ERROR": c["ERROR"], "OVER_ABSTAIN": c["OVER_ABSTAIN"],
            "ABSTAIN_OK": c["ABSTAIN_OK"], "HALLUCINATE": c["HALLUCINATE"]}


outcomes = []
for gname, cond, label in [("merge", "c2", "合并 · C2 旧 SQL"),
                           ("merge", "c3m", "合并 · C3m 列级"),
                           ("merge", "c3mp", "合并 · C3mp 段位"),
                           ("merge", "c_rw", "合并 · c_rw 改写"),
                           ("split", "c2", "拆分 · C2 旧 SQL"),
                           ("split", "c3m", "拆分 · C3m 列级"),
                           ("easy", "c3m", "改名 easy · C3m"),
                           ("easy", "c_rw", "改名 easy · c_rw")]:
    d = dist(gname, cond)
    d["label"] = label
    outcomes.append(d)

CLUSTER = json.load(open("results/stats_cluster.json", encoding="utf-8"))

COV_PATH = "results/paired_ci_coverage.json"
COV = json.load(open(COV_PATH, encoding="utf-8")) if os.path.exists(COV_PATH) else None

meta = {
    "version": VERSION,
    "supersedes": "stats_ch7_v4.py / results/stats_ch7_v4.json",
    "ci_single": "Wilson score, z=1.96",
    "estimate_paired": "delta = p_A - p_B = (a+b)/n - (a+c)/n（边际准确率差，百分点）",
    "ci_paired_primary": "Newcombe (1998) 方法 10，完整四格输入：边际比例 p_A=(a+b)/n 与 "
                         "p_B=(a+c)/n 各取 Wilson 区间，再用修正 phi 系数做相关性修正；"
                         "a=两条件都对、b=仅 A 对、c=仅 B 对、d=都错",
    "ci_paired_reference": "Newcombe RG. Improved confidence intervals for the difference "
                           "between binomial proportions based on paired data. "
                           "Stat Med 1998;17(22):2635-2650（方法 10 见 p.2639）",
    "ci_paired_validation": "实现与 statsmodels.stats.proportion.confint_proportions_paired"
                            "(method='newcomb') 的文档算例对齐到 1e-10"
                            "（table=[[10,5],[2,20]] → (-0.0643926050711, 0.2211160178958)）",
    "ci_paired_v4_defect": "v4 的 ci_newcombe(b,c,n) 只对两个不一致格各取 Wilson 后平方合成，"
                           "把 b、c 当作独立二项，未使用 a、d，也没有 phi 修正，故不是方法 10；"
                           "点估计恰好等于 delta 使错误未暴露。旧值保留为 ci_v4_retired。",
    "ci_paired_a_d_invariance": "方法 10 在 a<->d 交换下严格不变（f2 与 g3 互换、phi 不变）："
                                "区间只通过 phi 与边际 Wilson 宽度使用 a、d，"
                                "故本方法对「两条件都强/都弱」不区分，属方法的已知限制",
    "ci_paired_deprecated": "exact-conditional（v3 做法）：条件于 n_d，"
                            "未覆盖 n_d 的抽样变异，不是 delta 的置信区间；"
                            "保留为 ci_conditional_v3 仅供对照",
    "ci_paired_sensitivity": "Wald with McNemar variance（仅敏感性，不用于判方向）",
    "ci_paired_unconditional_note": "亦实现无条件精确反演（对不一致率取上确界），"
                                    "宽度约为方法 10 的 1.5-4 倍，过度保守，未采用",
    "coverage_sim_file": COV_PATH,
    "coverage_scope_note": "覆盖率证据分两段：五个与实际档位对应的情景 + 一个参数网格；"
                           "结论只对网格覆盖的 (p_A, p_B, phi, n) 区域成立，"
                           "不外推到全参数空间",
    "bootstrap_removed": "仅指配对差值的条目级百分位 bootstrap 已弃用；"
                         "聚类 bootstrap 仍在使用（5,000 次重采样），见表 7-19",
    "holm_family": [f"{g}:{a}-{b}" for g, a, b in HOLM_FAMILY],
    "holm_family_size": len(HOLM_FAMILY),
    "holm_scope_note": "仅 c_rw 相关的 5 项对照进入 Holm 族；"
                       "与 C2 的对照不校正，亦不携带 p_holm 字段",
    "tost_margins_pt": MARGINS,
    "tost_note": "margin 非预注册，为结果可见后设定；故同时报 ±3/±5/±10 三档",
    "tost_degenerate_note": "条件于不一致对数，差值上限为 n_d/n；当该上限不超过 margin 时，"
                            "精确判据必然通过，属退化情形，不能作为等价证据",
    "tost_ci_basis": "等价判据用 90% 方法 10 区间（完整四格），不用 v4 的旧区间",
    "inputs": USED,
    "input_sha1": {f"results/{f}": sha1_of(os.path.join("results", f)) for f in USED},
    "generation_chain": [
        "stats_ch7_v5.py 从 results/*.jsonl（冻结的逐题判定）重算全部计数与检验",
        "配对区间实现来自 paired_ci.py（v5 四格方法 10，含 selftest）",
        "覆盖率来自 paired_coverage.py → results/paired_ci_coverage.json",
        "make_chapter7_v5.py 只读 stats_ch7_v5.json / paired_ci.json / "
        "paired_ci_coverage.json，正文数值不自带常量",
    ],
    "derived_inputs": {
        "results/stats_cluster.json": {
            "script": "stats_cluster.py",
            "bootstrap_resamples": 5000,
            "bootstrap_seed": 7,
            "used_by": "表 7-19（基础题级聚类检验）",
            "note": "该文件由 stats_cluster.py 从同一批 jsonl 生成；"
                    "其 bootstrap 为簇级重采样，与已弃用的条目级 bootstrap 不同",
        },
        COV_PATH: {
            "script": "paired_coverage.py",
            "content": "方法 10（四格）/ v4 旧实现 / v3 条件区间 / Wald 的覆盖率："
                       "五个情景 + 参数网格",
            "used_by": "表 7-6、7.1.5 节覆盖率段",
        },
    },
    "cluster_stats": {
        "tests": len(CLUSTER.get("tests", [])),
        "bootstrap": 5000,
        "seed": 7,
    },
}
if COV:
    meta["coverage_headline"] = COV.get("headline")
out = {"main": main, "pairs": pairs, "scope": scope, "units": units,
       "meta": meta, "xmodel": xmodel, "outcomes": outcomes}
os.makedirs("results", exist_ok=True)
json.dump(out, open("results/stats_ch7_v5.json", "w", encoding="utf-8"),
          ensure_ascii=False, indent=1)

print("=== 主结果 ===")
for k, v in main.items():
    s = " ".join(f"{c}={v[c]['rate']}({v[c]['k']}/{v[c]['n']}){v[c]['ci']}"
                 for c in ("c1", "c2", "c3m", "c3mp", "c_rw") if c in v)
    print(f"  {k:11s} n={v['n']:3d} {s}")
print("\n=== 配对（四格 / 方法 10 / TOST） ===")
print(f"{'group':11s} {'cmp':14s} {'a':>3s} {'b':>3s} {'c':>2s} {'d':>3s} {'phi':>6s} "
      f"{'Δ':>7s} {'CI_M10':>16s} {'CI_v4':>15s} {'dev':>5s} {'p':>9s} {'p_holm':>9s}")
for p in pairs:
    print(f"{p['group']:11s} {p['a']+'-'+p['b']:14s} {p['both_right']:3d} "
          f"{p['a_only']:3d} {p['b_only']:2d} {p['both_wrong']:3d} {p['phi']:+6.3f} "
          f"{p['delta']:+7.1f} {str(p['ci']):>16s} {str(p['ci_v4_retired']):>15s} "
          f"{p['ci_dev_v4_pt']:5.2f} {p['p']:9.3g} {p.get('p_holm','—'):>9}")
print("\n=== 等价判据：90% 方法 10 区间 vs 界值 ===")
for p in pairs:
    if p["b"] != "c_rw":
        continue
    line = f"  {p['group']:11s} CI90={p['ci90_newcombe']}"
    for t in p["tost_newcombe"]:
        line += f"  ±{t['margin_pt']}:{'PASS' if t['pass'] else 'fail'}"
    line += f"   | Wald TOST ±5 p={[t for t in p['tost'] if t['margin_pt']==5][0]['p_tost']}"
    print(line)
if COV:
    print("\n=== 覆盖率 ===")
    print("  情景：", COV.get("headline"))
