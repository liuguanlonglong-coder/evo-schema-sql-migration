# -*- coding: utf-8 -*-
"""M1 与 schema matching 基线在「段位档 / 非段位档」上的分段对照。

回答审稿人唯一的创新性质询：
    自动谱系抽取的增益，是来自 M1 的设计，还是任何合理的列匹配都能达到？

做法：全部方法走同一个 evaluate()（m1_lineage.py），同一份 gold、同一套 op 键，
差异只在「谱系从哪来」。因此不存在「基线用自己的评测器」的不公平对照。

分组依据（本文件的核心设计）：
    段位档   = merge_col + split_col  —— gold 要求输出「新列 + 段位表达式」，
                                        列对形式本身信息不足
    非段位档 = 其余六个演化类型       —— gold 是列对，通用匹配器的主场

两个指标：
    目标列恢复率  recovered —— 预测的列集合与 gold 列集合相同（不论类型）
    op-type 一致率 type_ok  —— 类型也对得上；在段位档上等价于「段位是否恢复」

输出 results/ablation_matching.json，并打印可直接进论文的表。
"""
import collections
import json
import os
import sys

BASE = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE)
sys.path.insert(0, BASE)

from m1_lineage import gold_by_keys, evaluate  # noqa: E402

LIN = os.path.join(BASE, "data", "lineage_infer")
OUT = os.path.join(BASE, "results", "ablation_matching.json")
MANIFEST = "data/manifest_v3.jsonl"

# (显示名, 文件名后缀, 类别)
METHS = [
    ("B1 列名相似",        "name",    "baseline"),
    ("B2 值域重叠",        "vover",   "baseline"),
    ("B3 嵌入语义",        "emb",     "baseline"),
    ("M1 值域指纹",        "profile", "ours"),
    ("M1+形态指纹",        "shape",   "ours"),
    ("P0 位置先验（消融）", "pos",     "ablation"),
]

EMBED_LEVELS = ["merge_col", "split_col"]
ALL_LEVELS = ["rename", "rename_medium", "rename_hard", "merge_col", "split_col",
              "value_shift", "value_shift_named", "dep_abstain", "dep_successor"]


def load_infer(suffix, keys):
    """按 (evo, db) 读逐库谱系文件；缺文件记 None，不静默当空。"""
    got, missing = {}, []
    for evo, db in keys:
        p = os.path.join(LIN, f"{evo}_{db}_{suffix}.json")
        if os.path.exists(p):
            got[(evo, db)] = json.load(open(p, encoding="utf-8"))
        else:
            missing.append(f"{evo}/{db}")
    return got, missing


def collect_type_ok(detail):
    """把 evaluate() 的 detail 转成 {(evo, table, old): type_ok}，供 op 级配对。"""
    out = {}
    for d in detail:
        preds = d["pred"]
        gold = d["gold"]
        ok = any(p["type"] == gold or
                 (gold in ("deprecate_with_successor", "rename_and_recode")
                  and p["type"] == "rename_column")
                 for p in preds)
        out[(d["evo"], d["table"].lower(), d["old"].lower())] = bool(ok)
    return out


def sign_test(a_only, b_only):
    """精确双侧符号检验，只在不一致对上算。"""
    n = a_only + b_only
    if n == 0:
        return 1.0, 0
    k = min(a_only, b_only)
    tail = sum(_comb(n, i) for i in range(k + 1)) / (2 ** n)
    return min(1.0, 2 * tail), n


def _comb(n, k):
    r = 1
    for i in range(k):
        r = r * (n - i) // (i + 1)
    return r


def main():
    rows = [json.loads(l) for l in open(MANIFEST, encoding="utf-8")]
    gk = gold_by_keys(rows)
    keys = sorted(gk)

    table = {}
    type_ok_map = {}
    for label, suf, cat in METHS:
        infer, missing = load_infer(suf, keys)
        if missing and cat != "ablation":
            print(f"[warn] {label} 缺 {len(missing)} 个文件，例：{missing[:3]}")
        if missing and cat == "ablation":
            # 消融项只在部分演化上跑过，逐库文件不全；改读它自己的 eval 产物，
            # 否则会把「没跑」误记成「全错」（上一版就是这样得出 0/162 的）。
            ev = os.path.join(LIN, f"eval_{suf}.json")
            if os.path.exists(ev):
                src = json.load(open(ev, encoding="utf-8"))
                stat = {k: [v["recovered"], v["total"], v["type_ok"]]
                        for k, v in src["by_evo"].items()}
                detail = src.get("detail") or []
                print(f"[info] {label} 逐库文件不全，改读 {os.path.basename(ev)}"
                      f"（{sum(v[1] for v in stat.values())} op）")
            else:
                print(f"[warn] {label} 无可用产物，跳过")
                continue
        else:
            stat, detail = evaluate(rows, infer)
        rec = {}
        for evo in ALL_LEVELS:
            v = stat.get(evo)
            if v and v[1]:
                rec[evo] = {"recovered": v[0], "total": v[1], "type_ok": v[2]}
        table[label] = {"cat": cat, "suffix": suf, "missing": len(missing),
                        "by_evo": rec}
        type_ok_map[label] = collect_type_ok(detail)

    # ---- 分组汇总 ----
    def agg(label, levels):
        rec = table[label]["by_evo"]
        n = sum(rec[e]["total"] for e in levels if e in rec)
        r = sum(rec[e]["recovered"] for e in levels if e in rec)
        t = sum(rec[e]["type_ok"] for e in levels if e in rec)
        return {"n": n, "recovered": r, "type_ok": t,
                "recall_pct": round(r / n * 100, 1) if n else None,
                "type_pct": round(t / n * 100, 1) if n else None}

    groups = {"段位档": EMBED_LEVELS, "非段位档": [e for e in ALL_LEVELS
                                                   if e not in EMBED_LEVELS],
              "全部": ALL_LEVELS}
    summary = {}
    for label, _, _ in METHS:
        summary[label] = {g: agg(label, lv) for g, lv in groups.items()}

    # ---- op 级配对：M1 与三条基线，只在段位档 ----
    ref = "M1+形态指纹"
    paired = {}
    emb_ops = [(e, t, o) for (e, t, o) in type_ok_map[ref] if e in EMBED_LEVELS]
    for label, suf, cat in METHS:
        if cat != "baseline":
            continue
        a_only = b_only = 0
        for k in emb_ops:
            a = type_ok_map[ref].get(k)
            b = type_ok_map[label].get(k)
            if a is None or b is None:
                continue
            if a and not b:
                a_only += 1
            elif b and not a:
                b_only += 1
        p, n = sign_test(a_only, b_only)
        paired[f"{ref} vs {label}"] = {
            "scope": "段位档 op-type 一致",
            "ours_only": a_only, "base_only": b_only,
            "discordant": n, "p": round(p, 6),
            "n_ops": len(emb_ops),
        }

    out = {"version": "ablation-matching-v1",
           "manifest": MANIFEST,
           "note": "同一 evaluate() 口径；段位档 = merge_col + split_col（15 op），"
                   "非段位档 = 其余 147 op",
           "groups": summary, "by_evo": table, "paired": paired}
    json.dump(out, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)

    # ---- 打印 ----
    print("=" * 78)
    print("谱系抽取质量：同一 evaluate() 口径，op 级")
    print("=" * 78)
    for g in ("段位档", "非段位档", "全部"):
        print(f"\n[{g}]  n = {summary[METHS[0][0]][g]['n']}")
        print(f"  {'方法':<22}{'目标列恢复':>12}{'type 一致':>12}")
        for label, _, _ in METHS:
            s = summary[label][g]
            print(f"  {label:<22}{s['recovered']:>4}/{s['n']:<4}"
                  f"{s['recall_pct']:>5.1f}%{s['type_ok']:>7}/{s['n']:<4}"
                  f"{s['type_pct']:>5.1f}%")

    print("\n" + "=" * 78)
    print("op 级配对（段位档，type 一致）")
    print("=" * 78)
    for k, v in paired.items():
        print(f"  {k:<34}{v['ours_only']:>3} : {v['base_only']:<3}"
              f"  p = {v['p']}")
    print(f"\nSAVED: {OUT}")


if __name__ == "__main__":
    main()
