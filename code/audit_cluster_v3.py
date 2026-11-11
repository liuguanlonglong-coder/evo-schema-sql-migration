# -*- coding: utf-8 -*-
"""audit_cluster_v3.py - 数据库聚类稳健性（回应 MO-1 / MO-3）。

审稿意见的核心：条目之间共享基础题、库与演化事件，不能按记录数当独立样本；
扩充集 p=0.035（现为 14:4，n=19）处在边界，且同一库内事件相关。

本脚本只读已有逐条判定，不新增调用，输出四件事：
  A 条目级 2x2 + 精确 McNemar（仅供对照，不作主结论）
  B 事件级符号检验：主检验限定在「合格事件」（≥3 条题通过四步验证），
    「进入检验事件」版本作为敏感性并列出差异
  C 数据库级：逐库事件数与方向、Δpt、库级符号检验；leave-one-database-out
  D 按库聚类 bootstrap（重采样库，不是重采样条目）给出事件级平均 Δ 的区间

输出 results/audit_cluster_v3.json
"""
import collections
import json
import math
import os
import random

BASE = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE)

OUT = "results/audit_cluster_v3.json"
VERSION = "audit-cluster-v3"
SEED = 20261111
NBOOT = 10000


def comb(n, k):
    return math.comb(n, k)


def sign_p(pos, neg):
    n = pos + neg
    if n == 0:
        return 1.0
    k = min(pos, neg)
    tail = sum(comb(n, i) for i in range(k + 1)) / 2 ** n
    return min(1.0, 2 * tail)


def mcnemar(b, c):
    n = b + c
    if n == 0:
        return 1.0
    k = min(b, c)
    return min(1.0, 2 * sum(comb(n, i) for i in range(k + 1)) / 2 ** n)


def wilson(k, n, z=1.96):
    if n == 0:
        return [0.0, 0.0]
    p = k / n
    d = 1 + z * z / n
    c = p + z * z / (2 * n)
    h = z * math.sqrt(p * (1 - p) / n + z * z / (4 * n * n))
    return [round(100 * (c - h) / d, 1), round(100 * (c + h) / d, 1)]


# ------------------------------------------------------------------ 载入
def load_ext():
    rows = [json.loads(l) for l in
            open("results/ext_eval_merge_col.jsonl", encoding="utf-8") if l.strip()]
    per = json.load(open("results/evolution_ext_merge_col.json",
                        encoding="utf-8"))["per_event"]
    info = {e["tag"]: e for e in per}
    items = collections.defaultdict(dict)
    for r in rows:
        items[r["id"]][r["cond"]] = r["outcome"]
    meta = {}
    for r in rows:
        meta[r["id"]] = {"db": r["db_id"], "tag": r["tag"], "table": r["table"]}
    return items, meta, info


def load_main(event_map_ids):
    """主实验合并档：230 条，逐条判定来自 e1_ms_ds_p.jsonl（V3.2）。"""
    rows = []
    for f in ("results/e1_ms_ds.jsonl", "results/e1_ms_ds_p.jsonl"):
        rows += [json.loads(l) for l in open(f, encoding="utf-8") if l.strip()]
    items = collections.defaultdict(dict)
    for r in rows:
        if r["evo_type"] != "merge_col":
            continue
        # c3mp 只在 _p 文件里；同一 (id,cond) 出现两次时后写覆盖，值必然相同
        items[r["id"]][r["cond"]] = r["outcome"]
    meta = {}
    for r in rows:
        if r["evo_type"] != "merge_col":
            continue
        meta[r["id"]] = {"db": r["db_id"],
                         "tag": event_map_ids.get(r["id"], r["db_id"]),
                         "table": None}
    return items, meta


def main_event_map():
    """从 manifest_v3 的 ops 反推每条题属于哪个合并事件（库|表|新列）。"""
    out = {}
    for line in open("data/manifest_v3.jsonl", encoding="utf-8"):
        r = json.loads(line)
        if r["evo_type"] != "merge_col":
            continue
        ops = r.get("ops")
        if isinstance(ops, str):
            ops = json.loads(ops)
        if not ops:
            continue
        tbl, expr = ops[0][1], ops[0][3]
        col = "attr_001"
        out[r["id"]] = f"{r['db_id']}|{str(tbl).lower()}|{col}"
    return out


# ------------------------------------------------------------------ 统计
def analyse(items, meta, eligible_tags=None, label=""):
    conds = sorted({c for v in items.values() for c in v})
    res = {"label": label, "n_items": len(items), "conditions": {}}
    # A 条目级
    for c in conds:
        k = sum(1 for v in items.values() if v.get(c) == "EX")
        n = sum(1 for v in items.values() if c in v)
        res["conditions"][c] = {"n": n, "EX": k,
                                "rate": round(100 * k / max(1, n), 1),
                                "wilson95": wilson(k, n)}
    if "c3mp" in conds and "c3m" in conds:
        b = c = d = 0
        for v in items.values():
            if "c3mp" not in v or "c3m" not in v:
                continue
            a1, a2 = v["c3mp"] == "EX", v["c3m"] == "EX"
            if a1 and not a2:
                b += 1
            elif a2 and not a1:
                c += 1
        n = sum(1 for v in items.values() if "c3mp" in v and "c3m" in v)
        res["item_level"] = {"n_pairs": n, "c3mp_only": b, "c3m_only": c,
                             "delta_pt": round(100 * (b - c) / max(1, n), 1),
                             "mcnemar_p": round(mcnemar(b, c), 6)}
    # B 事件级
    ev = collections.defaultdict(lambda: {"n": 0, "mp": 0, "cm": 0, "db": None})
    for iid, v in items.items():
        if "c3mp" not in v or "c3m" not in v:
            continue
        m = meta[iid]
        k = m["tag"]
        ev[k]["db"] = m["db"]
        ev[k]["n"] += 1
        ev[k]["mp"] += v["c3mp"] == "EX"
        ev[k]["cm"] += v["c3m"] == "EX"

    def ev_block(keys):
        rows = []
        pos = neg = tie = 0
        for k in keys:
            e = ev[k]
            if e["n"] == 0:
                continue
            d = e["mp"] - e["cm"]
            sign = 1 if d > 0 else (-1 if d < 0 else 0)
            pos += sign > 0
            neg += sign < 0
            tie += sign == 0
            rows.append({"event": k, "db": e["db"], "n": e["n"],
                         "c3m_ex": e["cm"], "c3mp_ex": e["mp"],
                         "delta_pt": round(100 * d / e["n"], 1),
                         "sign": sign})
        deltas = [r["delta_pt"] for r in rows]
        return {"n_events": len(rows), "pos": pos, "neg": neg, "tie": tie,
                "p": round(sign_p(pos, neg), 6),
                "mean_delta_pt": round(sum(deltas) / max(1, len(deltas)), 1),
                "min_delta_pt": min(deltas) if deltas else None,
                "max_delta_pt": max(deltas) if deltas else None,
                "events": sorted(rows, key=lambda x: (-x["delta_pt"], x["event"]))}

    all_tags = sorted(ev)
    res["event_level_all"] = ev_block(all_tags)
    if eligible_tags is not None:
        res["event_level_eligible"] = ev_block(
            [t for t in all_tags if t in eligible_tags])
        res["event_level_entered"] = ev_block(
            [t for t in all_tags if ev[t]["mp"] != ev[t]["cm"]])

    # C 数据库级
    bydb = collections.defaultdict(list)
    for k in all_tags:
        bydb[ev[k]["db"]].append(k)
    dbt = []
    for db, ks in sorted(bydb.items()):
        blk = ev_block(ks)
        n_items = sum(ev[k]["n"] for k in ks)
        mp = sum(ev[k]["mp"] for k in ks)
        cm = sum(ev[k]["cm"] for k in ks)
        dbt.append({"db": db, "n_items": n_items, "n_events": blk["n_events"],
                    "c3m_rate": round(100 * cm / max(1, n_items), 1),
                    "c3mp_rate": round(100 * mp / max(1, n_items), 1),
                    "delta_pt": round(100 * (mp - cm) / max(1, n_items), 1),
                    "pos": blk["pos"], "neg": blk["neg"], "tie": blk["tie"],
                    "event_mean_delta_pt": blk["mean_delta_pt"]})
    res["by_db"] = dbt
    dp = sum(1 for d in dbt if d["delta_pt"] > 0)
    dn = sum(1 for d in dbt if d["delta_pt"] < 0)
    res["db_level"] = {"n_dbs": len(dbt), "pos": dp, "neg": dn,
                       "p": round(sign_p(dp, dn), 6)}

    # leave-one-database-out
    loo = []
    for db in sorted(bydb):
        keys = [k for k in all_tags if ev[k]["db"] != db]
        blk = ev_block(keys)
        loo.append({"dropped": db, "n_events": blk["n_events"],
                    "pos": blk["pos"], "neg": blk["neg"],
                    "p": blk["p"], "mean_delta_pt": blk["mean_delta_pt"]})
    res["leave_one_db_out"] = loo

    # D 按库聚类 bootstrap（事件级平均 Δ 与条目级 Δ 各一套）
    rng = random.Random(SEED)
    dbs = sorted(bydb)
    boot_ev, boot_item = [], []
    for _ in range(NBOOT):
        pick = [rng.choice(dbs) for _ in dbs]
        keys = [k for db in pick for k in bydb[db]]
        if not keys:
            continue
        # 事件级平均 Δ（事件等权）
        ds = [100 * (ev[k]["mp"] - ev[k]["cm"]) / ev[k]["n"]
              for k in keys if ev[k]["n"]]
        if ds:
            boot_ev.append(sum(ds) / len(ds))
        mp = sum(ev[k]["mp"] for k in keys)
        cm = sum(ev[k]["cm"] for k in keys)
        nn = sum(ev[k]["n"] for k in keys)
        if nn:
            boot_item.append(100 * (mp - cm) / nn)
    boot_ev.sort()
    boot_item.sort()

    def pct(a, q):
        return round(a[int(q * (len(a) - 1))], 1) if a else None

    res["cluster_bootstrap"] = {
        "resamples": len(boot_ev), "seed": SEED, "unit": "database",
        "event_mean_delta_pt": {"point": res["event_level_all"]["mean_delta_pt"],
                                "ci95": [pct(boot_ev, 0.025), pct(boot_ev, 0.975)]},
        "item_pooled_delta_pt": {"point": res["item_level"]["delta_pt"]
                                 if "item_level" in res else None,
                                 "ci95": [pct(boot_item, 0.025),
                                          pct(boot_item, 0.975)]},
        "note": "重采样的是库（簇），同一库内全部事件一起进出，"
                "因此区间反映库间变异；6 个库时该区间偏宽属预期。"}
    return res


def main():
    items, meta, info = load_ext()
    eligible = {t for t, e in info.items() if (e.get("n_pass") or 0) >= 3}
    out = {"version": VERSION,
           "note": "扩充合并档 c3mp vs c3m；事件键 = 构造分片 tag（库__表__档_pN）。"
                   "主检验用合格事件（≥3 题通过四步验证）；进入检验事件为敏感性。",
           "ext_merge": analyse(items, meta, eligible, "ext_merge")}

    emap = main_event_map()
    mitems, mmeta = load_main(emap)
    out["main_merge"] = analyse(mitems, mmeta, None, "main_merge_v32")

    json.dump(out, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)

    for key in ("ext_merge", "main_merge"):
        r = out[key]
        print("=" * 60)
        print(key, r["n_items"], "条")
        for c, v in r["conditions"].items():
            print(f"  {c:5s} {v['EX']:4d}/{v['n']:<4d} {v['rate']:5.1f}%  "
                  f"CI {v['wilson95']}")
        if "item_level" in r:
            il = r["item_level"]
            print(f"  条目级 {il['c3mp_only']}:{il['c3m_only']} Δ{il['delta_pt']}pt "
                  f"p={il['mcnemar_p']}")
        for k in ("event_level_all", "event_level_eligible",
                  "event_level_entered"):
            if k in r:
                e = r[k]
                print(f"  {k:22s} n={e['n_events']:2d} {e['pos']}正/{e['neg']}负"
                      f"/{e['tie']}平 p={e['p']} 平均Δ{e['mean_delta_pt']}pt")
        print("  库级:", r["db_level"])
        for d in r["by_db"]:
            print(f"    {d['db']:24s} n={d['n_items']:3d} ev={d['n_events']:2d} "
                  f"{d['c3m_rate']:5.1f}->{d['c3mp_rate']:5.1f} "
                  f"Δ{d['delta_pt']:6.1f} ({d['pos']}正{d['neg']}负)")
        print("  LOO:", [(x["dropped"], x["pos"], x["neg"], x["p"])
                         for x in r["leave_one_db_out"]])
        print("  bootstrap:", r["cluster_bootstrap"]["event_mean_delta_pt"],
              r["cluster_bootstrap"]["item_pooled_delta_pt"])
    print("写入", OUT)


if __name__ == "__main__":
    main()
