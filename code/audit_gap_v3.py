# -*- coding: utf-8 -*-
"""audit_gap_v3.py - 主实验与扩充集落差的分解（回应 MO-4）。

审稿意见：主实验合并档 C3mp 79.6%，扩充集合并档 44.7%（现为 44.9%），
不解释就无法判断扩充集是在复现主实验还是在测一个更难的分布。

本脚本先剥掉一个混淆，再做分层：
  ① 模型混淆：主表 79.6% 来自 DeepSeek-V3.2，扩充集 44.9% 来自 V4.1-Flash。
     同一 230 题上 V4.1-Flash 的读数为 87.8%，故 87.8% → 44.9% 才是同模型落差，
     模型版本不是落差来源（V4.1F 在旧题上反而高 8.2pt）。
  ② 分层：按库、表宽、引用列频次、目标段位、SQL 结构分层报成功率。
  ③ 直接标准化：把扩充集按主实验的分层权重重新加权，得到「若难度分布相同」
     的预期成功率，差额即构成效应。

只读已有逐条判定与 BIRD dev 题面，不新增任何模型调用。
输出 results/audit_gap_v3.json
"""
import collections
import json
import os
import re
import sqlite3

import sqlglot
from sqlglot import exp

BASE = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE)

OUT = "results/audit_gap_v3.json"
VERSION = "audit-gap-v3"

MAIN_RUN = "results/e1_xm_ds4f_ms_merged.jsonl"   # V4.1-Flash, 同 230 题
EXT_RUN = "results/ext_eval_merge_col.jsonl"       # V4.1-Flash, 扩充 479 题


# ------------------------------------------------------------------ 频次
def dev_freq():
    d = json.load(open("data/bird/dev_20240627/dev.json", encoding="utf-8"))
    per = collections.defaultdict(collections.Counter)
    for r in d:
        try:
            tree = sqlglot.parse_one(r["SQL"], read="sqlite")
        except Exception:
            continue
        for c in tree.find_all(exp.Column):
            if c.name:
                per[r["db_id"]][c.name] += 1
    return per


def table_width(path, table):
    if not os.path.exists(path):
        return None
    try:
        con = sqlite3.connect(f"file:{path}?mode=ro", uri=True)
        n = len(con.execute(f'PRAGMA table_info("{table}")').fetchall())
        con.close()
        return n
    except Exception:
        return None


def sql_feats(sql):
    try:
        t = sqlglot.parse_one(sql, read="sqlite")
    except Exception:
        return {"n_joins": None, "has_group": None, "has_subq": None,
                "has_agg": None, "has_order": None, "sql_len": len(sql)}
    joins = len(list(t.find_all(exp.Join)))
    return {
        "n_joins": joins,
        "has_group": any(True for _ in t.find_all(exp.Group)),
        "has_subq": any(True for _ in t.find_all(exp.Subquery)),
        "has_agg": any(True for _ in t.find_all(exp.AggFunc)),
        "has_order": any(True for _ in t.find_all(exp.Order)),
        "sql_len": len(sql),
    }


def seg_idx(expr_sql):
    """从 gold 表达式判断目标段位：0 = 前缀段，1 = 其后段。"""
    # 必须锁定 SUBSTR 的第 2 个参数：只写 "1,INSTR" 会误匹配
    # "attr_001,INSTR(...)"（001 里的 1），把所有段位都判成 0。
    if re.search(r",\s*1\s*,\s*INSTR\s*\(", expr_sql or ""):
        return 0
    if "INSTR" in (expr_sql or "") and "+1" in (expr_sql or ""):
        return 1
    return None


# ------------------------------------------------------------------ 载入
def load_rows(path):
    rows = [json.loads(l) for l in open(path, encoding="utf-8") if l.strip()]
    return rows


def build_items():
    items = []
    # ---- 主实验 230
    man = {}
    for line in open("data/manifest_v3.jsonl", encoding="utf-8"):
        r = json.loads(line)
        if r["evo_type"] == "merge_col":
            man[r["id"]] = r
    res = collections.defaultdict(dict)
    for r in load_rows(MAIN_RUN):
        res[r["id"]][r["cond"]] = r["outcome"]
    for iid, m in man.items():
        ops = m["ops"]
        if isinstance(ops, str):
            ops = json.loads(ops)
        if not ops:
            continue
        v = res.get(iid, {})
        if "c3mp" not in v:
            continue
        idx = seg_idx(m["new_sql"])
        col = ops[min(idx or 0, len(ops) - 1)][2]
        fo = sql_feats(m["old_sql"])
        items.append({
            "set": "main", "id": iid, "db": m["db_id"], "table": ops[0][1],
            "col": col, "seg": idx, "n_ops": len(ops),
            "c3mp": v.get("c3mp"), "c3m": v.get("c3m"),
            "db_path": f"data/evolved_merge/{m['db_id']}.sqlite", **fo})

    # ---- 扩充 479
    ext = collections.defaultdict(dict)
    for r in load_rows(EXT_RUN):
        ext[r["id"]][r["cond"]] = r["outcome"]
    cols_by_tag = {e["tag"]: e["cols"] for e in
                   json.load(open("results/evolution_ext_merge_col.json",
                                  encoding="utf-8"))["per_event"]}
    meta = {}
    for r in load_rows(EXT_RUN):
        meta[r["id"]] = r
    for line in open("data/manifest_ext_merge_col.jsonl", encoding="utf-8"):
        m = json.loads(line)
        v = ext.get(m["id"], {})
        if "c3mp" not in v:
            continue
        idx = seg_idx(m["new_sql"])
        cols = cols_by_tag.get(meta[m["id"]]["tag"]) or [None, None]
        col = cols[min(idx or 0, len(cols) - 1)]
        fo = sql_feats(m["old_sql"])
        items.append({
            "set": "ext", "id": m["id"], "db": m["db_id"], "table": m["table"],
            "col": col, "seg": idx, "n_ops": 2,
            "c3mp": v.get("c3mp"), "c3m": v.get("c3m"),
            "db_path": m["ext_db_path"].replace("\\", "/"), **fo})
    return items


# ------------------------------------------------------------------ 分层
def rate(xs):
    n = len(xs)
    k = sum(1 for x in xs if x["c3mp"] == "EX")
    return k, n, round(100 * k / n, 1) if n else None


def by(items, key, order=None):
    g = collections.defaultdict(list)
    for x in items:
        g[key(x)].append(x)
    out = []
    keys = order if order is not None else sorted(g, key=lambda z: (z is None, z))
    for k in keys:
        if k not in g:
            continue
        kk, nn, rr = rate(g[k])
        out.append({"stratum": k, "n": nn, "ex": kk, "rate": rr})
    return out


def main():
    freq = dev_freq()
    items = build_items()
    for x in items:
        x["ref"] = freq[x["db"]].get(x["col"], 0)
        x["width"] = table_width(x["db_path"], x["table"])
        x["ref_bin"] = ("0" if x["ref"] == 0 else
                        "1-4" if x["ref"] < 5 else
                        "5-19" if x["ref"] < 20 else "20+")
        x["width_bin"] = ("<10" if (x["width"] or 0) < 10 else
                          "10-19" if (x["width"] or 0) < 20 else "20+")
        x["join_bin"] = "0" if x["n_joins"] == 0 else ("1" if x["n_joins"] == 1 else "2+")

    M = [x for x in items if x["set"] == "main"]
    E = [x for x in items if x["set"] == "ext"]
    out = {"version": VERSION,
           "note": "两集均为 DeepSeek-V4.1-Flash 逐条判定；主实验另一版为 V3.2（79.6%），"
                   "同模型对比如本文件所示。",
           "overall": {"main": dict(zip(("ex", "n", "rate"), rate(M))),
                       "ext": dict(zip(("ex", "n", "rate"), rate(E)))},
           "by_db": {"main": by(M, lambda x: x["db"]),
                     "ext": by(E, lambda x: x["db"])},
           "by_ref": {"main": by(M, lambda x: x["ref_bin"],
                                 ["0", "1-4", "5-19", "20+"]),
                      "ext": by(E, lambda x: x["ref_bin"],
                                ["0", "1-4", "5-19", "20+"])},
           "by_width": {"main": by(M, lambda x: x["width_bin"],
                                   ["<10", "10-19", "20+"]),
                        "ext": by(E, lambda x: x["width_bin"],
                                  ["<10", "10-19", "20+"])},
           "by_join": {"main": by(M, lambda x: x["join_bin"], ["0", "1", "2+"]),
                       "ext": by(E, lambda x: x["join_bin"], ["0", "1", "2+"])},
           "by_seg": {"main": by(M, lambda x: x["seg"], [0, 1]),
                      "ext": by(E, lambda x: x["seg"], [0, 1])},
           "by_group": {"main": by(M, lambda x: bool(x["has_group"]), [False, True]),
                        "ext": by(E, lambda x: bool(x["has_group"]), [False, True])},
           "by_subq": {"main": by(M, lambda x: bool(x["has_subq"]), [False, True]),
                       "ext": by(E, lambda x: bool(x["has_subq"]), [False, True])},
           }

    # 直接标准化：按 (库, 引用频次分段) 用主实验的权重给扩充集重新加权
    def std(target, ref, keys):
        num = den = 0.0
        wsum = sum(1 for _ in ref)
        tg = collections.defaultdict(list)
        for x in target:
            tg[keys(x)].append(x)
        rg = collections.defaultdict(list)
        for x in ref:
            rg[keys(x)].append(x)
        covered = 0
        for k, xs in tg.items():
            if k not in rg:
                continue            # 该层主实验没有样本 -> 权重为 0，不参与
            w = len(rg[k]) / wsum
            r = sum(1 for y in xs if y["c3mp"] == "EX") / len(xs)
            num += w * r
            den += w
            covered += len(xs)
        return {"standardized_rate": round(100 * num / den, 1) if den else None,
                "weight_mass": round(den, 3),
                "strata_used": sum(1 for k in tg if k in rg),
                "strata_total": len(tg),
                "ext_items_in_overlapping_strata": covered}

    out["standardization"] = {
        "by_db": std(E, M, lambda x: x["db"]),
        "by_db_ref": std(E, M, lambda x: (x["db"], x["ref_bin"])),
        "by_db_ref_width": std(E, M, lambda x: (x["db"], x["ref_bin"], x["width_bin"])),
    }
    # 反方向：主实验按扩充集权重（看主实验是否只是「简单题」）
    out["standardization"]["main_to_ext_by_ref"] = std(M, E, lambda x: x["ref_bin"])

    json.dump(out, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)

    for side in ("main", "ext"):
        o = out["overall"][side]
        print(f"总体 {side}: {o['ex']}/{o['n']} = {o['rate']}%")
    for k in ("by_ref", "by_width", "by_join", "by_seg", "by_group", "by_subq"):
        print("\n" + k)
        for side in ("main", "ext"):
            print(" ", side, [(r["stratum"], r["n"], r["rate"])
                              for r in out[k][side]])
    print("\n库")
    for side in ("main", "ext"):
        print(" ", side, [(r["stratum"], r["n"], r["rate"])
                          for r in out["by_db"][side]])
    print("\n标准化", json.dumps(out["standardization"], ensure_ascii=False))
    print("写入", OUT)


if __name__ == "__main__":
    main()
