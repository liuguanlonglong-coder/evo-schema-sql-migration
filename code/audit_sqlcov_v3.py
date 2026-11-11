# -*- coding: utf-8 -*-
"""audit_sqlcov_v3.py - SQL 结构覆盖的实测统计（回应 MO-7）。

审稿意见：评测集中在列引用与基础 WHERE 谓词，JOIN / CTE / 子查询 /
GROUP BY / HAVING / UNION / 窗口 等覆盖不足，且未说明系统是否支持。

本脚本不做新的迁移实验，只把「现有 610 条到底覆盖了哪些结构」逐条数清，
使限制从形容词变成可核对的表。判据全部来自 sqlglot 解析（sqlite 方言），
解析失败单独计数，不计入任何一类。

输出 results/sql_coverage_v3.json
"""
import collections
import json
import os

import sqlglot
from sqlglot import exp

BASE = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE)

OUT = "results/sql_coverage_v3.json"
VERSION = "sql-coverage-v3"

FEATS = [
    ("join_any", lambda t: any(True for _ in t.find_all(exp.Join))),
    ("join_inner", lambda t: any(j.args.get("kind", "").upper() == "INNER"
                                 or (not j.args.get("kind") and not j.args.get("side"))
                                 for j in t.find_all(exp.Join))),
    ("join_left", lambda t: any((j.args.get("side") or "").upper() == "LEFT"
                                for j in t.find_all(exp.Join))),
    ("join_using_or_natural", lambda t: any(
        j.args.get("using") or j.args.get("method") == "NATURAL"
        for j in t.find_all(exp.Join))),
    ("cte", lambda t: any(True for _ in t.find_all(exp.CTE))),
    ("subquery_any", lambda t: any(True for _ in t.find_all(exp.Subquery))),
    ("subquery_correlated", lambda t: any(
        q.args.get("correlated") for q in t.find_all(exp.Subquery))),
    ("group_by", lambda t: any(True for _ in t.find_all(exp.Group))),
    ("having", lambda t: any(True for _ in t.find_all(exp.Having))),
    ("order_by", lambda t: any(True for _ in t.find_all(exp.Order))),
    ("limit", lambda t: t.args.get("limit") is not None),
    ("distinct", lambda t: any(True for _ in t.find_all(exp.Distinct))),
    ("aggregate", lambda t: any(True for _ in t.find_all(exp.AggFunc))),
    ("union_like", lambda t: any(True for _ in
                                 t.find_all(exp.Union, exp.Intersect, exp.Except))),
    ("window", lambda t: any(True for _ in t.find_all(exp.Window))),
    ("case_expr", lambda t: any(True for _ in t.find_all(exp.Case))),
    # 裸 SELECT * 与 COUNT(*) 必须分开：后者不是星号投影，构造时也不会被丢弃
    ("star_projection", lambda t: any(
        isinstance(p_, exp.Star) for sel in t.find_all(exp.Select)
        for p_ in sel.expressions)),
    ("qualified_star", lambda t: any(True for _ in t.find_all(exp.Column)
                                     if isinstance(c.this, exp.Star))),
    ("cast", lambda t: any(True for _ in t.find_all(exp.Cast))),
    ("like", lambda t: any(True for _ in t.find_all(exp.Like))),
    ("in_list", lambda t: any(True for _ in t.find_all(exp.In))),
    ("between", lambda t: any(True for _ in t.find_all(exp.Between))),
    ("is_null", lambda t: any(True for _ in t.find_all(exp.Is))),
    ("date_fn", lambda t: any(True for _ in t.find_all(exp.Anonymous)
                              if str(_.this).lower() in
                              ("strftime", "date", "datetime", "julianday",
                               "substr", "substring"))),
    ("multi_table", lambda t: len(list(t.find_all(exp.Table))) > 1),
]


def feats(sql):
    try:
        t = sqlglot.parse_one(sql, read="sqlite")
    except Exception:
        return None
    if t is None:
        return None
    d = {}
    for name, fn in FEATS:
        try:
            d[name] = bool(fn(t))
        except Exception:
            d[name] = None
    d["n_joins"] = len(list(t.find_all(exp.Join)))
    d["n_tables"] = len(list(t.find_all(exp.Table)))
    return d


def target_position(rows):
    """目标列出现在哪个子句——比「SQL 里有没有 JOIN」更贴近迁移难度。

    一次改写发生在 SELECT 列表与发生在 ON 条件里，对模型是不同的任务；
    HAVING / 子查询位置若为 0，就必须在限制里写明「未覆盖」。
    """
    import re
    cnt = collections.Counter()
    per_item = {}
    for m in rows:
        ops = m["ops"]
        if isinstance(ops, str):
            ops = json.loads(ops)
        if m["evo_type"] in ("merge_col", "split_col"):
            if m["evo_type"] == "merge_col":
                idx = 0 if re.search(r",\s*1\s*,\s*INSTR\s*\(", m["new_sql"]) else 1
                cols = {ops[min(idx, len(ops) - 1)][2].lower()}
            else:
                cols = {ops[0][2].lower()}
        else:
            cols = {o[2].lower() for o in ops}
        try:
            t = sqlglot.parse_one(m["old_sql"], read="sqlite")
        except Exception:
            t = None
        if t is None:
            continue
        got = set()
        for c in t.find_all(exp.Column):
            if c.name.lower() not in cols:
                continue
            anc, in_sub = c.parent, False
            while anc is not None:
                if isinstance(anc, exp.Subquery):
                    in_sub = True
                if isinstance(anc, (exp.Where, exp.Join, exp.Group,
                                    exp.Having, exp.Order, exp.Select)):
                    break
                anc = anc.parent
            kind = ("subquery" if in_sub else
                    "SELECT" if isinstance(anc, exp.Select) else
                    type(anc).__name__ if anc is not None else "?")
            got.add(kind)
        for g in got:
            cnt[g] += 1
        per_item[m["id"]] = sorted(got)
    n = len(rows)
    return {"n": n,
            "clauses": {k: {"k": v, "pct": round(100 * v / max(1, n), 1)}
                        for k, v in cnt.most_common()},
            "items_without_position": sum(1 for k, v in per_item.items() if not v)}


def tally(rows, label):
    c = collections.Counter()
    n = 0
    parse_fail = 0
    njoins = collections.Counter()
    for r in rows:
        f = feats(r["old_sql"])
        if f is None:
            parse_fail += 1
            continue
        n += 1
        for k, v in f.items():
            if k.startswith("n_"):
                continue
            if v:
                c[k] += 1
        njoins[f["n_joins"]] += 1
    out = {"label": label, "n": len(rows), "n_parsed": n,
           "parse_fail": parse_fail,
           "features": {k: {"k": c[k], "pct": round(100 * c[k] / max(1, n), 1)}
                        for k, _ in FEATS},
           "joins_hist": dict(sorted(njoins.items()))}
    return out


def main():
    man = [json.loads(l) for l in open("data/manifest_v3.jsonl", encoding="utf-8")]
    paper_ids = set()
    for f in ("data/ids_xm_full.txt", "data/ids_ms.txt"):
        paper_ids |= {l.strip() for l in open(f, encoding="utf-8") if l.strip()}
    paper = [r for r in man if r["id"] in paper_ids]
    by_evo = collections.defaultdict(list)
    for r in paper:
        by_evo[r["evo_type"]].append(r)

    ext = [json.loads(l) for l in open("data/manifest_ext_merge_col.jsonl",
                                       encoding="utf-8")]

    out = {"version": VERSION,
           "note": "逐条 sqlglot 解析 old_sql（sqlite 方言）；解析失败单列，"
                   "不计入任何一类，避免把「解析不了」算成「不支持」。",
           "paper": tally(paper, "610 主实验"),
           "by_evo": {k: tally(v, k) for k, v in sorted(by_evo.items())},
           "ext_merge": tally(ext, "扩充合并 479"),
           "target_position": target_position(paper)}
    json.dump(out, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)

    for key in ("paper", "ext_merge"):
        t = out[key]
        print(f"\n{key}: n={t['n']} parsed={t['n_parsed']} fail={t['parse_fail']}")
        for k, v in t["features"].items():
            print(f"   {k:22s} {v['k']:4d}  {v['pct']:5.1f}%")
        print("   joins:", t["joins_hist"])
    print("\n按档：")
    for k, t in out["by_evo"].items():
        f = t["features"]
        print(f"   {k:12s} n={t['n']:3d} join={f['join_any']['pct']:5.1f}% "
              f"group={f['group_by']['pct']:5.1f}% subq={f['subquery_any']['pct']:5.1f}% "
              f"cte={f['cte']['pct']:4.1f}% union={f['union_like']['pct']:4.1f}% "
              f"star={f['star_projection']['pct']:4.1f}%")
    print("\n目标列所在子句：", out["target_position"])
    print("写入", OUT)


if __name__ == "__main__":
    main()
