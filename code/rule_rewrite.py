# -*- coding: utf-8 -*-
"""确定性 SQL 重写基线 v2（回答第七轮第四节 + 第八轮第 1 条）。

v1 的两个缺陷（第八轮评审指出其一，实查后发现两个都在）
------------------------------------------------------
1. **投影星号未支持**：v1 对任何 `*` 直接判 STAR_PROJ 失败。评审指出可按
   旧表列清单展开再逐列替换。
2. **`COUNT(*)` 被误判为投影星号**（v1 用 `find_all(exp.Star)` 通吃）。
   实查 v1 判失败的 21 条**全部**是 `COUNT(*)`，没有一条是投影星号。
   `COUNT(*)` 在列替换下应保持不变，这一误判系统性压低了规则基线，
   并直接导致 v8 正文「LLM 只在星号投影上补上规则结果」这一互补结论
   建立在错误之上。

v2 的修正
---------
- 投影星号检测只看 `Select.expressions` 位上的 `*` / `t.*`，
  `COUNT(*)` 内的 `*` 不再触发。
- 新增星号展开：按**原始库**（演化前）的列清单、按 cid 序号展开，
  逐列交给统一替换通道。两个模式：
    --star expand  展开旧 `*`（评审建议的做法，语义忠实）
    --star keep    保留 `*` 原样（与构造生成的 gold new_sql 口径一致）
  当前 559 条样本中投影星号出现 0 次，故两模式数值完全相同；
  保留双模式是为了把「口径选择」显式化，而不是藏在实现里。

构造层的固有事实（必须写进正文，不能当成方法缺陷）
------------------------------------------------
本构造的 gold `new_sql` 由「机械替换列引用」生成，而 `*` 不是列引用，
故 gold 会原样保留 `*`，在演化后库上返回**新的列集合**——这与旧 SQL
的语义（返回旧列集合）并不等价。也就是说，含投影星号的题目在本构造下
**无法被忠实评价**。样本中恰好没有这类题目，故不影响已报告的数字；
此事实在 7.4.4 与 7.10 中明确列出。

失败分类
--------
  PARSE                  旧 SQL 无法被 sqlglot 解析
  STAR_NO_SCHEMA         需展开星号但取不到旧表列清单
  AMBIG_COL              列名多表同名且未限定，无法判定归属
  DB_ERROR               重写后执行报错
  SILENT                 执行成功但结果与 gold 不一致
  EX                     一致

输出 results/rule_rewrite[_<star>].json + results/rule_rewrite_items[_<star>].jsonl
"""
import argparse
import collections
import json
import os
import re
import sqlite3
import time

import sqlglot
from sqlglot import exp

BASE = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE)

VERSION = "rule-rewrite-v2"
TIMEOUT = 6.0

DIRS = {"merge_col": "data/evolved_merge",
        "split_col": "data/evolved_split",
        "merge_col_holdout": "data/evolved_merge_holdout",
        "split_col_holdout": "data/evolved_split_holdout"}
SRC = {"merge_col": "data/evolution_merge.jsonl",
       "split_col": "data/evolution_split.jsonl",
       "merge_col_holdout": "data/manifest_holdout_sample.jsonl",
       "split_col_holdout": "data/manifest_holdout_sample.jsonl"}
# 演化前的原始库（星号展开需要旧表列清单）
ORIG = {"merge_col": "data/bird/dev_20240627/dev_databases/%s/%s.sqlite",
        "split_col": "data/bird/dev_20240627/dev_databases/%s/%s.sqlite",
        "merge_col_holdout": "data/bird_train/train/train_databases/%s/%s.sqlite",
        "split_col_holdout": "data/bird_train/train/train_databases/%s/%s.sqlite"}


# --------------------------------------------------------------- 库与执行

def q(path, sql, limit=TIMEOUT):
    if not sql:
        return None, "EMPTY"
    if not os.path.exists(path):
        return None, "DB_MISSING"
    con = sqlite3.connect("file:%s?mode=ro" % path, uri=True)
    t0 = time.time()
    con.set_progress_handler(lambda: 1 if time.time() - t0 > limit else 0, 2000)
    try:
        return con.execute(sql).fetchall(), None
    except Exception as e:
        return None, str(e)[:200]
    finally:
        con.close()


def norm(rows):
    if rows is None:
        return None
    return sorted(tuple("" if v is None else str(v) for v in r) for r in rows)


def old_schema_of(path):
    """{table_lower: [col, ...]}，按 cid 顺序（与 SQLite `SELECT *` 展开顺序一致）。"""
    if not os.path.exists(path):
        return None
    con = sqlite3.connect("file:%s?mode=ro" % path, uri=True)
    out = {}
    try:
        tabs = [r[0] for r in con.execute(
            "select name from sqlite_master where type='table' "
            "and name not like 'sqlite_%'")]
        for t in tabs:
            try:
                cols = con.execute('pragma table_info("%s")' % t).fetchall()
            except Exception:
                continue
            out[t.lower()] = [(r[1], r[1]) for r in sorted(cols, key=lambda r: r[0])]
    finally:
        con.close()
    return out


# --------------------------------------------------------------- 谱系载入

def map_gt(item):
    """drift_ops（或留出清单的 ops 列表）-> {(table, old_col): new_expr}"""
    m = {}
    ops = item.get("drift_ops") or []
    if not ops:
        ops = [{"table": o[1], "old": o[2], "new": o[3]}
               for o in (item.get("ops") or [])
               if isinstance(o, list) and len(o) >= 4]
    for op in ops:
        t = (op.get("table") or "").lower()
        o = (op.get("old") or "").strip()
        n = (op.get("new") or "").strip()
        if o and n and t:
            m[(t, o.lower())] = n
    return m


def map_infer(evo, db, mode="part"):
    cands = [mode, "profile"] if mode == "part" else [mode]
    obj = None
    for md in cands:
        p = os.path.join("data", "lineage_infer", "%s_%s_%s.json" % (evo, db, md))
        if os.path.exists(p):
            obj = json.load(open(p, encoding="utf-8"))
            break
    if not obj:
        return None, None
    m = {}
    for op in obj.get("ops") or []:
        t = (op.get("table") or "").lower()
        o = (op.get("old") or "").strip()
        n = (op.get("new") or "").strip()
        if o and n and t:
            m[(t, o.lower())] = n
    return m, obj.get("mode") or cands[0]


def map_infer_name(evo, db):
    p = os.path.join("data", "lineage_infer", "%s_%s_profile.json" % (evo, db))
    if not os.path.exists(p):
        return None
    obj = json.load(open(p, encoding="utf-8"))
    m = {}
    for op in obj.get("ops") or []:
        t = (op.get("table") or "").lower()
        o = (op.get("old") or "").strip()
        n = (op.get("new") or "").strip()
        if o and n and t:
            m[(t, o.lower())] = n
    return m


# --------------------------------------------------------------- 重写

def _aliases(tree):
    amap = {}
    for t in tree.find_all(exp.Table):
        name = (t.name or "").lower()
        al = (t.alias or "").lower()
        if name:
            amap[name] = name
        if al:
            amap[al] = name
    return amap


def _own_tables(sel):
    """属于本层 SELECT 的 FROM/JOIN 表（不进入子查询）。"""
    out = []
    for t in sel.find_all(exp.Table):
        try:
            anc = t.find_ancestor(exp.Select)
        except Exception:
            anc = None
        if anc is sel:
            out.append(t)
    return out


def projection_stars(tree):
    """投影位上的星号：[(select, index, table_alias_or_None)]。`COUNT(*)` 不计入。"""
    out = []
    for sel in tree.find_all(exp.Select):
        for i, e in enumerate(list(sel.expressions)):
            if isinstance(e, exp.Star):
                out.append((sel, i, None))
            elif isinstance(e, exp.Column) and isinstance(e.this, exp.Star):
                out.append((sel, i, (e.table or "").lower() or None))
    return out


def expand_stars(tree, old_schema):
    """按旧表列清单展开投影星号；返回 (n_expanded, err)。"""
    stars = projection_stars(tree)
    if not stars:
        return 0, None
    if not old_schema:
        return 0, "旧表列清单不可得"
    n = 0
    for sel, idx, only in stars:
        tabs = _own_tables(sel)
        picked = []
        if only:
            for t in tabs:
                if ((t.alias or "").lower() or (t.name or "").lower()) == only:
                    picked.append(t)
        else:
            picked = tabs
        if not picked:
            return n, "星号限定的表不在本层 FROM/JOIN 中"
        new_exprs = []
        for t in picked:
            tname = (t.name or "").lower()
            alias = (t.alias or "").lower() or tname
            cols = old_schema.get(tname)
            if not cols:
                return n, "旧 schema 缺少表 %s" % tname
            for orig, _ in cols:
                new_exprs.append(exp.column(orig, table=alias))
        if not new_exprs:
            return n, "展开结果为空"
        # 用展开结果替换原星号投影（保持位置）
        head = list(sel.expressions)[:idx]
        tail = list(sel.expressions)[idx + 1:]
        sel.set("expressions", head + new_exprs + tail)
        n += 1
        return n, None   # 结构已变，后续另起一轮
    return n, None


def rewrite(sql, mapping, old_schema, star="expand"):
    """返回 (new_sql, status, detail)。status 为 None 表示正常重写。"""
    if not mapping:
        return None, "NO_MAP", "谱系为空"
    try:
        tree = sqlglot.parse_one(sql, read="sqlite")
    except Exception as e:
        return None, "PARSE", str(e)[:120]
    if tree is None:
        return None, "PARSE", "parse 返回空"

    # ---- 星号：先按需展开（可能多层，逐轮展开直到没有剩余）
    n_star = 0
    if star == "expand":
        for _ in range(8):
            k, err = expand_stars(tree, old_schema)
            if err:
                return None, "STAR_NO_SCHEMA", err
            if not k:
                break
            n_star += k

    amap = _aliases(tree)
    in_query = {v for v in amap.values() if v}
    n_hit, n_amb = 0, 0
    plan = []
    for col in list(tree.find_all(exp.Column)):
        if isinstance(col.this, exp.Star):
            continue          # 保留原样的 t.*（keep 模式）
        cn = (col.name or "").lower()
        qn = (col.table or "").lower()
        if not cn:
            continue
        if qn:
            tbl = amap.get(qn, qn)
            key = (tbl, cn)
            if key in mapping:
                plan.append((col, mapping[key]))
        else:
            cand = [(t, c) for (t, c) in mapping if c == cn
                    and (not in_query or t in in_query)]
            if len(cand) == 1:
                plan.append((col, mapping[cand[0]]))
            elif len(cand) > 1:
                n_amb += 1

    for col, expr_s in plan:
        try:
            new = sqlglot.parse_one("SELECT " + expr_s, read="sqlite").selects[0]
        except Exception:
            n_amb += 1
            continue
        col.replace(exp.paren(new))
        n_hit += 1

    if n_hit == 0:
        if n_amb:
            return None, "AMBIG_COL", "列名归属无法判定"
        return sql, "NO_HIT", "旧 SQL 未引用被演化列（恒等）"

    out = tree.sql(dialect="sqlite")
    return out, None, "hit=%d amb=%d star=%d" % (n_hit, n_amb, n_star)


# --------------------------------------------------------------- 主流程

def load_items(evo):
    rows = []
    for l in open(SRC[evo], encoding="utf-8"):
        r = json.loads(l)
        if evo.endswith("_holdout") and r.get("evo_type") != evo:
            continue
        rows.append(r)
    return rows


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--evo", default="merge_col,split_col,merge_col_holdout,split_col_holdout")
    ap.add_argument("--star", default="expand", choices=["expand", "keep"])
    args = ap.parse_args()

    evos = [x.strip() for x in args.evo.split(",") if x.strip()]
    out_rows = []
    summary = {}
    audit = {"version": VERSION, "star_mode": args.star,
             "projection_star_items": 0, "count_star_items": 0,
             "star_expanded": 0, "per_evo": {}}

    for evo in evos:
        d = DIRS[evo]
        items = load_items(evo)
        schema_cache, orig_cache = {}, {}
        conds = ["c3r_gt", "c3r_infer", "c3r_name"]
        stat = {c: collections.Counter() for c in conds}
        details = {c: [] for c in conds}
        n_proj_star = n_cnt_star = n_star_expanded = 0

        for it in items:
            db = it["db_id"]
            dbp = os.path.join(d, db + ".sqlite")
            if db not in orig_cache:
                orig_cache[db] = old_schema_of(ORIG[evo] % (db, db))
            old_schema = orig_cache[db]

            gold_rows, gold_err = q(dbp, it.get("new_sql") or "")
            if gold_err:
                continue

            old_sql = it.get("old_sql") or ""
            # 审计：投影星号 / COUNT(*) 各自出现次数
            try:
                tr = sqlglot.parse_one(old_sql, read="sqlite")
                if tr is not None:
                    ps = projection_stars(tr)
                    if ps:
                        n_proj_star += 1
                    if any(True for _ in tr.find_all(exp.Count)):
                        if any(isinstance(c.this, exp.Star)
                               for c in tr.find_all(exp.Count)):
                            n_cnt_star += 1
            except Exception:
                pass

            maps = {
                "c3r_gt": map_gt(it),
                "c3r_infer": (map_infer(evo, db) or (None, None))[0],
                "c3r_name": map_infer_name(evo, db),
            }

            for c in conds:
                mp = maps[c]
                iid = it.get("id") or "%s:%s:%s" % (evo, db, it.get("question_id"))
                if mp is None:
                    stat[c]["NO_MAP"] += 1
                    details[c].append({"id": iid, "db": db, "status": "NO_MAP", "pred": None})
                    continue
                sql2, status, det = rewrite(old_sql, mp, old_schema, star=args.star)
                if sql2 is None:
                    stat[c][status or "FAIL"] += 1
                    details[c].append({"id": iid, "db": db, "status": status or "FAIL",
                                       "detail": det, "pred": None})
                    continue
                if status is None and "star=" in (det or "") and "star=0" not in det:
                    n_star_expanded += 1
                rows, err = q(dbp, sql2)
                if err:
                    stat[c]["DB_ERROR"] += 1
                    details[c].append({"id": iid, "db": db, "status": "DB_ERROR",
                                       "detail": err, "note": status, "pred": sql2})
                    continue
                if norm(rows) == norm(gold_rows):
                    stat[c]["EX"] += 1
                    st = "EX"
                else:
                    stat[c]["SILENT"] += 1
                    st = "SILENT"
                if status == "NO_HIT":
                    stat[c]["NO_HIT_IDENT"] += 1
                details[c].append({"id": iid, "db": db, "status": st, "note": status,
                                   "pred": sql2,
                                   "mirror_gold": _same_sql(sql2, it.get("new_sql") or "")})

        n_all = len(items)
        summary[evo] = {"n_items": n_all,
                        "projection_star_items": n_proj_star,
                        "count_star_items": n_cnt_star,
                        "star_expanded_hits": n_star_expanded,
                        "conds": {c: {"counts": dict(stat[c]),
                                      "ex_rate": round(stat[c]["EX"] / n_all * 100, 1)
                                      if n_all else None}
                                  for c in conds}}
        audit["projection_star_items"] += n_proj_star
        audit["count_star_items"] += n_cnt_star
        audit["star_expanded"] += n_star_expanded
        audit["per_evo"][evo] = {"n": n_all, "projection_star": n_proj_star,
                                 "count_star": n_cnt_star, "expanded": n_star_expanded}
        for c in conds:
            for r in details[c]:
                r["evo_type"] = evo
                r["cond"] = c
                out_rows.append(r)
        print("【%s】n=%d  投影星号=%d  COUNT(*)=%d  展开=%d"
              % (evo, n_all, n_proj_star, n_cnt_star, n_star_expanded), flush=True)
        for c in conds:
            print("  %-10s EX %d/%d (%.1f%%)  %s"
                  % (c, stat[c]["EX"], n_all,
                     stat[c]["EX"] / n_all * 100 if n_all else 0,
                     dict(stat[c])), flush=True)

    suffix = "" if args.star == "expand" else "_" + args.star
    json.dump({"version": VERSION, "summary": summary, "audit": audit},
              open("results/rule_rewrite%s.json" % suffix, "w", encoding="utf-8"),
              ensure_ascii=False, indent=1)
    with open("results/rule_rewrite_items%s.jsonl" % suffix, "w", encoding="utf-8") as fh:
        for r in out_rows:
            fh.write(json.dumps(r, ensure_ascii=False) + "\n")
    print("\n-> results/rule_rewrite%s.json" % suffix)
    print("审计：投影星号 %d 条 / COUNT(*) %d 条 / 展开命中 %d 次"
          % (audit["projection_star_items"], audit["count_star_items"],
             audit["star_expanded"]))


def _same_sql(a, b):
    f = lambda s: re.sub(r"\s+", " ", (s or "").strip().lower())
    return f(a) == f(b)


if __name__ == "__main__":
    main()
