# -*- coding: utf-8 -*-
"""在 EvoSchema 公开的列级合并/拆分评测集上做外部对照。

为什么做这个对照
----------------
第五轮评审指出「新颖性相对近邻工作不够突出」。仅靠文字划界说服力有限：
需要证明本章的机制结论在**别人构造的样本**上是否也成立，而不只是在自建评测集上。

EvoSchema 的 eval_benchmark/bird_split_or_merge_columns.json 提供 84 条样本
（列级合并 47、拆分 37），字段含旧 SQL、旧 DDL、新 DDL、新 gold SQL。
它与本章的差别在任务定义：EvoSchema 固定 NLQ、不给旧 SQL；
本章以旧 SQL 为证据、要求结果等价。因此这里不评执行等价，只评**列引用正确性**
（EvoSchema 自有的 Column Match 指标），回答一个更窄的问题：

    在他人构造的列级合并/拆分样本上，把列级谱系显式注入，
    是否能降低「仍引用已删除旧列」的失败率、提高列引用 F1？

三个条件
--------
  c1   只给新 schema + 问题（无历史 SQL 基线）
  c2   旧 SQL + 新 schema（本章的「直接改写」）
  c3m  旧 SQL + 新 schema + 列级谱系（old→new）

不评 c3mp：EvoSchema 的样本不含「一列内多段编码」，没有段位信息可注入。
这一点本身是两者场景不同的证据，写进报告而不是掩盖。

数据可得性限制（如实记录）
--------------------------
  1. 41/84 条的 column_changes 字段为空（全部为 student_club 的 merge），
     映射改由新旧 DDL 的列集合差恢复；
  2. 评测集只给 DDL，不给施加演化后的数据库文件，也没有数据迁移规则，
     因此无法执行 SQL 求 EX——本脚本只做离线列引用匹配；
  3. 部分 new_gold_sql 疑似被重新生成而非机械适配（数值字面量与 WHERE 条件有变动），
     故 gold 只用于「列集合」参照，不用于语义等价判断。

输出 results/evoschema_eval.json
"""
import argparse
import collections
import json
import os
import re
import sys

BASE = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE)

VERSION = "evoschema-eval-v1"
SRC = "data/evoschema/bird_split_or_merge_columns.json"
OUT = "results/evoschema_eval.json"

SYSTEM = ("You are an expert SQLite analyst. You write SQL that runs on the CURRENT "
          "database schema. Output only SQL, no explanation, no markdown fences.")

SQL_KW = {"select", "from", "where", "and", "or", "as", "on", "join", "inner", "left",
          "right", "outer", "group", "by", "order", "having", "limit", "distinct",
          "count", "sum", "avg", "min", "max", "case", "when", "then", "else", "end",
          "not", "null", "in", "like", "between", "is", "asc", "desc", "union", "all",
          "cast", "real", "integer", "text", "round", "abs", "length", "substr",
          "instr", "coalesce", "ifnull", "trim", "lower", "upper", "replace", "printf",
          "with", "over", "partition", "row_number", "rank", "dense_rank", "exists"}


def ddl_cols(ddl):
    """从 'create table x( a text, b integer, ...)' 里取列名（跳过约束）。"""
    m = re.search(r"\((.*)\)", ddl, re.S)
    if not m:
        return []
    out = []
    for part in m.group(1).split(","):
        part = part.strip()
        if not part:
            continue
        low = part.lower()
        if low.startswith(("primary key", "foreign key", "unique", "constraint", "check")):
            continue
        name = part.split()[0].strip('`"[]')
        if name:
            out.append(name)
    return out


def schema_map(rec):
    """返回 (old_ddl, new_ddl, changes, kind)。

    changes 为 [(kind, old_cols, new_cols)]，kind ∈ {merge, split, rename}。
    优先用 column_changes；为空时用同名表的列集合差恢复。
    """
    old_t = {t: v["ddl"] for t, v in (rec.get("relevant_table") or {}).items()}
    new_t = {t: v["ddl"] for t, v in (rec.get("new_relevant_table") or {}).items()}
    changes = []
    cc = rec.get("column_changes")
    if isinstance(cc, dict) and cc:
        if rec["split_or_merge"] == "merge":
            for newcol, olds in cc.items():
                changes.append(("merge", list(olds), [newcol]))
        else:
            for oldcol, news in cc.items():
                changes.append(("split", [oldcol], list(news)))
        kind = rec["split_or_merge"]
    else:
        kind = "unknown"
        for t in set(old_t) & set(new_t):
            gone = [c for c in ddl_cols(old_t[t]) if c not in ddl_cols(new_t[t])]
            add = [c for c in ddl_cols(new_t[t]) if c not in ddl_cols(old_t[t])]
            if not gone or not add:
                continue
            if len(gone) > 1 and len(add) <= 2:
                changes.append(("merge", gone, add))
            elif len(gone) <= 2 and len(add) > 1:
                changes.append(("split", gone, add))
            elif len(gone) == 1 and len(add) == 1:
                changes.append(("rename", gone, add))
            kind = (changes[-1][0] if changes else "unknown")
    old_ddl = "\n\n".join(f"-- {t}\n{v}" for t, v in old_t.items())
    new_ddl = "\n\n".join(f"-- {t}\n{v}" for t, v in new_t.items())
    return old_ddl, new_ddl, changes, kind


def lineage_lines(changes):
    out = []
    for kind, olds, news in changes:
        if kind == "merge":
            out.append(f"- Column {', '.join('`'+c+'`' for c in olds)} "
                       f"were merged into a single column {', '.join('`'+c+'`' for c in news)}.")
        elif kind == "split":
            out.append(f"- Column {', '.join('`'+c+'`' for c in olds)} "
                       f"was split into {', '.join('`'+c+'`' for c in news)}.")
        else:
            out.append(f"- Column `{olds[0]}` was renamed to `{news[0]}`.")
    return out


def build_prompt(rec, cond, old_ddl, new_ddl, changes):
    parts = [f"### Question\n{rec['question'].strip()}"]
    if cond != "c1":
        parts.append(f"### Old SQL (written against the OLD schema)\n{rec['query'].strip()}")
    parts.append(f"### New schema\n{new_ddl}")
    if cond == "c3m":
        lines = lineage_lines(changes)
        if lines:
            parts.append("### Schema lineage (what changed since the old SQL was written)\n"
                         + "\n".join(lines))
    parts.append("Write one SQLite query that answers the question on the new schema. "
                 "Output only the SQL.")
    return "\n\n".join(parts)


_COL_RE = re.compile(r"\b(?:[A-Za-z_]\w*\.)?([A-Za-z_][\w ]*?)(?=\s*(?:,|\)|\bfrom\b|\bwhere\b|$))",
                     re.I)


def sql_cols(sql):
    """优先用 sqlglot 取列名；失败时退回正则。返回 (set, 解析方式)。"""
    sql = re.sub(r"```sql|```", "", sql or "").strip()
    try:
        import sqlglot
        from sqlglot import exp
        tree = sqlglot.parse_one(sql, read="sqlite")
        cols = {c.name.lower() for c in tree.find_all(exp.Column) if c.name}
        for s in tree.find_all(exp.Star):
            cols.add("*")
        return {c for c in cols if c and c not in SQL_KW}, "sqlglot"
    except Exception:
        body = re.sub(r"'[^']*'", " ", sql)
        toks = {t.lower() for t in re.findall(r"\b[A-Za-z_]\w*\b", body)}
        return {t for t in toks if t not in SQL_KW}, "regex"


def f1(pred, gold):
    if not pred and not gold:
        return None
    tp = len(pred & gold)
    p = tp / len(pred) if pred else 0.0
    r = tp / len(gold) if gold else 0.0
    return (2 * p * r / (p + r)) if (p + r) else 0.0


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--limit", type=int, default=0)
    ap.add_argument("--cond", default="c1,c2,c3m")
    ap.add_argument("--model", default="llm_config.json")
    ap.add_argument("--out", default=OUT)
    ap.add_argument("--shard", type=int, default=0, help="分片编号（0 起）")
    ap.add_argument("--of", type=int, default=1, help="分片总数；>1 时按索引取模")
    a = ap.parse_args()

    os.environ["LLM_CFG_PATH"] = a.model
    import llm_client as L
    cfg = L.load_cfg()
    print(f"模型: {cfg.get('model')} @ {cfg.get('base')}", flush=True)

    recs = json.load(open(SRC, encoding="utf-8"))
    # 同一 (db, question) 可能因多次扰动重复出现；保留首条，避免同一题多次计权
    seen, uniq = set(), []
    for r in recs:
        k = (r["db_id"], r["question"].strip())
        if k in seen:
            continue
        seen.add(k)
        uniq.append(r)
    if a.limit:
        uniq = uniq[:a.limit]
    if a.of > 1:
        uniq = uniq[a.shard::a.of]
    print(f"原始 {len(recs)} 条，去重后 {len(seen)} 条，本分片 {len(uniq)} 条 "
          f"({a.shard}/{a.of})", flush=True)

    conds = [c.strip() for c in a.cond.split(",") if c.strip()]
    rows = []
    for i, rec in enumerate(uniq, 1):
        old_ddl, new_ddl, changes, kind = schema_map(rec)
        gcols, ghow = sql_cols(rec["new_gold_sql"])
        gone = {c.lower() for _, olds, _ in changes for c in olds}
        rec_out = {"db": rec["db_id"], "kind": rec.get("split_or_merge"),
                   "recovered_kind": kind,
                   "n_changes": len(changes),
                   "changes": [{"kind": k, "old": o, "new": n} for k, o, n in changes],
                   "gold_cols": sorted(gcols), "gold_parse": ghow, "conds": {}}
        for cond in conds:
            user = build_prompt(rec, cond, old_ddl, new_ddl, changes)
            try:
                text, meta = L.chat(SYSTEM, user, cfg=cfg)
            except Exception as e:
                rec_out["conds"][cond] = {"error": str(e)[:160]}
                continue
            pcols, phow = sql_cols(text)
            rec_out["conds"][cond] = {
                "sql": text.strip()[:1200],
                "cols": sorted(pcols),
                "parse": phow,
                "f1": round(f1(pcols, gcols), 4) if f1(pcols, gcols) is not None else None,
                "exact_cols": pcols == gcols,
                "n_old_deleted_used": len(pcols & gone),
                "uses_deleted_old_col": bool(pcols & gone),
                "cached": bool(meta.get("cached")),
            }
        rows.append(rec_out)
        if i % 10 == 0 or i == len(uniq):
            print(f"  {i}/{len(uniq)}", flush=True)

    # ---- 汇总 ----
    summary = {}
    for cond in conds:
        vals = [r["conds"][cond] for r in rows if cond in r["conds"] and "error" not in r["conds"][cond]]
        f1s = [v["f1"] for v in vals if v["f1"] is not None]
        summary[cond] = {
            "n": len(vals),
            "mean_f1": round(sum(f1s) / len(f1s), 4) if f1s else None,
            "exact_cols_rate": round(sum(1 for v in vals if v["exact_cols"]) / len(vals), 4) if vals else None,
            "uses_deleted_old_col_rate": round(
                sum(1 for v in vals if v["uses_deleted_old_col"]) / len(vals), 4) if vals else None,
            "n_cached": sum(1 for v in vals if v["cached"]),
        }
    # 按变异类型分层（用恢复出的 kind）
    by_kind = {}
    for cond in conds:
        d = collections.defaultdict(list)
        for r in rows:
            v = r["conds"].get(cond) or {}
            if "error" in v or v.get("f1") is None:
                continue
            d[r["recovered_kind"]].append(v)
        by_kind[cond] = {k: {"n": len(v),
                             "mean_f1": round(sum(x["f1"] for x in v) / len(v), 4),
                             "uses_deleted_old_col_rate": round(
                                 sum(1 for x in v if x["uses_deleted_old_col"]) / len(v), 4)}
                         for k, v in d.items()}

    out_path = a.out if a.of == 1 else a.out.replace(".json", f".shard{a.shard}.json")
    out = {"version": VERSION, "source": SRC, "model": cfg.get("model"),
           "n_unique": len(uniq), "n_raw": len(recs),
           "shard": a.shard, "of": a.of,
           "note": "EvoSchema 只提供 DDL 与 gold SQL，未提供演化后的库文件，"
                   "故不评执行等价（EX），只评列引用 F1；gold 仅作列集合参照。",
           "summary": summary, "by_kind": by_kind, "rows": rows}
    os.makedirs("results", exist_ok=True)
    json.dump(out, open(out_path, "w", encoding="utf-8"), ensure_ascii=False, indent=1)

    print("\n==== 汇总 ====")
    for cond, s in summary.items():
        print(f"  {cond:4s} n={s['n']:3d}  列F1={s['mean_f1']}  "
              f"列集合全对={s['exact_cols_rate']}  仍用已删列={s['uses_deleted_old_col_rate']}")
    print("\n按变异类型:")
    for cond, d in by_kind.items():
        for k, v in d.items():
            print(f"  {cond:4s} {k:8s} n={v['n']:3d} F1={v['mean_f1']:.4f} "
                  f"用旧列={v['uses_deleted_old_col_rate']}")
    print(f"\n-> {out_path}")


if __name__ == "__main__":
    main()
