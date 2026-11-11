# -*- coding: utf-8 -*-
"""E3：直接 LLM 改写基线（回答审稿意见「B1/B2/B3 不是端到端迁移系统基线」）

现有 C2 只给「新 schema + 旧 SQL」，模型并不知道旧 SQL 写的那套列名在新库里
叫什么、旧库长什么样。审稿人要求补一类真正的端到端迁移对照：

  c_rw   旧 schema DDL + 新 schema DDL + 旧 SQL + 问题
         -> 让模型自己对比两份 DDL 完成映射（传统 schema-evolution
            query rewriting 工具的标准输入形态）
  c_rwv  c_rw + 旧/新列的取值样本（每列前 5 个非空 distinct）
         -> 补上值级证据，这是 DBA 手工迁移时的常规做法

与 c3m 的唯一差别：c3m 给的是「抽取好的 old->new 映射行」，
c_rw/c_rwv 给的是「原始两份 DDL（+取值）」让模型自己推断。
因此它测的是：谱系抽取这一层，究竟有没有被 LLM 的内建对比能力替代掉。

用法：
  python run_rw_baseline.py --evo rename --limit 12 --cond c_rw,c_rwv --tag rw_smoke
  python run_rw_baseline.py --evo rename,rename_medium,rename_hard \
         --cond c_rw,c_rwv --tag rw_main
"""
import argparse
import collections
import json
import os
import re
import sys
import time

BASE = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE)
sys.path.insert(0, BASE)

from run_c1c2c3 import (CANNOT, EVOLVED_DIR, OUTDIR, SYSTEM, db_path,
                        judge, load_schema, parse_sql, run_sql)

CANNOT = "CANNOT_ANSWER"
OLD_DB_FMT = "data/bird/dev_20240627/dev_databases/{db}/{db}.sqlite"
MAX_OLD_SCHEMA = 5000
N_SAMPLE_VALUES = 5

IDENT = re.compile(r"[A-Za-z_][A-Za-z0-9_]*")


# ---------------------------------------------------------------- 旧库访问
_SQLITE = {}


def old_conn(db_id):
    if db_id not in _SQLITE:
        import sqlite3
        p = os.path.join(BASE, OLD_DB_FMT.format(db=db_id))
        _SQLITE[db_id] = sqlite3.connect(p, timeout=5) if os.path.exists(p) else None
    return _SQLITE[db_id]


def old_schema_text(db_id):
    con = old_conn(db_id)
    if con is None:
        return ""
    out = []
    try:
        tabs = [r[0] for r in con.execute(
            "SELECT name FROM sqlite_master WHERE type='table' "
            "AND name NOT LIKE 'sqlite_%'").fetchall()]
        for t in tabs:
            cols = con.execute(f'PRAGMA table_info("{t}")').fetchall()
            body = ", ".join(f"{c[1]} {c[2]}" for c in cols)
            out.append(f"CREATE TABLE {t} ({body});")
    except Exception:
        return ""
    return "\n".join(out)


def sample_values(con, table, col, k=N_SAMPLE_VALUES):
    if con is None:
        return []
    try:
        rs = con.execute(
            f'SELECT DISTINCT "{col}" FROM "{table}" '
            f'WHERE "{col}" IS NOT NULL LIMIT {k}').fetchall()
        return [str(r[0])[:40] for r in rs]
    except Exception:
        return []


def col_value_samples(db_id, evo, table, cols):
    """旧库 + 新库同一列的取值样本，用于 c_rwv。"""
    lines = []
    ocon = old_conn(db_id)
    ncon = None
    try:
        import sqlite3
        np_ = db_path(evo, db_id)
        ncon = sqlite3.connect(np_, timeout=5) if os.path.exists(np_) else None
    except Exception:
        ncon = None
    for c in cols:
        ov = sample_values(ocon, table, c)
        nv = sample_values(ncon, table, c)      # 同名新列（改名后通常不存在）
        if ov or nv:
            lines.append(f'{table}.{c}: OLD values {ov}'
                         + (f" | NEW values {nv}" if nv else ""))
    return lines


# ---------------------------------------------------------------- prompt
def build_prompt(r, cond, new_schema):
    q = (r["question"] or "").strip()
    old_sql = (r["old_sql"] or "").strip()
    blocks = [
        "### Database schema at the time the example SQL was written (OLD)",
        old_schema_text(r["db_id"]),
        "",
        "### Current database schema (NEW)",
        new_schema,
        "",
        "### Question",
        q,
        "",
        "### Example SQL written against the OLD schema (it no longer runs)",
        old_sql,
        "",
    ]

    if cond == "c_rwv":
        # 只给旧 SQL 里真实出现过的列，避免把整库取值都塞进提示
        idents = {t.lower() for t in IDENT.findall(old_sql)}
        pairs = []
        tabs = re.findall(r"FROM\s+[`\"]?(\w+)", old_sql, re.I) + \
               re.findall(r"JOIN\s+[`\"]?(\w+)", old_sql, re.I)
        for t in dict.fromkeys(tabs):
            cols = [c for c in idents]
            got = col_value_samples(r["db_id"], r["evo_type"], t, cols)
            pairs += got
        if pairs:
            blocks += ["### Column value samples (OLD vs NEW)", "\n".join(pairs), ""]

    blocks += [
        "### Instruction",
        "Rewrite the Example SQL so that it answers the Question against the "
        "CURRENT schema. Use the two schema listings to work out how each "
        "referenced column/tables maps between OLD and NEW.",
        f"If the Question can no longer be answered with the CURRENT schema "
        f"(the needed column/table was removed and has no successor), "
        f"output exactly {CANNOT} and nothing else.",
        "Output only the SQL (or the token above).",
    ]
    return "\n".join(blocks)


# ---------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--evo", default="rename")
    ap.add_argument("--cond", default="c_rw,c_rwv")
    ap.add_argument("--model", default="llm_config.ds.json")
    ap.add_argument("--manifest", default="data/manifest_v3.jsonl")
    ap.add_argument("--limit", type=int, default=0)
    ap.add_argument("--tag", default="rw_main")
    ap.add_argument("--ids", default="")
    ap.add_argument("--eval_only", action="store_true",
                    help="只跑评测设计集（992 条），排除构造层的其余条目")
    args = ap.parse_args()

    os.environ["LLM_CFG_PATH"] = args.model
    import llm_client as L
    cfg = L.load_cfg()
    print(f"实际模型: {cfg.get('model')} @ {cfg.get('base')}", flush=True)

    conds = [c.strip().lower() for c in args.cond.split(",") if c.strip()]
    evos = {x.strip() for x in args.evo.split(",") if x.strip()}
    rows = [json.loads(l) for l in open(os.path.join(BASE, args.manifest),
                                        encoding="utf-8")]
    rows = [r for r in rows if r["evo_type"] in evos]
    if args.eval_only:
        from stats_revise import EVAL as _EVAL
        rows = [r for r in rows if r["id"] in _EVAL]
    if args.ids:
        keep = {l.strip() for l in open(os.path.join(BASE, args.ids),
                                        encoding="utf-8") if l.strip()}
        rows = [r for r in rows if r["id"] in keep]
    if args.limit and len(rows) > args.limit:
        by = collections.defaultdict(list)
        for r in rows:
            by[r["evo_type"]].append(r)
        per = max(1, args.limit // max(1, len(by)))
        picked = []
        for e, lst in by.items():
            picked += lst[:per]
        rows = picked

    print(f"样本 {len(rows)} | 条件 {conds} | evo {dict(collections.Counter(r['evo_type'] for r in rows))}",
          flush=True)

    gold = {}
    for r in rows:
        g_rows, g_err = run_sql(db_path(r["evo_type"], r["db_id"]),
                                r["new_sql"] or "")
        gold[r["id"]] = (g_rows, g_err)

    os.makedirs(OUTDIR, exist_ok=True)
    out_path = os.path.join(OUTDIR, f"e3_{args.tag}.jsonl")
    done = set()
    if os.path.exists(out_path):
        for l in open(out_path, encoding="utf-8"):
            try:
                done.add(json.loads(l)["key"])
            except Exception:
                pass
        print(f"续跑：已有 {len(done)} 条", flush=True)

    fout = open(out_path, "a", encoding="utf-8")
    stat = collections.defaultdict(collections.Counter)
    t0 = time.time()
    n_call = 0

    for i, r in enumerate(rows, 1):
        schema = load_schema(db_path(r["evo_type"], r["db_id"]))
        g_rows, g_err = gold[r["id"]]
        for cond in conds:
            key = f"{cond}|{r['id']}"
            if key in done:
                continue
            prompt = build_prompt(r, cond, schema)
            try:
                text, meta = L.chat(SYSTEM, prompt)
            except Exception as e:
                print(f"  [llm-err] {key}: {type(e).__name__} {str(e)[:100]}",
                      flush=True)
                stat[cond]["LLM_ERROR"] += 1
                continue
            n_call += 1
            sql, is_ab = parse_sql(text)
            outcome = judge(r, sql, is_ab, g_rows, r["evo_type"])
            stat[cond][outcome] += 1
            rec = {"key": key, "cond": cond, "id": r["id"],
                   "evo_type": r["evo_type"], "expected": r["expected_action"],
                   "outcome": outcome, "sql": sql, "raw": (text or "")[:2500],
                   "db_id": r["db_id"], "question_id": r["question_id"],
                   "difficulty": r["difficulty"], "model": cfg.get("model"),
                   "prompt_chars": len(prompt)}
            if meta:
                rec["meta"] = meta
            fout.write(json.dumps(rec, ensure_ascii=False) + "\n")
            fout.flush()

        if i % 10 == 0 or i == len(rows):
            el = time.time() - t0
            print(f"  [{i}/{len(rows)}] {el/60:.1f}min "
                  f"{n_call/max(1e-9,el)*60:.1f} call/min", flush=True)

    fout.close()
    print("=== 汇总（EX=执行且结果正确）===", flush=True)
    for c in conds:
        s = stat[c]
        n = sum(v for k, v in s.items() if k.isupper())
        print(f"  {c:8} n={n:4} " + " ".join(f"{k}={v}" for k, v in sorted(s.items())),
              flush=True)
    print(f"输出 {out_path}", flush=True)


if __name__ == "__main__":
    main()
