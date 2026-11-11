# -*- coding: utf-8 -*-
"""真实数据上的编码可逆性边界（回答第七轮评审第五节）。

反事实探针（probe_counterfactual.py）在内存库里人为造了各类边界值，说明
「合并-再切分」只在「两分量都非 NULL 且不含分隔符」时才是恒等的。本脚本回答
一个更硬的问题：**真实 BIRD 数据上这个前提成立吗？**

两类事件各查一侧：
  合并事件 -> 查**原始库**的源列：IS NULL、值内含分隔符、首尾空白、空串。
     任一项 > 0 即说明该事件在构造阶段就已不可逆。
  拆分事件 -> 查**演化后库**的分量列：单个分量为 NULL（拼接后整体为 NULL）、
     分量内含分隔符（切分错位）。任一项 > 0 即说明重建不恒等。

输入 data/evolution_{merge,split}[ _holdout].jsonl、原始 / 演化后库
输出 results/data_bounds.json
"""
import collections
import glob
import json
import os
import re
import sqlite3

BASE = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE)

VERSION = "data-bounds-v1"
SEP_FUNCS = ["CHAR(31)", "CHAR(30)", "CHAR(29)", "CHAR(28)"]

CASES = [
    ("merge_col", "data/evolution_merge.jsonl", "data/evolved_merge",
     "data/bird/dev_20240627/dev_databases", "merge"),
    ("split_col", "data/evolution_split.jsonl", "data/evolved_split",
     "data/bird/dev_20240627/dev_databases", "split"),
    ("merge_col_holdout", "data/evolution_merge_holdout.jsonl",
     "data/evolved_merge_holdout", "data/bird_train/train/train_databases", "merge"),
    ("split_col_holdout", "data/evolution_split_holdout.jsonl",
     "data/evolved_split_holdout", "data/bird_train/train/train_databases", "split"),
]


def find(db, root):
    for p in glob.glob(os.path.join(root, "**", db + ".sqlite"), recursive=True):
        return p
    return None


def cols(con, t):
    try:
        return [r[1] for r in con.execute('PRAGMA table_info("%s")' % t)]
    except Exception:
        return []


def open_ro(p):
    return sqlite3.connect("file:%s?mode=ro" % p, uri=True)


def _sep_of(expr):
    """从 new 表达式里取出真正的分隔符字面量（INSTR 可直接用的形式）。"""
    for f in SEP_FUNCS:
        if f in (expr or ""):
            return f
    for m in re.finditer(r"\|\|\s*'([^']*)'\s*\|\|", expr or ""):
        return "'%s'" % m.group(1).replace("'", "''")
    return None


def merge_event(orig_db, ops):
    t = ops[0]["table"]
    olds = [o["old"] for o in ops]
    con = open_ro(orig_db)
    cs = cols(con, t)
    if any(o not in cs for o in olds):
        con.close()
        return None
    n = con.execute('SELECT COUNT(*) FROM "%s"' % t).fetchone()[0]
    out = {"table": t, "old_cols": olds, "n_rows": n, "per_col": {}}
    bad = 0
    for o in olds:
        c = '"%s"' % o
        nu = con.execute('SELECT COUNT(*) FROM "%s" WHERE %s IS NULL' % (t, c)).fetchone()[0]
        em = con.execute("SELECT COUNT(*) FROM \"%s\" WHERE %s=''" % (t, c)).fetchone()[0]
        ws = con.execute('SELECT COUNT(*) FROM "%s" WHERE %s<>TRIM(%s)' % (t, c, c)).fetchone()[0]
        sepexpr = _sep_of(ops[0].get("new"))
        sp = (con.execute("SELECT COUNT(*) FROM \"%s\" WHERE INSTR(%s,%s)>0"
                          % (t, c, sepexpr)).fetchone()[0] if sepexpr else 0)
        out["per_col"][o] = {"null": nu, "empty": em,
                             "ws_trim_diff": ws, "has_sep_char": sp}
        bad = max(bad, nu, sp)
    out["lossy_rows"] = bad
    out["lossy_pct"] = round(bad / n * 100, 2) if n else None
    con.close()
    return out


def split_event(evo_db, ops):
    t = ops[0]["table"]
    src = ops[0]["old"]
    new = ops[0]["new"] or ""
    parts = [x.strip() for x in new.split("||")]
    comps = [x for x in parts if re.fullmatch(r"attr_\d+", x)]
    con = open_ro(evo_db)
    cs = cols(con, t)
    if not comps or any(c not in cs for c in comps):
        con.close()
        return None
    n = con.execute('SELECT COUNT(*) FROM "%s"' % t).fetchone()[0]
    out = {"table": t, "old_col": src, "components": comps,
           "separator": [x for x in parts if x not in comps], "n_rows": n}
    null_any = con.execute(
        'SELECT COUNT(*) FROM "%s" WHERE %s'
        % (t, " OR ".join('"%s" IS NULL' % c for c in comps))).fetchone()[0]
    all_null = con.execute(
        'SELECT COUNT(*) FROM "%s" WHERE %s'
        % (t, " AND ".join('"%s" IS NULL' % c for c in comps))).fetchone()[0]
    sep_in = con.execute(
        'SELECT COUNT(*) FROM "%s" WHERE %s'
        % (t, " OR ".join("INSTR(\"%s\",CHAR(31))>0" % c for c in comps))).fetchone()[0]
    empty = con.execute(
        'SELECT COUNT(*) FROM "%s" WHERE %s'
        % (t, " OR ".join("\"%s\"=''" % c for c in comps))).fetchone()[0]
    out.update({"null_any": null_any, "all_null": all_null,
                "sep_in_component": sep_in, "empty_component": empty})
    out["lossy_rows"] = max(null_any - all_null, sep_in, 0)
    out["lossy_pct"] = round(out["lossy_rows"] / n * 100, 2) if n else None
    con.close()
    return out


out = {"version": VERSION, "cases": {}}

for evo, srcfile, evodir, origroot, side in CASES:
    if not os.path.exists(srcfile):
        continue
    items = [json.loads(l) for l in open(srcfile, encoding="utf-8")]
    EV = collections.defaultdict(list)
    for it in items:
        ops = it.get("drift_ops") or []
        if not ops:
            continue
        sig = (it["db_id"], ops[0].get("table"),
               tuple(o.get("old") or "" for o in ops))
        EV[sig].append(it)

    per = []
    for sig, its in sorted(EV.items(), key=lambda kv: -len(kv[1])):
        db, tbl, _ = sig
        ops = its[0].get("drift_ops")
        dbp = find(db, origroot if side == "merge" else evodir)
        if not dbp:
            per.append({"db": db, "table": tbl, "n_items": len(its),
                        "status": "DB_NOT_FOUND"})
            continue
        try:
            r = merge_event(dbp, ops) if side == "merge" else split_event(dbp, ops)
        except Exception as e:
            r = None
            per.append({"db": db, "table": tbl, "n_items": len(its),
                        "status": "ERROR", "detail": str(e)[:140]})
            continue
        if not r:
            per.append({"db": db, "table": tbl, "n_items": len(its),
                        "status": "COL_NOT_FOUND"})
            continue
        r.update({"db": db, "n_items": len(its), "status": "OK"})
        per.append(r)

    ok = [x for x in per if x["status"] == "OK"]
    tot = sum(x["n_rows"] for x in ok)
    lossy = sum(x["lossy_rows"] for x in ok)
    out["cases"][evo] = {
        "side": side, "n_items": len(items), "n_events": len(EV),
        "n_events_measured": len(ok), "n_rows": tot,
        "lossy_rows": lossy,
        "lossy_row_pct": round(lossy / tot * 100, 3) if tot else None,
        "n_events_with_loss": sum(1 for x in ok if x["lossy_rows"] > 0),
        "events": per}

json.dump(out, open("results/data_bounds.json", "w", encoding="utf-8"),
          ensure_ascii=False, indent=1)

for evo, v in out["cases"].items():
    print("%-20s %-5s 事件 %d/%d 实测 | 行 %7d | 越界行 %6d (%.2f%%) | 有损失事件 %d"
          % (evo, v["side"], v["n_events_measured"], v["n_events"],
             v["n_rows"], v["lossy_rows"], v["lossy_row_pct"] or 0,
             v["n_events_with_loss"]))
    for x in v["events"]:
        if x.get("lossy_rows"):
            if v["side"] == "merge":
                d = {o: {k: s[k] for k in ("null", "has_sep_char") if s[k]}
                     for o, s in x["per_col"].items() if s["null"] or s["has_sep_char"]}
            else:
                d = {k: x[k] for k in ("null_any", "all_null", "sep_in_component")
                     if x.get(k)}
            print("    %-24s %-20s %6d/%6d  %s"
                  % (x["db"], x["table"], x["lossy_rows"], x["n_rows"], d))
print("\n-> results/data_bounds.json")
