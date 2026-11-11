# -*- coding: utf-8 -*-
"""探查「独立演化事件」的扩容空间（不改任何数据，只读）。

当前瓶颈：build_evolution_v2.py 的 run() 对每个库只调用一次 pick_merge_pair /
pick_split_col，即「每库每档 1 个事件」。于是合并档 230 条只对应 5 个事件、
拆分档 104 条对应 5 个事件，段位证据的统计功效被压在 10 个单元。

本脚本量化两件事：
  1) 若改为「每表 1 个事件」，每个库还能多出多少可用事件（表必须有 gold SQL 引用，
     否则演化后没有题可测）；
  2) 若把库范围从当前 5 个扩到 dev 全部 11 个，还能多出多少。

判定与 build_evolution_v2.pick_merge_pair / pick_split_col 完全一致，
差别只是不再取 top-1 而是全量枚举。输出 results/event_expansion_probe.json。

用法：python probe_event_expansion.py
"""
import collections
import importlib.util
import json
import os
import sqlite3

BASE = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE)


def _load_v2():
    spec = importlib.util.spec_from_file_location(
        "bev2", os.path.join(BASE, "build_evolution_v2.py"))
    m = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(m)
    return m


v2 = _load_v2()

# 与 audit_independence.py 的 paper_set 完全同一批题
ITEMS = json.load(open("data/bird/dev_20240627/dev.json", encoding="utf-8"))
MAN = [json.loads(l) for l in open("data/manifest_v3.jsonl", encoding="utf-8")]
_used = set()
for f in ("data/ids_xm_full.txt", "data/ids_ms.txt"):
    _used |= {l.strip() for l in open(f, encoding="utf-8") if l.strip()}
USED_DBS = sorted({r["db_id"] for r in MAN if r["id"] in _used})

rows_by_db = collections.defaultdict(list)
for it in ITEMS:
    rows_by_db[it["db_id"]].append(it)


def merge_candidates(conn, rows, cols):
    """全量枚举：返回每个可用表的 (table, [列...])，按被引用次数降序。"""
    freq = v2.freq_map(rows)
    owners, pk = collections.defaultdict(set), {}
    for t, c, _typ, _nn, p in cols:
        owners[c].add(t)
        pk[c] = p
    bytable = collections.defaultdict(list)
    for t, c, typ, _nn, p in cols:
        if p or len(owners[c]) != 1:
            continue
        bytable[t].append((freq.get(c, 0), c))
    out = []
    for t, lst in bytable.items():
        # 必须至少有一列被 gold SQL 引用，否则演化完没有题
        if not any(f >= 1 for f, _c in lst):
            continue
        lst.sort(reverse=True)
        txt = []
        for _f, c in lst:
            if v2.all_text(conn, t, c) and v2.no_null(conn, t, c):
                txt.append(c)
            if len(txt) >= 2:
                break
        if len(txt) < 2:
            continue
        ref = sum(f for f, c in lst if c in txt[:2])
        out.append({"table": t, "cols": txt[:2], "ref": ref})
    out.sort(key=lambda d: -d["ref"])
    return out


def split_candidates(conn, rows, cols):
    freq = v2.freq_map(rows)
    owners, pk = collections.defaultdict(set), {}
    for t, c, _typ, _nn, p in cols:
        owners[c].add(t)
        pk[c] = p
    out = []
    for t, c, typ, _nn, p in cols:
        if p or len(owners[c]) != 1 or freq.get(c, 0) < 1:
            continue
        if not v2.all_text(conn, t, c):
            continue
        try:
            nd = conn.execute(f'SELECT COUNT(DISTINCT "{c}") FROM "{t}"').fetchone()[0]
        except Exception:
            continue
        if nd < 5:
            continue
        for sep in v2.SPLIT_SEPS:
            # 注意：split 方向不能按 merge 的标准拒绝 NULL。merge 时 a||sep||b
            # 在有 NULL 分量处会塌缩（不可逆）；split 后 gold 是 a||sep||b，
            # 原值 NULL 则两段皆 NULL、拼回仍是 NULL，语义等价成立，故 NULL 可放行。
            try:
                bad = conn.execute(
                    f'SELECT 1 FROM "{t}" WHERE '
                    f'INSTR(COALESCE("{c}",""), {sep}) = 0 LIMIT 1').fetchone()
            except Exception:
                continue
            if bad is None:
                out.append({"table": t, "col": c, "sep": sep, "ref": freq[c]})
                break
    out.sort(key=lambda d: -d["ref"])
    return out


res = {"version": "event-expansion-probe-v1", "used_dbs": USED_DBS,
       "by_db": {}, "totals": {}}
cur = {"merge": collections.Counter(), "split": collections.Counter()}
new = {"merge": collections.Counter(), "split": collections.Counter()}

for db in USED_DBS:
    src = v2.sqlite_path(db)
    size_mb = os.path.getsize(src) / 1024 / 1024
    tabs, cols = v2.list_schema(src)
    conn = v2.connect(src, ro=True)
    mg = merge_candidates(conn, rows_by_db[db], cols)
    sp = split_candidates(conn, rows_by_db[db], cols)
    conn.close()
    # 现有构造实际只用了 top-1 事件落在哪张表，用于算「新增」
    res["by_db"][db] = {
        "size_mb": round(size_mb, 1),
        "n_items": len(rows_by_db[db]),
        "merge_tables": len(mg), "split_cols": len(sp),
        "merge_detail": mg, "split_detail": sp,
        "cap_ok": size_mb <= v2.MAX_DROP_MB,
    }
    if size_mb <= v2.MAX_DROP_MB:
        cur["merge"][db] = 1 if mg else 0
        cur["split"][db] = 1 if sp else 0
        new["merge"][db] = len(mg)
        new["split"][db] = len(sp)

res["totals"] = {
    "current_events": {"merge": sum(cur["merge"].values()),
                       "split": sum(cur["split"].values())},
    "expanded_events": {"merge": sum(new["merge"].values()),
                        "split": sum(new["split"].values())},
    "dbs_with_candidates": {"merge": [d for d in new["merge"] if new["merge"][d]],
                            "split": [d for d in new["split"] if new["split"][d]]},
}

json.dump(res, open("results/event_expansion_probe.json", "w", encoding="utf-8"),
          ensure_ascii=False, indent=1)

print("%-24s %6s %5s %6s %6s" % ("db", "MB", "items", "merge", "split"))
for db, d in res["by_db"].items():
    print("%-24s %6.1f %5d %6d %6d%s"
          % (db, d["size_mb"], d["n_items"], d["merge_tables"], d["split_cols"],
             "" if d["cap_ok"] else "  (超 130MB，DROP 不可行)"))
print()
print("当前事件：", res["totals"]["current_events"])
print("扩容后事件：", res["totals"]["expanded_events"])
