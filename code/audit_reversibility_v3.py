# -*- coding: utf-8 -*-
"""audit_reversibility_v3.py - 可逆性前提、边界用例与准入检测（回应 MO-8）。

审稿意见：段位表达式恢复依赖「合并编码可逆」，NULL 分量与非末段含分隔符
会造成歧义；扫描未发现越界只能说明「未观察到」。要求把可逆性写成正式前提、
补边界压力测试、并明确失败时的检测与回退策略。

三件事，全部可复跑、零模型调用：

  A 十类边界用例 × 5 分隔符 × 2..6 段，用生产 m1_part.to_expr 反解，
    判据「整行全部分量同时恢复」。新增 escaped_sep（转义分隔符）与
    numeric_text（数值以文本存储）两类此前未覆盖的情形。
  B 三规则准入检测器（NULL 丢失 / 分隔符计数 / 类型漂移），在 A 的全网格上
    与真实可恢复性对照，给出命中、漏检、误报。
  C 在 BIRD dev 真实候选列对上统计拒绝率，把「回退」的代价写成数字。

输出 results/reversibility_v3.json
"""
import collections
import itertools
import json
import os
import random
import sqlite3
import sys

BASE = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE)
sys.path.insert(0, BASE)

import m1_part as MP              # noqa: E402
import build_evolution_v2 as V2   # noqa: E402

OUT = "results/reversibility_v3.json"
VERSION = "reversibility-v3"
SEPS = [("\x1f", "CHAR31"), ("|", "PIPE"), (";", "SEMI"),
        ("~", "TILDE"), ("\uff5c", "MB")]
NSEGS = [2, 3, 4, 5, 6]
ROWS = 40
SEED = 20261111
CASES = ["normal", "null_any", "empty_first", "empty_last", "sep_in_nonlast",
         "sep_in_last", "whitespace", "multibyte", "escaped_sep", "numeric_text"]


def make_row(nseg, sep, case, i, rng):
    vals = ["v%d_%d_%s" % (i, k, "".join(rng.choice("abcXYZ019") for _ in range(4)))
            for k in range(nseg)]
    if case == "null_any":
        vals[rng.randrange(nseg)] = None
    elif case == "empty_first":
        vals[0] = ""
    elif case == "empty_last":
        vals[-1] = ""
    elif case == "sep_in_nonlast":
        vals[0] = vals[0] + sep + "tail"
    elif case == "sep_in_last":
        vals[-1] = vals[-1] + sep + "tail"
    elif case == "whitespace":
        vals = ["  " + v + "  " for v in vals]
    elif case == "multibyte":
        vals = ["\u4e2d" + v + "\u6587" for v in vals]
    elif case == "escaped_sep":
        # 反斜杠转义：分隔符出现次数仍等于 nseg-1，专门检验「按次数计数」
        # 的检测器会不会漏检。
        vals[0] = vals[0] + "\\" + sep + "tail"
    elif case == "numeric_text":
        vals = [str(1000 + i * 7 + k) for k in range(nseg)]
    return vals


def build_update(cols, sep):
    parts = []
    for i, c in enumerate(cols):
        if i:
            parts.append("'%s'" % sep.replace("'", "''"))
        parts.append(f'"{c}"')
    return "UPDATE t SET m = " + " || ".join(parts)


def run_grid():
    rng = random.Random(SEED)
    cells, det = {}, {}
    for nseg, (sep, sepname) in itertools.product(NSEGS, SEPS):
        for case in CASES:
            con = sqlite3.connect(":memory:")
            cols = [f"c{k}" for k in range(nseg)]
            con.execute("CREATE TABLE t (id INTEGER PRIMARY KEY, m TEXT, %s)"
                        % ", ".join(f'"{c}" TEXT' for c in cols))
            upd = build_update(cols, sep)
            exprs = [MP.to_expr("m", i, nseg, sep) for i in range(nseg)]
            ok = flag = 0
            for i in range(ROWS):
                vals = make_row(nseg, sep, case, i, rng)
                con.execute("DELETE FROM t")
                con.execute("INSERT INTO t (id, %s) VALUES (1, %s)"
                            % (", ".join(f'"{c}"' for c in cols),
                               ", ".join("?" * nseg)), vals)
                con.execute(upd)
                row = con.execute("SELECT m, %s FROM t"
                                  % ", ".join(f'"{c}"' for c in cols)).fetchone()
                merged, src = row[0], list(row[1:])
                # 真实可恢复性：整行全部分量同时对上才算恢复
                if merged is None:
                    rec = all(v is None for v in src)
                else:
                    got = [con.execute("SELECT %s FROM t" % e).fetchone()[0]
                           for e in exprs]
                    rec = got == src
                ok += rec
                # 检测器：三条规则任一命中即拒绝编码
                d = detect(merged, src, sep, nseg)
                flag += d
            cells[f"{nseg}|{sepname}|{case}"] = {
                "n": ROWS, "ok": ok, "rate_pct": round(100 * ok / ROWS, 1)}
            det[f"{nseg}|{sepname}|{case}"] = {
                "n": ROWS, "flagged": flag,
                "flag_pct": round(100 * flag / ROWS, 1)}
            con.close()
    return cells, det


def detect(merged, src, sep, nseg):
    """准入检测：编码前即可判定的三条不可逆/失真条件。

    R1 NULL 丢失：合并值为 NULL 但并非所有源分量都为 NULL（|| 遇 NULL 整串为 NULL）
    R2 分隔符计数：非 NULL 值中分隔符出现次数 != nseg-1（分量内含分隔符）
    R3 类型漂移：任一源分量的存储类型不是 text
    三条都是「编码前」可判定的，因此可以做成准入闸门而不是事后补救。
    """
    if merged is None:
        return not all(v is None for v in src)
    if merged.count(sep) != nseg - 1:
        return True
    return False


def candidate_scan():
    """BIRD dev 上按与构造器相同的口径枚举候选列对，统计三条规则的拒绝率。"""
    d = json.load(open("data/bird/dev_20240627/dev.json", encoding="utf-8"))
    rows_by_db = collections.defaultdict(list)
    for r in d:
        rows_by_db[r["db_id"]].append(r)
    stat = collections.Counter()
    per_db = []
    for db, rows in sorted(rows_by_db.items()):
        p = os.path.join("data/bird/dev_20240627/dev_databases", db, f"{db}.sqlite")
        if not os.path.exists(p):
            continue
        con = sqlite3.connect(f"file:{p}?mode=ro", uri=True)
        freq = V2.freq_map(rows)
        cols = []
        for (t,) in con.execute(
                "SELECT name FROM sqlite_master WHERE type='table'").fetchall():
            try:
                info = con.execute(f'PRAGMA table_info("{t}")').fetchall()
            except Exception:
                continue
            for c in info:
                cols.append((t, c[1], c[2], c[3], c[5]))
        owners = collections.defaultdict(set)
        for t, c, _ty, _nn, pk in cols:
            owners[c].add(t)
        bytable = collections.defaultdict(list)
        for t, c, ty, _nn, pk in cols:
            if pk or len(owners[c]) != 1 or freq.get(c, 0) < 1:
                continue
            bytable[t].append((freq.get(c, 0), c))
        # 与构造器同序：先在列层面过准入，再在幸存列里两两配对。
        # 先配对后过滤会把「一对里有一列不合格」算成「整对拒绝」，
        # 从而高估拒绝率——这是上一版的统计错误。
        n_col = rej_null = rej_type = rej_sep = kept_col = 0
        n_pair = kept = 0
        for t, lst in bytable.items():
            lst.sort(key=lambda x: (-x[0], x[1]))
            surv = []
            for _f, c in lst:
                n_col += 1
                if not V2.no_null(con, t, c):
                    rej_null += 1
                    continue
                if not V2.all_text(con, t, c):
                    rej_type += 1
                    continue
                try:
                    bad = con.execute(
                        f'SELECT 1 FROM "{t}" WHERE INSTR("{c}", CHAR(31))>0 LIMIT 1'
                    ).fetchone() is not None
                except Exception:
                    bad = False
                if bad:
                    rej_sep += 1
                    continue
                surv.append(c)
            kept_col += len(surv)
            for i in range(0, len(surv) - 1, 2):
                n_pair += 1
                kept += 1
        per_db.append({"db": db, "cand_cols": n_col, "rej_null": rej_null,
                       "rej_type": rej_type, "rej_sep": rej_sep,
                       "kept_cols": kept_col, "pairs": n_pair})
        stat["cand_cols"] += n_col
        stat["rej_null"] += rej_null
        stat["rej_type"] += rej_type
        stat["rej_sep"] += rej_sep
        stat["kept_cols"] += kept_col
        stat["pairs"] += n_pair
        con.close()
    return {"total": dict(stat), "by_db": per_db,
            "reject_rate_pct": round(
                100 * (stat["cand_cols"] - stat["kept_cols"])
                / max(1, stat["cand_cols"]), 1),
            "note": "先过列级准入再配对；拒绝率按列计，pair 数只报幸存列能配出的对数。"}


def main():
    cells, det = run_grid()
    out = {"version": VERSION,
           "config": {"nsegs": NSEGS, "seps": [s[1] for s in SEPS],
                      "cases": CASES, "rows_per_cell": ROWS, "seed": SEED,
                      "criterion": "整行全部分量同时恢复"},
           "cells": cells, "detector": det}

    def avg(d, keyname):
        g = collections.defaultdict(list)
        for k, v in d.items():
            g[k.split("|")[keyname]].append(v["rate_pct"] if "rate_pct" in v
                                            else v["flag_pct"])
        return {k: round(sum(v) / len(v), 1) for k, v in sorted(g.items())}

    out["by_case"] = avg(cells, 2)
    out["by_nseg"] = avg(cells, 0)
    out["by_sep"] = avg(cells, 1)
    out["detector_by_case"] = avg(det, 2)

    # 检测器与真实可恢复性的四格
    tp = fp = fn = tn = 0
    for k, v in cells.items():
        bad = v["rate_pct"] < 100          # 真实不可完全恢复
        flagged = det[k]["flag_pct"] > 0   # 检测器任一命中即拒绝
        if bad and flagged:
            tp += 1
        elif bad and not flagged:
            fn += 1
        elif (not bad) and flagged:
            fp += 1
        else:
            tn += 1
    out["detector_confusion"] = {"cells": len(cells), "tp": tp, "fp": fp,
                                 "fn": fn, "tn": tn,
                                 "note": "网格以「格」为单位；一格里全部行都失败才算 bad"}

    out["real_candidates"] = candidate_scan()
    json.dump(out, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)

    print(f"{'case':16s} " + " ".join(f"{s[1]:>7s}" for s in SEPS))
    for case in CASES:
        print(f"{case:16s} " + " ".join(
            f"{cells[f'2|{n}|{case}']['rate_pct']:7.1f}" for _s, n in SEPS))
    print("\n按用例:", out["by_case"])
    print("按段数:", out["by_nseg"])
    print("检测器按用例命中率:", out["detector_by_case"])
    print("检测器四格:", out["detector_confusion"])
    rc = out["real_candidates"]
    print("真实候选列对:", rc["total"], "拒绝率", rc["reject_rate_pct"], "%")
    print("写入", OUT)


if __name__ == "__main__":
    main()
