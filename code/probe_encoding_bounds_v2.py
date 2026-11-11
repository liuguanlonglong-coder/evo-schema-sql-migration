# -*- coding: utf-8 -*-
"""probe_encoding_bounds_v2.py - 段位编码压力测试 v2（P1 第 5 项）。

v1（probe_encoding_bounds.py）只测到 nseg=3、每类 1-3 行、单一分隔符 CHAR(31)。
本版补齐审稿意见点名的「三段以上编码」：
  · nseg = 2..6
  · 分隔符 4 种 ASCII/控制字符 + 1 个多字节字符
  · 每格 60 行固定种子随机值
  · 8 类边界情形
判据沿用 v1：整行全部分量同时恢复，不以单列恢复计。
只用生产 m1_part.to_expr 反解，零 API 调用。
输出 results/encoding_bounds_v2.json
"""
import collections
import json
import os
import random
import sqlite3
import sys

BASE = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE)
import m1_part as mp

VERSION = "encoding-bounds-v2"
# --legacy：复现修正前的 to_expr 末段写法，用作「修正前 0%」的可复现留档。
# 默认模式只走生产 mp.to_expr，确保测的就是正文所用实现。
LEGACY = "--legacy" in sys.argv
OUT = ("results/encoding_bounds_v2_legacy.json" if LEGACY
       else "results/encoding_bounds_v2.json")
NSEGS = [2, 3, 4, 5, 6]
SEPS = [("CHAR31", "\x1f"), ("PIPE", "|"), ("SEMI", ";"),
        ("TILDE", "~"), ("MB", "\uff5c")]
CASES = ["normal", "null_any", "empty_first", "empty_last",
         "sep_in_nonlast", "sep_in_last", "whitespace", "multibyte"]
ROWS = 60
SEED = 20261111
rng = random.Random(SEED)


def make_values(nseg, sep, case, i):
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
    return vals


def build_exprs(nseg, sep):
    """默认走生产 to_expr；legacy 复现修正前末段多切一次分隔符的写法。"""
    if not LEGACY:
        return [mp.to_expr("m", i, nseg, sep) for i in range(nseg)]
    sch = "'%s'" % sep if sep.isprintable() else "CHAR(%d)" % ord(sep)
    # 修正前的 to_expr 对 nseg<=2 走独立分支，末段直接取剩余全部（可恢复）；
    # 只有 nseg>=3 才多切一次。legacy 必须精确复现这一点，否则会把两段也判成 0%。
    if nseg <= 2:
        return ["SUBSTR(m,1,INSTR(m,%s)-1)" % sch,
                "SUBSTR(m,INSTR(m,%s)+1)" % sch]

    def rest(k):
        e = "m"
        for _ in range(k):
            e = "SUBSTR(%s,INSTR(%s,%s)+1)" % (e, e, sch)
        return e

    out = []
    for i in range(nseg):
        if i == 0:
            out.append("SUBSTR(m,1,INSTR(m,%s)-1)" % sch)
        else:
            head = rest(i)
            if i >= nseg - 1:
                out.append("SUBSTR(%s,INSTR(%s,%s)+1)" % (head, head, sch))
            else:
                out.append("SUBSTR(%s,1,INSTR(%s,%s)-1)" % (head, head, sch))
    return out


cells = {}
for nseg in NSEGS:
    for sname, sep in SEPS:
        for case in CASES:
            ok = 0
            con = sqlite3.connect(":memory:")
            con.execute("CREATE TABLE t(m TEXT)")
            exprs = build_exprs(nseg, sep)
            sql = "SELECT " + ", ".join(exprs) + " FROM t"
            for i in range(ROWS):
                vals = make_values(nseg, sep, case, i)
                merged = None if any(v is None for v in vals) else sep.join(vals)
                con.execute("DELETE FROM t")
                con.execute("INSERT INTO t VALUES (?)", (merged,))
                got = con.execute(sql).fetchone()
                if all(g == w for g, w in zip(tuple(got), tuple(vals))):
                    ok += 1
            con.close()
            cells["%d|%s|%s" % (nseg, sname, case)] = {
                "n": ROWS, "ok": ok, "rate_pct": round(ok / ROWS * 100, 1)}


def agg(keyfn):
    g = collections.defaultdict(list)
    for k, v in cells.items():
        nseg, sname, case = k.split("|")
        g[keyfn(int(nseg), sname, case)].append(v["rate_pct"])
    return {str(k): round(sum(v) / len(v), 1) for k, v in sorted(g.items(), key=lambda x: str(x[0]))}


by_case = agg(lambda n, s, c: c)
by_nseg = agg(lambda n, s, c: n)
by_sep = agg(lambda n, s, c: s)

# 逐格列出 nseg>=4 与 nseg<=3 的差异，供正文界定「三段以上是否改变结论」
diff = []
for sname, _ in SEPS:
    for case in CASES:
        lo = [cells["%d|%s|%s" % (n, sname, case)]["rate_pct"] for n in (2, 3)]
        hi = [cells["%d|%s|%s" % (n, sname, case)]["rate_pct"] for n in (4, 5, 6)]
        if round(sum(lo) / len(lo), 1) != round(sum(hi) / len(hi), 1):
            diff.append({"sep": sname, "case": case,
                         "nseg_2_3_pct": round(sum(lo) / len(lo), 1),
                         "nseg_4_6_pct": round(sum(hi) / len(hi), 1)})

report = {
    "version": VERSION,
    "config": {"mode": "legacy_prefix" if LEGACY else "production_to_expr",
               "nsegs": NSEGS, "seps": [s for s, _ in SEPS], "cases": CASES,
               "rows_per_cell": ROWS, "seed": SEED,
               "criterion": "all components of a row recovered simultaneously"},
    "cells": cells,
    "by_case": by_case,
    "by_nseg": by_nseg,
    "by_sep": by_sep,
    "nseg_ge4_changes": diff,
    "expr_sample": {"nseg": 4, "sep": "CHAR31",
                    "exprs": build_exprs(4, "\x1f")},
}
json.dump(report, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("saved", OUT, "cells", len(cells))
print("\n== by_case ==")
for k, v in by_case.items():
    print("  %-18s %6.1f%%" % (k, v))
print("\n== by_nseg ==")
for k, v in by_nseg.items():
    print("  nseg=%-3s %6.1f%%" % (k, v))
print("\n== by_sep ==")
for k, v in by_sep.items():
    print("  %-8s %6.1f%%" % (k, v))
print("\n== nseg>=4 相对 nseg<=3 发生变化的格 ==")
for x in diff:
    print("  ", x)
