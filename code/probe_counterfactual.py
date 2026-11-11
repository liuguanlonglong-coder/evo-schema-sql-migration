# -*- coding: utf-8 -*-
"""反事实边界行验证（回答第七轮评审第五节）。

评审指出：构造只用「两段 + 已知分隔符 + 脚本产生的干净值」，
在真实值上不保证等价。本节把边界情形显式枚举，用**真实 SQLite 语义**
检验两件事：

  R1 往返恒等   segment(encode(a, b)) 是否还原 (a, b)
                （split 方向：reconstruct(attrA, attrB) 是否还原 c）
  R2 谓词保真   同一逻辑行上，旧列谓词与新表达式谓词的真值是否一致

R2 是 R1 在查询层的直接后果：任何对旧列的引用在重写后都化为
`<段位表达式> <op> v`，故只需覆盖 `= v`、`<> v`、`IS NULL`、`LIKE`。

边界类（评审点名：NULL、空串、分隔符冲突、重复值、边界字符）
  CLEAN        正常两段
  NULL_SIB     一个分量为 NULL
  BOTH_NULL    两分量都 NULL
  EMPTY        分量为空串（单侧 / 双侧）
  SEP_IN_A     分隔符出现在第一分量内部
  SEP_IN_B     分隔符出现在第二分量内部
  SEP_REPEAT   值内出现多个分隔符（split 视角）
  NO_SEP       值内不含分隔符（split 视角）
  WS           首尾空白
  DUP          重复值（同值多行）
  UNICODE      非 ASCII / 控制字符边界

输出 results/probe_counterfactual.json
"""
import json
import os
import sqlite3

BASE = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE)

VERSION = "probe-counterfactual-v1"
SEP_MERGE = "|| CHAR(31) ||"        # 与 build_evolution_v2 的 SEP 一致
SEP_SPLIT = "' '"                    # 与 SPLIT_SEPS 首选一致（逐库实际值另取）

# 逐库实际用于 split 的分隔符，取自 evolution 清单的 drift_ops
def split_seps(evo, dbs):
    out = {}
    for fn in ("data/evolution_split.jsonl", "data/manifest_holdout_sample.jsonl"):
        if not os.path.exists(fn):
            continue
        for l in open(fn, encoding="utf-8"):
            r = json.loads(l)
            if r.get("evo_type") not in ("split_col", "split_col_holdout"):
                continue
            for op in r.get("drift_ops") or []:
                if op.get("op") == "split_column":
                    out[r["db_id"]] = op
            for op in r.get("ops") or []:
                if isinstance(op, list) and op[0] == "split_column":
                    out[r["db_id"]] = {"table": op[1], "old": op[2], "new": op[3]}
    return out


def q(con, sql):
    try:
        return con.execute(sql).fetchall()
    except Exception as e:
        raise RuntimeError("SQL 失败: %s | %s" % (sql[:160], str(e)[:120]))


def truth(con, col_expr, val, frm):
    """在指定表上求 `<expr> <op> v` 的真值；NULL 用 IS NULL。"""
    if val is None:
        return q(con, "SELECT (%s) IS NULL FROM %s" % (col_expr, frm))[0][0]
    return q(con, "SELECT (%s) = '%s' FROM %s"
             % (col_expr, str(val).replace("'", "''"), frm))[0][0]


def same(x, y):
    """NULL 与 0 不等价：谓词真值按 SQL 三值逻辑比较。"""
    return (x is None and y is None) or (x is not None and y is not None and int(x) == int(y))


# ---------------------------------------------------------------- 合并档
MERGE_CASES = [
    # (class, a, b)
    ("CLEAN", "Alameda", "Active"),
    ("CLEAN", "Los Angeles", "Merged"),
    ("NULL_SIB", "Alameda", None),
    ("NULL_SIB", None, "Active"),
    ("BOTH_NULL", None, None),
    ("EMPTY", "", "Active"),
    ("EMPTY", "Alameda", ""),
    ("EMPTY", "", ""),
    ("SEP_IN_A", "Ala\x1fmeda", "Active"),
    ("SEP_IN_B", "Alameda", "Act\x1five"),
    ("SEP_IN_A", "\x1f", "Active"),
    ("WS", "  Alameda  ", " Active "),
    ("DUP", "Alameda", "Active"),
    ("UNICODE", "阿拉米达", "启用"),
    ("UNICODE", "Åland", "Ø"),
]


def merge_probe():
    con = sqlite3.connect(":memory:")
    con.execute('CREATE TABLE ok(a TEXT, b TEXT, attr TEXT)')
    con.execute('CREATE TABLE nk(attr TEXT)')
    S1 = "SUBSTR(attr,1,INSTR(attr,CHAR(31))-1)"      # 与 drift_ops 逐字一致
    S2 = "SUBSTR(attr,INSTR(attr,CHAR(31))+1)"
    rows = []
    for cls, a, b in MERGE_CASES:
        va = "NULL" if a is None else "'%s'" % a.replace("'", "''")
        vb = "NULL" if b is None else "'%s'" % b.replace("'", "''")
        con.execute('DELETE FROM ok')
        con.execute('DELETE FROM nk')
        con.execute("INSERT INTO ok VALUES (%s, %s, %s %s %s)" % (va, vb, va, SEP_MERGE, vb))
        con.execute("INSERT INTO nk SELECT attr FROM ok")
        s1 = q(con, "SELECT %s FROM nk" % S1)[0][0]
        s2 = q(con, "SELECT %s FROM nk" % S2)[0][0]
        attr = q(con, "SELECT attr FROM nk")[0][0]
        # R2：四种谓词形态的真值是否一致
        r2 = {"a_eq": [truth(con, "a", a, "ok"), truth(con, S1, a, "nk")],
              "b_eq": [truth(con, "b", b, "ok"), truth(con, S2, b, "nk")],
              "a_null": [q(con, "SELECT (a) IS NULL FROM ok")[0][0],
                         q(con, "SELECT (%s) IS NULL FROM nk" % S1)[0][0]],
              "b_null": [q(con, "SELECT (b) IS NULL FROM ok")[0][0],
                         q(con, "SELECT (%s) IS NULL FROM nk" % S2)[0][0]]}
        r2_ok = all(same(v[0], v[1]) for v in r2.values())
        rows.append({"kind": "merge", "class": cls,
                     "a": a, "b": b,
                     "stored": attr,
                     "seg1": s1, "seg2": s2,
                     "roundtrip_a": s1 == a, "roundtrip_b": s2 == b,
                     "roundtrip": bool(s1 == a and s2 == b),
                     "pred_ok": r2_ok, "pred": r2})
    return rows


# ---------------------------------------------------------------- 拆分档
SPLIT_CASES = [
    ("CLEAN", "555 1234"),
    ("NULL_SIB", None),
    ("EMPTY", ""),
    ("SEP_REPEAT", "555 1234 5678"),
    ("NO_SEP", "5551234"),
    ("WS", " 555 1234 "),
    ("DUP", "555 1234"),
    ("UNICODE", "五五五 一二三四"),
]


def split_probe(sep="' '"):
    con = sqlite3.connect(":memory:")
    con.execute('CREATE TABLE ok(c TEXT)')
    con.execute('CREATE TABLE nk(x TEXT, y TEXT)')
    rows = []
    for cls, c in SPLIT_CASES:
        vc = "NULL" if c is None else "'%s'" % c.replace("'", "''")
        con.execute("DELETE FROM ok")
        con.execute("DELETE FROM nk")
        con.execute("INSERT INTO ok VALUES (%s)" % vc)
        # 迁移：x = 第一段，y = 分隔符之后的全部
        con.execute("INSERT INTO nk SELECT SUBSTR(c,1,INSTR(c,%s)-1), "
                    "SUBSTR(c,INSTR(c,%s)+LENGTH(%s)) FROM ok" % (sep, sep, sep))
        x, y = q(con, "SELECT x, y FROM nk")[0]
        rec = q(con, "SELECT x || %s || y FROM nk" % sep)[0][0]
        rec_expr = "x || %s || y" % sep
        rows.append({"kind": "split", "class": cls, "c": c,
                     "attrA": x, "attrB": y, "recon": rec,
                     "roundtrip": rec == c,
                     "pred_ok": same(truth(con, "c", c, "ok"),
                                     truth(con, rec_expr, c, "nk")),
                     "pred": {"c_eq": [truth(con, "c", c, "ok"),
                                       truth(con, rec_expr, c, "nk")],
                              "c_null": [q(con, "SELECT (c) IS NULL FROM ok")[0][0],
                                         q(con, "SELECT (%s) IS NULL FROM nk" % rec_expr)[0][0]]}})
    return rows


def summarize(rows):
    out = {}
    for r in rows:
        d = out.setdefault(r["class"], {"n": 0, "roundtrip": 0, "pred_ok": 0})
        d["n"] += 1
        d["roundtrip"] += int(r["roundtrip"])
        d["pred_ok"] += int(r["pred_ok"])
    for k, v in out.items():
        v["roundtrip_rate"] = round(v["roundtrip"] / v["n"] * 100, 1)
        v["pred_rate"] = round(v["pred_ok"] / v["n"] * 100, 1)
    tot = {"n": len(rows), "roundtrip": sum(r["roundtrip"] for r in rows),
           "pred_ok": sum(r["pred_ok"] for r in rows)}
    tot["roundtrip_rate"] = round(tot["roundtrip"] / tot["n"] * 100, 1)
    tot["pred_rate"] = round(tot["pred_ok"] / tot["n"] * 100, 1)
    return out, tot


def main():
    mg = merge_probe()
    sp = split_probe()
    mg_s, mg_t = summarize(mg)
    sp_s, sp_t = summarize(sp)

    out = {"version": VERSION, "merge": {"cases": mg, "by_class": mg_s, "total": mg_t},
           "split": {"cases": sp, "by_class": sp_s, "total": sp_t}}
    json.dump(out, open("results/probe_counterfactual.json", "w", encoding="utf-8"),
              ensure_ascii=False, indent=1)

    for name, s, t in (("merge", mg_s, mg_t), ("split", sp_s, sp_t)):
        print("【%s】往返恒等 %d/%d (%.1f%%)  谓词保真 %d/%d (%.1f%%)"
              % (name, t["roundtrip"], t["n"], t["roundtrip_rate"],
                 t["pred_ok"], t["n"], t["pred_rate"]))
        for k, v in sorted(s.items()):
            print("   %-11s n=%d  往返 %5.1f%%  谓词 %5.1f%%"
                  % (k, v["n"], v["roundtrip_rate"], v["pred_rate"]))
    print("\n-> results/probe_counterfactual.json")


if __name__ == "__main__":
    main()
