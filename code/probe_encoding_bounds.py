"""probe_encoding_bounds.py — 编码边界情形的可逆性探针（v1）

审稿意见（第十二轮）把「关键边界情况缺失」列为数据集质量的最低分项之一。
此前正文只写「构造未覆盖 NULL 与值内分隔符」，这是承认缺失，不是界定缺失。
本脚本把缺失变成可判定的问题：对每一类边界情形，在真实 SQLite 上按构造规则
编码、再用生产代码里的 to_expr 反解，逐行判定能否恢复，并区分失败的性质。

四类情形
    B0  对照：普通取值
    B1  任一分量为 NULL      —— 编码规则使整体为 NULL
    B2  分量为空串 ''        —— 真实数据中出现过，必须实测
    B3  取值本身含分隔符     —— 编码产生歧义
    S3  三段合并（nseg=3）   —— to_expr 的 nseg>=3 分支，此前从未在库上验证

判据
    不是「这一列能否恢复」，而是「整行全部分量能否同时恢复」。
    B1 若按单列判会被误判为可逆（NULL 反解仍是 NULL），
    实际是另一分量的信息随整体 NULL 一起丢失——这是本探针最要紧的一处口径。

失败类型
    fail_null   反解为 NULL 而原值非 NULL  → 显式丢失，下游表现为空结果集
    fail_value  反解出值但不等于原值       → 静默错值，下游不报错
    二者在五分类里分别落到 SILENT 的两端，危害不同，不能合并计数。

用法
    python probe_encoding_bounds.py            # 输出 results/encoding_bounds.json
"""
import json
import os
import sqlite3
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from m1_part import to_expr  # noqa: E402  生产代码，不复制粘贴

SEP = "\x1f"
OUT = "results/encoding_bounds.json"
VERSION = "encoding-bounds-v1"
SCH = "CHAR(31)"


def _classify(got, want):
    """单分量比较：区分显式丢失与静默错值。"""
    if got == want:
        return "ok"
    if got is None and want is not None:
        return "fail_null"
    return "fail_value"


def run_two(cols_pairs):
    """两段合并：返回逐情形判定。"""
    con = sqlite3.connect(":memory:")
    con.execute("CREATE TABLE t(a TEXT, b TEXT, merged TEXT)")
    con.executemany("INSERT INTO t(a,b) VALUES(?,?)", cols_pairs)
    # 编码规则与构造脚本一致：直接拼接，NULL 传播为整体 NULL
    con.execute("UPDATE t SET merged = a || %s || b" % SCH)
    e0 = to_expr("merged", 0, 2, SEP)
    e1 = to_expr("merged", 1, 2, SEP)
    rows = con.execute("SELECT a,b,merged,%s,%s FROM t" % (e0, e1)).fetchall()
    stat = {"n": len(rows), "ok": 0, "fail_null": 0, "fail_value": 0,
            "merged_null": 0, "detail": []}
    for a, b, m, r0, r1 in rows:
        if m is None:
            stat["merged_null"] += 1
        c0, c1 = _classify(r0, a), _classify(r1, b)
        # 整行判据：两个分量都对才算恢复
        whole = "ok" if (c0 == "ok" and c1 == "ok") else (
            "fail_null" if "fail_null" in (c0, c1) else "fail_value")
        stat[whole] += 1
        stat["detail"].append({"a": a, "b": b, "merged": m,
                               "rec_a": r0, "rec_b": r1,
                               "col_a": c0, "col_b": c1, "row": whole})
    con.close()
    stat["rate_pct"] = round(100.0 * stat["ok"] / stat["n"], 1) if stat["n"] else 0.0
    return stat


def run_three(triples):
    """三段合并：验证 to_expr 的 nseg>=3 分支。"""
    con = sqlite3.connect(":memory:")
    con.execute("CREATE TABLE t3(a TEXT, b TEXT, c TEXT, merged TEXT)")
    con.executemany("INSERT INTO t3(a,b,c) VALUES(?,?,?)", triples)
    con.execute("UPDATE t3 SET merged = a || %s || b || %s || c" % (SCH, SCH))
    ex = [to_expr("merged", i, 3, SEP) for i in range(3)]
    rows = con.execute("SELECT a,b,c,%s,%s,%s FROM t3" % tuple(ex)).fetchall()
    stat = {"n": len(rows), "ok": 0, "fail_null": 0, "fail_value": 0,
            "merged_null": 0, "exprs": ex, "detail": []}
    for a, b, c, r0, r1, r2 in rows:
        cls = [_classify(r0, a), _classify(r1, b), _classify(r2, c)]
        whole = "ok" if all(x == "ok" for x in cls) else (
            "fail_null" if "fail_null" in cls else "fail_value")
        stat[whole] += 1
        stat["detail"].append({"want": [a, b, c], "rec": [r0, r1, r2], "row": whole})
    con.close()
    stat["rate_pct"] = round(100.0 * stat["ok"] / stat["n"], 1) if stat["n"] else 0.0
    return stat


def real_data():
    """真实数据的边界出现频次，来自 data_bounds.json 的逐列扫描。

    两个口径必须分开报：逐列扫描只对合并侧做过（拆分侧不涉及分量拼接），
    故 null/empty/sep 的分母是合并侧的 13 个事件，
    而全量事件数（合并 + 拆分）为 23 个、215,216 行。
    """
    p = "results/data_bounds.json"
    if not os.path.exists(p):
        return None
    d = json.load(open(p, encoding="utf-8"))
    tot = {"null": 0, "empty": 0, "ws_trim_diff": 0, "has_sep_char": 0}
    m_events = m_rows = lossy = 0
    all_events = all_rows = 0
    for case in d["cases"].values():
        all_events += case.get("n_events", 0)
        all_rows += case.get("n_rows", 0)
        for ev in case.get("events") or []:
            if "per_col" not in ev:
                continue
            m_events += 1
            m_rows += ev.get("n_rows", 0)
            lossy += ev.get("lossy_rows", 0)
            for st in ev["per_col"].values():
                for k in tot:
                    tot[k] += st.get(k, 0)
    return {"merge_side": {"events": m_events, "rows": m_rows, "lossy_rows": lossy, **tot},
            "all_sides": {"events": all_events, "rows": all_rows},
            "source": "results/data_bounds.json"}


def main():
    two_cases = [
        ("B0",  "对照：普通取值（左, 右）", [("x", "y"), ("Acme", "Smith"), ("2020-01-01", "M")]),
        ("B1a", "左分量为 NULL", [("x", None)]),
        ("B1b", "右分量为 NULL", [(None, "q")]),
        ("B2a", "左分量为空串", [("", "q")]),
        ("B2b", "右分量为空串", [("x", "")]),
        ("B3a", "左分量内含分隔符", [("a" + SEP + "b", "c")]),
        ("B3b", "右分量内含分隔符", [("a", "b" + SEP + "c")]),
    ]
    out = {"version": VERSION, "sep": SCH,
           "note": "判据为整行全部分量同时恢复，非单列恢复",
           "two_seg": {}, "three_seg": None, "real_data": real_data()}

    for cid, desc, data in two_cases:
        st = run_two(data)
        st["desc"] = desc
        out["two_seg"][cid] = st

    three_cases = [
        ("S3a", "三段全部普通值", [("x", "y", "z"), ("Acme", "Co", "Ltd")]),
        ("S3b", "三段中段含分隔符", [("x", "y" + SEP + "w", "z")]),
        ("S3c", "三段首段为 NULL", [(None, "y", "z")]),
    ]
    out["three_seg"] = {}
    for cid, desc, data in three_cases:
        st = run_three(data)
        st["desc"] = desc
        out["three_seg"][cid] = st

    os.makedirs("results", exist_ok=True)
    json.dump(out, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)

    print("== 两段 ==")
    for cid, st in out["two_seg"].items():
        print(" %-4s %-22s n=%d ok=%d null=%d val=%d  %s" %
              (cid, st["desc"], st["n"], st["ok"], st["fail_null"],
               st["fail_value"], st["rate_pct"]))
    print("== 三段 ==")
    for cid, st in out["three_seg"].items():
        print(" %-4s %-22s n=%d ok=%d null=%d val=%d  %s" %
              (cid, st["desc"], st["n"], st["ok"], st["fail_null"],
               st["fail_value"], st["rate_pct"]))
    rd = out["real_data"]
    if rd:
        m = rd["merge_side"]
        print("== 真实数据（合并侧，逐列扫描）= 事件 %d / 行 %d / null %d / empty %d / "
              "内含分隔符 %d / 有损行 %d"
              % (m["events"], m["rows"], m["null"], m["empty"],
                 m["has_sep_char"], m["lossy_rows"]))
        a = rd["all_sides"]
        print("== 真实数据（全量）= 事件 %d / 行 %d" % (a["events"], a["rows"]))
    print("saved:", OUT)


if __name__ == "__main__":
    main()
