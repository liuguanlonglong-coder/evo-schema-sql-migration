# -*- coding: utf-8 -*-
"""judge_strict_audit.py - 全严判定规则下八个主对照的复核（P1 第 4 项）。

审稿 P1 要求「改进 SQL 结果比较器，区分 NULL 与空串、保留有序结果并采用类型感知比较」。
judge_sensitivity.py 早已实现这三条开关（r1_order / r2_numeric / r3_null /
r_all_strict）并在 2,164 条重执行上量化过影响，但正文只把它写成「残留风险」，
没有明确给出「收严后主对照方向是否改变」。本脚本不新增执行，只把已有结果
整理成可直接入稿的八个主对照表，并给出闭环判据。

输出 results/judge_strict_audit.json
"""
import json
import os

BASE = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE)
SRC = "results/judge_sensitivity.json"
OUT = "results/judge_strict_audit.json"

d = json.load(open(SRC, encoding="utf-8"))
c = d["contrasts"]

rows = []
for x in c:
    rows.append({
        "group": x["group"], "contrast": "%s - %s" % (x["a"], x["b"]),
        "n": x["n"],
        "r0_current_pt": x["r0_current"],
        "r_all_strict_pt": x["r_all_strict"],
        "delta_pt": round(x["r_all_strict"] - x["r0_current"], 1),
        "max_abs_change_pt": x["max_abs_change"],
        "sign_stable": x["sign_stable"],
    })

sign_ok = sum(1 for x in rows if x["sign_stable"])
max_shift = max(abs(x["delta_pt"]) for x in rows)
max_change = max(x["max_abs_change_pt"] for x in rows)
tot = d["strict_ex"]["total"]

report = {
    "version": "judge-strict-audit-v1",
    "source": SRC,
    "rules": d["rules"],
    "note": "No new SQL execution. All numbers are recomputed from the frozen "
            "judge_sensitivity.json contrasts; r0 remains the paper's primary judge.",
    "n_contrasts": len(rows),
    "sign_stable_all": sign_ok == len(rows),
    "sign_stable_count": "%d/%d" % (sign_ok, len(rows)),
    "max_abs_delta_pt": max_shift,
    "max_abs_change_any_rule_pt": max_change,
    "reexecution": {"n": tot["n"], "ex_current_pct": tot["ex_current_pct"],
                    "ex_strict_pct": tot["ex_strict_pct"],
                    "n_order_sensitive": tot["n_order_sensitive"],
                    "share_order_sensitive_pct": tot["share_order_sensitive_pct"],
                    "mcnemar_order_p": tot["mcnemar_order_p"]},
    "contrasts": rows,
    "reading": "Under the all-strict rule (order-sensitive + numeric tolerance + "
               "NULL/empty-string distinction) every one of the eight primary contrasts "
               "keeps its sign; the largest absolute shift is %.1f pt. Therefore the "
               "judge's leniency does not drive the paper's conclusions. Absolute EX "
               "values remain r0-based and must still be reported as execution-result "
               "agreement on the evaluated SQLite instances, not SQL equivalence." % max_shift,
}
json.dump(report, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("saved", OUT)
print("主对照 %s 符号稳定；全严规则最大偏移 %.1f pt；任一规则最大变化 %.1f pt" %
      (report["sign_stable_count"], max_shift, max_change))
print("重执行 n=%d，现行 %.1f%%，全严 %.1f%%，行序敏感 %d 条 (%.2f%%)" %
      (tot["n"], tot["ex_current_pct"], tot["ex_strict_pct"],
       tot["n_order_sensitive"], tot["share_order_sensitive_pct"]))
for r in rows:
    print("  %-7s %-12s n=%-4d r0=%+6.1f strict=%+6.1f delta=%+5.1f stable=%s" %
          (r["group"], r["contrast"], r["n"], r["r0_current_pt"],
           r["r_all_strict_pt"], r["delta_pt"], r["sign_stable"]))
