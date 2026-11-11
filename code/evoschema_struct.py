# -*- coding: utf-8 -*-
"""EvoSchema 公开列级样本的结构性统计（不涉及模型调用）。

只回答四个事实问题，供第 7 章 7.9 与 7.9.1 引用：
1. 样本按变异类型与库的分布；
2. column_changes 字段的可用性；
3. 新旧 gold SQL 是否使用段位切分（SUBSTR/INSTR）来恢复被合并列的原值；
4. 合并档的投影列数与来源列数之比（用于说明「合并后整串返回」）。

输出 results/evoschema_struct.json。
"""
import json
import os
import re
from collections import Counter

BASE = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE)

SRC = "data/evoschema/bird_split_or_merge_columns.json"
OUT = "results/evoschema_struct.json"

DDL_RE = re.compile(r"create\s+table\s+(\w+)\s*\((.*)\)\s*$", re.I | re.S)


def cols_from_ddl(ddl):
    """从 BIRD 风格的 DDL 里取列名。

    以顶层逗号切分（括号内不切），每段取第一个词作为列名，
    并剔除 primary key / foreign key / constraint 这三类约束段。
    """
    m = DDL_RE.search(ddl.strip())
    body = m.group(2) if m else ddl
    parts, depth, buf = [], 0, []
    for ch in body:
        if ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
        if ch == "," and depth == 0:
            parts.append("".join(buf))
            buf = []
        else:
            buf.append(ch)
    parts.append("".join(buf))
    out = []
    for p in parts:
        p = p.strip()
        if not p:
            continue
        low = p.lower()
        if low.startswith(("primary key", "foreign key", "constraint", "unique")):
            continue
        out.append(p.split()[0])
    return out


def table_cols(tbl):
    d = {}
    for name, v in (tbl or {}).items():
        d[name] = cols_from_ddl(v.get("ddl", ""))
    return d


def seg_calls(sql):
    """返回 SQL 中 SUBSTR/INSTR 的参数文本列表。"""
    out = []
    for fn in ("substr", "instr"):
        for m in re.finditer(fn + r"\s*\(", sql, re.I):
            i, depth = m.end(), 1
            while i < len(sql) and depth:
                if sql[i] == "(":
                    depth += 1
                elif sql[i] == ")":
                    depth -= 1
                i += 1
            out.append(sql[m.end():i - 1])
    return out


def main():
    items = json.load(open(SRC, encoding="utf-8"))
    uniq, seen = [], set()
    for it in items:
        k = (it["db_id"], it["question"].strip())
        if k in seen:
            continue
        seen.add(k)
        uniq.append(it)

    rec = []
    for it in uniq:
        old_cols = table_cols(it.get("relevant_table"))
        new_cols = table_cols(it.get("new_relevant_table"))
        added = {c for t in new_cols for c in new_cols[t]
                 if c not in {x.lower() for x in old_cols.get(t, [])}}
        removed = {c for t in old_cols for c in old_cols[t]
                   if c.lower() not in {x.lower() for x in new_cols.get(t, [])}}
        cc = it.get("column_changes") or {}
        # 合并目标列：新出现且不在旧列集合里的列
        targets = set(cc.keys()) if cc else set(added)
        # 合并目标列是否被段位切分函数引用
        new_sql = it.get("new_gold_sql") or ""
        used_seg = [a for a in seg_calls(new_sql)
                    if any(t.lower() in a.lower() for t in targets)]
        rec.append({
            "db": it["db_id"],
            "kind": it["split_or_merge"],
            "has_cc": bool(cc),
            "n_added": len(added),
            "n_removed": len(removed),
            "targets": sorted(targets),
            "new_sql_seg_any": bool(seg_calls(new_sql)),
            "seg_on_target": len(used_seg),
        })

    m = [r for r in rec if r["kind"] == "merge"]
    s = [r for r in rec if r["kind"] == "split"]
    out = {
        "version": "evoschema-struct-v1",
        "source": SRC,
        "n_raw": len(items),
        "n_uniq": len(uniq),
        "by_kind": {"merge": len(m), "split": len(s)},
        "by_kind_db": {
            "merge": dict(Counter(r["db"] for r in m)),
            "split": dict(Counter(r["db"] for r in s)),
        },
        "column_changes_empty": {
            "merge": sum(1 for r in m if not r["has_cc"]),
            "split": sum(1 for r in s if not r["has_cc"]),
        },
        "seg_on_target": {
            "merge": sum(1 for r in m if r["seg_on_target"]),
            "split": sum(1 for r in s if r["seg_on_target"]),
        },
        "seg_any_in_new_sql": {
            "merge": sum(1 for r in m if r["new_sql_seg_any"]),
            "split": sum(1 for r in s if r["new_sql_seg_any"]),
        },
        "merge_targets": dict(Counter(t for r in m for t in r["targets"])),
        "rows": rec,
    }
    os.makedirs("results", exist_ok=True)
    json.dump(out, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print("saved:", OUT)
    print("原始", out["n_raw"], "去重", out["n_uniq"],
          "合并", len(m), "拆分", len(s))
    print("column_changes 为空：merge", out["column_changes_empty"]["merge"],
          "split", out["column_changes_empty"]["split"])
    print("对合并列做段位切分：merge", out["seg_on_target"]["merge"],
          "split", out["seg_on_target"]["split"])
    print("新 gold SQL 出现任一 SUBSTR/INSTR：merge",
          out["seg_any_in_new_sql"]["merge"], "split",
          out["seg_any_in_new_sql"]["split"])
    print("合并目标列：", out["merge_targets"])


if __name__ == "__main__":
    main()
