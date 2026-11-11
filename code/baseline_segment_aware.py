# -*- coding: utf-8 -*-
"""段位感知的非 LLM 公平基线（B4）。

回应审稿意见 2.3 / 最终意见 (3)：B1–B3 的输出类型空间只有 rename/deprecate，
不含段位表达式，因此 0/15 只能证明「表达能力不对称」，不能回答
「在具备同等输出能力的算法面前，M1 是否仍有真实算法增益」。

设计（与 M1 对齐的部分刻意复用同一实现，保证公平）：
  * 复用 m1_part 的 detect_sep / jac / col_values / read_schema / to_expr
    —— 分隔符检测、值域 Jaccard、SUBSTR 表达式渲染完全一致；
  * 去掉 M1 的三处专属设计：
      ① SPLIT_THRESH = 0.55 的门控（低于阈值不产出段位）；
      ② 低基数列（|V| < 3）把阈值抬到 0.95 的守卫；
      ③ MIN_ROWS / 二次采样等数据门限；
  * 用最朴素的 argmax：每个被删旧列直接分给 Jaccard 最大的段位，
    分数再低也产出表达式；不做名字 tie-break，不读任何标注或 gold。

因此 B4 与 M1 的差别只剩下「是否使用阈值/守卫」，而输出空间完全相同，
这才是对报告所要求的公平比较。另报一个「贪心冲突消解」变体作为上界。

评估口径：与第 13 轮 D_segment_expr 完全相同——直接与 gold 段位表达式
做归一化后的逐字符比对，不经 evaluate() 的 type_ok。

输出：results/baseline_segment_aware.json
"""
import collections
import json
import os
import re
import sys

BASE = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE)
sys.path.insert(0, BASE)

import m1_part as mp  # noqa: E402  复用同一套检测/渲染，保证输出能力对齐

VERSION = "segment-aware-baseline-v1"
OUT = "results/baseline_segment_aware.json"
MANIFEST = "data/manifest_v3.jsonl"
OLD_ROOT = os.environ.get("OLD_DBDIR") or os.path.join(mp.BIRD, "dev_databases")
NEW_DIR = os.path.join("data", "evolved_merge")

# 段位档只来自这两个演化类型；拆分档 gold 是拼接表达式而非 SUBSTR，另列说明
SEG_EVOS = ("merge_col", "split_col")


def _norm(s):
    return re.sub(r"\s+", "", str(s)).upper()


def gold_ops(rows):
    """{(evo, db, table, old): 归一化 gold 表达式}，只取有段位/拼接表达式的 op。"""
    gold = {}
    for r in rows:
        if r["evo_type"] not in SEG_EVOS:
            continue
        for op in r["ops"]:
            if len(op) < 4 or not op[3]:
                continue
            k = (r["evo_type"], r["db_id"], op[1].lower(), op[2].lower())
            gold[k] = _norm(op[3])
    return gold


def infer_naive(db_id):
    """B4：无阈值、无守卫、纯 argmax 的段位分配。

    返回 {table: {attr_col: {"sep", "nseg", "assign": {old: (idx, score)}}}}
    """
    old_p = os.path.join(OLD_ROOT, db_id, f"{db_id}.sqlite")
    new_p = os.path.join(NEW_DIR, f"{db_id}.sqlite")
    if not (os.path.exists(old_p) and os.path.exists(new_p)):
        return None, "missing_db"
    old_s, new_s = mp.read_schema(old_p), mp.read_schema(new_p)
    result = {}
    for tab, new_cols in new_s.items():
        if tab not in old_s:
            continue
        old_cols = set(old_s[tab])
        merged = [c for c in new_cols if c.startswith("attr_") and c not in old_cols]
        dropped = [c for c in old_cols if c not in set(new_cols)]
        if not merged or not dropped:
            continue
        for mc in merged:
            vals = mp.col_values(new_p, tab, mc, limit=20000)
            sep = mp.detect_sep(vals)
            if not sep:
                continue
            segs = [v.split(sep) for v in vals]
            nseg = max(len(s) for s in segs) if segs else 0
            if nseg < 2:
                continue
            per_idx = collections.defaultdict(set)
            for s in segs:
                for i, piece in enumerate(s):
                    if piece != "":
                        per_idx[i].add(piece)
            # 纯 argmax：不看阈值、不看基数、不做冲突消解
            assign = {}
            for oc in dropped:
                ov = mp.col_values(old_p, tab, oc)
                scores = [(mp.jac(ov, per_idx[i]), i) for i in range(nseg)]
                scores.sort(key=lambda x: (-x[0], x[1]))
                if scores:
                    assign[oc] = (scores[0][1], round(scores[0][0], 3))
            if assign:
                result.setdefault(tab, {})[mc] = {
                    "sep": sep, "nseg": nseg, "assign": assign,
                    "n_dropped": len(dropped),
                }
    return result, None


def greedy_resolve(entry):
    """上界变体：同段位冲突时按 score 贪心，一个段位只留给得分最高的旧列。"""
    taken, out = set(), {}
    pairs = sorted(entry["assign"].items(), key=lambda kv: (-kv[1][1], kv[0]))
    for oc, (idx, sc) in pairs:
        if idx in taken:
            continue
        taken.add(idx)
        out[oc] = (idx, sc)
    return out


def render(attr, idx, nseg, sep):
    return _norm(mp.to_expr(attr, idx, nseg, sep))


def main():
    rows = [json.loads(l) for l in open(MANIFEST, encoding="utf-8")]
    gold = gold_ops(rows)
    dbs = sorted({k[1] for k in gold})

    naive_rows, greedy_rows = [], []
    per_db = {}
    errors = []
    for db in dbs:
        got, err = infer_naive(db)
        if got is None:
            errors.append({"db": db, "error": err})
            continue
        per_db[db] = got
        for (evo, g_db, tab, old), gexpr in gold.items():
            if g_db != db or evo != "merge_col":
                continue  # 拆分档是拼接表达式，SUBSTR 段位基线不适用，另计
            # 定位该 old 列在该库哪张表的哪个 attr 下。
            # gold 的 old 列名为小写，而库里是原始大小写（County/A2/SEX）；
            # 比对必须做大小写归一化，否则会把「检测成功」误报成「未产出段位」。
            # 这是评测侧的归一化，B4 本身不读 gold。
            hit = None
            for t, cols in got.items():
                for attr, e in cols.items():
                    low = {k.lower(): v for k, v in e["assign"].items()}
                    if old in low:
                        hit = (t, attr, e, low[old])
                        break
                if hit:
                    break
            if hit is None:
                naive_rows.append((db, tab, old, None, gexpr, False, None))
                greedy_rows.append((db, tab, old, None, gexpr, False, None))
                continue
            t, attr, e, (idx, sc) = hit
            expr = render(attr, idx, e["nseg"], e["sep"])
            naive_rows.append((db, tab, old, expr, gexpr, expr == gexpr, sc))
            gmap = {k.lower(): v for k, v in greedy_resolve(e).items()}
            if old in gmap:
                gidx, gsc = gmap[old]
                gexpr2 = render(attr, gidx, e["nseg"], e["sep"])
                greedy_rows.append((db, tab, old, gexpr2, gexpr, gexpr2 == gexpr, gsc))
            else:
                greedy_rows.append((db, tab, old, None, gexpr, False, None))

    def summarize(rw):
        n = len(rw)
        ok = sum(1 for x in rw if x[5])
        # 事件级：按 (db, table, 新列) 聚类，事件内全部 op 对才算成功
        ev = collections.defaultdict(list)
        for db, tab, old, expr, gexpr, good, sc in rw:
            m = re.search(r"(ATTR_\d+)", gexpr or "")
            ev[(db, tab, m.group(1) if m else "")].append(good)
        ev_ok = sum(1 for v in ev.values() if all(v))
        return {"n_ops": n, "expr_match": ok, "pct": round(100.0 * ok / max(1, n), 1),
                "n_events": len(ev), "events_ok": ev_ok,
                "detail": [{"db": x[0], "table": x[1], "old": x[2],
                            "pred": x[3], "gold": x[4], "match": x[5], "score": x[6]}
                           for x in rw]}

    # M1 与三条基线的同口径读数（来自第 13 轮审计产物，供并列）
    adi = json.load(open("results/audit_independence.json", encoding="utf-8"))
    m1 = adi["D_segment_expr"]
    b_type = adi["B_segment"]["methods"]

    out = {
        "version": VERSION,
        "note": ("B4 与 M1 共用 detect_sep/jac/to_expr，输出空间同为 SUBSTR 段位表达式；"
                 "差别仅在 M1 的阈值门控、低基数守卫与采样门限。"
                 "评估与 D_segment_expr 同口径：归一化后逐字符比对 gold，不经 evaluate()。"
                 "拆分档 gold 为拼接表达式(attr_001||sep||attr_002)，非段位 SUBSTR，"
                 "不计入本基线的分母，单独说明。"),
        "B4_naive": summarize(naive_rows),
        "B4_greedy": summarize(greedy_rows),
        "M1_part": {"n_ops": m1["part_covered_ops"], "expr_match": m1["part_expr_match"],
                    "pct": m1["part_expr_pct"], "n_events": m1["events_covered"],
                    "events_ok": m1["events_ok"]},
        "B1_B3_type_ok_segment": {
            k: {"op_ok": v["op_ok"], "op_n": v["op_n"], "event_ok": v["event_ok"],
                "event_n": v["event_n"]}
            for k, v in b_type.items() if k.startswith("B")
        },
        "split_ops_excluded": sorted(
            [{"db": k[1], "table": k[2], "old": k[3], "gold": v}
             for k, v in gold.items() if k[0] == "split_col"],
            key=lambda x: (x["db"], x["old"])),
        "dbs": dbs,
        "errors": errors,
    }
    json.dump(out, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)

    print("=== B4 段位感知公平基线 ===")
    for key in ("B4_naive", "B4_greedy"):
        s = out[key]
        print(f"  {key:<10} 表达式 {s['expr_match']}/{s['n_ops']} ({s['pct']}%) "
              f"事件 {s['events_ok']}/{s['n_events']}")
    m = out["M1_part"]
    print(f"  M1_part   表达式 {m['expr_match']}/{m['n_ops']} ({m['pct']}%) "
          f"事件 {m['events_ok']}/{m['n_events']}")
    print("  B1–B3 段位档 type_ok:", {k: f"{v['op_ok']}/{v['op_n']}"
                                    for k, v in out["B1_B3_type_ok_segment"].items()})
    print("  拆分档排除:", len(out["split_ops_excluded"]), "个 op（拼接表达式，非段位）")
    print("  B4_naive 逐 op:")
    for d in out["B4_naive"]["detail"]:
        print(f"    {d['db']:<24}{d['old']:<14} score={d['score']} match={d['match']}")
    if errors:
        print("  错误:", errors)
    print("->", OUT)


if __name__ == "__main__":
    main()
