# -*- coding: utf-8 -*-
"""audit_units_v3.py - 四层口径统一（回应审稿意见 MO-6 / 表 1 第 1 条）。

问题：主文、补充材料、摘要、修订说明里同时存在 610 / 437 / 90 / 61 / 11、
992 / 527、28 / 22 / 18、25 / 19 / 18 等数字，读者无法判断每个数字是哪一层。
本脚本把每一层的定义写死，并对每个数据集只按定义重算，不读正文常量。

五层定义（自下而上）：
  items    逐题记录数（一条 SQL 算一条）
  base     基础题数 = distinct (db, question_id)
  ops      原子操作数 = distinct (evo, db, table, 旧列)；一次合并作用两列即为 2 个 op
  events   独立演化事件数 = (evo, db, table, 新列) 或 (db, table, evo)
  dbs      库数

事件在扩充集上再分四档（主文必须同时给出，不能只报其一）：
  constructed  构造出来的事件数（build_evolution_ext 写入 per_event 的条数）
  evaluated    在评测输出里出现过的事件数（有题进入四步验证）
  effective    ≥1 条题通过四步验证
  eligible     ≥3 条题通过四步验证（事件级检验的预设门槛）
  tested       实际进入符号检验的事件数（存在至少一对不一致判定）

另出运行清单（run inventory）：表号 → 结果文件 → 模型 → 条数，用于解释
DeepSeek-V3.2 与 DeepSeek-V4.1-Flash 并存不是笔误而是不同阶段。
输出 results/audit_units_v3.json
"""
import collections
import glob
import json
import os
import re
import sys

BASE = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE)

OUT = "results/audit_units_v3.json"
VERSION = "audit-units-v3"
SEG_EVOS = ("merge_col", "split_col", "merge_col_holdout", "split_col_holdout")


# ------------------------------------------------------------------ 事件键
def event_key(evo, db, table, old, new):
    """与 audit_independence.py v2 完全一致：段位档键必须含 evo 与新列名。"""
    if evo in SEG_EVOS:
        m = re.search(r"(attr_\d+)", str(new))
        col = m.group(1) if m else str(new)
        return (evo, db, (table or "").lower(), col)
    # 非段位档与 audit_independence.py v2 逐字一致：改名/值改写一次作用于整表，
    # 不区分旧列（否则一次批量改名会被拆成十几个伪事件，61 -> 85）；
    # 依赖两档一次只改一列，键里必须保留旧列名（否则 23 -> 15）。
    if evo.startswith("rename") or evo.startswith("value_shift"):
        return (db, (table or "").lower(), evo)
    return (db, (table or "").lower(), evo, str(old).lower())


def op_key(evo, db, table, old):
    return (evo, db, (table or "").lower(), str(old).lower())


# ------------------------------------------------------------------ 主清单
def load_manifest(path, evos=None):
    rows = []
    for line in open(path, encoding="utf-8"):
        line = line.strip()
        if not line:
            continue
        r = json.loads(line)
        if evos and r.get("evo_type") not in evos:
            continue
        rows.append(r)
    return rows


def unit_stats(rows, id_filter=None):
    if id_filter is not None:
        rows = [r for r in rows if r["id"] in id_filter]
    items = len(rows)
    base = len({(r["db_id"], str(r["question_id"])) for r in rows})
    ops, evs, dbs = set(), set(), set()
    for r in rows:
        evo, db = r["evo_type"], r["db_id"]
        dbs.add(db)
        # 扩充清单：ops 只是代表 op，完整列表在 ops_all（第 22 轮补齐）
        ops_list = r.get("ops_all") or r.get("ops") or []
        if isinstance(ops_list, str):
            try:
                ops_list = json.loads(ops_list)
            except Exception:
                ops_list = []
        if not ops_list:
            # 无 ops 字段的记录退化为「(库, 题) 一个操作」，并在 note 里标出
            ops.add(op_key(evo, db, r.get("table"), r["question_id"]))
            evs.add(event_key(evo, db, r.get("table"), r["question_id"],
                              r["question_id"]))
            continue
        for o in ops_list:
            # o = [kind, table, old_col, new_expr]
            if len(o) < 4:
                continue
            ops.add(op_key(evo, db, o[1], o[2]))
            evs.add(event_key(evo, db, o[1], o[2], o[3]))
    return {"items": items, "base": base, "ops": len(ops),
            "events": len(evs), "dbs": len(dbs)}


# ------------------------------------------------------------------ 扩充集
def ext_stats(name):
    evo = json.load(open(f"results/evolution_ext_{name}.json", encoding="utf-8"))
    per = evo["per_event"]
    man = load_manifest(f"data/manifest_ext_{name}.jsonl")
    items = len(man)
    base = len({(r["db_id"], str(r["question_id"])) for r in man})
    # op 数按定义重算：一次合并/拆分作用到几列就有几个 op
    ops = sum(len(e.get("cols") or []) for e in per)
    evs = {e["tag"] for e in per}
    dbs = {e["db"] for e in per}
    evaluated = effective = eligible = 0
    tested = 0
    evj = None
    fp = f"results/ext_eval_{name}.json"
    rows = None
    if os.path.exists(f"results/ext_eval_{name}.jsonl"):
        rows = [json.loads(l) for l in open(f"results/ext_eval_{name}.jsonl",
                                            encoding="utf-8") if l.strip()]
    if rows:
        tag_items = collections.Counter(r["tag"] for r in rows)
        evaluated = len(tag_items)
        # 有效 / 合格按「通过四步验证的题数」计，与 ext_summary.py 一致
        # 逐条 jsonl 只记生成结果，不记四步验证；通过数从构造期 per_event 读取
        effective = sum(1 for e in per if (e.get("n_pass") or 0) >= 1)
        eligible = sum(1 for e in per if (e.get("n_pass") or 0) >= 3)
    if os.path.exists(fp):
        evj = json.load(open(fp, encoding="utf-8"))
        pb = evj["paired_by_event"].get("c3mp-c3m")
        tested = pb["n_events"] if pb else 0
    return {"name": name, "items": items, "base": base, "ops": ops,
            "events": len(evs), "dbs": len(dbs),
            "tables": len({e["table"] for e in per}),
            "constructed": len(per), "evaluated": evaluated,
            "effective": effective, "eligible": eligible, "tested": tested,
            "items_per_event": round(items / max(1, len(evs)), 1),
            "max_event_share_pct": round(
                100.0 * max(len(e.get("cols") or []) for e in per) /
                max(1, items), 2),
            "note": "ops 按 (evo,库,表,旧列) 重算；manifest 的 ops 字段只登记"
                    "了一个代表 op（构造时 applied 的首列），故不与本表相等。"}


# ------------------------------------------------------------------ 运行清单
def run_inventory():
    out = []
    for f in sorted(glob.glob("results/*.jsonl")):
        ms, n = collections.Counter(), 0
        for line in open(f, encoding="utf-8"):
            line = line.strip()
            if not line:
                continue
            try:
                r = json.loads(line)
            except Exception:
                continue
            n += 1
            m = r.get("model") or r.get("model_id")
            if m:
                ms[m] += 1
        if ms:
            out.append({"file": os.path.basename(f), "rows": n,
                        "model": "/".join(sorted(ms)),
                        "mtime": __import__("time").strftime(
                            "%Y-%m-%d", __import__("time").localtime(
                                os.path.getmtime(f)))})
    return out


# ------------------------------------------------------------------ 留出集
def holdout_stats():
    rows = []
    for f in sorted(glob.glob("results/e1_holdout_*.jsonl")):
        rows += [json.loads(l) for l in open(f, encoding="utf-8") if l.strip()]
    ids = {r["id"] for r in rows}
    conds = collections.Counter(r["cond"] for r in rows)
    dbs = {r["id"].split(":")[1] for r in rows if r.get("id")}
    man = load_manifest("data/manifest_holdout.jsonl")
    held = [m for m in man if m["id"] in ids]
    hs = unit_stats(held)
    return {"items": len(ids), "rows": len(rows), "conds": dict(conds),
            "dbs": len(dbs), "db_list": sorted(dbs), "units": hs,
            "manifest_holdout_items": len(man)}


def main():
    man_all = load_manifest("data/manifest_v3.jsonl")
    by_evo = collections.defaultdict(list)
    for r in man_all:
        by_evo[r["evo_type"]].append(r)

    paper_ids = set()
    for f in ("data/ids_xm_full.txt", "data/ids_ms.txt"):
        paper_ids |= {l.strip() for l in open(f, encoding="utf-8") if l.strip()}
    paper = [r for r in man_all if r["id"] in paper_ids]
    seg = [r for r in paper if r["evo_type"] in ("merge_col", "split_col")]

    res = {
        "version": VERSION,
        "definitions": {
            "items": "逐题记录数",
            "base": "基础题数 = distinct (库, 题号)",
            "ops": "原子操作数 = distinct (演化档, 库, 表, 旧列)",
            "events": "独立演化事件数 = (演化档, 库, 表, 新列)；非段位档为 (库, 表, 演化档)",
            "dbs": "库数",
            "event_funnel": "constructed ≥ evaluated ≥ effective ≥ eligible ≥ tested",
        },
        "main": {},
        "ext": {},
        "holdout": holdout_stats(),
        "runs": run_inventory(),
    }
    for evo, rows in sorted(by_evo.items()):
        res["main"][evo] = unit_stats(rows)
    res["main"]["paper_set"] = unit_stats(paper)
    res["main"]["segment_set"] = unit_stats(seg)
    res["main"]["paper_set"]["_note"] = \
        "610 条 = ids_xm_full(276) + ids_ms(334)；base 437 为 distinct (库,题号)"

    for name in ("merge_col", "split_col", "value_shift"):
        if os.path.exists(f"results/evolution_ext_{name}.json"):
            res["ext"][name] = ext_stats(name)

    json.dump(res, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)

    print(f"{'set':22s} {'items':>6s} {'base':>6s} {'ops':>5s} {'events':>7s} {'dbs':>4s}")
    for k, v in res["main"].items():
        print(f"{k:22s} {v['items']:6d} {v['base']:6d} {v['ops']:5d} "
              f"{v['events']:7d} {v['dbs']:4d}")
    print()
    for k, v in res["ext"].items():
        print(f"{k:22s} {v['items']:6d} {v['base']:6d} {v['ops']:5d} "
              f"{v['events']:7d} {v['dbs']:4d}  "
              f"构造{v['constructed']}/评测{v['evaluated']}/有效{v['effective']}"
              f"/合格{v['eligible']}/检验{v['tested']}")
    print()
    h = res["holdout"]
    hu = h["units"]
    print(f"holdout: {h['items']} 条 / {hu['base']} 基础题 / {hu['ops']} op / "
          f"{hu['events']} 事件 / {h['dbs']} 库  {h['conds']}")
    print("写入", OUT)


if __name__ == "__main__":
    main()
