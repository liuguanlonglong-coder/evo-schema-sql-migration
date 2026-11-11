# -*- coding: utf-8 -*-
"""统计独立性审计：把「条目数 / op 数 / 独立演化事件数」三层口径算清。

审稿人指出的核心问题：
    「M1 与基线的比较存在输出能力不对称，且 15 个 op 的统计独立性需要修正。」

本脚本回答三件事，全部只读已有产物，不新增实验：

    A. 610 条到底由什么构成 —— 每个演化档的条目数 / 唯一 op 数 /
       独立演化事件数 / 涉及库数，四者并列，禁止只报其一。
    B. 段位档的 15 个 op 背后有几个独立事件 —— 一次合并会产生两个 op
       （两个旧列各对应一个段位），它们由同一个 gold 表达式配对生成，
       不是两个独立观测。按事件重算符号检验。
    C. 非段位档「M1+形态 91.8% < B3 100%」这个 8.2pt 缺口 ——
       12 个差异 op 落在几个事件上？若只落在少数事件，则该差距的事件级
       证据不足以支撑「M1 劣于 B3」的方法归属。

事件的定义（保守，可复核）：
    段位档  event = (db_id, table, 新列名)      —— 一次列合并/拆分作用于一列
    改名档  event = (db_id, table, evo_type)    —— 一次批量改名作用于一表
事件级成功判据：该事件下的全部 op 都 type_ok 才算成功（最严）。

输出 results/audit_independence.json
"""
import collections
import json
import os
import re
import sys

BASE = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE)
sys.path.insert(0, BASE)

from m1_lineage import gold_by_keys, evaluate  # noqa: E402

MANIFEST = "data/manifest_v3.jsonl"
LIN = os.path.join(BASE, "data", "lineage_infer")
OUT = os.path.join(BASE, "results", "audit_independence.json")
VERSION = "audit-independence-v2"

SEG_EVOS = ("merge_col", "split_col")
# 论文主实验口径：610 条 = 改名 276 + 合并 230 + 拆分 104
PAPER_EVOS = ("rename", "merge_col", "split_col")
# 全部演化档（含后续阶段的 value_shift 等）
ALL_EVOS = ("rename", "rename_medium", "rename_hard", "merge_col", "split_col",
            "value_shift", "value_shift_named", "dep_abstain", "dep_successor")


def _comb(n, k):
    r = 1
    for i in range(k):
        r = r * (n - i) // (i + 1)
    return r


def sign_test(a_only, b_only):
    """精确双侧符号检验，只在不一致对上算。"""
    n = a_only + b_only
    if n == 0:
        return 1.0, 0
    k = min(a_only, b_only)
    tail = sum(_comb(n, i) for i in range(k + 1)) / (2 ** n)
    return min(1.0, 2 * tail), n


def event_key(evo, db, table, old, new):
    """把 op 映射到独立演化事件。"""
    if evo in SEG_EVOS:
        m = re.search(r"(attr_\d+)", str(new))
        col = m.group(1) if m else str(new)
        # v2 修正：键必须含 evo。合并与拆分在同一 (库, 表) 上会产出同名新列 attr_001，
        # 旧键忽略 evo，把两类相反的编码改变并成同一事件（5 库中 3 库中招），
        # 段位档 10 次事件被塌缩为 7 次，610 集 61 次被塌缩为 58 次。
        return (evo, db, table.lower(), col)
    if evo.startswith("rename") or evo.startswith("value_shift"):
        return (db, table.lower(), evo)
    return (db, table.lower(), evo, old.lower())


def op_rows(rows, evo):
    """把 manifest 里某演化档的 op 展开成 (op键, 事件键) 列表。"""
    out = []
    for r in rows:
        if r["evo_type"] != evo:
            continue
        for op in r["ops"]:
            kind, table, old, new = op[0], op[1], op[2], op[3]
            k = (r["evo_type"], table.lower(), old.lower())
            ek = event_key(r["evo_type"], r["db_id"], table, old, new)
            out.append((k, ek))
    return out


def load_infer(suffix, keys):
    got = {}
    for evo, db in keys:
        p = os.path.join(LIN, f"{evo}_{db}_{suffix}.json")
        if os.path.exists(p):
            got[(evo, db)] = json.load(open(p, encoding="utf-8"))
    return got


def type_ok_map(rows, suffix):
    """{(evo, table, old): bool}，走与正文同一个 evaluate()。"""
    keys = sorted({(r["evo_type"], r["db_id"]) for r in rows})
    infer = load_infer(suffix, keys)
    _, detail = evaluate(rows, infer)
    out = {}
    for d in detail:
        preds = d["pred"]
        gold = d["gold"]
        ok = any(p["type"] == gold or
                 (gold in ("deprecate_with_successor", "rename_and_recode")
                  and p["type"] == "rename_column")
                 for p in preds)
        out[(d["evo"], d["table"].lower(), d["old"].lower())] = bool(ok)
    return out


def paper_set(rows):
    """610 条主实验集的真实成员：ids_xm_full（276）+ ids_ms（334）。"""
    idx = {r["id"]: r for r in rows}
    ids = []
    for f in ("data/ids_xm_full.txt", "data/ids_ms.txt"):
        if os.path.exists(f):
            ids += [l.strip() for l in open(f, encoding="utf-8") if l.strip()]
    return [idx[i] for i in ids if i in idx]


def section_a(rows):
    """A：三层口径并列表。"""
    tab = []
    for evo in ALL_EVOS:
        sub = [r for r in rows if r["evo_type"] == evo]
        if not sub:
            continue
        pairs = op_rows(rows, evo)
        ops = {k for k, _ in pairs}
        evs = {ek for _, ek in pairs}
        tab.append({
            "evo": evo,
            "items": len(sub),
            "ops": len(ops),
            "events": len(evs),
            "dbs": len({r["db_id"] for r in sub}),
        })

    # 610 条主实验集（论文正文实际口径）
    sub610 = paper_set(rows)
    pairs610 = []
    for r in sub610:
        for op in r["ops"]:
            k = (r["evo_type"], op[1].lower(), op[2].lower())
            ek = event_key(r["evo_type"], r["db_id"], op[1], op[2], op[3])
            pairs610.append((k, ek))
    ops610 = {k for k, _ in pairs610}
    ev610 = {ek for _, ek in pairs610}
    # 旧三元组键（缺 evo）的事件数，仅供正文说明「上一版报了多少」，不作推断
    ev610_legacy = set()
    for r in sub610:
        for op in r["ops"]:
            if r["evo_type"] in SEG_EVOS:
                m = re.search(r"(attr_\d+)", str(op[3]))
                ev610_legacy.add((r["db_id"], op[1].lower(),
                                  m.group(1) if m else str(op[3])))
            elif r["evo_type"].startswith("rename") or r["evo_type"].startswith("value_shift"):
                ev610_legacy.add((r["db_id"], op[1].lower(), r["evo_type"]))
            else:
                ev610_legacy.add((r["db_id"], op[1].lower(), r["evo_type"], op[2].lower()))
    seg610 = [(k, ek) for k, ek in pairs610 if k[0] in SEG_EVOS]
    tot = {
        "items": len(sub610),
        "ops": len(ops610),
        "events": len(ev610),
        "events_v13_triplet": len(ev610_legacy),
        "dbs": len({r["db_id"] for r in sub610}),
        "by_evo": dict(collections.Counter(r["evo_type"] for r in sub610)),
    }
    seg = {
        "items": sum(1 for r in sub610 if r["evo_type"] in SEG_EVOS),
        "ops": len({k for k, _ in seg610}),
        "events": len({ek for _, ek in seg610}),
        "dbs": len({r["db_id"] for r in sub610 if r["evo_type"] in SEG_EVOS}),
    }
    return {"by_evo": tab, "paper_set": tot, "segment_set": seg}


def section_b(rows, maps, suffixes):
    """B：段位档按 op 级与事件级分别重算，并测「输出空间是否含段位表达式」。"""
    pairs = []
    for evo in SEG_EVOS:
        pairs += op_rows(rows, evo)
    op2ev = {}
    for k, ek in pairs:
        op2ev[k] = ek

    res = {"n_ops": len(op2ev), "n_events": len(set(op2ev.values())), "methods": {}}
    for label, m in maps.items():
        ok = {k: m.get(k, False) for k in op2ev}
        ev_stat = collections.defaultdict(list)
        for k, ek in op2ev.items():
            ev_stat[ek].append(ok[k])
        ev_ok = {ek: all(v) for ek, v in ev_stat.items()}
        res["methods"][label] = {
            "op_ok": sum(ok.values()),
            "op_n": len(ok),
            "event_ok": sum(ev_ok.values()),
            "event_n": len(ev_ok),
            "ops_per_event": sorted({len(v) for v in ev_stat.values()}),
        }

    # 输出空间检测：预测的 cols/new 里是否出现过段位表达式（SUBSTR 等）
    space = {}
    keys = sorted({(r["evo_type"], r["db_id"]) for r in rows
                   if r["evo_type"] in SEG_EVOS})
    for label, suf in suffixes.items():
        infer = load_infer(suf, keys)
        n_expr = n_col = 0
        for v in infer.values():
            for op in v.get("ops", []):
                blob = " ".join([str(op.get("new", ""))] +
                               [str(c) for c in op.get("cols", [])])
                if "SUBSTR" in blob.upper() or "INSTR" in blob.upper():
                    n_expr += 1
                else:
                    n_col += 1
        space[label] = {"expr_outputs": n_expr, "plain_outputs": n_col,
                        "can_express_segment": n_expr > 0}
    res["output_space"] = space

    ev_pairs = collections.defaultdict(list)
    for k, ek in op2ev.items():
        ev_pairs[ek].append(k)
    res["paired"] = {}
    for label in ("B1 列名相似", "B2 值域重叠", "B3 嵌入语义"):
        ours = maps["M1 值域指纹"]
        base = maps[label]
        a_only = sum(1 for k in op2ev if ours.get(k) and not base.get(k))
        b_only = sum(1 for k in op2ev if base.get(k) and not ours.get(k))
        p_op, n_op = sign_test(a_only, b_only)
        ea = eb = 0
        for ek, ks in ev_pairs.items():
            o = all(ours.get(k) for k in ks)
            b = all(base.get(k) for k in ks)
            if o and not b:
                ea += 1
            elif b and not o:
                eb += 1
        p_ev, n_ev = sign_test(ea, eb)
        res["paired"][label] = {
            "op_level": {"ours_only": a_only, "base_only": b_only,
                         "n": n_op, "p": p_op},
            "event_level": {"ours_only": ea, "base_only": eb,
                            "n": n_ev, "p": p_ev},
        }
    return res


def section_c(rows, maps):
    """C：非段位档 M1+形态 vs B3 —— 必须按演化档拆开看。

    混合统计（147 op）会把 610 条主实验集之外的 value_shift 两档混进来，
    得到「M1 91.8% < B3 100%」这种归因错误的方法结论。
    """
    def sub_block(evos, title):
        pairs = []
        for evo in evos:
            pairs += op_rows(rows, evo)
        op2ev = {}
        for k, ek in pairs:
            op2ev[k] = ek
        ours = maps["M1+形态指纹"]
        base = maps["B3 嵌入语义"]
        d_ours = [k for k in op2ev if ours.get(k) and not base.get(k)]
        d_base = [k for k in op2ev if base.get(k) and not ours.get(k)]
        ev_base = {op2ev[k] for k in d_base}
        ev_ours = {op2ev[k] for k in d_ours}
        p_op, n_op = sign_test(len(d_ours), len(d_base))
        p_ev, n_ev = sign_test(len(ev_ours), len(ev_base))
        detail = collections.Counter()
        for k in d_base:
            detail[op2ev[k]] += 1
        return {
            "title": title,
            "evos": list(evos),
            "n_ops": len(op2ev),
            "n_events": len(set(op2ev.values())),
            "ours_ok": sum(ours.get(k, False) for k in op2ev),
            "base_ok": sum(base.get(k, False) for k in op2ev),
            "op_level": {"ours_only": len(d_ours), "base_only": len(d_base),
                         "n": n_op, "p": p_op},
            "event_level": {"ours_only": len(ev_ours), "base_only": len(ev_base),
                            "n": n_ev, "p": p_ev},
            "base_only_events": [{"event": list(e), "n_ops": n}
                                 for e, n in detail.most_common()],
        }

    in610 = ("rename", "rename_medium", "rename_hard", "dep_abstain", "dep_successor")
    shift = ("value_shift", "value_shift_named")
    return {
        "mixed_147": sub_block(in610 + shift, "混合口径（上一版正文所用，会误导归因）"),
        "no_shift_121": sub_block(in610, "不含值改写的非段位档"),
        "shift_26": sub_block(shift, "值改写两档（610 条之外）"),
    }


def _norm(s):
    return re.sub(r"\s+", "", str(s)).upper()


def _gold_seg(rows):
    """段位档的 gold 段位表达式：{(evo, db, table, old): 归一化表达式}。"""
    gold = {}
    for r in rows:
        if r["evo_type"] not in SEG_EVOS:
            continue
        for op in r["ops"]:
            if len(op) < 4 or not op[3]:
                continue
            gold[(r["evo_type"], r["db_id"], op[1].lower(), op[2].lower())] = _norm(op[3])
    return gold


def _seg_event(k, expr):
    """段位档事件键：(db, table, 新列)，新列从表达式里取 attr_xxx。"""
    m = re.search(r"(ATTR_\d+)", _norm(expr))
    # 同 event_key：键必须含 evo，否则合并与拆分的同名新列会被并成一件事
    return (k[0], k[1], k[2], m.group(1) if m else _norm(expr))


def section_d(rows):
    """D：段位表达式的真实恢复率——直接与 gold 表达式比字符串，不经 evaluate()。

    必须独立成节。原因：evaluate() 的 type_ok 在段位档上判的是 op 类型
    （merge_columns / split_column），而 gold 的 cols 经 expr_cols() 已归一为纯列名，
    **段位表达式本身从未进入 evaluate() 的比较**。上一版把 type_ok 15/15 称为
    「段位恢复率」是口径错误，这里重算真正的段位一致率。
    """
    gold = _gold_seg(rows)
    keys = sorted({(r["evo_type"], r["db_id"]) for r in rows
                   if r["evo_type"] in SEG_EVOS})

    spaces = {}
    for suf in ("name", "vover", "emb", "profile", "shape", "part"):
        infer = load_infer(suf, keys)
        seen = collections.Counter()
        n_expr = 0
        for v in infer.values():
            for op in v.get("ops", []):
                seen[op.get("type")] += 1
                if re.search(r"SUBSTR|INSTR", str(op.get("new", "")), re.I):
                    n_expr += 1
        spaces[suf] = {"types": dict(seen), "expr_outputs": n_expr,
                       "n_files": len(infer)}

    part = load_infer("part", keys)
    okmap = {}
    for (evo, db), v in part.items():
        for op in v.get("ops", []):
            k = (evo, db, op["table"].lower(), op["old"].lower())
            if k in gold:
                okmap[k] = _norm(op.get("new", "")) == gold[k]

    hit = sum(1 for v in okmap.values() if v)
    covered = set(okmap)
    ev_ops = collections.defaultdict(list)
    for k, expr in gold.items():
        ev_ops[_seg_event(k, expr)].append(k)
    ev_covered = [ek for ek, ks in ev_ops.items() if all(k in okmap for k in ks)]
    ev_ok = [ek for ek in ev_covered if all(okmap[k] for k in ev_ops[ek])]
    p_ev, n_ev = sign_test(len(ev_ok), 0)

    return {
        "gold_segment_ops": len(gold),
        "gold_segment_events": len(ev_ops),
        "part_covered_ops": len(covered),
        "part_expr_match": hit,
        "part_expr_pct": round(100.0 * hit / max(1, len(covered)), 1),
        "uncovered_ops": sorted(list(k) for k in gold if k not in covered),
        "events_covered": len(ev_covered),
        "events_ok": len(ev_ok),
        "event_level_p": p_ev,
        "event_level_n": n_ev,
        "output_spaces": spaces,
    }


def main():
    rows = [json.loads(l) for l in open(MANIFEST, encoding="utf-8")]
    sufs = {"B1 列名相似": "name", "B2 值域重叠": "vover", "B3 嵌入语义": "emb",
            "M1 值域指纹": "profile", "M1+形态指纹": "shape"}
    maps = {k: type_ok_map(rows, v) for k, v in sufs.items()}
    out = {
        "version": VERSION,
        "manifest": MANIFEST,
        "note": ("三层口径并列：条目数(items) / 唯一 op 数(ops) / 独立演化事件数(events)。"
                 "段位档事件 = (演化档, db, table, 新列)；改名档事件 = (db, table, 演化档)。"
                 "v2 修正：段位档事件键补入 evo，此前缺 evo 会把合并与拆分的同名新列并为一事件。"
                 "事件级成功 = 该事件全部 op 都对（最严）。"
                 "paper_set 取 ids_xm_full.txt(276) + ids_ms.txt(334) = 610 条。"
                 "D 节独立于 evaluate()，直接比段位表达式字符串。"),
        "A_units": section_a(rows),
        "B_segment": section_b(rows, maps, sufs),
        "C_nonsegment": section_c(rows, maps),
        "D_segment_expr": section_d(rows),
    }
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    json.dump(out, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)

    a = out["A_units"]
    print("=== A 三层口径（全量 manifest）===")
    print(f"{'演化档':<22}{'条目':>6}{'唯一op':>8}{'独立事件':>9}{'库':>4}")
    for t in a["by_evo"]:
        print(f"{t['evo']:<22}{t['items']:>6}{t['ops']:>8}{t['events']:>9}{t['dbs']:>4}")
    ps, sg = a["paper_set"], a["segment_set"]
    print(f"\n  610 条主实验集 = {ps['items']} 条 / {ps['ops']} 唯一 op / "
          f"{ps['events']} 独立事件 / {ps['dbs']} 库")
    print(f"    构成 {ps['by_evo']}")
    print(f"  其中段位档   = {sg['items']} 条 / {sg['ops']} op / "
          f"{sg['events']} 事件 / {sg['dbs']} 库")

    b = out["B_segment"]
    print(f"\n=== B 段位档（evaluate 口径）：{b['n_ops']} op -> {b['n_events']} 事件 ===")
    for label, s in b["methods"].items():
        print(f"  {label:<14} op {s['op_ok']:>2}/{s['op_n']:<3} "
              f"事件 {s['event_ok']:>2}/{s['event_n']:<3} "
              f"每事件op数 {s['ops_per_event']}")
    print("  配对（M1 vs 基线）:")
    for label, s in b["paired"].items():
        o, e = s["op_level"], s["event_level"]
        print(f"    vs {label:<12} op级 {o['ours_only']}:{o['base_only']} "
              f"n={o['n']:<3} p={o['p']:.2e} | "
              f"事件级 {e['ours_only']}:{e['base_only']} n={e['n']} p={e['p']:.4f}")

    print("\n=== D 段位表达式真实恢复（不经 evaluate）===")
    d = out["D_segment_expr"]
    print(f"  gold 段位 op {d['gold_segment_ops']} / 事件 {d['gold_segment_events']}")
    print(f"  part 覆盖 op {d['part_covered_ops']}，表达式一致 "
          f"{d['part_expr_match']}/{d['part_covered_ops']} ({d['part_expr_pct']}%)")
    print(f"  事件级 {d['events_ok']}/{d['events_covered']} 覆盖，"
          f"符号检验 n={d['event_level_n']} p={d['event_level_p']:.4f}")
    if d["uncovered_ops"]:
        print(f"  未覆盖 op {len(d['uncovered_ops'])}: "
              f"{[u[0] for u in d['uncovered_ops']]}")
    print("  各方法在段位档的输出类型空间：")
    for suf, s in d["output_spaces"].items():
        print(f"    {suf:<8} 文件{s['n_files']:>2} 段位表达式输出 {s['expr_outputs']:>3} "
              f"类型 {s['types']}")

    print("\n=== C 非段位档：必须拆开看 ===")
    for key in ("mixed_147", "no_shift_121", "shift_26"):
        c = out["C_nonsegment"][key]
        o, e = c["op_level"], c["event_level"]
        print(f"  [{c['title']}] {c['n_ops']} op / {c['n_events']} 事件")
        print(f"    M1+形态 {c['ours_ok']}/{c['n_ops']} "
              f"({100.0*c['ours_ok']/max(1,c['n_ops']):.1f}%) | "
              f"B3 {c['base_ok']}/{c['n_ops']} "
              f"({100.0*c['base_ok']/max(1,c['n_ops']):.1f}%)")
        print(f"    配对 op级 {o['ours_only']}:{o['base_only']} n={o['n']} p={o['p']:.3g} | "
              f"事件级 {e['ours_only']}:{e['base_only']} n={e['n']} p={e['p']:.4f}")
    print(f"\n-> {OUT}")


if __name__ == "__main__":
    main()
