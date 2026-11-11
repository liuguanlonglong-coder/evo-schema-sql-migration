# -*- coding: utf-8 -*-
"""扩充独立 Schema 演化事件：从「每库 1 个事件」改为「每（库, 表）1 个事件」。

为什么必须改构造而不是改统计
  build_evolution_v2.py 的 run() 对每个库只取 pick_merge_pair / pick_split_col 的
  top-1 候选，因此合并档只有 5 个独立事件、拆分档 5 个。610 条题共享这 10 个事件，
  事件级检验的 n 就只能是 10 —— 这是构造决定的，任何统计方法都无法改变。
  本脚本按表分片：每个候选表各复制一份演化库，各自施加一次 DDL 变更。
  好处是三件事同时成立：
    ① 两个片段之间不共享任何被改列 → 事件独立性没有争议；
    ② 每个事件有自己的题集 → 不存在「一套题重复充数」；
    ③ 沿用 v2 的四步验证（原库可执行 / 旧 SQL 在新库失效 / 新 SQL 可执行 /
       结果等价），因此新增事件的构造质量与主实验同标准。

与主实验的关系
  写独立目录 data/evolved_ext/ 与独立清单 data/manifest_ext.jsonl，
  不触碰已冻结的 data/evolved_merge、data/evolved_split 与 manifest_v3.jsonl。
  id 前缀 `ext:` 避免与主实验 id 冲突。

用法
  python build_evolution_ext.py --task merge_col
  python build_evolution_ext.py --task split_col --max-mb 130
"""
import argparse
import collections
import filecmp
import importlib.util
import json
import os
import re
import shutil
import time

BASE = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE)


def _load_v2():
    spec = importlib.util.spec_from_file_location(
        "bev2", os.path.join(BASE, "build_evolution_v2.py"))
    m = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(m)
    return m


v2 = _load_v2()

OUTDIR = os.path.join(BASE, "data", "evolved_ext")
BIRD_JSON = os.path.join(BASE, "data", "bird", "dev_20240627", "dev.json")


# ------------------------------------------------------------------ 候选枚举
def merge_candidates(conn, rows, cols):
    """每表一对列。判定与 v2.pick_merge_pair 完全一致，只是不取 top-1。"""
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
        if not any(f >= 1 for f, _c in lst):
            continue          # 该表没有被 gold SQL 引用 → 演化后无题可测
        lst.sort(key=lambda x: (-x[0], x[1]))
        txt = [(f, c) for f, c in lst
               if v2.all_text(conn, t, c) and v2.no_null(conn, t, c)]
        if len(txt) < 2:
            continue
        # 同表可以有多对互不重叠的列组合，每一对是一次独立的列重整事实。
        # 这是审稿意见点名的「不同字段组合」——但同表的两对列共享行集，
        # 事件级报告须同时给出表数与库数，不能只报对数。
        for i in range(0, len(txt) - 1, 2):
            fa, a = txt[i]
            fb, b = txt[i + 1]
            out.append({"table": t, "a": a, "b": b, "pair_idx": i // 2,
                        "ref": fa + fb})
    out.sort(key=lambda d: (-d["ref"], d["table"]))
    return out


def split_candidates(conn, rows, cols):
    """每列一个候选。NULL 放行：split 后 gold 是 a||sep||b，原值 NULL 拼回仍 NULL。"""
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
            nd = conn.execute(
                f'SELECT COUNT(DISTINCT "{c}") FROM "{t}"').fetchone()[0]
        except Exception:
            continue
        if nd < 5:
            continue
        for sep in v2.SPLIT_SEPS:
            try:
                # NULL 必须放行：split 后 gold 是 a||sep||b，原值 NULL 时两段皆
                # NULL、拼回仍是 NULL，语义等价成立。用 COALESCE('') 会把 NULL 行
                # 误判成「不含分隔符」而拒绝整列——thrombosis.Birthday 就是这样被
                # 漏掉的（1193/1238 非空且全部含 '-'）。
                bad = conn.execute(
                    f'SELECT 1 FROM "{t}" WHERE "{c}" IS NOT NULL AND '
                    f'INSTR("{c}", {sep}) = 0 LIMIT 1').fetchone()
            except Exception:
                continue
            if bad is None:
                out.append({"table": t, "col": c, "sep": sep, "ref": freq[c]})
                break
    out.sort(key=lambda d: (-d["ref"], d["table"], d["col"]))
    # 同表的不同列是不同的列拆分事实，各自一个事件。但 (库, 表) 会被复用，
    # 必须给序号：否则 tag 相同、库文件互相覆盖，事件表去重时也会把
    # student_club.major 的 college 与 department 并成一个。
    per_table = collections.Counter()
    for d in out:
        d["pair_idx"] = per_table[d["table"]]
        per_table[d["table"]] += 1
    return out


def vshift_candidates(conn, rows, cols):
    """每个 (表, 列) 一个值域重编码事件。

    事件按列而不是按表计：值域重编码的作用对象是「一列的取值集合」，
    同表两列的值域是两个独立事实（两套映射、两份字面量），
    审稿意见要求的「复杂值改写 10–20 事件」只有在列粒度上才数得出来。
    判据与 v2.apply_value_shift 内部的过滤保持一致（nd 3–200、值长 ≤60、
    非主键、全库唯一归属、被 gold SQL 引用过）。
    """
    freq = v2.freq_map(rows)
    owners = collections.defaultdict(set)
    for t, c, _typ, _nn, p in cols:
        owners[c].add(t)
    out = []
    for t, c, typ, _nn, p in cols:
        if p or len(owners[c]) != 1 or freq.get(c, 0) < 1:
            continue
        try:
            # 用 LENGTH(c) > 0 代替 c != ''：后者在 f-string 里紧邻三引号，
            # 会把 SQL 在空串处截断，异常被 except 吞掉后表现为「全库无候选」。
            vals = conn.execute(
                f'SELECT COUNT(DISTINCT "{c}") FROM "{t}" '
                f'WHERE "{c}" IS NOT NULL AND LENGTH("{c}") > 0').fetchone()[0]
            mx = conn.execute(
                f'SELECT MAX(LENGTH("{c}")) FROM "{t}"').fetchone()[0] or 0
        except Exception:
            continue
        if vals < 3 or vals > 200 or mx > 60:
            continue
        out.append({"table": t, "col": c, "pair_idx": 0,
                    "ref": freq[c], "n_distinct": vals})
    out.sort(key=lambda d: (-d["ref"], d["table"], d["col"]))
    per_table = collections.Counter()
    for d in out:
        d["pair_idx"] = per_table[d["table"]]
        per_table[d["table"]] += 1
    return out


# ------------------------------------------------------------------ 单事件构造
def build_one(task, db, spec, rows, src, tag):
    """对单个 (库, 表) 施加一次演化并验证，返回记录列表。"""
    os.makedirs(OUTDIR, exist_ok=True)
    dst = os.path.join(OUTDIR, f"{tag}.sqlite")
    if os.path.exists(dst):
        os.remove(dst)
    v2.copy_db(src, dst)

    conn = v2.connect(dst)
    try:
        if task == "merge_col":
            applied = v2.apply_merge(conn, spec)
        elif task == "split_col":
            applied = v2.apply_split(conn, spec)
        else:                       # value_shift：单列值域重编码
            _cols = v2.list_schema(dst)[1]
            applied = v2.apply_value_shift(conn, _cols, [(spec["col"], spec["ref"])],
                                           named=False)
        conn.commit()
    finally:
        conn.close()
    if not applied:
        return [], None, "apply-failed"

    _c = v2.connect(src, ro=True)
    old_ddl = v2.ddl_map(_c)
    _c.close()
    _c = v2.connect(dst, ro=True)
    new_ddl = v2.ddl_map(_c)
    _c.close()

    # value_shift 下新库的值是 'X-原值'，结果等价必须在归一化空间比较，
    # 否则每一条都会被判成失败。
    lit_map = {}
    for _v in applied.values():
        lit_map.update(_v.get("lit_map") or {})
    n_aff = n_ok = n_bad = n_to = 0
    recs = []
    for it in rows:
        old_sql = it["SQL"]
        if not v2.touched_cols(old_sql, applied):
            continue
        if re.search(r"\bselect\s+\*|\w\.\*", old_sql, re.I):
            n_to += 1
            continue
        orig_rows, orig_err = v2.run_sql(src, old_sql)
        if orig_err is not None:
            n_to += 1
            continue
        old_on_new_rows, old_on_new_err = v2.run_sql(dst, old_sql)
        if old_on_new_err == "TIMEOUT":
            n_to += 1
            continue
        n_aff += 1
        new_sql = v2.rewrite(old_sql, {k: v["new"] for k, v in applied.items()}, lit_map)
        if not new_sql:
            n_bad += 1
            continue
        new_rows, new_err = v2.run_sql(dst, new_sql)
        same = (v2.norm_rows(orig_rows, lit_map) == new_rows
                if lit_map else orig_rows == new_rows)
        ok = (new_err is None and old_on_new_err is not None and same)
        if ok:
            n_ok += 1
        else:
            n_bad += 1
        recs.append({
            "id": f"ext:{task}:{db}:{spec['table']}:{it['question_id']}",
            "evo_type": task,
            "db_id": db,
            "question_id": it["question_id"],
            "question": it["question"],
            "difficulty": it.get("difficulty"),
            "expected_action": "rewrite",
            "n_ops": 1,
            "table": spec["table"],
            "sep": spec.get("sep"),
            "ops": [[{"merge_col": "merge_columns", "split_col": "split_column"}
                     .get(task, "rename_and_recode"),
                     spec["table"],
                     (spec.get("a") or spec.get("col")),
                     applied[[k for k in applied][0]]["new"]]],
            "old_sql": old_sql,
            "new_sql": new_sql,
            "old_schema_ddl": old_ddl,
            "new_schema_ddl": new_ddl,
            "check_pass": ok,
            "ext_db_path": os.path.relpath(dst, BASE),
        })
    info = {"tag": tag, "table": spec["table"],
            "pair_idx": spec.get("pair_idx"), "n_affected": n_aff,
            "n_pass": n_ok, "n_fail": n_bad, "n_dropped": n_to,
            "cols": list(applied)}
    return recs, info, None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--task", required=True,
                    choices=["merge_col", "split_col", "value_shift"])
    ap.add_argument("--max-mb", type=float, default=130)
    ap.add_argument("--dbs", default="")
    a = ap.parse_args()

    only = {x.strip() for x in a.dbs.split(",") if x.strip()} or None
    items = json.load(open(BIRD_JSON, encoding="utf-8"))
    by_db = collections.defaultdict(list)
    for it in items:
        by_db[it["db_id"]].append(it)

    all_recs, per_event, skipped = [], [], []
    t0 = time.time()

    for db in sorted(by_db):
        if only and db not in only:
            continue
        src = v2.sqlite_path(db)
        size_mb = os.path.getsize(src) / 1024 / 1024
        if size_mb > a.max_mb:
            skipped.append([db, round(size_mb, 1), f">{a.max_mb}MB"])
            print(f"[skip] {db} {size_mb:.0f}MB", flush=True)
            continue
        tabs, cols = v2.list_schema(src)
        c0 = v2.connect(src, ro=True)
        if a.task == "merge_col":
            cands = merge_candidates(c0, by_db[db], cols)
        elif a.task == "split_col":
            cands = split_candidates(c0, by_db[db], cols)
        else:
            cands = vshift_candidates(c0, by_db[db], cols)
        c0.close()
        if not cands:
            skipped.append([db, round(size_mb, 1), "no-candidate"])
            print(f"[skip] {db} 无候选", flush=True)
            continue

        for spec in cands:
            _pi = spec.get("pair_idx")
            tag = (f"{db}__{spec['table']}__{a.task}"
                   + ("" if _pi in (None, 0) else f"_p{_pi}"))
            if spec.get("col"):
                tag += "__" + re.sub(r"[^0-9A-Za-z]", "_", str(spec["col"]))
            recs, info, err = build_one(a.task, db, spec, by_db[db], src, tag)
            if err:
                skipped.append([db, spec["table"], err])
                print(f"  [warn] {tag} {err}", flush=True)
                # 失败的库副本立刻删掉，避免目录里留下「列已改但不计入」的文件
                p = os.path.join(OUTDIR, f"{tag}.sqlite")
                if os.path.exists(p):
                    os.remove(p)
                continue
            info.update({"db": db, "size_mb": round(size_mb, 1),
                         "n_items": len(recs),
                         "pass_rate": round(info["n_pass"] / max(1, info["n_affected"]) * 100, 1)})
            per_event.append(info)
            all_recs += recs
            print(f"  [{a.task}] {db:22s} {spec['table']:24s} "
                  f"受影响 {info['n_affected']:3d} 通过 {info['n_pass']:3d} "
                  f"列 {info['cols']}", flush=True)

    out_jsonl = os.path.join(BASE, "data", "manifest_ext_%s.jsonl" % a.task)
    # 合并写入，不覆盖：同一任务可能分批跑（先 130MB 以内，再单独放大库），
    # 若每次全量重写，第二批会把第一批的事件抹掉——本轮第一次跑 card_games 时
    # 就是这样把已有 19 个拆分事件冲成 5 个的。按 id 去重，新记录覆盖同 id 旧记录。
    merged = {}
    if os.path.exists(out_jsonl):
        for line in open(out_jsonl, encoding="utf-8"):
            line = line.strip()
            if line:
                r = json.loads(line)
                merged[r["id"]] = r
    for r in all_recs:
        merged[r["id"]] = r
    with open(out_jsonl, "w", encoding="utf-8") as f:
        for r in merged.values():
            f.write(json.dumps(r, ensure_ascii=False) + "\n")
    all_recs = list(merged.values())

    # 事件表同样合并：按 (db, table, pair_idx) 去重
    stats_path = os.path.join(BASE, "results", "evolution_ext_%s.json" % a.task)
    old_events = []
    if os.path.exists(stats_path):
        try:
            old_events = json.load(open(stats_path, encoding="utf-8")).get("per_event", [])
        except Exception:
            old_events = []
    ev_map = {(e["db"], e["table"], e.get("pair_idx")): e for e in old_events}
    for e in per_event:
        ev_map[(e["db"], e["table"], e.get("pair_idx"))] = e
    per_event = sorted(ev_map.values(), key=lambda e: (e["db"], e["table"],
                                                       e.get("pair_idx") or 0))

    stats = {
        "version": "evolution-ext-v1",
        "task": a.task,
        "max_mb": a.max_mb,
        "events": len(per_event),
        "events_effective": sum(1 for e in per_event if e["n_pass"] > 0),
        "dbs": len({e["db"] for e in per_event}),
        "tables": len({(e["db"], e["table"]) for e in per_event}),
        "items": len(all_recs),
        "pass": sum(1 for r in all_recs if r["check_pass"]),
        "elapsed_s": round(time.time() - t0, 1),
        "skipped": skipped,
        "per_event": per_event,
    }
    json.dump(stats, open(stats_path, "w", encoding="utf-8"),
              ensure_ascii=False, indent=1)

    print("\n==== 汇总 ====")
    print(f"{a.task}: 独立事件 {stats['events']} / 库 {stats['dbs']} / "
          f"题 {stats['items']} / 通过 {stats['pass']} / {stats['elapsed_s']}s")
    print("输出:", out_jsonl)


if __name__ == "__main__":
    main()
