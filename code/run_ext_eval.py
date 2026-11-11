# -*- coding: utf-8 -*-
"""在扩充事件集上跑端到端重写实验（C1 / C2 / C3m / C3mp）。

为什么要有这一层
  扩充事件集（build_evolution_ext.py）解决的是「独立事件数」，
  但事件数本身不构成证据：必须证明在**更多、彼此独立的列重整事实上**，
  段位谱系相对列级谱系的优势依然成立。本脚本给出这组读数。

条件（与主实验同一提示模板、同一判定器）
  c1    只给当前 schema + 问题
  c2    当前 schema + 旧 SQL（过期证据）
  c3m   再加 M1 抽取的列级谱系（值域指纹）
  c3mp  再加合并列的段位表达式（m1_part 无监督推断，仅 merge_col 有定义）
  值改写档用 M1 的 shape 模式（与主实验 c3s 一致）；拆分档只有 c3m。

判定
  EX     新 SQL 在分片库执行结果 == 原题在原始库的执行结果
  SILENT 执行成功但结果不等（最危险：不报错的错答案）
  ERROR  执行失败
  ABSTAIN 输出 CANNOT_ANSWER

输出 results/ext_eval_<task>.jsonl（逐题）与 results/ext_eval_<task>.json（汇总）

用法
  python run_ext_eval.py --task merge_col --cond c1,c2,c3m,c3mp --workers 6
  python run_ext_eval.py --task merge_col --limit 20        # 冒烟
"""
import argparse
import collections
import json
import os
import sys
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed

BASE = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE)
sys.path.insert(0, BASE)

import llm_client as LLM                        # noqa: E402
from run_c1c2c3 import (SYSTEM, CANNOT, parse_sql,  # noqa: E402
                        lineage_line, run_sql, norm,
                        load_schema)

INFER_DIR = os.path.join(BASE, "data", "lineage_infer_ext")
BIRD_DB = os.path.join(BASE, "data", "bird", "dev_20240627", "dev_databases")

MODE = {"c3m": "profile", "c3mp": "part", "c3s": "shape"}
COND_ALL = {
    "merge_col": ["c1", "c2", "c3m", "c3mp"],
    "split_col": ["c1", "c2", "c3m"],
    "value_shift": ["c1", "c2", "c3s"],
}
_LOCK = threading.Lock()
_INF = {}


def tag_of(r):
    return os.path.basename(r["ext_db_path"])[:-len(".sqlite")]


def load_inf(task, tag, mode):
    k = (task, tag, mode)
    with _LOCK:
        if k in _INF:
            return _INF[k]
    for m in ([mode, "profile"] if mode == "part" else [mode]):
        p = os.path.join(INFER_DIR, f"{task}_{tag}_{m}.json")
        if os.path.exists(p):
            obj = json.load(open(p, encoding="utf-8")).get("ops") or []
            with _LOCK:
                _INF[k] = obj
            return obj
    with _LOCK:
        _INF[k] = None
    return None


def render(ops, r, natural_only=False):
    if not ops:
        return []
    sql_l = (r.get("old_sql") or "").lower()
    out = []
    for op in ops:
        if op["old"].lower() not in sql_l:
            continue
        ln = lineage_line(op["type"], op["table"], op["old"], op.get("new") or "")
        if ln:
            out.append(ln)
    return out


def build_prompt(r, cond, schema, task):
    blocks = ["### Current database schema (SQLite)", schema, "",
              "### Question", r["question"].strip(), ""]
    if cond != "c1":
        blocks += ["### Retrieved example from an EARLIER version of this database",
                   "Question: " + r["question"].strip(),
                   "SQL: " + (r["old_sql"] or "").strip(), ""]
    if cond in ("c3m", "c3mp", "c3s"):
        ops = load_inf(task, tag_of(r), MODE[cond])
        lines = render(ops, r)
        if lines:
            blocks += ["### Schema lineage (what changed since the example was written)",
                       "\n".join(lines), ""]
    blocks += [
        "### Instruction",
        "Write one SQLite query that answers the Question against the CURRENT schema.",
        f"If the Question can no longer be answered with the current schema "
        f"(the needed column/table was removed and has no successor), "
        f"output exactly {CANNOT} and nothing else.",
        "Output only the SQL (or the token above).",
    ]
    return "\n".join(blocks)


def gold_rows(r):
    db = r["db_id"]
    p = os.path.join(BIRD_DB, db, f"{db}.sqlite")
    return run_sql(p, r["old_sql"])


def judge(sql, rows, gold, err):
    if rows is None and err == "ABSTAIN":
        return "ABSTAIN", None
    if err is not None:
        return "ERROR", err
    if norm(rows) == norm(gold):
        return "EX", None
    return "SILENT", None


_TL = threading.local()


def get_cfg():
    if not hasattr(_TL, "cfg"):
        _TL.cfg = LLM.load_cfg()
    return _TL.cfg


def one(r, cond, task):
    shard = os.path.join(BASE, r["ext_db_path"])
    schema = load_schema(shard)
    gold, gerr = gold_rows(r)
    if gerr is not None:
        return dict(id=r["id"], cond=cond, tag=tag_of(r), evo=r["evo_type"],
                    db_id=r["db_id"], table=r.get("table"),
                    outcome="GOLD_FAIL", sql=None, raw=None)
    prompt = build_prompt(r, cond, schema, task)
    for attempt in range(3):
        try:
            text, meta = LLM.chat(SYSTEM, prompt, cfg=get_cfg())
            break
        except Exception as e:
            if attempt == 2:
                return dict(id=r["id"], cond=cond, tag=tag_of(r), evo=r["evo_type"],
                            db_id=r["db_id"], table=r.get("table"),
                            outcome="LLM_FAIL", sql=None, raw=str(e)[:200])
            time.sleep(5 * (attempt + 1))
    sql, abstain = parse_sql(text)
    if abstain or not sql:
        return dict(id=r["id"], cond=cond, tag=tag_of(r), evo=r["evo_type"],
                    db_id=r["db_id"], table=r.get("table"),
                    outcome="ABSTAIN", sql=None, raw=text[:400])
    rows, err = run_sql(shard, sql)
    outcome, note = judge(sql, rows, gold, err)
    return dict(id=r["id"], cond=cond, tag=tag_of(r), evo=r["evo_type"],
                db_id=r["db_id"], table=r.get("table"),
                outcome=outcome, sql=sql, raw=text[:400],
                err=None if outcome != "ERROR" else err)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--task", required=True,
                    choices=["merge_col", "split_col", "value_shift"])
    ap.add_argument("--cond", default="")
    ap.add_argument("--limit", type=int, default=0)
    ap.add_argument("--workers", type=int, default=6)
    ap.add_argument("--fresh", action="store_true", help="忽略已有结果重跑")
    a = ap.parse_args()

    conds = (a.cond.split(",") if a.cond else COND_ALL[a.task])
    man = [json.loads(l) for l in
           open(os.path.join(BASE, "data", f"manifest_ext_{a.task}.jsonl"),
                encoding="utf-8")]
    if a.limit:
        man = man[:a.limit]

    p_out = os.path.join(BASE, "results", f"ext_eval_{a.task}.jsonl")
    # 断点续跑：读已有行，跳过已完成的 (cond, id)。后台任务可能被外部终止，
    # 没有这一步就会反复消耗同一批调用。
    done_keys, prev = set(), []
    if os.path.exists(p_out) and not a.fresh:
        for line in open(p_out, encoding="utf-8"):
            line = line.strip()
            if not line:
                continue
            try:
                q = json.loads(line)
            except Exception:
                continue
            prev.append(q)
            done_keys.add((q["cond"], q["id"]))
    jobs = [(r, c) for c in conds for r in man
            if (c, r["id"]) not in done_keys]
    print(f"{a.task}: 题 {len(man)} × 条件 {len(conds)} = {len(man)*len(conds)} 次；"
          f"已完成 {len(done_keys)}，待跑 {len(jobs)}", flush=True)

    t0 = time.time()
    p = p_out
    # 增量写：本任务曾出现「进程静默退出、结果全部丢失」的情况，
    # 因此每完成一条立即 append，并把异常写成 _ERROR 行的 traceback，
    # 便于区分「模型失败」与「脚本自身出错」。
    fh = open(p, "a" if prev else "w", encoding="utf-8")
    out = list(prev)
    try:
        with ThreadPoolExecutor(max_workers=a.workers) as ex:
            futs = {ex.submit(one, r, c, a.task): (r, c) for r, c in jobs}
            done = 0
            for f in as_completed(futs):
                done += 1
                try:
                    rec = f.result()
                except Exception as e:
                    import traceback
                    rec = dict(id=futs[f][0]["id"], cond=futs[f][1],
                               tag=tag_of(futs[f][0]), evo=a.task,
                               db_id=futs[f][0]["db_id"], table=None,
                               outcome="_ERROR",
                               sql=None, raw=traceback.format_exc()[-500:])
                out.append(rec)
                fh.write(json.dumps(rec, ensure_ascii=False) + "\n")
                fh.flush()
                if done % 50 == 0:
                    print(f"  {done}/{len(jobs)}  {time.time()-t0:.0f}s", flush=True)
    finally:
        fh.close()

    # 汇总：按条件；再按事件（分片）做配对
    summ = {}
    for c in conds:
        xs = [r for r in out if r["cond"] == c]
        cnt = collections.Counter(r["outcome"] for r in xs)
        n = len(xs)
        summ[c] = {"n": n, "EX": cnt["EX"],
                   "rate": round(cnt["EX"] / max(1, n) * 100, 1),
                   "by_outcome": dict(cnt),
                   "events": len({r["tag"] for r in xs})}
    # 事件级配对（同题同事件下 c3mp/c3m 谁对）
    paired = {}
    for hi, lo in (("c3mp", "c3m"), ("c3m", "c2"), ("c2", "c1"),
                   ("c3s", "c3m")):
        if hi not in conds or lo not in conds:
            continue
        A = {(r["tag"], r["id"]): r["outcome"] for r in out if r["cond"] == hi}
        B = {(r["tag"], r["id"]): r["outcome"] for r in out if r["cond"] == lo}
        per_ev = collections.defaultdict(lambda: [0, 0])
        for k in set(A) & set(B):
            a_ok, b_ok = A[k] == "EX", B[k] == "EX"
            if a_ok and not b_ok:
                per_ev[k[0]][0] += 1
            elif b_ok and not a_ok:
                per_ev[k[0]][1] += 1
        pos = sum(1 for v in per_ev.values() if v[0] > v[1])
        neg = sum(1 for v in per_ev.values() if v[1] > v[0])
        tie = len(per_ev) - pos - neg
        import math
        n_ = pos + neg
        pv = (min(1.0, 2.0 * sum(math.comb(n_, j) for j in range(min(pos, neg) + 1))
                  / 2 ** n_) if n_ else 1.0)
        paired[f"{hi}-{lo}"] = {"pos_events": pos, "neg_events": neg,
                                "tie": tie, "n_events": len(per_ev),
                                "p": round(pv, 6)}

    js = {"version": "ext-eval-v1", "task": a.task, "n_items": len(man),
          "elapsed_s": round(time.time() - t0, 1),
          "conditions": summ, "paired_by_event": paired}
    json.dump(js, open(os.path.join(BASE, "results", f"ext_eval_{a.task}.json"),
                       "w", encoding="utf-8"), ensure_ascii=False, indent=1)

    print()
    for c, v in summ.items():
        print(f"  {c:6s} n={v['n']:4d} EX={v['EX']:4d} ({v['rate']:5.1f}%) "
              f"事件 {v['events']}  {v['by_outcome']}")
    print("  事件级配对:", json.dumps(paired, ensure_ascii=False))


if __name__ == "__main__":
    main()
