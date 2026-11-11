"""用原始数据复算两处：①图9（合并档四条件失败类型构成）；②扩充集四条件计数。

口径与主实验判定器完全一致（见 run_c1c2c3.py 的 run_sql / norm / judge）：
  - 结果集每单元格转字符串，None 统一为空串，整行排序后逐元组比较；
  - 预测为空（弃答）→ 弃答；执行抛异常 → ERROR；与 gold 相等 → EX；否则 → SILENT；
  - gold 由清单里的 new_sql 在演化后库上现场执行得到，不沿用任何已存标签。

输出：results/audit_recompute.json
"""
import json
import os
import sqlite3
import time
from collections import Counter, defaultdict

BASE = os.path.abspath(".")
TIMEOUT = 6.0
ROOT_OUT = os.path.join("公开材料包", "outputs.jsonl")
EXT_OUT = os.path.join("公开材料包", "outputs_ext.jsonl")

MERGE_DB_DIR = os.path.join(BASE, "data", "evolved_merge")


def run_sql(path, sql):
    """返回 (rows, err)，行为与主实验判定器的 run_sql 一致。"""
    if not sql:
        return None, "EMPTY_SQL"
    p = path.replace("\\", "/")
    if not os.path.exists(p):
        return None, "DB_MISSING"
    con = sqlite3.connect(f"file:{p}?mode=ro", uri=True)
    t0 = time.time()
    con.set_progress_handler(lambda: 1 if time.time() - t0 > TIMEOUT else 0, 2000)
    try:
        return con.execute(sql).fetchall(), None
    except Exception as e:
        return None, str(e)[:200]
    finally:
        con.close()


def norm(rows):
    if rows is None:
        return None
    return sorted(tuple("" if v is None else str(v) for v in r) for r in rows)


def label(pred_sql, gold_rows, gold_err):
    """复算标签；gold 无法执行时单独记为 GOLD_ERR，不归入任何条件。"""
    if gold_err:
        return "GOLD_ERROR"
    if not pred_sql:
        return "ABSTAIN"
    rows, err = run_sql_current(pred_sql)
    if err:
        return "ERROR"
    return "EX" if norm(rows) == norm(gold_rows) else "SILENT"


# 预测 SQL 需要在各自的演化库上执行，故用全局变量传当前库路径
_CUR_DB = {"path": None}

def run_sql_current(sql):
    return run_sql(_CUR_DB["path"], sql)


def load_jsonl(p):
    return [json.loads(l) for l in open(p, encoding="utf-8")]


# ---------------- A. 图9：合并档四条件 ----------------
def part_a():
    man = [r for r in load_jsonl("data/manifest_v3.jsonl") if r["evo_type"] == "merge_col"]
    man = {r["id"]: r for r in man}
    preds = [r for r in load_jsonl(ROOT_OUT) if r["id"].split(":")[0] == "merge_col"]
    conds = ["c1", "c2", "c3m", "c3f"]
    by = defaultdict(dict)
    for r in preds:
        if r["cond"] in conds:
            by[r["cond"]][r["id"]] = r

    ids = sorted(man)
    gold = {}
    for i in ids:
        _CUR_DB["path"] = os.path.join(MERGE_DB_DIR, f"{man[i]['db_id']}.sqlite")
        rows, err = run_sql(_CUR_DB["path"], man[i]["new_sql"])
        gold[i] = (rows, err)

    out = {"n_items": len(ids), "per_cond": {}, "mismatch": []}
    for c in conds:
        cnt, stored_cnt = Counter(), Counter()
        missing = 0
        for i in ids:
            rec = by[c].get(i)
            if rec is None:
                missing += 1
                continue
            _CUR_DB["path"] = os.path.join(MERGE_DB_DIR, f"{man[i]['db_id']}.sqlite")
            lab = label(rec["sql"], *gold[i])
            cnt[lab] += 1
            st = "ABSTAIN" if rec["outcome"] in ("OVER_ABSTAIN", "ABSTAIN") else rec["outcome"]
            stored_cnt[st] += 1
            if st != lab:
                out["mismatch"].append({"cond": c, "id": i, "stored": st, "recomputed": lab})
        out["per_cond"][c] = {
            "n": sum(cnt.values()),
            "missing": missing,
            "recomputed": dict(cnt),
            "stored": dict(stored_cnt),
            "agree": sum(cnt[k] for k in cnt if stored_cnt.get(k) == cnt[k]) and
                     (dict(cnt) == dict(stored_cnt)),
        }

    # C1 一列的内部构成（供复核报告使用）：静默错误里空结果集 vs 非空错值，
    # 报错里列不存在 vs 6 秒超时中断。
    c1 = Counter()
    for i in ids:
        rec = by["c1"].get(i)
        if rec is None:
            continue
        d = os.path.join(MERGE_DB_DIR, f"{man[i]['db_id']}.sqlite")
        g_rows, g_err = gold[i]
        if not rec["sql"]:
            c1["abstain"] += 1
            continue
        rows, err = run_sql(d, rec["sql"])
        if err:
            if "no such column" in err or "no such table" in err:
                c1["error_schema"] += 1
            elif "interrupted" in err:
                c1["error_timeout"] += 1
            else:
                c1["error_other"] += 1
            continue
        if norm(rows) == norm(g_rows):
            c1["ex"] += 1
        elif len(rows) == 0:
            c1["silent_empty"] += 1
        else:
            c1["silent_nonempty"] += 1
    out["c1_detail"] = dict(c1)

    # 与图 9 中硬编码的四组数对照
    fig = {"c1": [13, 198, 7, 12], "c2": [23, 179, 9, 19],
           "c3m": [6, 223, 0, 1], "c3f": [183, 47, 0, 0]}
    out["fig9_check"] = {}
    for c in conds:
        rc = out["per_cond"][c]["recomputed"]
        got = [rc.get("EX", 0), rc.get("SILENT", 0), rc.get("ERROR", 0), rc.get("ABSTAIN", 0)]
        out["fig9_check"][c] = {"fig": fig[c], "recomputed": got, "match": got == fig[c]}
    return out


# ---------------- B. 扩充集：479 条四条件 ----------------
def part_b():
    man = {r["id"]: r for r in load_jsonl("data/manifest_ext_merge_col.jsonl")}
    preds = load_jsonl(EXT_OUT)
    conds = ["c1", "c2", "c3m", "c3mp"]
    by = defaultdict(dict)
    for r in preds:
        by[r["cond"]][r["id"]] = r

    ids = sorted(man)
    assert len(ids) == 479, len(ids)
    # gold 口径必须与 run_ext_eval.py 的 gold_rows 一致：
    # 用 old_sql 在原始 BIRD 库上执行，而不是 new_sql 在演化库上执行。
    BIRD = os.path.join(BASE, "data", "bird", "dev_20240627", "dev_databases")
    gold = {}
    for i in ids:
        p = man[i]["ext_db_path"]
        if not os.path.isabs(p):
            p = os.path.join(BASE, p)
        gdb = os.path.join(BIRD, man[i]["db_id"], f"{man[i]['db_id']}.sqlite")
        rows, err = run_sql(gdb, man[i]["old_sql"])
        gold[i] = (rows, err, p)

    out = {"n_items": len(ids), "per_cond": {}, "per_cond_453": {},
           "mismatch": [], "coverage": {}}
    for c in conds:
        cnt = Counter()
        stored = Counter()
        for i in ids:
            rec = by[c].get(i)
            if rec is None:
                continue
            _CUR_DB["path"] = gold[i][2]
            lab = label(rec["sql"], gold[i][0], gold[i][1])
            cnt[lab] += 1
            stored[rec["outcome"]] += 1
            if rec["outcome"] != lab:
                out["mismatch"].append({"cond": c, "id": i,
                                        "stored": rec["outcome"], "recomputed": lab})
        out["per_cond"][c] = {"n": sum(cnt.values()), "recomputed": dict(cnt),
                              "stored": dict(stored),
                              "ex_pct": round(100 * cnt.get("EX", 0) / max(1, sum(cnt.values())), 2)}
        out["coverage"][c] = {
            "rows": len(by[c]),
            "ids_in_manifest": len([i for i in by[c] if i in man]),
            "missing_items": len([i for i in ids if i not in by[c]]),
            "extra_ids": len([i for i in by[c] if i not in man]),
        }

        # 453 良构子集（check_pass）
        good = [i for i in ids if str(man[i].get("check_pass")).lower() in ("true", "1", "1.0")]
        c2 = Counter()
        for i in good:
            rec = by[c].get(i)
            if rec is None:
                continue
            _CUR_DB["path"] = gold[i][2]
            c2[label(rec["sql"], gold[i][0], gold[i][1])] += 1
        n = sum(c2.values())
        out["per_cond_453"][c] = {"n": n, "recomputed": dict(c2),
                                  "ex_pct": round(100 * c2.get("EX", 0) / max(1, n), 2)}
    # gold 是否整体可执行
    out["gold_errors"] = sum(1 for i in ids if gold[i][1])
    return out


if __name__ == "__main__":
    t0 = time.time()
    a = part_a()
    print(f"[A] 图9 复算完成 {time.time()-t0:.1f}s, 不一致 {len(a['mismatch'])}")
    for c, v in a["fig9_check"].items():
        print("   ", c, "图", v["fig"], "复算", v["recomputed"], "一致" if v["match"] else "★不一致")
    b = part_b()
    print(f"[B] 扩充集复算完成 {time.time()-t0:.1f}s, 不一致 {len(b['mismatch'])}, gold 执行失败 {b['gold_errors']}")
    for c in ["c1", "c2", "c3m", "c3mp"]:
        v = b["per_cond"][c]
        w = b["per_cond_453"][c]
        print(f"    {c:5s} n={v['n']} EX={v['recomputed'].get('EX')} ({v['ex_pct']}%)   "
              f"453: EX={w['recomputed'].get('EX')} ({w['ex_pct']}%)")
    os.makedirs("results", exist_ok=True)
    json.dump({"fig9": a, "ext": b},
              open("results/audit_recompute.json", "w", encoding="utf-8"),
              ensure_ascii=False, indent=1)
    print("saved results/audit_recompute.json")
