# -*- coding: utf-8 -*-
"""自动谱系基线 B1/B2/B3（回答审稿人：增益来自 M1 设计，还是任何合理匹配都能达到？）

三个基线都输出与 M1 完全同格式的谱系文件，因此可走同一条 prompt/评测路径，
不存在「基线用自己的评测器」这种不公平对照。

  B1 name  ：只用列名相似度（difflib 字符级 + token 包含），最朴素的 schema matching
  B2 vover ：只用值域重合（集合 Jaccard + 双向子串包含），不含名字 tie-break、
             不含近唯一列守卫、不含段位推断 —— 即「有值证据但没有 M1 的三项设计」
  B3 emb   ：用嵌入模型（Qwen3-Embedding-8B）匹配「列名+值样例」的语义，
             代表通用 schema-matching 做法

设计上刻意对基线宽厚：
  - 阈值取常用值（0.55/0.50/0.50），不是调出来的最优值
  - 双向子串包含是 B2 的加分项（否则合并/拆列它一条都找不到）
  - 全部基线都带 1-1 贪心约束（避免一对多虚假命中）

输出：data/lineage_infer/<evo>_<db>_<name|vover|emb>.json
"""
import collections
import json
import math
import os
import sys
import atexit
import time
import urllib.request

BASE = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE)
sys.path.insert(0, BASE)

from m1_lineage import EVO_DIR, OLD_DB, gold_by_keys, name_sim, evaluate  # noqa: E402

OUTDIR = os.path.join(BASE, "data", "lineage_infer")
EVOS = ["rename", "rename_medium", "rename_hard", "merge_col", "split_col",
        "value_shift", "value_shift_named", "dep_abstain", "dep_successor"]

TH_NAME = 0.55
TH_VOVER = 0.50
TH_EMB = 0.50
SUB = 80                  # 覆盖率类指标在 80 条抽样上估计（profile_db 用 120）


def profile_light(path, max_rows=20000):
    """轻量列剖析：只要名字/基数/值集/值样例，不算法形态向量与频次表。

    基线不需要形态指纹（那是 M1 的第三类证据），省掉它是 128 次全表扫描里
    最贵的一项。
    """
    import sqlite3
    out = {}
    if not os.path.exists(path):
        return out
    con = sqlite3.connect(f"file:{path}?mode=ro", uri=True)
    try:
        tabs = [r[0] for r in con.execute(
            "select name from sqlite_master where type='table' "
            "and name not like 'sqlite_%'")]
        for t in tabs:
            try:
                cols = [(r[1], (r[2] or "")) for r in
                        con.execute(f'PRAGMA table_info("{t}")')]
            except Exception:
                continue
            if not cols:
                continue
            q = ("SELECT " + ",".join(f'"{c}"' for c, _ in cols) +
                 f' FROM "{t}" LIMIT {max_rows}')
            try:
                data = con.execute(q).fetchall()
            except Exception:
                continue
            info = {"tab": t, "n": len(data), "cols": {}}
            for j, (c, _typ) in enumerate(cols):
                vals = [("" if r[j] is None else str(r[j]).strip()) for r in data]
                nn = [v for v in vals if v != ""]
                s = set(nn)
                info["cols"][c.lower()] = {
                    "name": c, "n": len(data), "nn": len(nn), "ndist": len(s),
                    "set": s, "list": nn[:SUB]}
            out[t.lower()] = info
    finally:
        con.close()
    return out

EMB_MODEL = "Qwen/Qwen3-Embedding-8B"
# 密钥从环境变量读取，不写死在源码里（公开材料包会附带本脚本）。
# B3 的嵌入结果已缓存在 data/emb_cache.json，
# 复现该条基线**不需要**重新调用接口；只有清空缓存时才需设置 SILICONFLOW_API_KEY。
API_KEY = os.environ.get("SILICONFLOW_API_KEY", "")
EMB_CACHE = os.path.join(BASE, "data", "emb_cache.json")


# ---------------- B1：只用列名 ---------------- #

def map_name(old_cols, new_cols):
    """贪心 1-1：按名字相似度从高到低配对。"""
    cand = []
    for ok, op in old_cols.items():
        for nk, np_ in new_cols.items():
            s = name_sim(op["name"], np_["name"])
            if s >= TH_NAME:
                cand.append((s, ok, nk))
    cand.sort(reverse=True)
    used_o, used_n, out = set(), set(), []
    for s, ok, nk in cand:
        if ok in used_o or nk in used_n:
            continue
        used_o.add(ok)
        used_n.add(nk)
        out.append({"type": "rename_column", "old": op_name(old_cols, ok),
                    "new": new_cols[nk]["name"], "cols": [new_cols[nk]["name"]],
                    "score": round(s, 3)})
    for ok, op in old_cols.items():
        if ok not in used_o:
            out.append({"type": "deprecate_column", "old": op["name"],
                        "new": "", "cols": [], "score": 0.0})
    return out


def op_name(cols, k):
    return cols[k]["name"]


# ---------------- B2：只用值域重合 ---------------- #

def _vover(a, b, blob_a=None, blob_b=None, longs_b=None):
    """max(Jaccard, 正向包含, 反向包含)。

    正向：旧列的值以子串形式出现在新列的某个值里（合并列的典型形态）
    反向：新列的长值出现在旧列的值里（拆列时新列的值是旧值的一部分）
    反向只统计长度 >= 3 的去重值，避免单字符列（M/F）造成虚假高覆盖。
    """
    A, B = a["set"], b["set"]
    if not A or not B:
        return 0.0
    inter = len(A & B)
    jac = inter / len(A | B)
    if jac >= TH_VOVER:          # 已过阈值，无需再算包含度
        return jac
    if blob_a is None:
        blob_a = "\x00".join(a["list"])
    if blob_b is None:
        blob_b = "\x00".join(b["list"])
    fwd = sum(1 for v in a["list"] if v and v in blob_b) / max(1, len(a["list"]))
    longs = longs_b if longs_b is not None else [v for v in B if len(v) >= 3]
    bwd = (sum(1 for v in longs if v in blob_a) / len(longs)) if longs else 0.0
    return max(jac, fwd, bwd)


def map_vover(old_cols, new_cols):
    bo = {k: "\x00".join(v["list"]) for k, v in old_cols.items()}
    bn = {k: "\x00".join(v["list"]) for k, v in new_cols.items()}
    ln = {k: [v for v in v_["set"] if len(v) >= 3] for k, v_ in new_cols.items()}
    cand = []
    for ok, op in old_cols.items():
        for nk, np_ in new_cols.items():
            s = _vover(op, np_, bo[ok], bn[nk], ln[nk])
            if s >= TH_VOVER:
                cand.append((s, ok, nk))
    cand.sort(reverse=True)
    used_o, used_n, out = set(), set(), []
    for s, ok, nk in cand:
        if ok in used_o or nk in used_n:
            continue
        used_o.add(ok)
        used_n.add(nk)
        out.append({"type": "rename_column", "old": old_cols[ok]["name"],
                    "new": new_cols[nk]["name"], "cols": [new_cols[nk]["name"]],
                    "score": round(s, 3)})
    for ok in old_cols:
        if ok not in used_o:
            out.append({"type": "deprecate_column", "old": old_cols[ok]["name"],
                        "new": "", "cols": [], "score": 0.0})
    return out


# ---------------- B3：嵌入语义匹配 ---------------- #

def col_text(tab, prof, k=8):
    """把一列变成可嵌入的文本：列名 + 前 k 个去重值。

    把值也放进来是对基线有利的做法——否则中性名 attr_001 连嵌入都无从判断。
    """
    vals = list(dict.fromkeys([v for v in prof["list"] if v]))[:k]
    return f"{tab}.{prof['name']}" + (" | values: " + ", ".join(vals) if vals else "")


def _embed(texts):
    body = json.dumps({"model": EMB_MODEL, "input": texts,
                       "encoding_format": "float"}).encode()
    last = None
    for attempt in range(6):
        try:
            req = urllib.request.Request(
                "https://api.siliconflow.cn/v1/embeddings", data=body,
                headers={"Authorization": "Bearer " + API_KEY,
                         "Content-Type": "application/json"})
            r = json.load(urllib.request.urlopen(req, timeout=180))
            return [d["embedding"] for d in
                    sorted(r["data"], key=lambda x: x["index"])]
        except Exception as e:              # 502/503/超时都可能，退避重试
            last = e
            time.sleep(min(5 * (2 ** attempt), 60))
    raise RuntimeError(f"embedding 连续失败: {last}")


def _cos(a, b):
    s = sum(x * y for x, y in zip(a, b))
    na = math.sqrt(sum(x * x for x in a))
    nb = math.sqrt(sum(y * y for y in b))
    return s / (na * nb) if na and nb else 0.0


_CACHE = None
_CACHE_DIRTY = False


def _load_cache():
    global _CACHE
    if _CACHE is None:
        _CACHE = (json.load(open(EMB_CACHE, encoding="utf-8"))
                  if os.path.exists(EMB_CACHE) else {})
    return _CACHE


def _save_cache():
    """原子写盘：先写 .tmp 再 replace，避免中断后留下残缺 JSON。

    只在进程退出时调用一次——此前每处理一对就重写整个缓存（400MB），
    既慢又会在并发读时读到半截文件。
    """
    global _CACHE_DIRTY
    if _CACHE is None or not _CACHE_DIRTY:
        return
    tmp = EMB_CACHE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(_CACHE, f)
    os.replace(tmp, EMB_CACHE)
    _CACHE_DIRTY = False


atexit.register(_save_cache)


def embed_pair(db_id, evo_type):
    """取 (evo,db) 双端所有列的嵌入，带磁盘缓存。"""
    global _CACHE_DIRTY
    cache = _load_cache()
    old_p = profile_light(OLD_DB.format(db=db_id))
    new_p = profile_light(os.path.join(BASE, EVO_DIR[evo_type], f"{db_id}.sqlite"))
    need = []
    for p, src in ((old_p, "old"), (new_p, "new")):
        for t, info in p.items():
            for ck, prof in info["cols"].items():
                key = f"{src}|{db_id}|{evo_type if src=='new' else '-'}|{t}|{ck}"
                need.append((key, col_text(t, prof)))
    todo = [(k, tx) for k, tx in need if k not in cache]
    for i in range(0, len(todo), 32):
        batch = todo[i:i + 32]
        vecs = _embed([tx for _, tx in batch])
        for (k, _), v in zip(batch, vecs):
            cache[k] = v
        _CACHE_DIRTY = True
        print(f"    embed {db_id} {evo_type if True else ''} {i+len(batch)}/{len(todo)}", flush=True)
    return old_p, new_p, cache


def map_emb(old_cols, new_cols, vold, vnew):
    cand = []
    for ok, op in old_cols.items():
        for nk, np_ in new_cols.items():
            cand.append((_cos(vold[ok], vnew[nk]), ok, nk))
    cand.sort(reverse=True)
    used_o, used_n, out = set(), set(), []
    for s, ok, nk in cand:
        if s < TH_EMB or ok in used_o or nk in used_n:
            continue
        used_o.add(ok)
        used_n.add(nk)
        out.append({"type": "rename_column", "old": old_cols[ok]["name"],
                    "new": new_cols[nk]["name"], "cols": [new_cols[nk]["name"]],
                    "score": round(s, 3)})
    for ok in old_cols:
        if ok not in used_o:
            out.append({"type": "deprecate_column", "old": old_cols[ok]["name"],
                        "new": "", "cols": [], "score": 0.0})
    return out


# ---------------- 驱动 ---------------- #

def run_baseline(kind, evo_type, db_id):
    old_p = profile_light(OLD_DB.format(db=db_id))
    new_p = profile_light(os.path.join(BASE, EVO_DIR[evo_type], f"{db_id}.sqlite"))
    if kind == "emb":
        _, _, cache = embed_pair(db_id, evo_type)
    ops = []
    for t, info in old_p.items():
        if t not in new_p:
            continue
        ac, bc = info["cols"], new_p[t]["cols"]
        # 与 M1 完全相同的候选集：同名稳定列不参与匹配（名字没变就是没变）。
        # 不这么做的话基线会对全表每一列都输出判定，精确率没法比。
        oc = {k: v for k, v in ac.items() if k not in bc}
        nc = {k: v for k, v in bc.items() if k not in ac}
        if not oc:
            continue
        if kind == "name":
            sub = map_name(oc, nc)
        elif kind == "vover":
            sub = map_vover(oc, nc)
        else:
            if not nc:
                sub = [{"type": "deprecate_column", "old": v["name"],
                        "new": "", "cols": [], "score": 0.0}
                       for v in oc.values()]
            else:
                vo = {ck: cache[f"old|{db_id}|-|{t}|{ck}"] for ck in oc}
                vn = {ck: cache[f"new|{db_id}|{evo_type}|{t}|{ck}"] for ck in nc}
                sub = map_emb(oc, nc, vo, vn)
        for op in sub:
            op["table"] = info["tab"]
        ops.extend(sub)
    out = {"db_id": db_id, "evo_type": evo_type, "mode": kind, "ops": ops}
    p = os.path.join(OUTDIR, f"{evo_type}_{db_id}_{kind}.json")
    json.dump(out, open(p, "w", encoding="utf-8"), ensure_ascii=False)
    return out


def main():
    kinds = sys.argv[1].split(",") if len(sys.argv) > 1 else ["name", "vover", "emb"]
    rows = [json.loads(l) for l in
            open("data/manifest_v3.jsonl", encoding="utf-8")]
    pairs = sorted({(r["evo_type"], r["db_id"]) for r in rows})
    for kind in kinds:
        print(f"===== {kind} =====", flush=True)
        got = {}
        for evo, db in pairs:
            got[(evo, db)] = run_baseline(kind, evo, db)
            print(f"  {evo:<20}{db:<26}{len(got[(evo, db)]['ops']):>4} ops", flush=True)
        stat, detail = evaluate(rows, got)
        print(f"--- {kind} 列级恢复率（与 M1 同一口径）---")
        tot = [0, 0]
        for evo in EVOS:
            if evo in stat:
                rec, n, _ = stat[evo]
                tot[0] += rec
                tot[1] += n
                print(f"  {evo:<20}{rec:>4}/{n:<4}{rec/n*100:6.1f}%")
        if tot[1]:
            print(f"  {'TOTAL':<20}{tot[0]:>4}/{tot[1]:<4}"
                  f"{tot[0]/tot[1]*100:6.1f}%")
        json.dump({evo: {"recovered": v[0], "total": v[1], "type_ok": v[2]}
                   for evo, v in stat.items()},
                  open(f"results/base_acc_{kind}.json", "w", encoding="utf-8"),
                  ensure_ascii=False, indent=1)


if __name__ == "__main__":
    main()
