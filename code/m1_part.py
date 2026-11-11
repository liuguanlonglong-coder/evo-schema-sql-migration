# -*- coding: utf-8 -*-
"""合并列的段位推断（position-level reconstruction for merge_columns）。

问题：M1 值域指纹只能回答「attr_001 里装的是 County + StatusType」，
不能回答「County 在第 1 段、StatusType 在第 2 段」。而 merge 档的可答性
恰恰取决于段位——不给段位，模型无法写出 SUBSTR(...)。

方法（不读任何标注）：
  1. 边界探测：统计合并列取值中出现的非字母数字字符，按「出现在多少比例的值里」
     排序，取最高的作为分隔符（\x1f 这类控制字符自然会被检出）。
  2. 段匹配：把合并列按分隔符切段，对每个候选旧列，计算其值集合与每一段
     值集合的 Jaccard，取最大者。Jaccard 高即判定「该列在第 i 段」。
  3. 输出重构表达式 SUBSTR(attr, INSTR 链) 形式，供 prompt 渲染。

用法：
  python m1_part.py               # 全库推断 + 与 gold 比对
  python m1_part.py --write       # 同时写出 data/lineage_infer/merge_col_<db>_part.json
"""
import argparse
import collections
import json
import os
import sqlite3
import sys

BASE = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE)
sys.path.insert(0, BASE)

import build_evolution_v2 as v2  # noqa: E402

BIRD = v2.BIRD
SEP_CANDIDATES = ["\x1f", "\x1e", "\x1d", "|", ";", ",", "~", "\t"]
SPLIT_THRESH = 0.55      # 段 Jaccard 判定阈值
MIN_ROWS = 20


def read_schema(path):
    con = sqlite3.connect(f"file:{path}?mode=ro", uri=True)
    out = {}
    for (name,) in con.execute("select name from sqlite_master where type='table'"):
        cols = [r[1] for r in con.execute(f"PRAGMA table_info('{name}')")]
        out[name] = cols
    con.close()
    return out


def col_values(path, table, col, limit=4000):
    con = sqlite3.connect(f"file:{path}?mode=ro", uri=True)
    try:
        rs = con.execute(
            f'SELECT DISTINCT "{col}" FROM "{table}" '
            f'WHERE "{col}" IS NOT NULL AND "{col}" != "" LIMIT {limit}'
        ).fetchall()
    except sqlite3.Error:
        rs = []
    con.close()
    return {str(r[0]) for r in rs}


def detect_sep(values):
    """边界探测：返回最可能的拼接分隔符。"""
    if not values:
        return None
    n = len(values)
    best, best_cnt = None, 0
    for ch in SEP_CANDIDATES:
        cnt = sum(1 for v in values if ch in v)
        if cnt > best_cnt:
            best, best_cnt = ch, cnt
    # 自动发现：非字母数字、非空格的高频字符
    cnt = collections.Counter()
    for v in values:
        for ch in set(v):
            if not ch.isalnum() and ch != " ":
                cnt[ch] += 1
    for ch, c in cnt.most_common(3):
        if c > best_cnt:
            best, best_cnt = ch, c
    if best is None or best_cnt / n < 0.5:
        return None
    return best


def jac(a, b):
    if not a or not b:
        return 0.0
    return len(a & b) / len(a | b)


def infer_db(db_id):
    """返回 {table: {attr_col: {"sep":..., "parts": {old_col: (idx, score)}}}}"""
    old_root = os.environ.get("OLD_DBDIR") or os.path.join(BIRD, "dev_databases")
    new_dir = os.environ.get("EVO_DIR") or os.path.join("data", "evolved_merge")
    old_p = os.path.join(old_root, db_id, f"{db_id}.sqlite")
    new_p = os.path.join(new_dir, f"{db_id}.sqlite")
    if not (os.path.exists(old_p) and os.path.exists(new_p)):
        return None
    old_s, new_s = read_schema(old_p), read_schema(new_p)

    result = {}
    for tab, new_cols in new_s.items():
        if tab not in old_s:
            continue
        old_cols = set(old_s[tab])
        # 新出现的 attr_* 列 = 合并产物；旧库里消失的列 = 被合并的列
        merged = [c for c in new_cols if c.startswith("attr_") and c not in old_cols]
        dropped = [c for c in old_cols if c not in set(new_cols)]
        if not merged or not dropped:
            continue
        for mc in merged:
            vals = col_values(new_p, tab, mc)
            if len(vals) < MIN_ROWS:
                vals = col_values(new_p, tab, mc, limit=20000)
            sep = detect_sep(vals)
            if not sep:
                continue
            segs = [v.split(sep) for v in vals]   # 必须是 list，顺序即段号
            nseg = max(len(s) for s in segs) if segs else 0
            if nseg < 2:
                continue
            # 按段号聚合（不同行段数可能不一致，按位置收集）
            # 空段不入集合（与 col_values 排除空串口径一致），但绝不压缩段号
            per_idx = collections.defaultdict(set)
            for s in segs:
                for i, piece in enumerate(s):
                    if piece != "":
                        per_idx[i].add(piece)
            parts = {}
            used = set()
            for oc in dropped:
                ov = col_values(old_p, tab, oc)
                if len(ov) < 2:
                    continue
                # 低基数列（如 M/F）与任意短段都容易撞车，要求更高置信度
                need = 0.95 if len(ov) < 3 else SPLIT_THRESH
                scores = []
                for i in range(nseg):
                    sc = jac(ov, per_idx[i])
                    scores.append((sc, i))
                scores.sort(reverse=True)
                if not scores:
                    continue
                sc, idx = scores[0]
                if sc >= need:
                    parts[oc] = {"idx": idx, "score": round(sc, 3),
                                 "nseg": nseg, "sep": sep}
                    used.add(oc)
            if parts:
                result.setdefault(tab, {})[mc] = {"sep": sep, "nseg": nseg,
                                                  "parts": parts}
    return result


def to_expr(col, idx, nseg, sep):
    """把「第 idx 段（0 基，共 nseg 段）」渲染成 SQLite 表达式。

    SQLite 没有 split_part，只能用「取前缀 / 剥前缀」组合出各段：
        rest_0     = col
        rest_{k+1} = SUBSTR(col, INSTR(rest_k, sep) + 1)
        段 0       = SUBSTR(col, 1, INSTR(col, sep) - 1)
        段 i（中间）= SUBSTR(rest_i, 1, INSTR(rest_i, sep) - 1)
        段 nseg-1  = rest_{nseg-1}（剩余全部，不可再按分隔符切）

    旧实现把 idx > 0 一律当作末段，nseg >= 3 时第 1 段与第 2 段给出完全相同的
    表达式（都取最后一个分隔符之后），三段合并会出现重复表达式——这是此前
    「只验证两段」的实现原因，不是数据选择。
    """
    sch = f"'{sep}'" if sep.isprintable() else f"CHAR({ord(sep)})"

    def rest(k):
        # 必须对上一段剩余串再取子串，不能用绝对位置——INSTR 返回的是相对该串的位置
        e = col
        for _ in range(k):
            e = f"SUBSTR({e},INSTR({e},{sch})+1)"
        return e

    if nseg <= 2:
        # 两段时与旧实现逐字一致，保证历史数值不受影响
        return (f"SUBSTR({col},1,INSTR({col},{sch})-1)" if idx == 0
                else f"SUBSTR({col},INSTR({col},{sch})+1)")
    if idx == 0:
        return f"SUBSTR({col},1,INSTR({col},{sch})-1)"
    head = rest(idx)
    if idx >= nseg - 1:
        # 末段：rest(nseg-1) 已经是「最后一段及其后全部内容」，必须整段取回。
        # 旧实现再按分隔符切一次（SUBSTR(head, INSTR(head,sep)+1)），末段内部
        # 含分隔符时只保留其后的尾部，nseg>=3 会静默丢失信息——
        # probe_encoding_bounds_v2 实测 nseg>=4 的 sep_in_last 为 0%。
        # 两段走上面 nseg<=2 的独立分支，逐字不受本次修正影响。
        return head
    return f"SUBSTR({head},1,INSTR({head},{sch})-1)"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--write", action="store_true")
    ap.add_argument("--manifest", default="data/manifest_v3.jsonl")
    ap.add_argument("--evo", default="merge_col")
    ap.add_argument("--outdir", default="data/lineage_infer")
    args = ap.parse_args()

    rows = [json.loads(l) for l in open(args.manifest, encoding="utf-8")]
    gold = collections.defaultdict(set)   # (db) -> {(tab, old)}
    for r in rows:
        if r["evo_type"] != args.evo:
            continue
        for op in r.get("ops") or []:
            if isinstance(op, list) and len(op) >= 3:
                gold[r["db_id"]].add((str(op[1]), str(op[2])))

    total = hit = 0
    wrong = []
    for db in sorted(gold):
        inf = infer_db(db)
        if not inf:
            print(f"{db:<28} 无推断")
            continue
        flat = {}
        for tab, d in inf.items():
            for mc, meta in d.items():
                for oc, p in meta["parts"].items():
                    flat[(tab, oc)] = (mc, p)
        g = gold[db]
        ok = sum(1 for k in g if k in flat)
        total += len(g)
        hit += ok
        print(f"{db:<28} gold={len(g)}  段位复原={ok}")
        for k in sorted(g):
            if k not in flat:
                wrong.append((db, k))
                continue
            mc, p = flat[k]
            for op in rows:
                pass
        if args.write:
            out = {"db_id": db, "evo_type": args.evo, "mode": "part",
                   "sep_detected": True, "ops": []}
            for tab, d in inf.items():
                for mc, meta in d.items():
                    for oc, p in meta["parts"].items():
                        out["ops"].append({
                            "type": "merge_columns", "table": tab, "old": oc,
                            "new": to_expr(mc, p["idx"], p["nseg"], meta["sep"]),
                            "cols": [mc], "score": p["score"],
                            "part_idx": p["idx"], "nseg": p["nseg"],
                        })
            os.makedirs(args.outdir, exist_ok=True)
            json.dump(out, open(os.path.join(args.outdir, f"{args.evo}_{db}_part.json"),
                                "w", encoding="utf-8"),
                      ensure_ascii=False, indent=1)

    print(f"\n段位复原合计 {hit}/{total} = {hit/total*100:.1f}%")
    if wrong:
        print("未复原：", wrong)


if __name__ == "__main__":
    main()
