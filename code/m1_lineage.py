# -*- coding: utf-8 -*-
"""M1：自动谱系抽取（不读 drift_ops）

目的：把 E1 里 C3 的 oracle 谱系换成本文方法自己抽出来的谱系，证明
     “对齐”不是靠 ground truth 堆出来的上界，而是可落地的模块。

输入只有三样东西（部署时全都拿得到）：
    旧库 sqlite、演化后库 sqlite、旧 SQL 日志（C2 已给）

三类证据：
  E1 值域指纹  列改名/拆分/合并都不产生新值，只搬运/拼接值。
               因此「值集合的交并比」对 attr_001 这类中性名依然有效，
               而名字相似度在这种情况下完全失效（M0 基线）。
  E2 名字相似  字符级 ratio + token 包含，仅在 E1 弱（基数极低）时做 tie-break。
  E3 位置先验  SQLite 的 ALTER RENAME COLUMN 保序；仅作消融，默认关闭。

输出：data/lineage_infer/<evo>_<db>.json
    {"db_id":..., "evo_type":..., "ops":[{"type","table","old","new","cols","score"}]}

用法：
    python m1_lineage.py --all                 # 全部 (evo, db) 组合
    python m1_lineage.py --eval --all          # 抽取 + 与 gold 对比
"""
import argparse
import collections
import json
import math
import os
import re
import sqlite3
import sys

BASE = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE)

_OLD_ROOT = os.environ.get("OLD_DBDIR") or os.path.join(
    BASE, "data", "bird", "dev_20240627", "dev_databases")
OLD_DB = os.path.join(_OLD_ROOT, "{db}", "{db}.sqlite")
EVO_DIR = {
    "rename": "data/evolved",
    "rename_medium": "data/evolved_medium",
    "rename_hard": "data/evolved_hard",
    "dep_abstain": "data/evolved_dep_abstain",
    "dep_successor": "data/evolved_dep_successor",
    "merge_col": "data/evolved_merge",
    "split_col": "data/evolved_split",
    "value_shift": "data/evolved_vshift",
    "value_shift_named": "data/evolved_vshift_named",
    # 留出库（BIRD train 划分另建的库）
    "merge_col_holdout": os.environ.get("EVO_DIR_MERGE_HOLDOUT",
                                        "data/evolved_merge_holdout"),
    "split_col_holdout": os.environ.get("EVO_DIR_SPLIT_HOLDOUT",
                                        "data/evolved_split_holdout"),
}
OUTDIR = os.path.join(BASE, "data", "lineage_infer")

MAX_ROWS = 20000          # 每表最多扫这么多行做指纹（新旧库同一 rowid 顺序，可比）
# 采样策略（审稿意见要求检验「前 20,000 行是否代表性」）
#   head      —— 当前主结果用的策略：按 rowid 顺序取前 20,000 行
#   rand<N>   —— 按 rowid 派生的确定性伪随机顺序取 20,000 行（同 seed 新旧库一致，
#                因此仍可在同一 rowid 上比较取值）
# 新旧库必须用同一策略，否则值域比对失效。
SAMPLE_MODE = "head"


def _limit_clause(max_rows):
    m = SAMPLE_MODE
    if m.startswith("rand"):
        try:
            seed = int(m[4:] or 1)
        except ValueError:
            seed = 1
        return (f"ORDER BY (rowid*2654435761+{seed})%2147483647 "
                f"LIMIT {max_rows}")
    return f"LIMIT {max_rows}"


SUB_SAMPLE = 120          # 子串包含度抽样条数
LOW_CARD = 5              # 低于此基数，值证据不可靠，转名字 tie-break

IDENT = re.compile(r"[A-Za-z_][A-Za-z0-9_]*")
SQL_FUNC = {"substr", "substring", "instr", "char", "replace", "trim", "ltrim",
            "rtrim", "cast", "printf", "length", "upper", "lower", "coalesce",
            "ifnull", "nullif", "round", "abs", "group_concat", "concat"}


def expr_cols(expr):
    """从 gold 的目标表达式里取出被引用的列名（排除函数名与字面量）。"""
    s = str(expr)
    out = set()
    for m in IDENT.finditer(s):
        tok = m.group(0)
        if s[m.end():].lstrip().startswith("("):
            continue
        if tok.lower() in SQL_FUNC:
            continue
        out.add(tok)
    return sorted(out)


# ---------------- profiling ---------------- #

def _norm(v):
    return "" if v is None else str(v)


def shape_vec(nn, k=256):
    """值的频率曲线：按「频次降序、同频按值字典序」排序后的归一化频次向量。

    这是值域指纹之外的第三类证据。当演化同时改写取值（去重编码、加前缀、
    单位换算）时，新旧列的值集合可以完全不相交——Jaccard 严格为 0，
    但「哪个值多、哪个值少」的结构通常不变，因为重编码几乎总是一一映射。
    因此频率曲线对「值改了但语义结构没变」的演化免疫。
    """
    if not nn:
        return []
    cnt = collections.Counter(nn)
    tot = float(len(nn))
    return [c / tot for _v, c in
            sorted(cnt.items(), key=lambda kv: (-kv[1], kv[0]))][:k]


def shape_sim(a, b):
    """频率曲线相似度：余弦 × 长度接近度。

    长度项不能省：58 个取值的列和 3 个取值的列，前 3 个频次可能碰巧接近，
    但它们的「值域宽度」明显不同，单靠余弦会把它们判成同一列。
    """
    va, vb = a.get("shape") or [], b.get("shape") or []
    if not va or not vb:
        return 0.0
    m = max(len(va), len(vb))
    x = va + [0.0] * (m - len(va))
    y = vb + [0.0] * (m - len(vb))
    dot = sum(p * q for p, q in zip(x, y))
    nx = math.sqrt(sum(p * p for p in x))
    ny = math.sqrt(sum(q * q for q in y))
    if nx == 0 or ny == 0:
        return 0.0
    cos = dot / (nx * ny)
    lr = min(len(va), len(vb)) / m
    return cos * (0.5 + 0.5 * lr)


def profile_db(path, max_rows=MAX_ROWS):
    """返回 {tab_lower: {"tab": 原名, "cols": {col_lower: prof}}}"""
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
            q = "SELECT " + ",".join(f'"{c}"' for c, _ in cols) + \
                f' FROM "{t}" {_limit_clause(max_rows)}'
            try:
                data = con.execute(q).fetchall()
            except Exception:
                continue
            n = len(data)
            info = {"tab": t, "n": n, "cols": {}}
            for j, (c, typ) in enumerate(cols):
                vals = [_norm(r[j]) for r in data]
                nn = [v for v in vals if v != ""]
                s = set(nn)
                info["cols"][c.lower()] = {
                    "name": c, "pos": j, "type": typ.upper()[:12],
                    "n": n, "nn": len(nn), "ndist": len(s),
                    "set": s, "list": nn[:SUB_SAMPLE],
                    "shape": shape_vec(nn),
                    "freq": dict(collections.Counter(nn)),
                }
            out[t.lower()] = info
    finally:
        con.close()
    return out


# ---------------- 相似度 ---------------- #

def _sims(a, b):
    A, B = a["set"], b["set"]
    if not A or not B:
        return 0.0, 0.0, 0.0
    inter = len(A & B)
    jac = inter / len(A | B)
    return jac, inter / len(A), inter / len(B)


def name_sim(x, y):
    x, y = x.lower(), y.lower()
    if x == y:
        return 1.0
    a = re.sub(r"[^a-z0-9]", "", x)
    b = re.sub(r"[^a-z0-9]", "", y)
    if not a or not b:
        return 0.0
    if a in b or b in a:
        return 0.85
    ta = set(re.split(r"[^a-z0-9]+", x)) - {""}
    tb = set(re.split(r"[^a-z0-9]+", y)) - {""}
    if ta & tb:
        return 0.75
    # 字符级（difflib 调用便宜，只在不同名时走）
    import difflib
    return difflib.SequenceMatcher(None, a, b).ratio()


def sub_cover(a, b):
    """a 的值有多大比例「作为子串」出现在 b 的值里（拆列/并列用）"""
    blob = "\x00".join(b["list"])
    if not blob or not a["list"]:
        return 0.0
    hit = sum(1 for v in a["list"] if v and v in blob)
    return hit / len(a["list"])


def part_order(old_p, first, second):
    """判断 first 是否在旧值里先于 second 出现（0=是，1=否，2=不定）"""
    for v in old_p["list"]:
        f = next((x for x in first["list"] if x and x in v), None)
        s = next((x for x in second["list"] if x and x in v), None)
        if f is None or s is None:
            continue
        if v.find(f) < v.find(s):
            return 0
        if v.find(f) > v.find(s):
            return 1
    return 2


def detect_sep(old_p, first, second, limit=400):
    """推断拆列后重组用的分隔符：旧值 = 前段 + sep + 后段，取 sep 的众数。

    不同库的拆分分隔符不同（空格、'-'、'/' 等），写死任何一个都会让
    下游 SQL 拼不出原值——曾导致 split 档只有 29.8%。
    完全无监督：只用前后段候选值在旧值中的位置关系。
    """
    cnt = collections.Counter()
    fs = sorted({x for x in first["list"] if x}, key=len, reverse=True)[:limit]
    ss = sorted({x for x in second["list"] if x}, key=len, reverse=True)[:limit]
    for v in old_p["list"][:limit]:
        if not v:
            continue
        for p in fs:
            if not v.startswith(p):
                continue
            rest = v[len(p):]
            for s in ss:
                if rest.endswith(s):
                    sep = rest[:len(rest) - len(s)]
                    if 0 < len(sep) <= 4:
                        cnt[sep] += 1
                    break
            if cnt:
                break
    return cnt.most_common(1)[0][0] if cnt else " "


def sql_sep(sep):
    """把分隔符渲染成 SQLite 字面量。"""
    if all(32 <= ord(c) < 127 for c in sep) and "'" not in sep:
        return "'" + sep + "'"
    return " || ".join(f"CHAR({ord(c)})" for c in sep)


# ---------------- 推断 ---------------- #

def pair_values(a, b):
    """按「频率降序、同频按值字典序」配对两侧取值，得到 旧值 -> 新值。

    一一重编码不改变频率结构，所以排序后的位次就是对应关系。
    只有当两侧基数相同才给映射——基数不同说明不是一一重编码，宁可不给。
    """
    fa, fb = a.get("freq") or {}, b.get("freq") or {}
    if not fa or len(fa) != len(fb):
        return {}
    oa = sorted(fa.items(), key=lambda kv: (-kv[1], kv[0]))
    ob = sorted(fb.items(), key=lambda kv: (-kv[1], kv[0]))
    return {ka: kb for (ka, _), (kb, _) in zip(oa, ob)}


def usable_shape(p):
    """形态指纹只对「有重复值」的列有区分力。

    近唯一列（ndist/n→1）的频率向量恒为 [1/n,...,1/n]，任意两个这种列
    形状完全相同，参与打分只会制造平局——实测把 member.first_name 和
    member.last_name 互换了配对。常量列同理。
    """
    n, nd = p.get("n") or 0, p.get("ndist") or 0
    return n > 0 and 2 <= nd and nd / n < 0.90


def infer_table(old_cols, new_cols, mode="profile", use_pos=False):
    """对单表推断 old->new 路由，返回 ops 列表。

    四步：① 同名稳定列跳过 ② 值域互为最佳且 jac≈1 → rename（不受基数限制，
    解决 Currency/SEX 这类低基数列）③ 阈值贪心兜底 ④ 子串方向判 split/merge
    ⑤ 剩余判 deprecate

    mode="shape" 时，把「值域 Jaccard」替换为「max(值域 Jaccard, 频率曲线相似度)」，
    用来覆盖「值被同步重编码」的演化——此时值域可以完全不相交，但频率曲线不变。
    """
    ops = []
    O = {c: p for c, p in old_cols.items() if c not in new_cols}
    N = {c: p for c, p in new_cols.items() if c not in old_cols}
    if not O and not N:
        return ops
    if not N:                      # 只删不增：整批判废弃
        return [{"type": "deprecate_column", "table": None, "old": O[o]["name"],
                 "new": None, "cols": [], "score": 0.0} for o in O]
    if not O:
        return ops

    use_shape = (mode == "shape")
    JAC, SHP, VAL = {}, {}, {}
    for o in O:
        for n in N:
            j = _sims(O[o], N[n])[0]
            s = (shape_sim(O[o], N[n])
                 if use_shape and usable_shape(O[o]) and usable_shape(N[n])
                 else 0.0)
            JAC[(o, n)] = j
            SHP[(o, n)] = s
            VAL[(o, n)] = max(j, s) if use_shape else j

    def score(o, n):
        p, q = O[o], N[n]
        jac = VAL[(o, n)]
        ns = name_sim(p["name"], q["name"])
        if use_pos:
            return (1.0 if p["pos"] == q["pos"] else 0.0), jac, ns
        if mode == "name":
            return ns, jac, ns
        lowcard = min(p["ndist"], q["ndist"]) < LOW_CARD
        if lowcard:
            return 0.55 * ns + 0.45 * jac, jac, ns
        return 0.88 * jac + 0.12 * ns, jac, ns

    used_o, used_n = set(), set()

    # ---- ② 值域（或频率曲线）互为最佳且 ≈1：低基数也认 ----
    if mode in ("profile", "shape") and not use_pos:
        best_for_o, best_for_n = {}, {}
        for o in O:
            mx = max(VAL[(o, n)] for n in N)
            best_for_o[o] = [n for n in N if VAL[(o, n)] >= mx - 1e-9]
        for n in N:
            mx = max(VAL[(o, n)] for o in O)
            best_for_n[n] = [o for o in O if VAL[(o, n)] >= mx - 1e-9]
        for o in O:
            if len(best_for_o[o]) != 1:
                continue
            n = best_for_o[o][0]
            if len(best_for_n[n]) != 1 or best_for_n[n][0] != o:
                continue
            # 方向性复核：拆分出来的新列，其值大量是旧值的「真子串」
            # （'1935-08' / '01' 都出现在 '1935-08-01' 里）；而重编码是
            # 反过来的方向（旧值 'Lake' 出现在新值 'X-Lake' 里）。
            # 形状指纹分不出这两者，必须靠方向判——否则 split 会被误判成 rename。
            # 前置 JAC<0.9 不可省：值域完全相同的纯改名，每个值都是自己的子串，
            # 不加这个条件会把所有 rename/dep_successor 一起误杀。
            if (SHP[(o, n)] > 0 and JAC[(o, n)] < 0.9
                    and sub_cover(N[n], O[o]) >= 0.5):
                continue
            if VAL[(o, n)] >= 0.98:
                used_o.add(o)
                used_n.add(n)
                ops.append({"type": "rename_column", "table": None,
                            "old": O[o]["name"], "new": N[n]["name"],
                            "cols": [N[n]["name"]],
                            "score": 0.98, "jac": round(JAC[(o, n)], 3),
                            "shape": round(SHP[(o, n)], 3),
                            "value_map": (pair_values(O[o], N[n]) if use_shape else {})})

    # ---- ③ 阈值贪心 ----
    cand = []
    for o in O:
        if o in used_o:
            continue
        for n in N:
            if n in used_n:
                continue
            s, jac, ns = score(o, n)
            cand.append((s, jac, ns, o, n))
    cand.sort(key=lambda x: -x[0])
    for s, jac, ns, o, n in cand:
        if o in used_o or n in used_n:
            continue
        # 与 ② 同一处方向性复核：新值大量是旧值的真子串 → 是拆分，留给 ④
        if (use_shape and SHP[(o, n)] > 0 and JAC[(o, n)] < 0.9
                and sub_cover(N[n], O[o]) >= 0.5):
            continue
        p, q = O[o], N[n]
        lowcard = min(p["ndist"], q["ndist"]) < LOW_CARD
        if mode == "name":
            ok = ns >= 0.55
        elif mode == "pos":
            ok = s >= 0.5
        else:
            ok = jac >= 0.90 if not lowcard else (jac >= 0.90 and ns >= 0.4) or ns >= 0.80
        if not ok:
            continue
        used_o.add(o)
        used_n.add(n)
        ops.append({"type": "rename_column", "table": None,
                    "old": p["name"], "new": q["name"], "cols": [q["name"]],
                    "score": round(s, 3), "jac": round(JAC[(o, n)], 3),
                    "shape": round(SHP[(o, n)], 3),
                    "value_map": (pair_values(p, q) if use_shape else {})})

    rest_o = [o for o in O if o not in used_o]
    rest_n = [n for n in N if n not in used_n]

    # ---- ④ 拆列：旧列的值被 2 个新列「各自装下一部分」→ 新值出现在旧值里 ----
    for o in list(rest_o):
        hits = [(sub_cover(N[n], O[o]), n) for n in rest_n]
        hits = sorted([h for h in hits if h[0] >= 0.30], reverse=True)
        if len(hits) >= 2 and sum(h[0] for h in hits[:2]) >= 1.60:
            a, b = hits[0][1], hits[1][1]
            # 拼接顺序：哪一段在旧值里先出现，就先拼
            if part_order(O[o], N[b], N[a]) == 0:
                a, b = b, a
            cols = [N[a]["name"], N[b]["name"]]
            sep = detect_sep(O[o], N[a], N[b])
            ops.append({"type": "split_column", "table": None,
                        "old": O[o]["name"],
                        "new": cols[0] + " || " + sql_sep(sep) + " || " + cols[1],
                        "cols": cols, "sep": sep,
                        "score": round(sum(h[0] for h in hits[:2]), 3)})
            rest_o.remove(o)
            for h in hits[:2]:
                if h[1] in rest_n:
                    rest_n.remove(h[1])

    # ---- ④' 并列：旧列的值出现在同一个新列里（拼接） ----
    for n in list(rest_n):
        hits = [(sub_cover(O[o], N[n]), o) for o in rest_o]
        hits = sorted([h for h in hits if h[0] >= 0.30], reverse=True)
        if len(hits) >= 2:
            for sc, o in hits:
                ops.append({"type": "merge_columns", "table": None,
                            "old": O[o]["name"], "new": N[n]["name"],
                            "cols": [N[n]["name"]], "score": round(sc, 3)})
                rest_o.remove(o)
            rest_n.remove(n)

    # ---- ⑤ 剩余：无后继 = 废弃 ----
    for o in rest_o:
        ops.append({"type": "deprecate_column", "table": None,
                    "old": O[o]["name"], "new": None, "cols": [],
                    "score": 0.0})
    return ops


def infer_db(db_id, evo_type, mode="profile", use_pos=False):
    old_p = profile_db(OLD_DB.format(db=db_id))
    evo_dir = EVO_DIR[evo_type]
    new_p = profile_db(os.path.join(BASE, evo_dir, f"{db_id}.sqlite"))
    ops = []
    for t, info in old_p.items():
        if t not in new_p:
            continue
        sub = infer_table(info["cols"], new_p[t]["cols"], mode=mode, use_pos=use_pos)
        for op in sub:
            op["table"] = info["tab"]
        ops.extend(sub)
    return {"db_id": db_id, "evo_type": evo_type, "mode": mode,
            "use_pos": use_pos, "ops": ops}


# ---------------- gold 与对齐评分 ---------------- #

def gold_ops(rows, evo_type, db_id):
    seen = {}
    for r in rows:
        if r["evo_type"] != evo_type or r["db_id"] != db_id:
            continue
        for op in r.get("ops") or []:
            if not isinstance(op, list) or len(op) < 3:
                continue
            typ, tab, old = op[0], op[1], str(op[2])
            new = str(op[3]) if len(op) > 3 and op[3] else None
            cols = expr_cols(new) if new else []
            seen[(tab.lower(), old.lower())] = {"type": typ, "table": tab,
                                                "old": old, "cols": cols}
    return list(seen.values())


def evaluate(rows, inferred_by_key):
    """按「目标列是否恢复」评：rename/dep_successor 看目标列，dep 看是否判无后继。"""
    stat = collections.defaultdict(lambda: [0, 0, 0])  # recovered, total, type_ok
    detail = []
    for (evo, db), gops in gold_by_keys(rows).items():
        inf = inferred_by_key.get((evo, db), {}).get("ops", [])
        idx = collections.defaultdict(list)
        for op in inf:
            idx[op["table"].lower() + "." + op["old"].lower()].append(op)
        for g in gops:
            key = (g["table"].lower() + "." + g["old"].lower())
            cands = idx.get(key, [])
            if g["type"] in ("rename_column", "rename_and_recode",
                             "deprecate_with_successor",
                             "merge_columns", "split_column"):
                tgt = {c.lower() for c in g["cols"]}
                ok = any({c.lower() for c in op["cols"]} == tgt for op in cands)
                same_type = any(op["type"] == g["type"] or
                                (g["type"] in ("deprecate_with_successor",
                                               "rename_and_recode")
                                 and op["type"] == "rename_column")
                                for op in cands)
            else:  # deprecate_column
                ok = any(op["type"] == "deprecate_column" and not op["cols"]
                         for op in cands)
                same_type = ok
            s = stat[evo]
            s[0] += 1 if ok else 0
            s[1] += 1
            s[2] += 1 if same_type else 0
            detail.append({"evo": evo, "db": db, "table": g["table"],
                           "old": g["old"], "gold": g["type"], "gold_cols": g["cols"],
                           "pred": [{"type": o["type"], "cols": o["cols"]}
                                    for o in cands], "ok": bool(ok)})
    return stat, detail


def gold_by_keys(rows):
    keys = sorted({(r["evo_type"], r["db_id"]) for r in rows})
    return {k: gold_ops(rows, k[0], k[1]) for k in keys}


# ---------------- main ---------------- #

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--all", action="store_true")
    ap.add_argument("--eval", action="store_true")
    ap.add_argument("--mode", default="profile",
                    choices=["profile", "name", "pos", "shape"])
    ap.add_argument("--manifest", default="data/manifest_v3.jsonl")
    ap.add_argument("--evo", default="", help="逗号分隔，只跑指定演化类型")
    ap.add_argument("--sample", default="head",
                    help="采样策略：head（主结果）或 rand<N>（确定性伪随机，做敏感性）")
    ap.add_argument("--db", default="", help="逗号分隔，只跑指定库（分批跑以控内存）")
    args = ap.parse_args()

    global SAMPLE_MODE
    SAMPLE_MODE = args.sample

    rows = [json.loads(l) for l in open(os.path.join(BASE, args.manifest),
                                        encoding="utf-8")]
    only_evo = {x.strip() for x in args.evo.split(",") if x.strip()}
    if only_evo:
        rows = [r for r in rows if r["evo_type"] in only_evo]
    only_db = {x.strip() for x in args.db.split(",") if x.strip()}
    if only_db:
        rows = [r for r in rows if r["db_id"] in only_db]
    gk = gold_by_keys(rows)
    os.makedirs(OUTDIR, exist_ok=True)

    results = {}
    for (evo, db) in sorted(gk):
        obj = infer_db(db, evo, mode=args.mode, use_pos=(args.mode == "pos"))
        tag = f"{args.mode}" + ("_pos" if args.mode == "pos" else "")
        if args.sample != "head":
            # 敏感性分析写入独立文件，绝不覆盖主结果
            tag += f"_{args.sample}"
        p = os.path.join(OUTDIR, f"{evo}_{db}_{tag}.json")
        json.dump(obj, open(p, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
        results[(evo, db)] = obj
        print(f"[{evo:14s}] {db:24s} gold_ops={len(gk[(evo, db)]):3d} "
              f"infer_ops={len(obj['ops']):3d} -> {os.path.basename(p)}", flush=True)

    if args.eval:
        stat, detail = evaluate(rows, results)
        print("\n" + "=" * 66)
        print(f"M1 谱系恢复（mode={args.mode}）  目标列恢复率 / op-type 一致率")
        print("=" * 66)
        tot = [0, 0, 0]
        for evo in sorted(stat):
            r, n, t = stat[evo]
            tot[0] += r
            tot[1] += n
            tot[2] += t
            print(f"{evo:<16}{n:>5} ops  目标恢复 {r:>4}/{n:<4} = {r/n*100:5.1f}%"
                  f"   type 一致 {t/n*100:5.1f}%")
        if tot[1]:
            print(f"{'TOTAL':<16}{tot[1]:>5} ops  目标恢复 {tot[0]:>4}/{tot[1]:<4} = "
                  f"{tot[0]/tot[1]*100:5.1f}%   type 一致 {tot[2]/tot[1]*100:5.1f}%")
        json.dump({"mode": args.mode,
                   "by_evo": {k: {"recovered": v[0], "total": v[1], "type_ok": v[2]}
                              for k, v in stat.items()},
                   "detail": detail},
                  open(os.path.join(OUTDIR, "eval_%s%s.json" % (
                      args.mode,
                      "" if args.sample == "head" else "_" + args.sample)), "w",
                       encoding="utf-8"), ensure_ascii=False, indent=1)
        print(f"\nSAVED: data/lineage_infer/eval_{args.mode}"
              f"{'' if args.sample == 'head' else '_' + args.sample}.json")


if __name__ == "__main__":
    main()
