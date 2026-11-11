# -*- coding: utf-8 -*-
"""实验 E1：过期证据是否有害（Stale Evidence Poisoning）

三个条件，唯一变量是「模型能看到什么证据」：

  C1 SchemaOnly   只给当前 schema + 问题          —— 不知道数据库演化过
  C2 Schema+Stale 当前 schema + 旧 SQL 示例        —— 模拟 RAG 检索到过期证据
  C3 OracleLineage C2 + 谱系映射（old->new 显式）  —— 上界，证明「对齐做对了能到多少」

核心假设：C2 < C1。检索到的过期证据不是「没用」，是「有害」——
模型会照着旧列名生成，或把旧列名猜成当前 schema 里的近似列（→ 静默错误）。

答案空间的两种形态必须分开计：
  rewrite 类（1,800 条）：应产出可执行且结果正确的 SQL
  abstain 类（601 条）：目标已不可达，应输出 CANNOT_ANSWER

用法：
  python run_c1c2c3.py --split dev --limit 120 --cond c1,c2,c3
  python run_c1c2c3.py --split dev --limit 120 --cond c2 --model llm_config.ds.json
"""
import argparse
import collections
import json
import os
import re
import sqlite3
import sys
import time

BASE = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE)
sys.path.insert(0, BASE)

MANIFEST = os.path.join(BASE, "data", "manifest.jsonl")
OUTDIR = os.path.join(BASE, "results")
CANNOT = "CANNOT_ANSWER"
TIMEOUT = 6.0
MAX_SCHEMA_CHARS = 12000
MAX_EVIDENCE = 3          # C2/C3 最多给几条历史示例（本题只有 1 条旧 SQL）

EVOLVED_DIR = {
    "rename": "data/evolved",
    "rename_medium": "data/evolved_medium",
    "rename_hard": "data/evolved_hard",
    "dep_abstain": "data/evolved_dep_abstain",
    "dep_successor": "data/evolved_dep_successor",
    "merge_col": "data/evolved_merge",
    "split_col": "data/evolved_split",
    "value_shift": "data/evolved_vshift",
    "value_shift_named": "data/evolved_vshift_named",
    # 留出库确认实验（BIRD train 划分上另建的 6 库，与 dev 的 11 库不重叠）
    "merge_col_holdout": "data/evolved_merge_holdout",
    "split_col_holdout": "data/evolved_split_holdout",
    "merge_col3_holdout": "data/evolved_merge3_holdout",
}

ORACLE_V_STATS = {
    "value_shift": "evo_stats_vshift.json",
    "value_shift_named": "evo_stats_vshift_named.json",
}

SYSTEM = (
    "You are an expert SQLite analyst. You write SQL that runs on the CURRENT "
    "database schema. Output only SQL, no explanation, no markdown fences."
)


# ---------------- schema / 执行 ---------------- #

def db_path(evo_type, db_id):
    return os.path.join(BASE, EVOLVED_DIR[evo_type], f"{db_id}.sqlite")


def load_schema(path, max_chars=MAX_SCHEMA_CHARS):
    """从 sqlite_master 读当前（演化后）建表语句。"""
    con = sqlite3.connect(f"file:{path}?mode=ro", uri=True)
    rs = con.execute(
        "select name, sql from sqlite_master where type='table' "
        "and name not like 'sqlite_%' order by name").fetchall()
    con.close()
    parts, total = [], 0
    for name, sql in rs:
        if not sql:
            continue
        block = sql.strip() + ";"
        if total + len(block) > max_chars:
            parts.append(f"-- (schema truncated: {len(rs)} tables total)")
            break
        parts.append(block)
        total += len(block)
    return "\n\n".join(parts)


def run_sql(path, sql, limit=TIMEOUT):
    if not sql:
        return None, "EMPTY_SQL"
    if not os.path.exists(path):
        return None, "DB_MISSING"
    con = sqlite3.connect(f"file:{path}?mode=ro", uri=True)
    t0 = time.time()
    con.set_progress_handler(lambda: 1 if time.time() - t0 > limit else 0, 2000)
    try:
        return con.execute(sql).fetchall(), None
    except Exception as e:
        return None, str(e)[:300]
    finally:
        con.close()


def norm(rows):
    if rows is None:
        return None
    return sorted(tuple("" if v is None else str(v) for v in r) for r in rows)


# ---------------- prompt ---------------- #

_INFER_CACHE = {}


def load_inferred(evo_type, db_id, mode="profile"):
    """读 M1 自动抽取的谱系（不是 ground truth）。缺失时返回 None。

    mode="profile" 用值域指纹版，mode="shape" 用值域+形态指纹版，
    mode="part" 在 profile 基础上补合并列的段位（merge_col 专有），
    缺失 _part 文件时回退到 profile——保证任何演化档都能取到谱系。
    """
    key = (evo_type, db_id, mode)
    if key in _INFER_CACHE:
        return _INFER_CACHE[key]
    obj = None
    for m in ([mode, "profile"] if mode == "part" else [mode]):
        p = os.path.join(BASE, "data", "lineage_infer",
                         f"{evo_type}_{db_id}_{m}.json")
        if os.path.exists(p):
            obj = json.load(open(p, encoding="utf-8")).get("ops") or []
            break
    _INFER_CACHE[key] = obj
    return obj


_ORACLE_V_CACHE = {}


def load_oracle_values(evo_type, db_id):
    """值改写的 ground-truth 值映射（只有值改写档有）。

    用来构造 c3v：列名来自 oracle、值映射也来自 oracle。
    与 c3s（两者都由 M1 自动抽取）配对，就能把「值映射」这一项的贡献单独拎出来。
    """
    key = (evo_type, db_id)
    if key in _ORACLE_V_CACHE:
        return _ORACLE_V_CACHE[key]
    vm = {}
    fn = ORACLE_V_STATS.get(evo_type)
    if fn:
        p = os.path.join(BASE, "data", fn)
        if os.path.exists(p):
            vm = (json.load(open(p, encoding="utf-8"))
                  .get("per_db", {}).get(db_id, {}).get("lit_map") or {})
    _ORACLE_V_CACHE[key] = vm
    return vm


def lineage_line(typ, tab, old, new):
    """把一条谱系事实渲染成自然语言。C3(orange) 与 C3m(inferred) 共用。"""
    if typ in ("rename_column", "rename_and_recode"):
        return f"- {tab}.{old} was renamed to {new}"
    if typ == "split_column":
        return f"- {tab}.{old} was split; reconstruct it as {new}"
    if typ in ("merge_columns", "merge_column"):
        if "(" in str(new):
            return f"- {tab}.{old} and friends were merged; {old} is now {new}"
        return f"- {tab}.{old} was merged into {new}"
    if typ == "deprecate_with_successor":
        return f"- {tab}.{old} was DEPRECATED but its data moved to {new}"
    if typ == "deprecate_column":
        return f"- {tab}.{old} was DEPRECATED and removed"
    return None


def _natural_part_line(op, new):
    """段位消融用：只描述「第几段 / 共几段 / 哪个列 / 分隔符字面量」，
    不给出可直接粘贴的 SUBSTR/INSTR 重构表达式。

    分隔符字面量从表达式中解析并保留，目的是让 c3mp_nat 与 c3mp 的
    **语义信息等价**，唯一差异是「是否给出可执行表达式」这一项。
    """
    cols = op.get("cols") or []
    col = cols[0] if cols else "the merged column"
    idx, nseg = op.get("part_idx"), op.get("nseg")
    m = re.search(r"CHAR\((\d+)\)", str(new))
    if m:
        sep_txt = f"CHAR({m.group(1)})"
    else:
        m2 = re.search(r"'([^']*)'", str(new))
        sep_txt = f"'{m2.group(1)}'" if m2 else "an internal separator"
    ordinal = {0: "first", 1: "second", 2: "third", 3: "fourth"}.get(idx)
    if ordinal is None:
        ordinal = f"segment {idx + 1}" if idx is not None else "one of the"
    tail = f" of {nseg} segments" if nseg else " delimiter-separated segments"
    return (f"- {op['table']}.{op['old']} was merged into {col} and is stored as "
            f"the {ordinal}{tail}; the segments are separated by {sep_txt}")


def _render_inferred(ops, r, with_values=False, natural_only=False):
    """把 M1 抽取的谱系渲染成 prompt 行，只保留旧 SQL 里真实出现过的列。

    用子串判断而非标识符切分：`School Name` 这类带空格的列名切开后对不上。
    with_values=True 时额外给出值级映射（仅当不是恒等映射才输出）。
    natural_only=True 时把含表达式的段位谱系降级为自然语言描述（消融条件）。
    """
    if not ops:
        return []
    sql_l = (r.get("old_sql") or "").lower()
    lines = []
    for op in ops:
        if op["old"].lower() not in sql_l:
            continue
        new = op.get("new") or ""
        if natural_only and "(" in str(new):
            ln = _natural_part_line(op, new)
        else:
            ln = lineage_line(op["type"], op["table"], op["old"], new)
        if ln:
            lines.append(ln)
        if not with_values:
            continue
        vm = op.get("value_map") or {}
        used = [k for k in vm if vm[k] != k and ("'" + k + "'").lower() in sql_l]
        if used:
            pairs = ", ".join(f"'{k}' -> '{vm[k]}'" for k in used[:8])
            lines.append(f"- values stored in {op['table']}.{op['old']} are now "
                         f"re-encoded, not equal to the example's literals: {pairs}")
    return lines


def _render_oracle_values(r):
    """c3v 用：列名给 oracle，值映射也给 oracle。"""
    vm = load_oracle_values(r["evo_type"], r["db_id"])
    sql_l = (r.get("old_sql") or "").lower()
    used = [k for k in vm if vm[k] != k and ("'" + k + "'").lower() in sql_l]
    if not used:
        return []
    pairs = ", ".join(f"'{k}' -> '{vm[k]}'" for k in used[:8])
    return [f"- values stored in the columns above are now re-encoded, "
            f"not equal to the example's literals: {pairs}"]


def build_prompt(r, cond, schema):
    """按条件拼证据。C2/C3/C3m 的「历史示例」就是本题在旧 schema 上的 gold SQL，
    这正是 RAG 检索会捞到的东西：同一问题、过期列名。

    C3  = 谱系来自 ground truth（上界，仅 DDL 层）
    C3m = 谱系来自 M1（值域指纹），且只给旧 SQL 里真实出现过的列
    C3s = 谱系来自 M1（值域+形态指纹），额外带值级映射
    C3v = 列名给 oracle、值映射也给 oracle（完整上界，用来分解「值映射」这一项）
    """
    blocks = [
        "### Current database schema (SQLite)",
        schema,
        "",
        "### Question",
        r["question"].strip(),
        "",
    ]

    if cond in ("c2", "c3", "c3f", "c3m", "c3mp", "c3mp_nat", "c3s", "c3v", "c3n70",
                "c3n40", "c3b1", "c3b2", "c3b3"):
        blocks += [
            "### Retrieved example from an EARLIER version of this database",
            "Question: " + r["question"].strip(),
            "SQL: " + (r["old_sql"] or "").strip(),
            "",
        ]

    if cond in ("c3", "c3f"):
        # c3   = oracle 谱系全量（把该库所有演化 op 都写进提示）
        # c3f  = oracle 谱系 + 引用列过滤（与 c3m 完全同一过滤规则，公平对照）
        # 不加 c3f 时 oracle 基线会被"未被任何问题引用的列"噪声拖累，
        # 曾出现 c3m > c3 的反常，属提示精度差异而非谱系质量差异。
        sql_l_c3 = (r.get("old_sql") or "").lower()
        lines = []
        for op in r["ops"] or []:
            if not isinstance(op, list) or len(op) < 4:
                continue
            if cond == "c3f" and str(op[2]).lower() not in sql_l_c3:
                continue
            ln = lineage_line(op[0], op[1], str(op[2]), str(op[3]))
            if ln:
                lines.append(ln)
        if lines:
            blocks += ["### Schema lineage (what changed since the example was written)",
                       "\n".join(lines), ""]

    if cond in ("c3m", "c3mp", "c3mp_nat", "c3n70", "c3n40", "c3s", "c3b1", "c3b2", "c3b3"):
        # c3nXX 按 XX% 的保留率随机丢弃若干条，用来测「谱系质量 → 下游性能」的
        # 敏感度曲线（不需要重跑 M1）
        # c3mp = c3m + 合并列段位（merge_col 档专有，由 m1_part.py 无监督推断），
        #        谱系行直接给出可执行的 SUBSTR/INSTR 重构表达式
        # c3mp_nat = 消融：与 c3mp 信息等价，但只给段位自然语言描述
        #            （段号/总段数/目标列/分隔符字面量），**不给可执行表达式**
        # c3b1/c3b2/c3b3 = 三个自动谱系基线（名字 / 值域重合 / 嵌入），
        #                 用于回答「增益来自 M1 设计还是任何合理匹配都能达到」
        import hashlib
        keep = {"c3m": 100, "c3mp": 100, "c3mp_nat": 100, "c3n70": 70, "c3n40": 40,
                "c3s": 100, "c3b1": 100, "c3b2": 100, "c3b3": 100}[cond]
        mode = {"c3s": "shape", "c3mp": "part", "c3mp_nat": "part", "c3b1": "name",
                "c3b2": "vover", "c3b3": "emb"}.get(cond, "profile")
        ops = load_inferred(r["evo_type"], r["db_id"], mode=mode)
        if keep < 100 and ops:
            ops = [op for op in ops
                   if int(hashlib.md5(f"{r['db_id']}|{op['old']}".encode())
                          .hexdigest()[:8], 16) % 100 < keep]
        lines = _render_inferred(ops, r, with_values=(cond == "c3s"),
                                 natural_only=(cond == "c3mp_nat"))
        if lines:
            blocks += ["### Schema lineage (what changed since the example was written)",
                       "\n".join(lines), ""]

    if cond == "c3v":
        lines = []
        for op in r["ops"] or []:
            if not isinstance(op, list) or len(op) < 4:
                continue
            ln = lineage_line(op[0], op[1], str(op[2]), str(op[3]))
            if ln:
                lines.append(ln)
        lines += _render_oracle_values(r)
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


# ---------------- 输出解析 ---------------- #

FENCE = re.compile(r"```(?:sql|sqlite)?\s*(.*?)```", re.S | re.I)


def parse_sql(text):
    """返回 (sql_or_None, is_abstain)"""
    if not text:
        return None, True
    t = text.strip()
    m = FENCE.search(t)
    if m:
        t = m.group(1).strip()
    # 去掉开头的 SQL: 前缀
    t = re.sub(r"^(?:sql|SQL)\s*[:：]\s*", "", t).strip()
    up = t.upper()
    if CANNOT in up:
        return None, True
    if up in ("NONE", "NULL", "N/A", "NA"):
        return None, True
    # 没有任何 SELECT/WITH 关键字 → 视为弃答性回答
    if not re.search(r"\b(select|with)\b", t, re.I):
        # 打印诊断（前 3 次）
        if parse_sql.n_diag < 3:
            print(f"    [diag] 非 SQL 输出: {t[:120]!r}", flush=True)
            parse_sql.n_diag += 1
        return None, True
    # 取第一段完整语句
    t = t.split(";")[0].strip()
    return t, False


parse_sql.n_diag = 0


# ---------------- 判定 ---------------- #

def judge(r, sql, is_abstain, gold_rows, evo):
    """返回 outcome 标签。answer space 分两类。"""
    expect = r["expected_action"]
    actual_abstain = is_abstain or not sql

    if expect == "abstain":
        if actual_abstain:
            return "ABSTAIN_OK"          # 正确弃答
        rows, err = run_sql(db_path(evo, r["db_id"]), sql)
        if err:
            return "ABSTAIN_BY_ERROR"     # 想答但答不出（也不算幻觉成灾）
        return "HALLUCINATE"              # 目标已不存在却给出了结果：最危险
    # rewrite 类
    if actual_abstain:
        return "OVER_ABSTAIN"             # 题目可答却弃答
    rows, err = run_sql(db_path(evo, r["db_id"]), sql)
    if err:
        return "ERROR"
    if gold_rows is not None and norm(rows) == norm(gold_rows):
        return "EX"
    return "SILENT"


# ---------------- 主流程 ---------------- #

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--split", default="dev", choices=["train", "dev", "test", "all"])
    ap.add_argument("--limit", type=int, default=0, help="0=全量；>0 按 evo 类型分层抽样")
    ap.add_argument("--cond", default="c1,c2,c3")
    ap.add_argument("--model", default="llm_config.ds.json")
    ap.add_argument("--tag", default="")
    ap.add_argument("--manifest", default="data/manifest.jsonl")
    ap.add_argument("--evo", default="", help="只跑指定 evo_type，逗号分隔")
    ap.add_argument("--cache", default="")
    ap.add_argument("--ids", default="", help="只跑这些 id（每行一个），用于与已有结果配对")
    ap.add_argument("--db", default="", help="只跑指定 db_id，逗号分隔（用于按库并行）")
    args = ap.parse_args()

    os.environ["LLM_CFG_PATH"] = args.model
    import llm_client as L
    _cfg = L.load_cfg()
    print(f"实际模型: {_cfg.get('model')} @ {_cfg.get('base')}", flush=True)

    conds = [c.strip().lower() for c in args.cond.split(",") if c.strip()]
    if args.cache:
        os.environ["LLM_CACHE_PATH"] = args.cache
    rows = [json.loads(l) for l in open(os.path.join(BASE, args.manifest),
                                        encoding="utf-8")]
    if args.evo:
        keep = {x.strip() for x in args.evo.split(",") if x.strip()}
        rows = [r for r in rows if r["evo_type"] in keep]
    if args.split != "all":
        rows = [r for r in rows if r["split"] == args.split]
    rows = [r for r in rows if r.get("check_pass", True)]

    if args.ids:
        keep_ids = {l.strip() for l in open(os.path.join(BASE, args.ids),
                                            encoding="utf-8") if l.strip()}
        rows = [r for r in rows if r["id"] in keep_ids]

    if args.db:
        keep_db = {x.strip() for x in args.db.split(",") if x.strip()}
        rows = [r for r in rows if r["db_id"] in keep_db]

    if args.limit:
        by = collections.defaultdict(list)
        for r in rows:
            by[r["evo_type"]].append(r)
        picked = []
        per = max(1, args.limit // len(by))
        for evo, lst in by.items():
            picked += lst[:per]
        rows = picked

    print(f"样本 {len(rows)} | 条件 {conds} | 模型 {args.model}", flush=True)
    print(f"evo 分布: {dict(collections.Counter(r['evo_type'] for r in rows))}", flush=True)

    # 预取 gold 结果（只算一次）
    gold = {}
    for r in rows:
        p = db_path(r["evo_type"], r["db_id"])
        g_rows, g_err = run_sql(p, r["new_sql"] or "")
        gold[r["id"]] = (g_rows, g_err)

    os.makedirs(OUTDIR, exist_ok=True)
    tag = args.tag or f"{args.split}_{len(rows)}"
    out_path = os.path.join(OUTDIR, f"e1_{tag}.jsonl")
    done = set()
    if os.path.exists(out_path):
        for l in open(out_path, encoding="utf-8"):
            try:
                done.add(json.loads(l)["key"])
            except Exception:
                pass
        print(f"续跑：已有 {len(done)} 条", flush=True)

    fout = open(out_path, "a", encoding="utf-8")
    stat = collections.defaultdict(collections.Counter)
    t_start = time.time()
    n_call = 0

    for i, r in enumerate(rows, 1):
        evo = r["evo_type"]
        p = db_path(evo, r["db_id"])
        schema = load_schema(p)
        g_rows, g_err = gold[r["id"]]

        for cond in conds:
            key = f"{cond}|{r['id']}"
            if key in done:
                continue
            prompt = build_prompt(r, cond, schema)
            try:
                text, meta = L.chat(SYSTEM, prompt)
            except Exception as e:
                print(f"  [llm-err] {key}: {type(e).__name__} {str(e)[:120]}", flush=True)
                stat[cond]["LLM_ERROR"] += 1
                continue
            n_call += 1
            sql, is_ab = parse_sql(text)
            outcome = judge(r, sql, is_ab, g_rows, evo)

            stat[cond][outcome] += 1
            stat[cond][f"__{r['expected_action']}"] += 1
            fout.write(json.dumps({
                "key": key, "cond": cond, "id": r["id"], "evo_type": evo,
                "expected": r["expected_action"], "outcome": outcome,
                "sql": sql, "raw": (text or "")[:1500],
                "db_id": r["db_id"], "question_id": r["question_id"],
                "difficulty": r["difficulty"],
                "model": meta.get("model") if isinstance(meta, dict) else None,
            }, ensure_ascii=False) + "\n")
            fout.flush()

        if i % 20 == 0:
            el = time.time() - t_start
            print(f"  ...{i}/{len(rows)} 调用 {n_call} 条 已用 {el/60:.1f}min "
                  f"({el/max(n_call,1):.1f}s/次)", flush=True)

    fout.close()

    # ---------------- 汇总 ----------------
    print("\n" + "=" * 82)
    print("E1  过期证据是否有害")
    print("=" * 82)
    print(f"{'条件':<6}{'样本':>6}{'EX':>6}{'SILENT':>8}{'ERROR':>7}"
          f"{'OVER_ABS':>9}{'ABST_OK':>9}{'HALLUC':>8}{'ABS_BY_ERR':>11}")

    summary = {}
    for cond in conds:
        c = stat[cond]
        n = sum(v for k, v in c.items() if not k.startswith("__"))
        if not n:
            continue
        print(f"{cond.upper():<6}{n:>6}{c['EX']:>6}{c['SILENT']:>8}{c['ERROR']:>7}"
              f"{c['OVER_ABSTAIN']:>9}{c['ABSTAIN_OK']:>9}{c['HALLUCINATE']:>8}"
              f"{c['ABSTAIN_BY_ERROR']:>11}")
        summary[cond] = {
            "n": n, "EX": c["EX"], "SILENT": c["SILENT"], "ERROR": c["ERROR"],
            "OVER_ABSTAIN": c["OVER_ABSTAIN"], "ABSTAIN_OK": c["ABSTAIN_OK"],
            "HALLUCINATE": c["HALLUCINATE"], "ABSTAIN_BY_ERROR": c["ABSTAIN_BY_ERROR"],
            "LLM_ERROR": c["LLM_ERROR"],
            "n_rewrite": c["__rewrite"], "n_abstain": c["__abstain"],
        }
        # 两个关键率
        nw = c["__rewrite"]
        na = c["__abstain"]
        if nw:
            summary[cond]["EX_rate"] = round(c["EX"] / nw, 4)
            summary[cond]["silent_rate"] = round(c["SILENT"] / nw, 4)
        if na:
            summary[cond]["abstain_acc"] = round(c["ABSTAIN_OK"] / na, 4)

    out = {"experiment": "E1_stale_evidence", "model": args.model,
           "split": args.split, "generated_at": time.strftime("%Y-%m-%d %H:%M:%S"),
           "by_cond": summary}
    jp = os.path.join(OUTDIR, f"e1_{tag}_summary.json")
    json.dump(out, open(jp, "w", encoding="utf-8"), ensure_ascii=False, indent=2)
    print(f"\nSAVED: {out_path}\nSAVED: {jp}")

    # 关键对比
    if "c1" in summary and "c2" in summary:
        d_ex = (summary["c2"].get("EX_rate", 0) - summary["c1"].get("EX_rate", 0)) * 100
        print(f"\nC2 - C1  EX 变化: {d_ex:+.1f} pt  "
              f"（负值＝过期证据有害，核心假设成立）")


if __name__ == "__main__":
    main()
