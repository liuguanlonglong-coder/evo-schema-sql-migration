# -*- coding: utf-8 -*-
"""在 BIRD dev 之上构建 schema 演化集（v2：rename + deprecate）。

三类切片：
  rename        L1  列改名      -> 正确动作 = 重写（映射可在 DDL diff 中发现）
  dep_abstain   L4a 列删除且无后继 -> 正确动作 = 弃答（不存在可写目标）
  dep_successor L4b 列删除但有后继（值完全相同的同义列）-> 正确动作 = 重写，
                    且后继列名与旧名无字面相似性，只能靠值重叠/谱系发现

流程（与 v1 一致）：
  1. 选候选列：gold SQL 中被引用 >=2 次、且列名全库唯一归属
  2. sqlite backup 复制库 -> 施加 DDL 演化
  3. sqlglot 重写旧 SQL（大小写不敏感）
  4. 三重验证：旧 SQL 在原库成功 / 旧 SQL 在新库失效 / 新 SQL 在新库成功且结果等价

用法：
  python build_evolution_v2.py --task rename --max-mb 700
  python build_evolution_v2.py --task dep_abstain
  python build_evolution_v2.py --task dep_successor
"""
import argparse, json, os, re, sqlite3, time, collections
import sqlglot
from sqlglot import exp

BASE = os.path.dirname(os.path.abspath(__file__))
# 默认仍是 BIRD dev；留出库确认实验通过环境变量指向 train 划分，避免复制脚本。
BIRD = os.environ.get("BIRD_ROOT") or os.path.join(BASE, "data", "bird", "dev_20240627")
DBDIR = os.path.join(BIRD, os.environ.get("BIRD_DBDIR") or "dev_databases")
BIRD_JSON = os.path.join(BIRD, os.environ.get("BIRD_JSON") or "dev.json")

# 留出实验必须写独立目录，否则会覆盖 dev 上已冻结的演化库。
_EVO_SUFFIX = os.environ.get("EVO_SUFFIX", "")


def _sfx(name):
    if not _EVO_SUFFIX:
        return name
    stem, ext = os.path.splitext(name)
    return stem + _EVO_SUFFIX + ext


OUT = {
    "rename":        (os.path.join(BASE, "data", "evolved" + _EVO_SUFFIX),
                      os.path.join(BASE, "data", _sfx("evolution_set.jsonl")),
                      os.path.join(BASE, "data", _sfx("evo_stats.json"))),
    # 注意：两类 dep 必须写不同目录。共用目录时后跑的会覆盖先跑的，
    # 导致另一类的 manifest 指向「列其实还在」的库（曾在 v2 里踩过）。
    "dep_abstain":   (os.path.join(BASE, "data", "evolved_dep_abstain" + _EVO_SUFFIX),
                      os.path.join(BASE, "data", _sfx("evolution_dep_abstain.jsonl")),
                      os.path.join(BASE, "data", _sfx("evo_stats_dep_abstain.json"))),
    "dep_successor": (os.path.join(BASE, "data", "evolved_dep_successor" + _EVO_SUFFIX),
                      os.path.join(BASE, "data", _sfx("evolution_dep_successor.jsonl")),
                      os.path.join(BASE, "data", _sfx("evo_stats_dep_successor.json"))),
    "merge_col":     (os.path.join(BASE, "data", "evolved_merge" + _EVO_SUFFIX),
                      os.path.join(BASE, "data", _sfx("evolution_merge.jsonl")),
                      os.path.join(BASE, "data", _sfx("evo_stats_merge.json"))),
    "split_col":     (os.path.join(BASE, "data", "evolved_split" + _EVO_SUFFIX),
                      os.path.join(BASE, "data", _sfx("evolution_split.jsonl")),
                      os.path.join(BASE, "data", _sfx("evo_stats_split.json"))),
    # 值改写档：改名 + 值重编码。M1 的值域指纹在此必然失效，用于划定方法边界。
    # 两档用于 2x2 消融：中性名（名字/值域双失效）vs 保留名（仅值域失效）。
    "value_shift":   (os.path.join(BASE, "data", "evolved_vshift" + _EVO_SUFFIX),
                      os.path.join(BASE, "data", _sfx("evolution_vshift.jsonl")),
                      os.path.join(BASE, "data", _sfx("evo_stats_vshift.json"))),
    "value_shift_named": (os.path.join(BASE, "data", "evolved_vshift_named" + _EVO_SUFFIX),
                      os.path.join(BASE, "data", _sfx("evolution_vshift_named.jsonl")),
                      os.path.join(BASE, "data", _sfx("evo_stats_vshift_named.json"))),
}
EXPR_TASKS = {"merge_col", "split_col"}
VSHIFT_TASKS = {"value_shift", "value_shift_named"}
SEP = "CHAR(31)"        # merge 用的不可见分隔符
SPACE = "' '"           # split 默认分隔符
SPLIT_SEPS = ["' '", "', '", "','", "'-'", "'/'", "'|'", "'_'"]   # 按真实感排序

K_COLS = 3
QUERY_TIMEOUT = 4.0
MAX_UPDATE_ROWS = 400_000     # dep_successor 需 UPDATE 全表，超大表跳过
MAX_DROP_MB = 130             # DROP COLUMN 会重写整表，大库跳过


def sqlite_path(db):
    return os.path.join(DBDIR, db, f"{db}.sqlite")


def connect(path, ro=False):
    if ro:
        return sqlite3.connect(f"file:{path}?mode=ro", uri=True)
    return sqlite3.connect(path)


def list_schema(path):
    con = connect(path, ro=True)
    tabs = [r[0] for r in con.execute(
        "select name from sqlite_master where type='table' and name not like 'sqlite_%'")]
    cols = []
    for t in tabs:
        for r in con.execute(f'PRAGMA table_info("{t}")'):
            cols.append((t, r[1], r[2], r[3], r[5]))   # table, name, type, notnull, pk
    con.close()
    return tabs, cols


def ddl_map(conn):
    out = {}
    for t in [r[0] for r in conn.execute(
            "select name from sqlite_master where type='table' and name not like 'sqlite_%'")]:
        r = conn.execute(
            "select sql from sqlite_master where type='table' and name=?", (t,)).fetchone()
        if r:
            out[t] = r[0]
    return out


def norm_name(col, table, used):
    """生成实质不同的新列名。SQLite 标识符大小写不敏感，仅改大小写不算演化。"""
    if re.search(r"[^0-9a-z_]", col):
        new = re.sub(r"[^0-9A-Za-z]+", "_", col).strip("_").lower()
    else:
        new = f"{table}_{col}"
    if new.lower() == col.lower():
        new = f"{table}_{new}"
    while new.lower() in used or new.lower() == col.lower():
        new += "_x"
    return new


def opaque_name(idx, used):
    """与旧名无字面相似性的中性名，迫使系统靠数据重叠而非字符串相似度发现后继。"""
    new = f"attr_{idx:03d}"
    while new.lower() in used:
        idx += 1
        new = f"attr_{idx:03d}"
    return new


def run_sql(path, sql, timeout=QUERY_TIMEOUT):
    con = connect(path, ro=True)
    t0 = time.time()
    con.set_progress_handler(lambda: 1 if time.time() - t0 > timeout else 0, 2000)
    try:
        rows = con.execute(sql).fetchall()
        con.close()
        return rows, None
    except Exception as e:
        con.close()
        msg = str(e)
        return None, "TIMEOUT" if "interrupted" in msg.lower() else msg


IDENT_RE = re.compile(r"[A-Za-z_][A-Za-z0-9_$]*")
PLAIN_IDENT_RE = re.compile(r"^[A-Za-z_][A-Za-z0-9_$]*$")


def _strip_trailing_qualifier(s):
    """把 'T1.' / '"T1".' 这类限定前缀去掉——表达式替换后限定符会变成语法错误。"""
    t = s.rstrip()
    if not t.endswith("."):
        return s
    t = t[:-1].rstrip()
    m = re.search(r'("[^"]*"|`[^`]*`|\[[^\]]*\]|[A-Za-z_][A-Za-z0-9_$]*)$', t)
    if m:
        t = t[:m.start()]          # 保留限定符之前的空白，避免 SELECT( 粘连
    return t


def rewrite(sql, col_map, lit_map=None):
    """标识符级重写：只替换列名 token，其余字节原样保留。

    两种替换形态：
      - 映射到普通标识符：原地换名（rename / dep_successor）
      - 映射到 SQL 表达式：整体替换并加括号，同时去掉限定前缀
        （merge_col / split_col —— 重写必须引入表达式，这是与纯改名任务的本质差别）

    lit_map 非空时同时重写字符串字面量（value_shift：值重编码后
    `WHERE county = 'Alameda'` 必须变成 `WHERE county = '1'`）。
    只做整串精确匹配，不做模糊替换——避免把不相关列的同名值改掉。

    不用 sqlglot 的 tree.sql()：它会做方言归一化（实测把 SUBSTR 改成 SUBSTRING），
    那不是 schema 演化，会污染重写任务的监督信号。
    单引号字面量整体跳过，避免改坏 'Fresno County Office of Education' 里的 County。
    """
    if not col_map and not lit_map:
        return sql
    cmap = {k.lower(): v for k, v in col_map.items()}
    lit_map = lit_map or {}
    res, i, n = [], 0, len(sql)
    while i < n:
        ch = sql[i]
        if ch == "'":                              # 字符串字面量
            j = i + 1
            while j < n:
                if sql[j] == "'":
                    if j + 1 < n and sql[j + 1] == "'":
                        j += 2
                        continue
                    j += 1
                    break
                j += 1
            lit = sql[i + 1:j - 1] if j - 1 > i else ""
            rep = lit_map.get(lit)
            if rep is not None and "''" not in lit:
                res.append("'" + rep + "'")
            else:
                res.append(sql[i:j])
            i = j
            continue
        if ch in '"`[':                            # 引号标识符
            close = '"' if ch == '"' else ("`" if ch == "`" else "]")
            j = sql.find(close, i + 1)
            if j == -1:
                res.append(sql[i:])
                break
            inner = sql[i + 1:j]
            rep = cmap.get(inner.lower())
            if rep is None:
                res.append(sql[i:j + 1])
            elif PLAIN_IDENT_RE.match(rep):
                res.append(ch + rep + close)
            else:
                res.append("(" + rep + ")")
            i = j + 1
            continue
        m = IDENT_RE.match(sql, i)
        if m:
            w = m.group(0)
            rep = cmap.get(w.lower())
            if rep is None:
                res.append(w)
            elif PLAIN_IDENT_RE.match(rep):
                res.append(rep)
            else:
                s = _strip_trailing_qualifier("".join(res))
                res = [s, "(" + rep + ")"]
            i = m.end()
            continue
        res.append(ch)
        i += 1
    return "".join(res)


def touched_cols(sql, keys):
    """返回该 SQL 引用到的、属于 keys 的列名（小写）"""
    try:
        tree = sqlglot.parse_one(sql, read="sqlite")
    except Exception:
        return set()
    kl = {k.lower() for k in keys}
    return {c.name.lower() for c in tree.find_all(exp.Column) if c.name and c.name.lower() in kl}


def pick_candidates(db, rows, cols, k, tabs):
    """列名全库唯一归属 + gold SQL 引用 >=2 次 + 非主键 + 不与任何表名同名"""
    freq = collections.Counter()
    for it in rows:
        try:
            tree = sqlglot.parse_one(it["SQL"], read="sqlite")
        except Exception:
            continue
        for c in tree.find_all(exp.Column):
            if c.name:
                freq[c.name] += 1
    owners = collections.defaultdict(set)
    pk = {}
    for t, c, _typ, _nn, p in cols:
        owners[c].add(t)
        pk[c] = p
    colset = {c for _, c, _, _, _ in cols}
    tabset = {t.lower() for t in tabs}
    cand = [(c, n) for c, n in freq.most_common()
            if c in colset and n >= 2 and len(owners[c]) == 1 and not pk.get(c)
            and c.lower() not in tabset]
    return cand[:k], owners, colset


def freq_map(rows):
    freq = collections.Counter()
    for it in rows:
        try:
            tree = sqlglot.parse_one(it["SQL"], read="sqlite")
        except Exception:
            continue
        for c in tree.find_all(exp.Column):
            if c.name:
                freq[c.name] += 1
    return freq


def is_text_affinity(typ):
    u = (typ or "").upper()
    return ("CHAR" in u) or ("TEXT" in u) or ("CLOB" in u)


def all_text(conn, t, c):
    """按实际存储类型判定：所有非空值 typeof 都是 text。
    DATE/DATETIME 这类声明在 SQLite 里常按文本存，也能无损拆分。"""
    try:
        r = conn.execute(
            f'SELECT DISTINCT typeof("{c}") FROM "{t}" WHERE "{c}" IS NOT NULL').fetchall()
    except Exception:
        return False
    return [x[0] for x in r] == ["text"]


def no_null(conn, t, c):
    try:
        return conn.execute(
            f'SELECT 1 FROM "{t}" WHERE "{c}" IS NULL LIMIT 1').fetchone() is None
    except Exception:
        return False


def pick_merge_pair(conn, rows, cols):
    """选同表两列合并。实际存储为文本、全库唯一归属、非主键、两列均无 NULL。"""
    freq = freq_map(rows)
    owners, pk = collections.defaultdict(set), {}
    for t, c, _typ, _nn, p in cols:
        owners[c].add(t)
        pk[c] = p
    bytable = collections.defaultdict(list)
    for t, c, typ, _nn, p in cols:
        if p or len(owners[c]) != 1 or freq.get(c, 0) < 1:
            continue
        bytable[t].append((freq[c], c))
    scored = []
    for t, lst in bytable.items():
        if len(lst) < 2:
            continue
        lst.sort(reverse=True)
        scored.append((lst[0][0] + lst[1][0], t, [c for _f, c in lst]))
    scored.sort(reverse=True)
    for score, t, cand_cols in scored:
        txt = [c for c in cand_cols if all_text(conn, t, c) and no_null(conn, t, c)]
        if len(txt) >= 2:
            return {"table": t, "a": txt[0], "b": txt[1]}
    return None


def pick_split_col(conn, rows, cols):
    """选可无损拆分的文本列：存在某个分隔符出现在所有非空值中，且不同值 >= 5。"""
    freq = freq_map(rows)
    owners, pk = collections.defaultdict(set), {}
    for t, c, _typ, _nn, p in cols:
        owners[c].add(t)
        pk[c] = p
    cands = sorted(
        [(freq[c], t, c) for t, c, typ, _nn, p in cols
         if not p and len(owners[c]) == 1 and freq.get(c, 0) >= 1],
        reverse=True)
    for _f, t, c in cands:
        if not all_text(conn, t, c):
            continue
        try:
            nd = conn.execute(f'SELECT COUNT(DISTINCT "{c}") FROM "{t}"').fetchone()[0]
        except Exception:
            continue
        if nd < 5:
            continue
        for sep in SPLIT_SEPS:
            try:
                bad = conn.execute(
                    f"""SELECT 1 FROM "{t}" WHERE "{c}" IS NOT NULL
                        AND INSTR("{c}", {sep}) = 0 LIMIT 1""").fetchone()
            except Exception:
                continue
            if bad:
                continue
            return {"table": t, "col": c, "sep": sep}
    return None


def copy_db(src, dst):
    if os.path.exists(dst):
        os.remove(dst)
    a = connect(src, ro=True)
    b = connect(dst)
    a.backup(b)
    a.close()
    b.close()


# ---------------- 三类切片的 DDL 施加 ---------------- #

def apply_rename(conn, cols, cand, colset):
    picked = [c for c, _ in cand]
    col_map, used, applied = {}, {c.lower() for c in colset}, {}
    for c in picked:
        table = [t for t, cc, *_ in cols if cc == c][0]
        new = norm_name(c, table, used)
        used.add(new.lower())
        col_map[c] = new
    for old, new in col_map.items():
        t = [tt for tt, cc, *_ in cols if cc == old][0]
        try:
            conn.execute(f'ALTER TABLE "{t}" RENAME COLUMN "{old}" TO "{new}"')
            applied[old] = {"table": t, "new": new, "op": "rename_column"}
        except Exception as e:
            print(f"  [warn] {t}.{old} 改名失败: {e}", flush=True)
    return applied


def apply_deprecate(conn, cols, cand, mode):
    """mode='dep_abstain' 直接删除；'dep_successor' 先建同值新列再删旧列。"""
    applied = {}
    used = {c.lower() for _, c, *_ in cols}
    idx = 0
    for c, _n in cand:
        metas = [(t, typ, nn, p) for t, cc, typ, nn, p in cols if cc == c]
        t, typ, nn, p = metas[0]
        if p:
            continue
        if mode == "dep_abstain":
            try:
                conn.execute(f'ALTER TABLE "{t}" DROP COLUMN "{c}"')
                applied[c] = {"table": t, "new": None, "op": "deprecate_column"}
            except Exception as e:
                print(f"  [warn] {t}.{c} 删除失败: {e}", flush=True)
        else:
            # 行数保护
            try:
                nrow = conn.execute(f'SELECT COUNT(*) FROM "{t}"').fetchone()[0]
            except Exception:
                continue
            if nrow > MAX_UPDATE_ROWS:
                print(f"  [skip] {t} 行数 {nrow} > {MAX_UPDATE_ROWS}", flush=True)
                continue
            idx += 1
            new = opaque_name(idx, used)
            used.add(new)
            try:
                conn.execute(f'ALTER TABLE "{t}" ADD COLUMN "{new}" {typ or ""}'.strip())
                conn.execute(f'UPDATE "{t}" SET "{new}" = "{c}"')
                conn.execute(f'ALTER TABLE "{t}" DROP COLUMN "{c}"')
                applied[c] = {"table": t, "new": new, "op": "deprecate_with_successor"}
            except Exception as e:
                print(f"  [warn] {t}.{c} 后继替换失败: {e}", flush=True)
    return applied


def apply_merge(conn, spec):
    """两列合并为一列（分隔符拼接），旧列删除。
    重写 A 需 SUBSTR(...INSTR...) —— 只换名字做不到。
    """
    t, a, b = spec["table"], spec["a"], spec["b"]
    try:
        bad = conn.execute(
            f'SELECT 1 FROM "{t}" WHERE "{a}" IS NULL OR "{b}" IS NULL LIMIT 1').fetchone()
        if bad:
            print(f"  [skip] {t}.{a}+{b} 含 NULL，无法无损合并", flush=True)
            return {}
        nrow = conn.execute(f'SELECT COUNT(*) FROM "{t}"').fetchone()[0]
        if nrow > MAX_UPDATE_ROWS:
            print(f"  [skip] {t} 行数 {nrow} > {MAX_UPDATE_ROWS}", flush=True)
            return {}
        m = opaque_name(1, {a.lower(), b.lower()})
        conn.execute(f'ALTER TABLE "{t}" ADD COLUMN "{m}" TEXT')
        conn.execute(f'UPDATE "{t}" SET "{m}" = "{a}" || {SEP} || "{b}"')
        conn.execute(f'ALTER TABLE "{t}" DROP COLUMN "{a}"')
        conn.execute(f'ALTER TABLE "{t}" DROP COLUMN "{b}"')
    except Exception as e:
        print(f"  [warn] merge {t}.{a}+{b} 失败: {e}", flush=True)
        return {}
    return {
        a: {"table": t, "new": f"SUBSTR({m},1,INSTR({m},{SEP})-1)", "op": "merge_columns"},
        b: {"table": t, "new": f"SUBSTR({m},INSTR({m},{SEP})+1)", "op": "merge_columns"},
    }


def apply_split(conn, spec):
    """一列拆为两列（按首个分隔符），旧列删除。
    重写 C 需 (a || sep || b) —— 只换名字做不到。
    """
    t, c, sep = spec["table"], spec["col"], spec["sep"]
    try:
        nrow = conn.execute(f'SELECT COUNT(*) FROM "{t}"').fetchone()[0]
        if nrow > MAX_UPDATE_ROWS:
            print(f"  [skip] {t} 行数 {nrow} > {MAX_UPDATE_ROWS}", flush=True)
            return {}
        a = opaque_name(1, {c.lower()})
        b = opaque_name(2, {c.lower(), a})
        conn.execute(f'ALTER TABLE "{t}" ADD COLUMN "{a}" TEXT')
        conn.execute(f'ALTER TABLE "{t}" ADD COLUMN "{b}" TEXT')
        conn.execute(
            f"""UPDATE "{t}" SET
                "{a}" = SUBSTR("{c}", 1, INSTR("{c}", {sep}) - 1),
                "{b}" = SUBSTR("{c}", INSTR("{c}", {sep}) + LENGTH({sep}))""")
        conn.execute(f'ALTER TABLE "{t}" DROP COLUMN "{c}"')
    except Exception as e:
        print(f"  [warn] split {t}.{c} 失败: {e}", flush=True)
        return {}
    return {c: {"table": t, "new": f"{a} || {sep} || {b}", "op": "split_column"}}


# ---------------- 值改写档：改名 + 值重编码 ---------------- #

def apply_value_shift(conn, cols, cand, named=False):
    """列改名 + 取值重编码：每个旧值 v 变成 'X-<v>'。

    为什么不映射成纯数字：值域一旦变成 '1','2',...，问题里提到的实体名
    （'Fresno'、'Alameda'）就再也对不上任何取值，任务从「需要谱系翻译」
    退化成「不可回答」，端到端测的就不是我们想测的东西了。
    加统一前缀后：字面量层面完全不可复用（值域 Jaccard 严格为 0），
    但语义仍可从值域读出，题目依然可回答。

    为什么按频次降序取映射：映射是一一对应的，所以「值的频率结构」
    （排序后的归一化频次向量）在重编码前后不变。这是形态指纹能救回来的前提。

    为什么候选列的值集合必须两两不相交：字面量重写靠整串精确匹配，
    两列若共享取值，同一个字面量就无法判定属于哪一列。
    """
    applied, used_vals = {}, set()
    used = {c.lower() for _, c, *_ in cols}
    idx = 0
    for c, _n in cand:
        metas = [(t, typ, nn, p) for t, cc, typ, nn, p in cols if cc == c]
        t, typ, nn, p = metas[0]
        if p:
            continue
        try:
            nrow = conn.execute(f'SELECT COUNT(*) FROM "{t}"').fetchone()[0]
        except Exception:
            continue
        if nrow > MAX_UPDATE_ROWS:
            continue
        try:
            vals = conn.execute(
                f'''SELECT "{c}", COUNT(*) FROM "{t}" WHERE "{c}" IS NOT NULL
                    GROUP BY "{c}" ORDER BY COUNT(*) DESC, "{c}"''').fetchall()
        except Exception:
            continue
        # 空串不参与重编码：指纹侧会把空串当缺失值过滤掉，若这里把它改成
        # 'X-'，新旧库的基数就不一致了（thrombosis.SEX 曾因此多出一个 1.1% 的值）
        vals = [(v, cnt) for v, cnt in vals if str(v) != ""]
        nd = len(vals)
        if nd < 3 or nd > 200:          # 过低歧义大，过高映射表臃肿
            continue
        if max((len(str(v)) for v, _ in vals), default=0) > 60:   # 超长值多为自由文本
            continue
        keys = {str(v) for v, _ in vals}
        if keys & used_vals:           # 与已改列的值域相交 → 字面量重写有歧义
            continue
        idx += 1
        new = norm_name(c, t, used) if named else opaque_name(idx, used)
        used.add(new.lower())
        lmap = {str(v): "X-" + str(v) for v, _ in vals}
        try:
            conn.execute(f'ALTER TABLE "{t}" ADD COLUMN "{new}" TEXT')
            cases = " ".join(f'WHEN "{c}" = ? THEN ?' for _ in vals)
            params = []
            for v, _cnt in vals:
                params += [v, lmap[str(v)]]
            conn.execute(
                f'UPDATE "{t}" SET "{new}" = CASE {cases} ELSE "{c}" END', params)
            conn.execute(f'ALTER TABLE "{t}" DROP COLUMN "{c}"')
        except Exception as e:
            print(f"  [warn] {t}.{c} 值重编码失败: {e}", flush=True)
            try:                        # 回滚脏列，避免库里留下半成品
                conn.execute(f'ALTER TABLE "{t}" DROP COLUMN "{new}"')
            except Exception:
                pass
            continue
        used_vals |= keys
        applied[c] = {"table": t, "new": new, "op": "rename_and_recode",
                      "lit_map": lmap, "n_distinct": nd}
    return applied


def norm_rows(rows, lit_map):
    """把旧结果集按正向映射归一化，用于与新库结果集比较。

    只替换「恰好等于某个旧值」的单元格，其余值原样保留（含类型）。
    若旧结果集里别的列碰巧出现同名字符串，会归一化过度导致判不等——
    方向是保守的：只会丢样本，不会放过错样本。
    """
    if rows is None:
        return None
    out = []
    for row in rows:
        out.append(tuple(
            lit_map[str(v)] if not isinstance(v, (bytes, bytearray)) and str(v) in lit_map else v
            for v in row))
    return out          # 必须返回 list：fetchall() 是 list，type 不一致会让 == 恒假


# ---------------- 主流程 ---------------- #

def run(task, max_mb, only_dbs=None):
    outdir, out_jsonl, out_stats = OUT[task]
    os.makedirs(outdir, exist_ok=True)

    items = json.load(open(BIRD_JSON, encoding="utf-8"))
    # BIRD train.json 只有 db_id/question/evidence/SQL，没有 question_id 与 difficulty；
    # 用文件内顺序作稳定标识（同一份 train.json 顺序固定，可复现）。
    for _k, _it in enumerate(items):
        _it.setdefault("question_id", _k)
        _it.setdefault("difficulty", None)
    by_db = collections.defaultdict(list)
    for it in items:
        by_db[it["db_id"]].append(it)

    records, per_db, skipped = [], {}, []
    t_all = time.time()

    for db, rows in sorted(by_db.items()):
        if only_dbs and db not in only_dbs:
            continue
        src = sqlite_path(db)
        size_mb = os.path.getsize(src) / 1024 / 1024
        cap = max_mb if task == "rename" else min(max_mb, MAX_DROP_MB)
        if size_mb > cap:
            skipped.append([db, round(size_mb, 1), f">{cap}MB"])
            print(f"[skip] {db} {size_mb:.0f}MB > {cap}MB", flush=True)
            continue

        tabs, cols = list_schema(src)
        spec = None
        if task in EXPR_TASKS:
            _c0 = connect(src, ro=True)
            spec = (pick_merge_pair(_c0, rows, cols) if task == "merge_col"
                    else pick_split_col(_c0, rows, cols))
            _c0.close()
            if not spec:
                skipped.append([db, round(size_mb, 1), "no-candidate"])
                print(f"[skip] {db} 无可用候选列", flush=True)
                continue
            cand, colset = [], {c for _, c, *_ in cols}
        else:
            k = 5 if task in VSHIFT_TASKS else K_COLS
            cand, owners, colset = pick_candidates(db, rows, cols, k, tabs)
            if not cand:
                skipped.append([db, round(size_mb, 1), "no-candidate"])
                print(f"[skip] {db} 无候选列", flush=True)
                continue

        old_ddl = {}
        _c = connect(src, ro=True)
        old_ddl = ddl_map(_c)
        _c.close()

        dst = os.path.join(outdir, f"{db}.sqlite")
        copy_db(src, dst)

        conn = connect(dst)
        if task == "rename":
            applied = apply_rename(conn, cols, cand, colset)
        elif task == "merge_col":
            applied = apply_merge(conn, spec)
        elif task == "split_col":
            applied = apply_split(conn, spec)
        elif task in VSHIFT_TASKS:
            applied = apply_value_shift(conn, cols, cand, named=(task == "value_shift_named"))
        else:
            applied = apply_deprecate(conn, cols, cand, task)
        conn.commit()
        new_ddl = ddl_map(conn)
        conn.close()

        if not applied:
            skipped.append([db, round(size_mb, 1), "apply-failed"])
            print(f"[skip] {db} 演化施加失败", flush=True)
            continue

        n_aff = n_ok = n_bad = n_to = 0
        lit_map = {}
        for _v in applied.values():
            lit_map.update(_v.get("lit_map") or {})
        for it in rows:
            old_sql = it["SQL"]
            if not touched_cols(old_sql, applied):
                continue

            orig_rows, orig_err = run_sql(src, old_sql)
            if orig_err is not None:
                # 原题本身执行不了（超时/报错）→ 无 EX 真值，整条剔除，不计入失败
                n_to += 1
                continue
            n_aff += 1

            old_on_new_rows, old_on_new_err = run_sql(dst, old_sql)
            if old_on_new_err == "TIMEOUT":
                # 旧 SQL 在新库跑不完 → 无法判定是否失效，剔除
                n_aff -= 1
                n_to += 1
                continue

            if task == "dep_abstain":
                # 正确动作是弃答：不需要新 SQL，只验证旧证据在新库必然失效
                ok = (orig_err is None and old_on_new_err is not None
                      and old_on_new_err != "TIMEOUT")
                new_sql = None
                new_err = None
                result_equal = None
                expected = "abstain"
            else:
                if task != "dep_abstain" and re.search(r"\bselect\s+\*|\w\.\*", old_sql, re.I):
                    continue      # SELECT * 的列数/列序会变，结果等价不可判，剔除
                new_sql = rewrite(old_sql, {k: v["new"] for k, v in applied.items()},
                                  lit_map)
                if not new_sql:
                    n_bad += 1
                    continue
                new_rows, new_err = run_sql(dst, new_sql)
                result_equal = (norm_rows(orig_rows, lit_map) == new_rows
                                if lit_map else orig_rows == new_rows)
                ok = (orig_err is None and new_err is None
                      and old_on_new_err is not None and result_equal)
                expected = "rewrite"

            if ok:
                n_ok += 1
            else:
                n_bad += 1

            records.append({
                "db_id": db,
                "question_id": it["question_id"],
                "question": it["question"],
                "evidence": it.get("evidence"),
                "difficulty": it.get("difficulty"),
                "evo_type": task,
                "expected_action": expected,
                "drift_ops": [{"op": v["op"], "table": v["table"], "old": k, "new": v["new"],
                               **({"n_distinct": v["n_distinct"]} if "n_distinct" in v else {})}
                              for k, v in applied.items()],
                "old_schema_ddl": old_ddl,
                "new_schema_ddl": new_ddl,
                "old_sql": old_sql,
                "new_sql": new_sql,
                "check": {
                    "old_in_orig": "ok" if orig_err is None else orig_err,
                    "old_in_new": "fails" if old_on_new_err else "STILL_VALID",
                    "new_in_new": ("-" if new_sql is None else
                                   ("ok" if new_err is None else new_err)),
                    "result_equal": result_equal,
                    "pass": ok,
                },
            })

        per_db[db] = {"n_total": len(rows), "n_affected": n_aff,
                      "n_pass": n_ok, "n_fail": n_bad, "n_dropped_orig_bad": n_to,
                      "map": {k: v["new"] for k, v in applied.items()}}
        if lit_map:
            per_db[db]["lit_map"] = lit_map
        print(f"[ok] {db:24} 受影响 {n_aff:3} 通过 {n_ok:3} 失败 {n_bad:3} "
              f"剔除 {n_to} 列 {list(applied)[:3]}", flush=True)

    with open(out_jsonl, "w", encoding="utf-8") as f:
        for r in records:
            f.write(json.dumps(r, ensure_ascii=False) + "\n")

    stats = {
        "task": task,
        "total_questions": len(items),
        "dbs_processed": len(per_db),
        "dbs_skipped": skipped,
        "records": len(records),
        "pass": sum(1 for r in records if r["check"]["pass"]),
        "fail": sum(1 for r in records if not r["check"]["pass"]),
        "elapsed_s": round(time.time() - t_all, 1),
        "per_db": per_db,
    }
    json.dump(stats, open(out_stats, "w", encoding="utf-8"), ensure_ascii=False, indent=2)
    print("\n==== 汇总 ====")
    print(f"{task}: 受影响 {stats['records']}，通过 {stats['pass']}，"
          f"失败 {stats['fail']}，耗时 {stats['elapsed_s']}s")
    print("输出:", out_jsonl)


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--task", default="rename",
                    choices=["rename", "dep_abstain", "dep_successor",
                             "merge_col", "split_col",
                             "value_shift", "value_shift_named"])
    ap.add_argument("--max-mb", type=float, default=700)
    ap.add_argument("--dbs", default="", help="逗号分隔，限定库")
    a = ap.parse_args()
    only = {x.strip() for x in a.dbs.split(",") if x.strip()} or None
    run(a.task, a.max_mb, only)
