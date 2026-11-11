# -*- coding: utf-8 -*-
"""生成可公开的材料包（第五轮评审第 4 项后半：题目清单、变换脚本、阈值、提示与逐题输出）。

材料包目录 公开材料包/ 含：
  README.md             —— 材料清单与复现步骤
  MANIFEST.json         —— 全部文件与生成脚本的 sha1、字节数
  items.csv             —— 逐题清单（题目、旧/新 SQL、演化 op、是否进入第 7 章面板）
  outputs.jsonl         —— 逐题输出（条件、判定、生成 SQL）
  prompts.md            —— 系统提示、各条件的证据块，以及同一道题的实际渲染示例
  prompt_lengths.md     —— 各条件的提示长度（字符），供核对「c_rw 提示更长」这一对照问题
  evolution_rules.md    —— 演化规则与全部阈值
  stats/                —— 本章统计产物副本

不新增模型调用；提示示例由现有 build_prompt 渲染，与采集时同一份代码。
"""
import collections
import hashlib
import importlib.util
import json
import os
import shutil

BASE = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE)
OUTDIR = "公开材料包"
os.makedirs(OUTDIR, exist_ok=True)
os.makedirs(os.path.join(OUTDIR, "stats"), exist_ok=True)

HLD_FILES = ["e1_holdout_%s.jsonl" % d for d in
             ("books", "movie_3", "movies_4", "olympics",
              "regional_sales", "retail_world", "soccer_2016")]

FILES = ["e1_dev_full.jsonl", "e1_diff_3lvl.jsonl", "e1_c3m_paired.jsonl",
         "e1_expand_med.jsonl", "e1_expand_hard.jsonl", "e1_ds_c3f.jsonl",
         "e1_ms_ds.jsonl", "e1_ms_ds_p.jsonl", "e1_ms_ds_p_split.jsonl",
         "e1_ds_one.jsonl", "e1_ds_one_b1.jsonl", "e1_ds_one_b2.jsonl",
         "e3_rw_main.jsonl", "e1_ab_nat_0.jsonl", "e1_ab_nat_1.jsonl",
         "e1_ab_nat_2.jsonl", "e1_ab_nat_3.jsonl", "e1_c3noise_paired.jsonl"] + HLD_FILES
STATS = ["stats_ch7_v5.json", "paired_ci.json", "paired_ci_coverage.json",
         "stats_cluster.json", "op_level.json", "judge_sensitivity.json",
         "holdout_audit.json", "holdout_stats.json", "evoschema_eval.json",
         "evoschema_struct.json", "evoschema_events.json", "holdout_funnel.json",
         "rule_rewrite.json", "rule_rewrite_keep.json", "rule_vs_llm.json",
         "probe_counterfactual.json", "probe_star_expand.json",
         "data_bounds.json", "encoding_bounds.json", "stats_rw.json",
         "base_acc_name.json", "base_acc_vover.json", "base_acc_emb.json",
         "base_eval.json", "ablation_matching.json",
         "audit_independence.json", "event_map_audit.json",
         "baseline_segment_aware.json", "m1_holdout_eval.json",
         "judge_strict_audit.json", "encoding_bounds_v2.json",
         "encoding_bounds_v2_legacy.json"]
SCRIPTS = ["paired_ci.py", "paired_coverage.py", "stats_ch7_v5.py",
           "stats_cluster.py", "judge_sensitivity.py", "op_level.py",
           "holdout_audit.py", "run_c1c2c3.py", "run_rw_baseline.py",
           "build_evolution_v2.py", "m1_lineage.py", "m1_part.py",
           "select_holdout_dbs.py", "holdout_manifest.py", "holdout_sample.py",
           "holdout_stats.py", "evoschema_eval.py", "evoschema_struct.py",
           "merge_evoschema.py", "rule_rewrite.py", "rule_vs_llm.py",
           "evoschema_events.py", "holdout_funnel.py", "probe_counterfactual.py",
           "probe_star_expand.py", "data_bounds.py", "probe_encoding_bounds.py", "pack_materials.py",
           "make_chapter7_v9.py", "make_ch7_response_v9.py",
           "make_front_matter_v9.py", "make_chapter7_v10.py",
           "make_ch7_response_v10.py", "make_front_matter_v10.py",
           "ablation_matching.py", "baselines_lineage.py",
           "make_chapter7_v11.py", "make_ch7_response_v11.py",
           "make_front_matter_v11.py", "audit_independence.py",
           "make_ch7_response_v12.py", "make_ch7_response_v13.py",
           "event_map_audit.py", "baseline_segment_aware.py",
           "make_ch7_response_v14.py", "slim_chapter7.py",
           "prepare_archive.py", "m1_holdout_eval.py",
           "judge_strict_audit.py", "probe_encoding_bounds_v2.py",
           "make_p1_closure_docs.py", "fill_doi.py"]


def sha1_of(p):
    h = hashlib.sha1()
    with open(p, "rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


MAN = {}
for _l in open("data/manifest_v3.jsonl", encoding="utf-8"):
    _r = json.loads(_l)
    MAN[_r["id"]] = _r

IDS_XM = set(l.strip() for l in open("data/ids_xm_full.txt", encoding="utf-8") if l.strip())
IDS_MS = set(l.strip() for l in open("data/ids_ms.txt", encoding="utf-8") if l.strip())

# ---------------------------------------------------------------- items.csv
import csv
cols = ["id", "panel", "evo_type", "level", "db_id", "question_id", "difficulty",
        "split", "expected_action", "n_ops", "ops_json", "question",
        "old_sql", "new_sql", "check_pass"]
panel_of = {}
for i in IDS_MS:
    panel_of[i] = MAN[i]["evo_type"].replace("merge_col", "merge").replace("split_col", "split")
for i in IDS_XM:
    panel_of[i] = MAN[i]["evo_type"]
n_row = 0
with open(os.path.join(OUTDIR, "items.csv"), "w", encoding="utf-8-sig", newline="") as fh:
    w = csv.writer(fh)
    w.writerow(cols)
    for i, r in MAN.items():
        w.writerow([i, panel_of.get(i, ""), r["evo_type"], r.get("level", ""), r["db_id"],
                    r["question_id"], r.get("difficulty", ""), r.get("split", ""),
                    r.get("expected_action", ""), r.get("n_ops", ""),
                    json.dumps(r.get("ops"), ensure_ascii=False),
                    (r.get("question") or "").strip(),
                    (r.get("old_sql") or "").strip(), (r.get("new_sql") or "").strip(),
                    r.get("check_pass", "")])
        n_row += 1
print(f"items.csv  {n_row} 行")

# ---------------------------------------------------------------- items_holdout.csv
n_hl = 0
with open(os.path.join(OUTDIR, "items_holdout.csv"), "w", encoding="utf-8-sig", newline="") as fh:
    w = csv.writer(fh)
    w.writerow(cols)
    for _l in open("data/manifest_holdout_sample.jsonl", encoding="utf-8"):
        r = json.loads(_l)
        w.writerow([r["id"], r["evo_type"], r["evo_type"], r.get("level", ""), r["db_id"],
                    r["question_id"], r.get("difficulty", ""), r.get("split", ""),
                    r.get("expected_action", ""), r.get("n_ops", ""),
                    json.dumps(r.get("ops"), ensure_ascii=False),
                    (r.get("question") or "").strip(),
                    (r.get("old_sql") or "").strip(), (r.get("new_sql") or "").strip(),
                    r.get("check_pass", "")])
        n_hl += 1
print(f"items_holdout.csv  {n_hl} 行")

# ---------------------------------------------------------------- evoschema_cols.csv
# EvoSchema 公开列级样本的逐条结果（只评列引用，无预测 SQL 全文）
EV = json.load(open("results/evoschema_eval.json", encoding="utf-8"))
n_ev = 0
with open(os.path.join(OUTDIR, "evoschema_cols.csv"), "w",
          encoding="utf-8-sig", newline="") as fh:
    w = csv.writer(fh)
    w.writerow(["db_id", "kind", "recovered_kind", "n_changes", "changes",
                "c1_f1", "c2_f1", "c3m_f1",
                "c1_exact", "c2_exact", "c3m_exact"])
    for r in EV.get("rows", []):
        cd = r.get("conds") or {}

        def g(c, k, d=0.0):
            return (cd.get(c) or {}).get(k, d)
        w.writerow([r.get("db") or r.get("db_id") or "", r.get("kind", ""),
                    r.get("recovered_kind", ""), r.get("n_changes", ""),
                    json.dumps(r.get("changes"), ensure_ascii=False, sort_keys=True),
                    f"{g('c1', 'f1'):.4f}", f"{g('c2', 'f1'):.4f}",
                    f"{g('c3m', 'f1'):.4f}",
                    int(bool(g('c1', 'exact_cols', False))),
                    int(bool(g('c2', 'exact_cols', False))),
                    int(bool(g('c3m', 'exact_cols', False)))])
        n_ev += 1
print(f"evoschema_cols.csv  {n_ev} 行")

# ---------------------------------------------------------------- outputs.jsonl
seen = set()
n_out = 0
with open(os.path.join(OUTDIR, "outputs.jsonl"), "w", encoding="utf-8") as fo:
    for f in FILES:
        p = os.path.join("results", f)
        if not os.path.exists(p):
            continue
        for l in open(p, encoding="utf-8"):
            r = json.loads(l)
            cond = r.get("cond")
            if not cond:
                continue
            k = (cond, r["id"])
            if k in seen:
                continue
            seen.add(k)
            fo.write(json.dumps({"cond": cond, "id": r["id"], "outcome": r.get("outcome"),
                                 "sql": r.get("sql") or "", "model": r.get("model", ""),
                                 "source": f}, ensure_ascii=False) + "\n")
            n_out += 1
print(f"outputs.jsonl  {n_out} 条 (cond,id)")

# ---------------------------------------------------------------- prompts.md
spec = importlib.util.spec_from_file_location("rc", "run_c1c2c3.py")
RC = importlib.util.module_from_spec(spec)
spec.loader.exec_module(RC)
spec2 = importlib.util.spec_from_file_location("rw", "run_rw_baseline.py")
RW = importlib.util.module_from_spec(spec2)
spec2.loader.exec_module(RW)

merge_ids = sorted(i for i in IDS_MS if MAN[i]["evo_type"] == "merge_col")
demo = merge_ids[0]
demo_r = MAN[demo]

COND_DESC = {
    "c1": "只给当前 schema，不给示例、不给谱系（下界）",
    "c2": "当前 schema + 检索到的旧版示例 SQL（旧列名，不能直接执行）",
    "c3": "C2 + oracle 谱系（该库全部演化 op，DDL 级）",
    "c3f": "C2 + oracle 谱系，但按旧 SQL 引用列过滤（与 C3m 同一过滤规则）",
    "c3m": "C2 + M1 自动抽取的谱系，只给旧 SQL 真实引用过的列",
    "c3mp": "C3m + 合并列的段位表达式（可执行的 SUBSTR/INSTR 重构）",
    "c3mp_nat": "C3mp 的信息等价消融：段位用自然语言描述，不给可执行表达式",
    "c_rw": "直给改写对照：同时给旧 schema 与当前 schema，要求直接改写，不给谱系",
}

lines = ["# 提示模板与渲染示例", "",
         "系统提示（全部条件共用）：", "", "```", RC.SYSTEM.strip(), "```", "",
         "各条件的证据块：", ""]
for c, d in COND_DESC.items():
    lines.append(f"- `{c}`：{d}")
lines += ["", "谱系行的渲染由 `lineage_line()` 完成；`c3mp` 的表达式直接取自 M1 段位推断"
          "（`m1_part.py`）的输出，不是人工撰写。", "",
          f"## 示例：{demo}（{demo_r['db_id']}）", ""]

for c in ("c2", "c3m", "c3mp", "c3mp_nat"):
    p = RC.build_prompt(demo_r, c, RC.load_schema(RC.db_path(demo_r["evo_type"], demo_r["db_id"])))
    lines += [f"### {c}", "", "```text", p.strip(), "```", ""]
p = RW.build_prompt(demo_r, "c_rw", RC.load_schema(RC.db_path(demo_r["evo_type"], demo_r["db_id"])))
lines += ["### c_rw", "", "```text", p.strip(), "```", ""]

# ---------------------------------------------------------------- prompt_lengths
panels = {"合并": [i for i in IDS_MS if MAN[i]["evo_type"] == "merge_col"],
          "拆分": [i for i in IDS_MS if MAN[i]["evo_type"] == "split_col"],
          "改名 easy": [i for i in IDS_XM if MAN[i]["evo_type"] == "rename"]}
plen = {}
schema_cache = {}
for pname, ids in panels.items():
    for c in ("c2", "c3m", "c3mp", "c_rw"):
        vals = []
        for i in ids:
            r = MAN[i]
            key = (r["evo_type"], r["db_id"])
            if key not in schema_cache:
                schema_cache[key] = RC.load_schema(RC.db_path(r["evo_type"], r["db_id"]))
            if c == "c_rw":
                vals.append(len(RW.build_prompt(r, c, schema_cache[key])))
            else:
                vals.append(len(RC.build_prompt(r, c, schema_cache[key])))
        vals.sort()
        plen.setdefault(pname, {})[c] = {
            "mean": round(sum(vals) / len(vals)),
            "median": vals[len(vals) // 2],
            "min": vals[0], "max": vals[-1], "n": len(vals)}

lines += ["## 提示长度（字符数，含 schema）", "",
          "| 面板 | 条件 | 均值 | 中位数 | 最小 | 最大 | n |",
          "|---|---|---|---|---|---|---|"]
for pname, d in plen.items():
    for c, v in d.items():
        lines.append(f"| {pname} | {c} | {v['mean']} | {v['median']} | "
                     f"{v['min']} | {v['max']} | {v['n']} |")
lines += ["", "`c_rw` 同时给出旧、新两份 schema，因此提示更长；这一长度差是"
          "「显式谱系 vs 直给改写」这一对照本身的组成，不是实现瑕疵，"
          "但它意味着结论只适用于该提示设计。", ""]
open(os.path.join(OUTDIR, "prompts.md"), "w", encoding="utf-8").write("\n".join(lines))
json.dump(plen, open(os.path.join(OUTDIR, "prompt_lengths.json"), "w", encoding="utf-8"),
          ensure_ascii=False, indent=1)
print("prompts.md / prompt_lengths.json 完成")

# ---------------------------------------------------------------- evolution_rules.md
rules = """# 演化规则与阈值

## 三类演化（均在 BIRD dev 数据库的副本上施加，原始文件不变）

1. **改名 rename**：把被 gold SQL 引用的列改成不可读名。三档由改名比例控制
   （easy / medium / hard），每档的题池相同，档间差异仅在于改了多少列。
2. **列合并 merge_col**：把同一张表的两列合并为一列，值以 `CHAR(31)` 为分隔符拼接；
   旧 SQL 必须用 `SUBSTR(..., INSTR(...))` 取回原来的段位。
3. **列拆分 split_col**：把一个复合列拆成两列，旧 SQL 必须重新拼接
   （如 `attr_001 || '-' || attr_002`）。

## 构建期阈值（build_evolution_v2.py）

| 参数 | 值 | 用途 |
|---|---|---|
| `K_COLS` | 3 | 每库每档参与改名的列数 |
| `QUERY_TIMEOUT` | 4.0 s | gold SQL 校验的单项超时 |
| `MAX_UPDATE_ROWS` | 400000 | 依赖型演化需全表 UPDATE 时的行数上限 |
| `MAX_DROP_MB` | 130 MB | DROP COLUMN 需重写整表，超过该体积的库跳过 |

每个演化项都通过 `check_pass` 校验：新 gold SQL 在演化库上可执行，且结果与旧 SQL
在旧库上的结果一致。

## 谱系抽取阈值（m1_lineage.py / m1_part.py）

| 参数 | 值 | 说明 |
|---|---|---|
| 值域指纹匹配阈值 | 0.30–0.95 扫描 | 结果在区间内全同（天花板效应），不作为稳健性证据 |
| 名字证据 | 仅 tie-break | 不参与主判定 |
| 形态指纹 | `--mode shape` | 值频率曲线余弦 × 长度项，用于值改写档 |
| 近唯一列禁用 | ndist/n ≥ 0.90 | 该区间无法从取值恢复，需外部元数据 |
| 方向判定前置 | JAC < 0.9 | 与 split 方向判定互斥 |

## 判定器（run_c1c2c3.judge）

结果行逐格转字符串、NULL 视为空串、排序后比较多元组集合；列序敏感、重复行保留；
`LLM_ERROR` 不计入分母。该规则对行序不敏感，影响量见第 7 章判定敏感性一节。
"""
open(os.path.join(OUTDIR, "evolution_rules.md"), "w", encoding="utf-8").write(rules)

# ---------------------------------------------------------------- stats 副本
copied = []
for s in STATS:
    p = os.path.join("results", s)
    if os.path.exists(p):
        shutil.copy2(p, os.path.join(OUTDIR, "stats", s))
        copied.append(s)
print("stats 副本:", len(copied))

# ---------------------------------------------------------------- code/ 脚本本体
# 评审第七节点名：只给 sha1 无法重算结果，必须把脚本一并提供。
CODEDIR = os.path.join(OUTDIR, "code")
os.makedirs(CODEDIR, exist_ok=True)
n_code = 0
for s_ in SCRIPTS:
    if os.path.exists(s_):
        shutil.copy2(s_, os.path.join(CODEDIR, s_))
        n_code += 1
print(f"code/  {n_code} 个脚本")

# ---------------------------------------------------------------- MANIFEST + README
man = {"version": "materials-v6", "files": {}, "scripts": {}, "inputs": {}}
for root, _dirs, fs in os.walk(OUTDIR):
    for f in sorted(fs):
        p = os.path.join(root, f)
        rel = os.path.relpath(p, OUTDIR)
        man["files"][rel] = {"sha1": sha1_of(p), "bytes": os.path.getsize(p)}
for s in SCRIPTS:
    if os.path.exists(s):
        man["scripts"][s] = sha1_of(s)
for f in FILES:
    p = os.path.join("results", f)
    if os.path.exists(p):
        man["inputs"][f] = sha1_of(p)
man["manifest_v3_sha1"] = sha1_of("data/manifest_v3.jsonl")
json.dump(man, open(os.path.join(OUTDIR, "MANIFEST.json"), "w", encoding="utf-8"),
          ensure_ascii=False, indent=1)

readme = f"""# 公开材料包 v3

本目录收录第 7 章全部可公开的复现材料。**不含数据库文件**（BIRD 以 CC BY-SA 4.0 发布，
请从其官方渠道获取），仅含清单、规则、提示、逐题输出与统计产物。

| 文件 | 内容 |
|---|---|
| `items.csv` | {n_row} 条演化项的逐题清单：题目、旧/新 SQL、演化 op、所属面板、`check_pass` |
| `items_holdout.csv` | {n_hl} 条留出库条目的逐题清单（BIRD train 划分，未参与开发） |
| `outputs.jsonl` | {n_out} 条 (条件, 题号) 的逐题输出：判定标签与生成 SQL（含主实验与留出库） |
| `evoschema_cols.csv` | 在 EvoSchema 公开列级样本上的逐条列引用 F1（C1/C2/C3m） |
| `prompts.md` | 系统提示、各条件证据块、同一道题在五个条件下的实际渲染示例 |
| `prompt_lengths.json` | 各面板各条件的提示长度（字符），用于核对「c_rw 提示更长」 |
| `evolution_rules.md` | 三类演化规则与构建期、谱系抽取、判定器的全部阈值 |
| `code/` | 全部生成脚本本体（评审第七节点名：只给 sha1 无法重算，故随包发布） |
| `ENVIRONMENT.md` | 运行环境、依赖版本、数据获取方式与逐产物复现命令 |
| `stats/*.json` | 全部统计产物：区间与覆盖率、聚类、四层分析单位与跨档池化、判定敏感性、留出审计、留出库结果、EvoSchema 外部对照与其结构统计、规则基线配对、EvoSchema 事件级重算、反事实探针、真实数据边界、留出筛选漏斗 |
| `MANIFEST.json` | 上述每个文件与每个生成脚本的 sha1 与字节数 |

复现顺序（主实验）：`paired_ci.py` → `paired_coverage.py` → `stats_cluster.py` →
`stats_ch7_v5.py` → `judge_sensitivity.py` → `op_level.py` → `holdout_audit.py`。

复现顺序（留出库确认）：`select_holdout_dbs.py` → `build_evolution_v2.py`（`BIRD_ROOT`
指向 train 划分、输出加 `_holdout` 后缀）→ `holdout_manifest.py` → `holdout_sample.py`
→ `run_c1c2c3.py`（逐库并行）→ `holdout_stats.py`。

复现顺序（EvoSchema 外部对照）：`evoschema_eval.py` → `merge_evoschema.py` →
`evoschema_struct.py`，产物为 `evoschema_eval.json` 与 `evoschema_struct.json`。

复现顺序（确定性重写基线）：`rule_rewrite.py`（四档一次跑完）→ `rule_vs_llm.py`，
产物为 `rule_rewrite.json`、`rule_rewrite_items.jsonl` 与 `rule_vs_llm.json`。

复现顺序（外部对照的事件级重算）：`evoschema_events.py` → `evoschema_events.json`。

复现顺序（编码边界）：`probe_counterfactual.py`（隔离库探针，不读真实库）→
`data_bounds.py`（真实库全表扫描）→ `probe_counterfactual.json`、`data_bounds.json`。
三段以上压力测试：`probe_encoding_bounds_v2.py`（默认走生产 `m1_part.to_expr`，
nseg 2–6 × 5 分隔符 × 8 情形 × 60 行）→ `encoding_bounds_v2.json`；
加 `--legacy` 复现修正前 `to_expr` 末段写法 → `encoding_bounds_v2_legacy.json`，
两者用于对照 `sep_in_last` 从 20.0% 到 100.0% 的实现修正。
留出库 M1 直接表达式评估：`m1_holdout_eval.py` → `m1_holdout_eval.json`
（需 BIRD train 划分的 `dev_databases` 同级目录，见 ENVIRONMENT.md）；
判定器全严复核：`judge_strict_audit.py` → `judge_strict_audit.json`。
注意：`m1_part.py` 的 `to_expr` 已修正 nseg>=3 末段分支，两段表达式逐字不变。

留出筛选漏斗：`holdout_funnel.py` → `holdout_funnel.json`。

## 文档生成链（v11 + 正文瘦身，顺序不可颠倒）

1. `make_chapter7_v11.py` → 生成 **49 表完整版**《投稿版_第7章_实验结果_v11.docx》；
2. `slim_chapter7.py` → 以完整版为输入，输出**主文 15 表**（文件名不变）、
   《补充材料_第7章_v11.docx》（S-1 至 S-34）与《第7章_v11_完整版_49表.docx》备份；
3. `make_front_matter_v11.py` → 摘要 / 标题 / 贡献点 / 结论 / 投稿信，
   其 `_T()` 同时反查主文与补充材料的表号。

只跑第 1 步会得到瘦身前的 49 表版本；只跑第 3 步会因缺少补充材料 docx 而报错。
所有数值均由脚本从 `results/*.jsonl` 与 `data/manifest_v3.jsonl`、
`data/manifest_holdout_sample.jsonl` 重算，章节文档不含手写常量；
`MANIFEST.json` 的 sha1 用于核对材料版本与正文一致。

## 归档与引用

| 文件 | 内容 |
|---|---|
| `CITATION.cff` | 引用元数据（作者、题名、年份、仓库 DOI 占位） |
| `LICENSE` | 数据与文档 CC BY 4.0 / 代码 MIT，发布前需作者确认 |
| `DAS.md` | Data Availability Statement 中英模板 |
| `UPLOAD_CHECKLIST.md` | GitHub 与 Zenodo 分步上传清单 |

归档完成前，`DAS.md` 中的 `[REPOSITORY_DOI]` 为占位符，**不得作为已发布声明**；
投稿信相应写为「将在发表前完成公开归档」。

**DOI 回填链（顺序不可颠倒）**：Zenodo 发布后文件与 DOI 都不可修改，因此必须
「草稿阶段预留 DOI → 回填 → 上传 → Publish」，不能反过来。
`fill_doi.py` 写 `archive_info.json` 并重建全部材料；`prepare_archive.py` 与
`make_front_matter_v11.py` 都读同一份文件，故 CITATION.cff、DAS.md、
UPLOAD_CHECKLIST.md、投稿信第四段、表 9-2、表 9-3 六处口径一次同步。

    python fill_doi.py --doi 10.5281/zenodo.XXXXXXX --repo-name NAME --repo-url URL --author-family 姓 --author-given 名
    python fill_doi.py --status deposited --date-released YYYY-MM-DD   # Publish 后
    python fill_doi.py --reset                       # 回到未归档状态

`status=reserved` 时 CITATION.cff 已写入真实 DOI（供上传），但稿件中的 DAS 与
投稿信仍保持「将归档」口径——这是刻意设计，不是漏改。

## 两组外部数据的边界

1. 留出库取自 BIRD **train** 划分的 7 个库（合并 6、拆分 5，其中 movie_3 的拆分因
   删列后外键失效而未施加），每库每档按 id 哈希取前 15 条，共 {n_hl} 条。
   它检验的是「方向是否可迁移」，样本量不足以把主实验升级为留出测试。
2. EvoSchema 公开样本（84 条，去重 81）只提供新旧 DDL 与 gold SQL，**不提供演化后的库文件**，
   因此只评列引用 F1，不评执行等价；其涉及的 8 个库全部与本章 dev 的 11 个库重叠，
   属「换构造者、不换库」的外部检查。
"""
open(os.path.join(OUTDIR, "README.md"), "w", encoding="utf-8").write(readme)

# ---------------------------------------------------------------- ENVIRONMENT.md
import platform
try:
    import sqlglot as _sg
    _sgv = getattr(_sg, "__version__", "?")
except Exception:
    _sgv = "未安装"
try:
    import docx as _dx
    _dxv = getattr(_dx, "__version__", "?")
except Exception:
    _dxv = "未安装"

env = f"""# 运行环境与复现方式

## 一、解释器与依赖

| 项 | 值 |
|---|---|
| 生成本材料包时的解释器 | Python {platform.python_version()}（{platform.system()}） |
| 必备第三方包 | `sqlglot`（解析与改写 SQL，本包测得 {_sgv}）、`python-docx`（{_dxv}，仅用于生成章节文档） |
| 标准库依赖 | `sqlite3`、`json`、`csv`、`hashlib`、`glob`、`math`、`collections` |
| 模型调用 | 通过 OpenAI 兼容的 HTTP 接口；模型标识与端点见 `evolution_rules.md` 末节。**本包不含任何 API 密钥**，复现需自行提供 |

## 二、数据获取

1. **BIRD dev**（主实验）：从 BIRD 官方渠道下载 `dev_20240627`，解压到
   `data/bird/dev_20240627/`，其中数据库位于 `dev_databases/<db_id>/<db_id>.sqlite`。
2. **BIRD train**（留出库）：下载 `train.zip`，解压到 `data/bird_train/train/`，
   数据库位于 `train_databases/<db_id>/<db_id>.sqlite`。
3. **EvoSchema 公开列级样本**：`data/evoschema/bird_split_or_merge_columns.json`，
   来自 EvoSchema（PVLDB 2025）的公开材料，随本包附带该文件本身（不含数据库）。

BIRD 以 CC BY-SA 4.0 发布，本包**不重新分发**任何数据库文件。

## 三、环境变量（均有默认值，可省略）

| 变量 | 作用 | 默认 |
|---|---|---|
| `BIRD_ROOT` | train 划分构造时的 BIRD 根目录 | `data/bird/dev_20240627` |
| `BIRD_DBDIR` | 数据库子目录名 | `dev_databases` |
| `OLD_DBDIR` | 谱系抽取读取原始库的根目录 | dev 的 `dev_databases` |
| `EVO_DIR` | 谱系抽取读取演化后库的根目录 | `data/evolved_*` |
| `SILICONFLOW_API_KEY` | 仅 `baselines_lineage.py` 的 B3 嵌入基线在**未命中缓存**时需要 | 空 |

关于 B3：该基线把「列名 + 值样例」送嵌入模型（`Qwen/Qwen3-Embedding-8B`）算相似度，
向量缓存在 `data/emb_cache.json`。**该缓存未随包发布**（属数据派生物，且体积较大），
因此完整重跑 B3 需要自备密钥；`results/base_acc_emb.json` 已包含其结果，
复核表 7-16 的 B3 一行不需要重新调用接口。B1、B2 与 M1 全部为本地计算，无需任何密钥。

## 四、最短复现路径（每步产物已在 `stats/` 内）

```bash
# 0) 环境
python -m pip install sqlglot python-docx

# 1) 主实验统计（读 results/*.jsonl 与 data/manifest_v3.jsonl）
python code/paired_ci.py
python code/paired_coverage.py
python code/stats_cluster.py
python code/stats_ch7_v5.py
python code/judge_sensitivity.py      # 含 ORDER BY 子集（strict_ex.*.order_by_subset）
python code/op_level.py               # 含跨档池化（pooled）
python code/holdout_audit.py
python code/holdout_funnel.py

# 2) 确定性重写基线（新增，对应 7.4.2/7.4.3）
python code/rule_rewrite.py --evo merge_col,split_col,merge_col_holdout,split_col_holdout
python code/rule_vs_llm.py            # 配对、因素分解、失败画像

# 3) EvoSchema 外部对照（新增事件级重算）
python code/evoschema_eval.py --shard 0 --of 1   # 需模型调用
python code/merge_evoschema.py
python code/evoschema_struct.py
python code/evoschema_events.py

# 4) 编码可逆性边界
python code/probe_counterfactual.py      # 隔离库，不读真实库
python code/data_bounds.py               # 真实库全表扫描
python code/probe_encoding_bounds.py     # 两段/三段逐类界定
python code/probe_encoding_bounds_v2.py  # 200 格压力测试（生产 to_expr）
python code/probe_encoding_bounds_v2.py --legacy   # 修正前对照留档
python code/m1_holdout_eval.py           # 留出库 M1 直接表达式（BIRD train）
python code/judge_strict_audit.py        # 八个主对照全严规则复核

# 5) 生成章节与修订说明（顺序不可颠倒）
python code/make_chapter7_v11.py
python code/slim_chapter7.py             # 主文 15 表 + 补充材料
python code/make_front_matter_v11.py
python code/make_ch7_response_v14.py
python code/make_p1_closure_docs.py
```

## 五、各类数字的出处对照

| 正文位置 | 产物文件 | 生成脚本 |
|---|---|---|
| 表 7-21~24（规则基线与因素分解） | `stats/rule_vs_llm.json`、`stats/rule_rewrite.json` | `rule_vs_llm.py`、`rule_rewrite.py` |
| 表 7-40（外部对照的事件级重算） | `stats/evoschema_events.json` | `evoschema_events.py` |
| 表 7-33/34（编码边界） | `stats/probe_counterfactual.json`、`stats/data_bounds.json` | `probe_counterfactual.py`、`data_bounds.py` |
| 表 7-30（留出筛选漏斗） | `stats/holdout_funnel.json` | `holdout_funnel.py` |
| 表 7-14（ORDER BY 子集） | `stats/judge_sensitivity.json` 的 `strict_ex` | `judge_sensitivity.py` |

## 六、已知的运行前提

- 规则基线的 `c3r_gt` 使用构造脚本的 `drift_ops`，与构造同源，**只能作为天花板检验**，
  不能读作方法性能；正文 7.4.2 末段已声明。
- `data_bounds.py` 需要原始 BIRD 库（合并方向查源列）与演化后库（拆分方向查分量列）同时存在。
- EvoSchema 公开样本不含演化后数据库，`evoschema_eval.py` 只评列引用，不评执行等价。
"""
open(os.path.join(OUTDIR, "ENVIRONMENT.md"), "w", encoding="utf-8").write(env)
print("ENVIRONMENT.md 已生成")

print("-> 公开材料包/ 完成")
