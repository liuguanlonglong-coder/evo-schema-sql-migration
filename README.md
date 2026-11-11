# 公开材料包 v3

本目录收录第 7 章全部可公开的复现材料。**不含数据库文件**（BIRD 以 CC BY-SA 4.0 发布，
请从其官方渠道获取），仅含清单、规则、提示、逐题输出与统计产物。

| 文件 | 内容 |
|---|---|
| `items.csv` | 4031 条演化项的逐题清单：题目、旧/新 SQL、演化 op、所属面板、`check_pass` |
| `items_holdout.csv` | 165 条留出库条目的逐题清单（BIRD train 划分，未参与开发） |
| `outputs.jsonl` | 4490 条 (条件, 题号) 的逐题输出：判定标签与生成 SQL（含主实验与留出库） |
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
   删列后外键失效而未施加），每库每档按 id 哈希取前 15 条，共 165 条。
   它检验的是「方向是否可迁移」，样本量不足以把主实验升级为留出测试。
2. EvoSchema 公开样本（84 条，去重 81）只提供新旧 DDL 与 gold SQL，**不提供演化后的库文件**，
   因此只评列引用 F1，不评执行等价；其涉及的 8 个库全部与本章 dev 的 11 个库重叠，
   属「换构造者、不换库」的外部检查。
