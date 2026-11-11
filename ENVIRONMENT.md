# 运行环境与复现方式

## 一、解释器与依赖

| 项 | 值 |
|---|---|
| 生成本材料包时的解释器 | Python 3.13.14（Windows） |
| 必备第三方包 | `sqlglot`（解析与改写 SQL，本包测得 30.30.0）、`python-docx`（1.2.0，仅用于生成章节文档） |
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
