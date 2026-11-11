# 已作废产物（第 22 轮归档，共 2 个文件）

## 是什么

`ext_eval_merge_col_v1_wrongtable.*` 是扩充事件集合并档在**注入缺陷期内**的评测产物，
对应正文提到的旧读数 C3mp = 44.9%（弃答 111 条）。

## 为什么作废

当时的实验脚本 `m1_ext.py` 只设置了模块属性 `MP.EVO_DIR`，而 `m1_part.infer_db()`
读的是环境变量 `EVO_DIR`，两者不一致时推断退回默认目录 `data/evolved_merge/<db>.sqlite`
（主实验的合并库）。结果 28 个构造事件中有 18 个、479 条题中有 223 条被注入了
**同一数据库里另一张表**的段位表达式。评测侧按旧列名过滤谱系行后为空，
C3mp 退化成 C2，模型大量输出 CANNOT_ANSWER。

## 怎么用

**不要用它算任何指标。** 它在本包中的唯一用途是事故核查：
读者可以把这份产物与 `data/manifest_ext_merge_col.jsonl` 逐条对照，
复核「注入表 ≠ 应注入表」的事件确实集中在缺陷期内，并验证修复后的
`stats/ext_eval_merge_col.json`（C3mp = 91.2%，437/479）与之不同源。

判定口径见 `stats/lineage_injection_audit.json` 与第 7 章 7.16 节。
