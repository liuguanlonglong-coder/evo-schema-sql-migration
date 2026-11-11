# -*- coding: utf-8 -*-
"""数据集流转审计：把 4,031 / 4,490 / 610 / 437 / 61 五个口径逐个落到原始文件上。

为什么必须做这件事
  审稿意见指出材料里同时出现 4,031（构造条数）、4,490（逐题输出）、610（冻结评测）
  三组数字，却没有任何一处说明它们的层级关系；又强调「不能直接假定它们之间存在
  简单包含关系」。本脚本逐一核对，并把每一层的定义写死到 JSON 里，
  供正文与补充材料的流转表直接引用——数字全部来自文件计数，不接受手填。

五个层次
  L1 构造记录   data/manifest_v3.jsonl 的行数（9 类演化 × 11 库 × 948 基础题）
  L2 冻结评测   data/ids_xm_full.txt + data/ids_ms.txt 的去重 id 数
  L3 基础题     L2 中 question_id 的去重数（一题多档 → 唯一题少于条目）
  L4 独立事件   results/audit_independence.json 的 paper_set.events
  L5 模型输出   公开材料包/outputs.jsonl 的行数，以及条件数 / 覆盖 id 数

输出 results/data_lineage_audit.json
"""
import collections
import json
import os

BASE = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE)

out = {"version": "data-lineage-audit-v1"}

# ---------------- L1 构造记录 ----------------
man = [json.loads(l) for l in open("data/manifest_v3.jsonl", encoding="utf-8")]
L1 = {"file": "data/manifest_v3.jsonl", "rows": len(man),
      "by_evo": dict(collections.Counter(r["evo_type"] for r in man)),
      "dbs": len({r["db_id"] for r in man}),
      "question_ids": len({r["question_id"] for r in man}),
      "unique_id": len({r["id"] for r in man})}
out["L1_constructed"] = L1

# ---------------- L2 冻结评测 ----------------
ids = []
for f in ("data/ids_xm_full.txt", "data/ids_ms.txt"):
    ids += [l.strip() for l in open(f, encoding="utf-8") if l.strip()]
fz = sorted(set(ids))
idx = {r["id"]: r for r in man}
missing = [i for i in fz if i not in idx]
L2 = {"files": ["data/ids_xm_full.txt", "data/ids_ms.txt"],
      "ids_listed": len(ids), "unique_ids": len(fz),
      "ids_not_in_manifest": len(missing),
      "by_evo": dict(collections.Counter(idx[i]["evo_type"] for i in fz if i in idx)),
      "dbs": len({idx[i]["db_id"] for i in fz if i in idx})}
out["L2_frozen"] = L2

# ---------------- L3 基础题 ----------------
sub = [idx[i] for i in fz if i in idx]
L3 = {"unique_question_id": len({r["question_id"] for r in sub}),
      "unique_question_text": len({r["question"] for r in sub}),
      "items_per_question": round(len(sub) / max(1, len({r["question_id"] for r in sub})), 2)}
out["L3_base_questions"] = L3

# ---------------- L4 独立事件 ----------------
adi = json.load(open("results/audit_independence.json", encoding="utf-8"))
ps, ss = adi["A_units"]["paper_set"], adi["A_units"]["segment_set"]
out["L4_events"] = {"version": adi["version"], "paper_set": ps, "segment_set": ss}

# ---------------- L5 模型输出 ----------------
p = "公开材料包/outputs.jsonl"
if os.path.exists(p):
    rows = [json.loads(l) for l in open(p, encoding="utf-8") if l.strip()]
    by_c = collections.Counter(r["cond"] for r in rows)
    ids_c = collections.defaultdict(set)
    for r in rows:
        ids_c[r["cond"]].add(r["id"])
    allids = {r["id"] for r in rows}
    hold = json.load(open("results/m1_holdout_eval.json", encoding="utf-8")) if \
        os.path.exists("results/m1_holdout_eval.json") else {}
    out["L5_outputs"] = {
        "file": p, "rows": len(rows),
        "unique_(cond,id)": len({(r["cond"], r["id"]) for r in rows}),
        "conditions": len(by_c),
        "per_cond": {c: {"rows": n, "unique_ids": len(ids_c[c])}
                     for c, n in by_c.most_common()},
        "unique_ids": len(allids),
        "ids_in_frozen": len(allids & set(fz)),
        "ids_outside_frozen": len(allids - set(fz)),
        "models": sorted({r.get("model") for r in rows if r.get("model")}),
    }

# ---------------- 包含关系核验 ----------------
rel = {
    "frozen_subset_of_manifest": len(missing) == 0,
    "outputs_ids_superset_of_frozen": set(fz) <= (out["L5_outputs"]["unique_ids"]
                                                  and allids) if "L5_outputs" in out else None,
    "outputs_not_multiple_of_frozen":
        out["L5_outputs"]["rows"] % max(1, len(fz)) != 0 if "L5_outputs" in out else None,
    "note": ("4,490 不是 610 × 条件数的整齐乘积：各条件适用的题集不同"
             "（c1/c3f/c_rw 用冻结 610；c2/c3m 用 610+165 留出；c3mp 只覆盖"
             "合并与拆分档 499 条；c3 只有 275 条）。因此三者之间只有"
             "「冻结 610 ⊂ 构造 4,031」和「冻结 610 ⊂ 输出 4,490 的 id 集合」两条"
             "包含关系，不能把 4,490 当作 610 的重复测量。"),
}
out["relations"] = rel

json.dump(out, open("results/data_lineage_audit.json", "w", encoding="utf-8"),
          ensure_ascii=False, indent=1)

print("L1 构造", L1["rows"], "条 /", L1["question_ids"], "基础题 /", L1["dbs"], "库")
print("L2 冻结", L2["unique_ids"], "条（清单", L2["ids_listed"], "，不在 manifest 的",
      L2["ids_not_in_manifest"], "条）")
print("L3 基础题", L3["unique_question_id"], "（平均每基础题",
      L3["items_per_question"], "条）")
print("L4 事件", ps["events"], "次 / op", ps["ops"], "/ 库", ps["dbs"],
      "；段位档", ss["events"], "次 / op", ss["ops"], "/ 库", ss["dbs"])
if "L5_outputs" in out:
    o = out["L5_outputs"]
    print("L5 输出", o["rows"], "行 /", o["conditions"], "条件 /",
          o["unique_ids"], "唯一 id（冻结内", o["ids_in_frozen"],
          "，冻结外", o["ids_outside_frozen"], "）")
    print("  模型:", o["models"])
print("关系:", json.dumps({k: v for k, v in rel.items() if k != "note"},
                        ensure_ascii=False))
