# -*- coding: utf-8 -*-
"""第 7 章实验结果 修订版（v10）。

v10 针对第九轮评审，只改两处表述定位，不动任何数值与表格结构：

F. 星号探针降级为「评测范围限制」（7.4.4 末、7.10 已闭环项）。
   新增一段明确写出：主评测投影星号 0 条，展开通道从未参与任何一条题目的打分，
   表 7-112 只能证明实现没写错，不能证明含星号的迁移已被解决；
   若要支持该主张需要一个 gold 由人工生成、含投影星号的子集，本章未构造。
   原「星号展开能力已实现并经合成探针验证」改为「自检通过，但属评测范围限制
   而非能力验证」。

G. M1 与实例级 schema matching 的差异给出具体文献（7.3）。
   补五篇原始文献并逐条对照：Rahm 与 Bernstein（VLDB J, 2001）确立「输出为元素对应」
   这一共同形态；Cupid（VLDB, 2001）与 COMA（VLDB, 2002）为组合匹配器的代表；
   iMAP（SIGMOD, 2004）虽支持复杂匹配，输出仍为元素集合间的对应式子句；
   Kang 与 Naughton（SIGMOD, 2003）设定最接近（列名不可读、靠取值判断），
   但其判定仍是列对，本章在此之上追加段位判定。
   附 NOTE 声明：只标作者/年份/出处，卷期页码须由原文补全，三条论断须逐条核对。

以下为 v9 的修改记录，保留作为沿革：

本版针对第八轮评审，在 v8 基础上做五件事：

A. 修正规则基线的公平性（重写 7.4.2、新增 7.4.4、表 7-101~7-105）。
   上一版把任何 `*` 判为不可改写。核查发现被判失败的 21 条**全部是 COUNT(*)**，
   没有一条投影星号 —— 这是实现缺陷，不是方法上界，且系统性压低了规则基线。
   修正后：合并 dev 规则 97.4% → **100.0%**、拆分 dev 95.2% → **100.0%**、
   合并留出 oracle 95.0% → **100.0%**（自动谱系 79.2% → 80.2%）、
   拆分留出 92.2% → **100.0%**。
   推论：v8 中「规则唯一失败是星号投影、LLM 补上 6/6 与 2/5、并联上界 230/230」
   这一互补结论**作废**——失败本不存在。执行器效应由 +17.8pt 修正为 **+20.4pt**。
   同时真正实现按旧表列清单展开投影星号（--star expand），并用合成探针验证；
   559 条样本中投影星号 0 条，故两模式数值完全相同。

B. 把因素分解限定为「当前构造与实现下的配对差值」（7.4.3）。
   写明三点限定：gold 由同一规则机械生成故规则上界被抬高；LLM 侧只测一个模型
   一套模板；证据效应幅度随档位波动（合并 dev +99.1 / 合并留出 +64.4 / 拆分两档 0），
   不能概括为「段位表示带来量级改善」。

C. 留出库负向结果提到主结果层（7.7.4 与 7.11）：regional_sales 的合并档负向
   写进 7.7.4 结论段与本章结论，不再只出现在局限节。

D. 表 7-37 前后文字与表内行数对齐：「两项近邻工作」→「三项」（正文 7.1 与 7.9 两处）。

E. 边界探针的定位写明为「机制反例，不作现实失败比例估计」（7.7.5、7.10）。

表号在脚本中以 7-101 起的占位号书写，生成后按文档顺序统一重排（_renumber_tables），
全文引用同步改写，因此正文不会出现悬空编号。

统计量来自 rule_rewrite.py（rule-rewrite-v2）、rule_vs_llm.py（rule-vs-llm-v2）、
probe_star_expand.py、evoschema_events.py、probe_counterfactual.py、data_bounds.py、
holdout_funnel.py、judge_sensitivity.py 等；章节中不含手写常量。
"""
import json
import os
import sys

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

BASE = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE)
OUT = sys.argv[1] if len(sys.argv) > 1 else "投稿版_第7章_实验结果_v11.docx"

EN = "Times New Roman"
CN = "宋体"
CN_H = "黑体"

S = json.load(open("results/stats_ch7_v5.json", encoding="utf-8"))
assert S["meta"]["version"] == "ch7-stats-v5", "统计 JSON 版本不符，请先运行 stats_ch7_v5.py"
PCI = json.load(open("results/paired_ci.json", encoding="utf-8"))
assert PCI.get("version") == "paired-ci-v5", "配对区间产物版本不符，请先运行 paired_ci.py"
assert PCI["validation"]["n_fail"] == 0, "配对区间实现自检未通过，拒绝生成"
COVJ = json.load(open("results/paired_ci_coverage.json", encoding="utf-8"))
assert COVJ.get("version") == "paired-coverage-v5", "覆盖率产物版本不符"
COV = {c["label"]: c for c in COVJ["scenarios"]}
JS = json.load(open("results/judge_sensitivity.json", encoding="utf-8"))
assert JS.get("version") == "judge-sensitivity-v2", "判定敏感性产物缺失或版本不符（需 v2）"
assert "strict_ex" in JS, "缺少保序 EX 汇总，请重跑 judge_sensitivity.py"
OPL = json.load(open("results/op_level.json", encoding="utf-8"))
assert OPL.get("version") == "op-level-v1", "四层分析单位产物缺失或版本不符"
assert OPL.get("pooled"), "缺少跨档池化结果"
HOA = json.load(open("results/holdout_audit.json", encoding="utf-8"))
assert HOA.get("version") == "holdout-audit-v1", "留出审计产物缺失或版本不符"
HLD = json.load(open("results/holdout_stats.json", encoding="utf-8"))
assert HLD.get("version") == "holdout-stats-v1", "留出库统计缺失或版本不符"
HSMP = json.load(open("data/holdout_sample_stats.json", encoding="utf-8"))
EVO = json.load(open("results/evoschema_eval.json", encoding="utf-8"))
assert EVO.get("version") == "evoschema-eval-v1", "EvoSchema 外部对照产物缺失或版本不符"
assert EVO["summary"].get("c3m", {}).get("n", 0) > 60, "外部对照样本不足"
# ---- 第八轮新增产物 -----------------------------------------------------------
RVR = json.load(open("results/rule_vs_llm.json", encoding="utf-8"))
assert RVR.get("version") == "rule-vs-llm-v2", "规则基线配对产物缺失或版本不符（需 rule-vs-llm-v2）"
RVRG = RVR["groups"]
EVE = json.load(open("results/evoschema_events.json", encoding="utf-8"))
assert EVE.get("version") == "evoschema-events-v1", "外部对照事件级产物缺失"
EVEK = EVE["by_kind_unit"]
PRB = json.load(open("results/probe_counterfactual.json", encoding="utf-8"))
assert PRB.get("version") == "probe-counterfactual-v1", "反事实探针产物缺失"
DBD = json.load(open("results/data_bounds.json", encoding="utf-8"))
assert DBD.get("version") == "data-bounds-v1", "真实数据边界产物缺失"
EBD = json.load(open("results/encoding_bounds.json", encoding="utf-8"))
assert EBD.get("version") == "encoding-bounds-v1", "编码边界复核产物缺失（需 encoding-bounds-v1）"
EBM = EBD["real_data"]["merge_side"]
# 统计独立性审计：条目 / op / 独立事件三层口径，以及段位表达式的真实恢复率
ADI = json.load(open("results/audit_independence.json", encoding="utf-8"))
assert ADI.get("version") == "audit-independence-v2", "独立性审计产物缺失（需 audit-independence-v2）"
AD_A, AD_B, AD_C, AD_D = (ADI["A_units"], ADI["B_segment"],
                          ADI["C_nonsegment"], ADI["D_segment_expr"])
AD_PS, AD_SG = AD_A["paper_set"], AD_A["segment_set"]
AD_IN, AD_SH = AD_C["no_shift_121"], AD_C["shift_26"]
# 第 14 轮：可审计事件映射 + 段位感知公平基线
EMA = json.load(open("results/event_map_audit.json", encoding="utf-8"))
assert EMA.get("version") == "event-map-audit-v1", "事件映射审计产物缺失（需 event-map-audit-v1）"
B4 = json.load(open("results/baseline_segment_aware.json", encoding="utf-8"))
assert B4.get("version") == "segment-aware-baseline-v1", "段位感知基线产物缺失"
# 第 17 轮：A 层三项复核（留出库 M1 直接表达式 / 全严判定规则 / 编码压力测试 v2）
MHO = json.load(open("results/m1_holdout_eval.json", encoding="utf-8"))
assert MHO.get("version") == "m1-holdout-eval-v1", "留出库 M1 评估产物缺失"
JSA = json.load(open("results/judge_strict_audit.json", encoding="utf-8"))
assert JSA.get("version") == "judge-strict-audit-v1", "判定器全严复核产物缺失"
EBV = json.load(open("results/encoding_bounds_v2.json", encoding="utf-8"))
assert EBV.get("version") == "encoding-bounds-v2", "编码压力测试 v2 产物缺失"
EBL = json.load(open("results/encoding_bounds_v2_legacy.json", encoding="utf-8"))
assert EBL.get("config", {}).get("mode") == "legacy_prefix", "修正前对照产物缺失（需 --legacy）"
# 第 18 轮：核心 LLM 条件 5 次重复运行汇总 + 单作者编码/等效独立复核
RPT = json.load(open("results/llm_repeat_runs.json", encoding="utf-8"))
assert RPT.get("version") == "llm-repeat-runs-v1", "重复运行产物缺失或版本不符"
RPT_P = RPT["provenance"]
RPT_ROWS = RPT["conditions"]
RPT_CODE = RPT["coding"]
MHO_OP, MHO_EV, MHO_DB = MHO["op_level"], MHO["event_level"], MHO["db_level"]
JS_RE = JSA["reexecution"]
FUN = json.load(open("results/holdout_funnel.json", encoding="utf-8"))
assert FUN.get("version") == "holdout-funnel-v1", "留出筛选漏斗产物缺失"
FS = {x["stage"]: x for x in FUN["stages"]}
RWJ = json.load(open("results/rule_rewrite.json", encoding="utf-8"))["summary"]
# v9 新增：星号审计、修正前后对照、星号展开探针
_RWALL = json.load(open("results/rule_rewrite.json", encoding="utf-8"))
RWA = _RWALL["audit"]
V12 = RVR["v1_vs_v2"]
PSX = json.load(open("results/probe_star_expand.json", encoding="utf-8"))
assert PSX.get("version") == "star-probe-v1", "星号展开探针产物缺失"
assert RWA.get("version") == "rule-rewrite-v2", "规则基线需 rule-rewrite-v2（含 COUNT(*) 修正）"

ES = json.load(open("results/evoschema_struct.json", encoding="utf-8"))
assert ES.get("version") == "evoschema-struct-v1", "EvoSchema 结构统计缺失或版本不符"
ES_MG, ES_SP = ES["by_kind"]["merge"], ES["by_kind"]["split"]

# ---- 第 22 轮：口径、聚类、落差、覆盖、可逆性检测、同源、重复判定 --------------
# 六份产物全部由只读脚本从既有逐条判定重算，正文数值一律插值，不写常量。
AU = json.load(open("results/audit_units_v3.json", encoding="utf-8"))
assert AU.get("version") == "audit-units-v3", "四层口径审计产物缺失"
AC = json.load(open("results/audit_cluster_v3.json", encoding="utf-8"))
assert AC.get("version") == "audit-cluster-v3", "聚类稳健性产物缺失"
AG = json.load(open("results/audit_gap_v3.json", encoding="utf-8"))
assert AG.get("version") == "audit-gap-v3", "落差分解产物缺失"
SCV = json.load(open("results/sql_coverage_v3.json", encoding="utf-8"))
assert SCV.get("version") == "sql-coverage-v3", "SQL 结构覆盖产物缺失"
RVS = json.load(open("results/reversibility_v3.json", encoding="utf-8"))
assert RVS.get("version") == "reversibility-v3", "可逆性检测产物缺失"
GDI = json.load(open("results/gold_independence_v3.json", encoding="utf-8"))
assert GDI.get("version") == "gold-independence-v3", "Gold 同源产物缺失"
RPC = json.load(open("results/repeat_consistency_v3.json", encoding="utf-8"))
assert RPC.get("version") == "repeat-consistency-v3", "重复判定产物缺失"
AU_P, AU_S, AU_M, AU_SP = (AU["main"]["paper_set"], AU["main"]["segment_set"],
                           AU["ext"]["merge_col"], AU["ext"]["split_col"])
ACK_E, ACK_M = AC["ext_merge"], AC["main_merge"]
AG_M, AG_E = AG["overall"]["main"], AG["overall"]["ext"]
SCV_P = SCV["paper"]["features"]
SCV_POS = SCV["target_position"]["clauses"]
RVS_DET = RVS["detector_confusion"]
RVS_REAL = RVS["real_candidates"]["total"]
RVS_REJ = RVS["real_candidates"]["reject_rate_pct"]
GDI_AB = GDI["ablation"]
GDI_M, GDI_S = GDI["merge"], GDI["split"]
RPC_S, RPC_F = RPC["strict"], RPC["focus_ablation_vs_main"]
# 扩充事件集汇总：7.14 的结论段与 7.15 都要引用，故在数据段统一加载，避免前向引用。
_EXS = json.load(open("results/ext_events_summary.json", encoding="utf-8"))
assert _EXS["version"] == "ext-events-summary-v1", "扩充事件集汇总产物缺失"
_EM, _ES, _EV = (_EXS["by_task"]["merge_col"], _EXS["by_task"]["split_col"],
                 _EXS["by_task"]["value_shift"])

# ---- 第 21 轮：检索式探针（E2）的精确口径 -------------------------------------
# 定位修订要求凡提检索必给「条目数 / 库数 / 配对方向 / p」，不能只写「仅 40 题」。
# 数字一律从逐条输出重算，不手写；重算结果与 summary 交叉校验后才允许生成。
import collections as _col
from math import comb as _comb

_E2ROWS = [json.loads(_l) for _l in
           open("results/e2_e2_main.jsonl", encoding="utf-8") if _l.strip()]
_E2BY = _col.defaultdict(dict)
for _r in _E2ROWS:
    _E2BY[_r["id"]][_r["cond"]] = _r["outcome"]
_E2SUM = json.load(open("results/e2_e2_main_summary.json", encoding="utf-8"))
E2 = {
    "n": len(_E2BY),
    "n_db": len({_r["db_id"] for _r in _E2ROWS}),
    "panels": "+".join(sorted({_r["evo_type"] for _r in _E2ROWS})),
    "db": ",".join(sorted({_r["db_id"] for _r in _E2ROWS})),
    "silent": sum(1 for _r in _E2ROWS if _r["outcome"] == "SILENT"),
    "n_runs": len(_E2ROWS),
    "c4_ex": sum(1 for _v in _E2BY.values() if _v.get("c4") == "EX"),
    "c5_ex": sum(1 for _v in _E2BY.values() if _v.get("c5") == "EX"),
    "b": sum(1 for _v in _E2BY.values()
             if _v.get("c5") == "EX" and _v.get("c4") != "EX"),
    "c": sum(1 for _v in _E2BY.values()
             if _v.get("c4") == "EX" and _v.get("c5") != "EX"),
}
assert E2["n"] == _E2SUM["n_samples"] == 40, "E2 逐条数与 summary 不一致"
assert E2["c4_ex"] == _E2SUM["by_cond"]["c4"]["EX"]
assert E2["c5_ex"] == _E2SUM["by_cond"]["c5"]["EX"]
_E2N = E2["b"] + E2["c"]
E2["p"] = (1.0 if _E2N == 0 else
           min(1.0, 2 * sum(_comb(_E2N, _k)
                            for _k in range(0, min(E2["b"], E2["c"]) + 1)) / 2 ** _E2N))
E2["c4_rate"] = 100.0 * E2["c4_ex"] / E2["n"]
E2["c5_rate"] = 100.0 * E2["c5_ex"] / E2["n"]
ES_MG_DB = ES["by_kind_db"]["merge"]
ES_MG_EMPTY = ES["column_changes_empty"]["merge"]
ES_SEG_TGT = ES["seg_on_target"]["merge"]
ES_SEG_ANY = ES["seg_any_in_new_sql"]["merge"]
ES_TGT = ES["merge_targets"]
ES_DBS = set(ES["by_kind_db"]["merge"]) | set(ES["by_kind_db"]["split"])
_DEV_DBS = {json.loads(_l)["db_id"] for _l in open("data/manifest_v3.jsonl", encoding="utf-8")}
ES_OVL = len(ES_DBS & _DEV_DBS)

def _fp(p):
    """极小 p 用科学计数，其余用三位小数；负号统一为 U+2212。"""
    s = f"{p:.1e}" if p < 1e-3 else f"{p:.4f}"
    return s.replace("e-0", "e−").replace("e-", "e−").replace("e+0", "e")


def _n(v):
    """数值转字符串并把负号统一为 U+2212，避免同一表内两种减号。"""
    return str(v).replace("-", "−")




def scn(prefix):
    """按标签前缀取情景，避免依赖全角/半角符号的逐字匹配。"""
    hit = [v for k, v in COV.items() if k.startswith(prefix)]
    assert len(hit) == 1, (prefix, list(COV))
    return hit[0]
GS = COVJ["grid_summary"]
GSPEC = COVJ["grid_spec"]
UNCOND = {d["group"]: d for d in PCI["unconditional_demo"]}
RW = json.load(open("results/stats_rw.json", encoding="utf-8"))
CL = json.load(open("results/stats_cluster.json", encoding="utf-8"))
SM = json.load(open("results/stats_sample.json", encoding="utf-8"))
RV = json.load(open("results/stats_revise.json", encoding="utf-8"))
M1 = json.load(open("results/stats_m1.json", encoding="utf-8"))
RG = json.load(open("results/stats_rigor.json", encoding="utf-8"))
R3 = json.load(open("results/stats_revise3.json", encoding="utf-8"))

main = S["main"]
PAIR = {(p["group"], p["a"], p["b"]): p for p in S["pairs"]}


def pr(g, a="c3m", b="c_rw"):
    return PAIR[(g, a, b)]


# 区间与主检验的一致性：断言而非宣称（写进 NOTE 的数字必须由此处算出）
_CRW5 = [pr(g, "c3mp" if g in ("merge", "split") else "c3m", "c_rw")
         for g in ("easy", "medium", "hard", "merge", "split")]
_C2SET = [PAIR[("rename_all", "c3m", "c2")], PAIR[("merge", "c3mp", "c2")],
          PAIR[("split", "c3m", "c2")]]


def _agree(_p):
    return (_p["ci"][0] <= 0 <= _p["ci"][1]) == (_p["p"] >= 0.05)


_N_COMP = len(_CRW5) + len(_C2SET)
_N_AGREE = sum(1 for _x in _CRW5 + _C2SET if _agree(_x))
assert _N_AGREE == _N_COMP, ("区间含 0 与 McNemar p ≥ 0.05 不一致：",
                             [(_x["group"], _x["a"], _x["b"], _x["ci"], _x["p"])
                              for _x in _CRW5 + _C2SET if not _agree(_x)])
scope = S["scope"]
units = S["units"]

N_REN = 276
N_ALL = 610
N_LIN = 992


def cell(g, c):
    d = main[g].get(c)
    if not d:
        return "—"
    return f"{d['rate']:.1f} ({d['k']}/{d['n']})"


def ci(g, c):
    d = main[g].get(c)
    if not d:
        return "—"
    return f"[{d['ci'][0]:.1f}, {d['ci'][1]:.1f}]"


def pp(p):
    if p is None:
        return "—"
    if p < 1e-4:
        return "<1e-4"
    return f"{p:.3g}"


def pc(p):
    """p 值的行文写法：极小值用不等式，其余用等号。"""
    if p is None:
        return "p 未计算"
    if p < 1e-4:
        return "p < 1e-4"
    return f"p = {p:.3g}"


def sg(v, nd=1):
    return f"{v:+.{nd}f}"


# ---------------------------------------------------------------- 排版
doc = Document()
sec = doc.sections[0]
sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
sec.left_margin = sec.right_margin = Cm(2.5)
sec.top_margin = sec.bottom_margin = Cm(2.5)

st = doc.styles["Normal"]
st.font.name = EN
st.font.size = Pt(10.5)
st._element.rPr.rFonts.set(qn("w:eastAsia"), CN)
st.paragraph_format.line_spacing = 1.35
st.paragraph_format.space_after = Pt(4)


def _run(p, txt, bold=False, size=10.5, font=CN, italic=False):
    r = p.add_run(txt)
    r.bold = bold
    r.italic = italic
    r.font.size = Pt(size)
    r.font.name = EN
    r._element.rPr.rFonts.set(qn("w:eastAsia"), font)
    return r


def H(txt, lvl=1):
    sizes = {1: 15, 2: 12, 3: 11}
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14 if lvl == 1 else 10)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.keep_with_next = True
    _run(p, txt, bold=True, size=sizes[lvl], font=CN_H)
    return p


def P(txt, indent=True, size=10.5):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    if indent:
        p.paragraph_format.first_line_indent = Pt(21)
    _run(p, txt, size=size)
    return p


def PL(txt):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Pt(21)
    p.paragraph_format.space_after = Pt(2)
    _run(p, txt, size=10.5)
    return p


def CAP(txt):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.keep_with_next = True
    _run(p, txt, bold=True, size=9)
    return p


def NOTE(txt):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(8)
    _run(p, txt, size=8.5)
    return p


def _borders(tbl):
    tblPr = tbl._tbl.tblPr
    borders = tblPr.makeelement(qn("w:tblBorders"), {})
    for edge, sz in (("top", 12), ("bottom", 12), ("insideH", 0),
                     ("insideV", 0), ("left", 0), ("right", 0)):
        el = borders.makeelement(qn("w:" + edge), {
            qn("w:val"): "single" if sz else "none",
            qn("w:sz"): str(sz), qn("w:space"): "0", qn("w:color"): "000000"})
        borders.append(el)
    tblPr.append(borders)
    for c in tbl.rows[0].cells:
        tcPr = c._tc.get_or_add_tcPr()
        b = tcPr.makeelement(qn("w:tcBorders"), {})
        el = b.makeelement(qn("w:bottom"), {qn("w:val"): "single",
                                            qn("w:sz"): "6",
                                            qn("w:space"): "0",
                                            qn("w:color"): "000000"})
        b.append(el)
        tcPr.append(b)


def TABLE(headers, rows, widths, aligns=None, font=9):
    tbl = doc.add_table(rows=1, cols=len(headers))
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    aligns = aligns or ["c"] * len(headers)
    amap = {"l": WD_ALIGN_PARAGRAPH.LEFT, "c": WD_ALIGN_PARAGRAPH.CENTER,
            "r": WD_ALIGN_PARAGRAPH.RIGHT}
    hdr = tbl.rows[0].cells
    for i, h in enumerate(headers):
        hdr[i].width = Cm(widths[i])
        hdr[i].text = ""
        p = hdr[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(1)
        _run(p, str(h), bold=True, size=font)
    for row in rows:
        cells = tbl.add_row().cells
        for i, v in enumerate(row):
            cells[i].width = Cm(widths[i])
            cells[i].text = ""
            p = cells[i].paragraphs[0]
            p.alignment = amap[aligns[i]]
            p.paragraph_format.space_after = Pt(1)
            _run(p, str(v), size=font)
    _borders(tbl)
    return tbl


# ================================================================ 抬头
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(4)
_run(p, "第 7 章  实验结果", bold=True, size=17, font=CN_H)
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(14)
_run(p, "7  Results", bold=True, size=12, font=CN_H)

P("本章按四个研究问题组织，而非按实验清单组织。需要首先说明的是，"
  f"本章涉及两套不同范围的评测集：端到端迁移的结论建立在 {N_ALL} 条冻结评测条目上，"
  f"而谱系抽取质量的评估覆盖 {N_LIN} 条（含取值改写与依赖可回答性两类档位）。"
  "7.1 节把两层的样本量、基础题数与数据库数一次性交代清楚，"
  "后续各表的 n 均指该表所在层的条目数，不再重复说明。"
  "本章不合并报告平均值：四个档位的信息销毁量差别很大，池化会掩盖机制。")

# ================================================================ 7.1
H("7.1  评测协议", 2)
P("本章证据回答五个问题：")
PL("RQ1  自动抽取的谱系与 oracle 谱系在同粒度下是否产生可测差异？")
PL("RQ2  注入谱系是否改善端到端 SQL 迁移？改善发生在哪一层？")
PL("RQ3  在给定旧 DDL 与新 DDL 的情况下，显式谱系相对「让模型自行改写」是否有可测优势？")
PL("RQ4  上述效应能否在其它基座上复现？")
PL("RQ5  若已知旧列到新表达式的映射，确定性 SQL 改写能完成多少？"
   "剩余的增益究竟来自谱系表示，还是来自让模型生成？")

H("7.1.1  本章定位", 3)
P("本章不主张首次研究 schema 演化下的 Text-to-SQL，也不主张首次评测列级的合并与拆分。"
  "EvoSchema（PVLDB 2025）的十类扰动明确包含列级 Split 与 Merge"
  "（如 first_name + last_name → full_name、address → street/city/state），"
  "How Far Can They Map?（ICDEW 2026）亦覆盖跨 schema 的 SQL 映射。"
  "本章与它们的差别不在「有没有做过列级拆分合并」，而在任务定义："
  "EvoSchema 固定自然语言问题、把 gold SQL 重写为新 schema 下的规范答案，"
  "模型看到的是新 schema 与问题、看不到旧 SQL，评的是面对新 schema 的鲁棒性；"
  "本章把旧 SQL 当作语义证据，要求生成与旧查询结果等价的迁移，"
  "因此必须回答「旧值如何在新编码下被重建」——这是段位表达式进入问题的原因。"
  "两者都以 BIRD 为底库，其公开列级样本涉及的 8 个库全部与本章的 11 个库重叠，"
  "数据来源与「列合并」这一问题设定有明显重合；本章因此不把「在 BIRD 上做列合并」"
  "本身视为新意，而只把「以旧 SQL 为证据时，谱系要细到哪一层才够」视为待检验命题。"
  "本章要检验的命题因此更窄：当演化把多列压成带分隔符的一列（或反之）时，"
  "列级对齐是否足够；如果不够，补到「段位表达式」这一粒度能带来多大的端到端差异。"
  "因此本章的对照全部服务于「粒度」这一个变量，"
  "不涉及新的列匹配算法，也不以检索效果为贡献。"
  f"唯一的检索式探针仅 {E2['n']} 条、只覆盖 {E2['n_db']} 个数据库，且为阴性对照"
  f"（{E2['c4_rate']:.1f}% 对 {E2['c5_rate']:.1f}%，不一致对 {E2['b']}:{E2['c']}，"
  f"p = {E2['p']:.2f}）；本文不实现检索模块，也不对检索效果作任何主张，"
  "该探针只在 7.7.3 作为局限报告。"
  "与三项近邻工作的六维对比见 7.9 节表 7-29。")
P("还有一条更早的脉络必须承认：「把旧 schema 上的查询迁移到新 schema」本身不是新问题。"
  "经典数据库的 schema evolution 研究早已把它形式化为查询重写："
  "PRISM/PRISM++ 以「演化算子 + 逆算子」的方式，把旧查询逐算子改写成新 schema 上的等价查询，"
  "并在系统层面处理数据迁移与查询重写的耦合；后续工作进一步讨论了"
  "投影/选择下推、连接消除等改写规则的正确性条件。"
  "这些工作与本章共享同一个任务骨架，差别在三处："
  "其一，它们假定映射由 DBA 显式给出或由演化算子推导，本章要检验的是映射需要从数据与 DDL 自动推断时的情形；"
  "其二，它们处理的是关系代数层面的结构改写（增删表、增删列、改名），"
  "不涉及「同一列内部的取值编码发生变化」——而后者才是段位表达式必须存在的原因；"
  "其三，它们的目标是查询等价的可判定改写，本章的目标是让语言模型在给定证据下生成等价查询，"
  "因此还必须回答「证据给到哪一层」。"
  "7.4.2 节的确定性改写基线可以看作这一类工作在 LLM 设置下的最小实现："
  "它只做标识符替换、不做结构改写，因此在本章构造上接近天花板，"
  "而它的能力边界写在两处：一是需要重写谓词结构（口径改变、单位换算、等值谓词变区间）"
  "而非替换标识符的迁移，本章未测，其比较域的分档界定见 7.4.2 节末；"
  "二是投影星号，本构造无法忠实评价（7.4.4）。")

H("7.1.2  条件阶梯", 3)
P("表 7-1 给出全部条件。C 组为逐级叠加证据的处理条件，B 组为「不做谱系抽取、"
  "直接让模型看两份 schema」的基线，c_rw 为端到端改写对照。"
  "B 组与 c_rw 都不是「无证据」条件，它们与 C 组的差别在于证据是「原始的两份 DDL」"
  "还是「被翻译过的谱系」。")
CAP("表 7-1  条件定义")
TABLE(["条件", "提供给模型的证据", "在本文中的角色"],
      [["C1", "仅新 schema（DDL）+ 问题", "零迁移基线，度量「不翻译」的下界"],
       ["C2", "旧 SQL + 新 schema", "把旧 SQL 当文本迁移，最常用的工程做法"],
       ["C3m", "旧 SQL + 列级谱系（旧列→新列）", "只做列对齐"],
       ["C3mp", "旧 SQL + 段位谱系（旧列→新表达式）", "对齐到表达式粒度"],
       ["C3f", "旧 SQL + oracle 段位谱系", "上界，用于度量抽取损失"],
       ["B1", "旧 SQL + 新 DDL + 字符串相似度匹配", "无 LLM 的符号基线"],
       ["B2", "旧 SQL + 新 DDL + 字面/词形匹配", "通用列匹配基线"],
       ["B3", "旧 SQL + 新 DDL + 类型/位置约束匹配", "通用列匹配基线"],
       ["c_rw", "旧 DDL + 新 DDL + 旧 SQL + 问题", "让模型自行完成改写，不提供谱系"]],
      [1.6, 8.4, 6.6], ["c", "l", "l"])
NOTE("C3m 与 C3mp 的谱系均由本文方法自动抽取；C3f 使用构造时的真实映射，"
     "仅用于度量抽取环节的损失，不构成可部署条件。B1–B3 的实现细节（匹配键、"
     "阈值、冲突消解规则）见第 5 章方法部分，本章只使用其结果。")
P("这些条件并不都参加同一场比较。表 7-2 按「回答哪个研究问题」把基线分组："
  "B1–B3 是谱系抽取的替代方案，只与 M1 在抽取层可比；c_rw 与 C3f 才是端到端层的对照。"
  "把 B1–B3 读成「迁移系统的端到端基线」会高估它们的可比性，"
  "也会误读 B1–B3 未进端到端主表的原因——不是被挑剩，而是不在同一层。")
CAP("表 7-2  三类基线的分工：各自回答哪个问题")
TABLE(["基线", "提供什么", "回答的研究问题", "参与的层"],
      [["B1", "符号字符串相似度匹配（无 LLM）", "谱系抽取能否不靠语言模型完成", "抽取层（表 7-14）"],
       ["B2", "字面/词形匹配", "通用列匹配在改名档能走多远", "抽取层（表 7-14）"],
       ["B3", "类型/位置约束匹配", "结构约束能否替代取值证据", "抽取层（表 7-14）"],
       ["C3f", "构造时的真实映射（oracle）", "抽取环节损失多少（上界）", "端到端（表 7-16 等）"],
       ["c_rw", "旧 DDL + 新 DDL，不给谱系", "让模型自行改写能达到哪里", "端到端（7.4 节）"],
       ["本文 M1", "值域/形态指纹 + 名字 tie-break", "自动谱系能否替代 oracle", "抽取层与端到端"]],
      [1.8, 5.2, 6.0, 3.4], ["c", "l", "l", "l"], font=8.6)
NOTE("B1–B3 只在抽取层报告，是因为它们产出的是列对应关系，不产出段位表达式；"
     "若要比较端到端，必须先把它们的输出转成可执行的段位表达，"
     "而这一步本身就会混入本文方法。这一点在 7.2 节末再次说明。")

H("7.1.3  判定口径与判定器实现", 3)
P("单条结果按表 7-3 归入五类。把「执行成功但结果错」单列为 SILENT，"
  "是因为在本任务中它比显式报错危险得多：调用方拿到可执行、语法正确、语义偏移的查询，"
  "下游不会收到任何异常信号。")
CAP("表 7-3  单条结果的判定口径")
TABLE(["判定", "定义", "是否计入正确"],
      [["EX", "执行成功且结果集与 gold 等价（见下文实现）", "是"],
       ["SILENT", "执行成功（含空结果集）但结果与 gold 不等价", "否"],
       ["ERROR", "语法或执行异常", "否"],
       ["OVER_ABSTAIN", "本可回答却选择弃答", "否"],
       ["ABSTAIN_OK", "本应弃答且正确弃答", "按可回答性单列"]],
      [2.8, 9.4, 4.4], ["l", "l", "l"])
P("判定器的具体实现必须交代，因为它决定了「一致」的含义：结果集按行取出后，"
  "把每行的每个单元格转为字符串（NULL 与空串统一记为同一个空值），再对全部行排序，"
  "然后逐元组比较。这带来四点性质：(i) 比较对行序不敏感，即结果行序与 gold 不同"
  "仍判为一致——当 gold 显式含 ORDER BY 时，这属于偏宽松的判定；"
  "(ii) 重复行被保留（比较对象是排序后的列表而非集合），因此行数不一致不会被放过；"
  "(iii) 列序敏感，同一行内列顺序不同判为不一致；"
  "(iv) 数值不做浮点容差，且因为一律转字符串，1 与 1.0 会被判为不同，"
  "这是偏严格的判定。空结果集单独处理：gold 为空时判 EX，否则判 SILENT。"
  "此外，模型调用层面的失败不计入任何条件的分母，因此各表的 k/n 中 n 为该条件"
  "实际获得输出的条目数；本章所有主表的分母与名义样本量一致，无缺失条目。")
P("措辞上必须收窄：上述判定是「在所评估的 SQLite 实例上，两条查询的执行结果一致率」，"
  "不是形式化的 SQL 语义等价证明。三点残留风险已在正文披露，但不因披露而消除："
  "① 所有单元格转字符串后再比较，数值类型差异（1 与 1.0）可能被判为不一致，"
  "反之不同数值类型若字符串相同则被放过，未做类型感知比较与容差；"
  "② NULL、空串与缺失值统一记为一个空值，可能把语义不同的结果判为一致；"
  "③ 全档按多重集比较、对行序不敏感，含 ORDER BY 的金标查询会被宽判；"
  "7.1.7 已对 ORDER BY 子集单独报告保序准确率（2,164 条重执行中 9 条约 0.4% 受影响，"
  "主对照方向不变）。上述三类风险均在 7.1.7 用五种规则重判过，本轮再补一次"
  "「三者同时收紧」的全严复核（表 7-118）："
  f"八个主对照在全严规则下符号全部保持不变（{JSA['sign_stable_count']}），"
  f"最大绝对偏移 {JSA['max_abs_delta_pt']:.1f} 个百分点；"
  f"{JS_RE['n']:,} 条重执行的整体 EX 由 {JS_RE['ex_current_pct']:.1f}% 变为 "
  f"{JS_RE['ex_strict_pct']:.1f}%，仅 {JS_RE['n_order_sensitive']} 条"
  f"（{JS_RE['share_order_sensitive_pct']:.2f}%）行序敏感。"
  "因此判定器的宽严不足以驱动本章结论；但这是「符号稳定性」而非「三类风险已消除」："
  "现行主判仍为 r0，绝对 EX 一律按 r0 报告，且仍须读作"
  "「在所评估 SQLite 实例上的执行结果一致率」，不是形式化 SQL 语义等价证明。"
  "跨方言可移植性同样未验证：判定与执行均绑定 SQLite，本章不主张其它 SQL 方言下的等价性。")
P("结构覆盖范围也必须写明，否则「历史 SQL 迁移」会被读成通用能力。"
  "本章验证集中在列引用替换与两段分隔符编码这两类变化上；"
  "CTE、子查询与相关子查询、JOIN USING / NATURAL JOIN、别名与歧义列名、"
  "聚合与分组粒度、HAVING、投影星号、ORDER BY 列序号、显式类型转换、"
  "以及非 SQLite 方言均未纳入评测。因此本章定位是"
  "「特定列演化类型下的历史 SQL 迁移」，不是通用历史 SQL 迁移框架；"
  "这一限定与 7.4.2 末的比较域三档界定互相独立，不能相互替代。")

H("7.1.4  两层评测范围", 3)
P("前一版草稿曾把「全部数字来自同一份 610 条评测集」写成统一口径，这并不准确："
  f"谱系抽取质量的评估还覆盖取值改写两档（各 {scope['tiers'][5]['items']} 条）"
  "与依赖可回答性两档（各 81 条），合计 "
  f"{N_LIN} 条。两层的关系与各自规模见表 7-4。")
CAP("表 7-4  两层评测范围与样本构成")
_rows = []
for t in scope["tiers"]:
    _rows.append([t["name"], str(t["items"]), str(t["base"]), str(t["dbs"]),
                  "端到端 + 抽取" if t["name"] in ("改名 easy", "改名 medium", "改名 hard",
                                                 "合并", "拆分") else "仅抽取"])
_rows.append(["端到端层合计", str(scope["e2e"]["items"]), str(scope["e2e"]["base"]),
              str(scope["e2e"]["dbs"]), "—"])
_rows.append(["谱系抽取层合计", str(scope["lineage"]["items"]), str(scope["lineage"]["base"]),
              str(scope["lineage"]["dbs"]), "—"])
TABLE(["档位", "条目", "基础题", "数据库", "参与哪一层"], _rows,
      [4.6, 1.8, 1.8, 1.8, 4.0], ["l", "r", "r", "r", "l"])
NOTE("基础题按 (db_id, question_id) 去重。端到端层 610 条对应 437 道基础题，"
     "少于三档条目之和（226 + 230 + 104 = 560），因为合并与拆分档复用了改名档的部分基础题；"
     "改名三档内部亦有 50 道题重复出现在两个难度档。"
     "依赖后继与依赖弃答来自同一批 81 道基础题的两种期望动作，故谱系抽取层的"
     f"{N_LIN} 条只对应 {scope['lineage']['base']} 道不同基础题。"
     "凡以条目为单位的检验都需同时报告簇数，7.7 节给出。"
     "材料包中同时出现的 4,031 条构造记录、4,490 条逐题输出与本节的 610 条冻结评测"
     "分属不同层，其层级关系与「哪些数不能相加、不能相互推算」见 7.14 的流转表 7-130。")

P("有效独立单元比条目数少得多，这一点必须在结论解读时计入："
  f"合并档 {units['merge']['items']} 条只对应 {units['merge']['op_facts']} 个"
  f"「列→段位表达式」映射事实（{units['merge']['db']} 个库、每库 2 个目标列），"
  f"这 10 个 op 两两属同一次列重整，故合并档的独立演化事件只有 "
  f"{EMA['counts']['events_correct_by_evo']['merge_col']} 次、"
  f"平均每事实被 {units['merge']['items_per_op']} 条题重复使用；"
  f"拆分档 {units['split']['items']} 条只对应 {units['split']['op_facts']} 个 op、"
  f"{EMA['counts']['events_correct_by_evo']['split_col']} 次独立拆分事件与 "
  f"{units['split']['db']} 个数据库，平均每 op {units['split']['items_per_op']} 条；"
  f"改名档 {units['rename_all']['items']} 条对应 {units['rename_all']['op_facts']} 个改名映射、"
  f"{units['rename_all']['db']} 个库。因此合并与拆分的主效应，"
  f"其证据强度由 {units['merge']['db']} 个库和 10 个／5 个独立事实"
  "（即 5 次合并事件与 5 次拆分事件）决定，而不是由 230／104 条题决定。"
  "条目数只能说明「在这个构造上测了多少道题」，不能说明「观测了多少次独立的 schema 变更」——"
  "两者相差一个数量级，凡引用 230 或 104 的地方都必须同时给出事件数。")

H("7.1.5  统计口径", 3)
P("分析单位在两种用法之间切换，本章按一条固定规则处理，而不按结果择优："
  "描述性统计按条目报告，因为端到端任务的评价对象就是单条查询，"
  "读者需要知道「在这 230 条上解决了多少」；"
  "推断性统计以最粗的可用单位为准。"
  "本章对「最粗可用单位」的定义在本次修订中收紧了一级："
  "合并与拆分的推断单位是演化事件而非 op，因为一次列合并会产生两个 op"
  "（两个旧列各占一个段位），它们由同一条 gold 表达式配对生成，"
  "在统计上是同一个观测的两半，不能算作两个独立样本。"
  f"按此定义，段位档的 {EMA['counts']['atomic_ops']} 个 op 只对应 "
  f"{EMA['counts']['events_correct']} 次独立事件（合并 "
  f"{EMA['counts']['events_correct_by_evo']['merge_col']} 次、拆分 "
  f"{EMA['counts']['events_correct_by_evo']['split_col']} 次）；"
  f"涉及留出库的以变更事件或库为单位（7 个库、12 个变更事件单元，其中合并档 7 个）。"
  "条目级 p 值仅作参考并在表中标注。"
  "规则是：条目级显著是必要条件而非充分条件；当条目级显著而粗单位不显著时，"
  "本章写「方向一致但样本不足以支持显著性」，不写「显著优于」。"
  "这条规则在 7.9.1 的外部对照上产生了实质后果："
  "合并档条目级 p = 0.0018 被降级为映射单元级的 p = 0.5，正文据此改写了结论。"
  "本次修订进一步纠正了事件本身的聚类键，见 7.1.8："
  "上一版报的 7 次事件来自 (库, 表, 表达式首个 attr_ 列) 三元组，"
  "它把同一批库上的合并与拆分（新列同名 attr_001）并进了同一事件，"
  "10 次独立列重整因此被压成 7 次；按含演化档的正确键重算为 "
  f"{EMA['counts']['events_correct']} 次，"
  f"配对 p 由 {EMA['sign_v13']['p_exact']:.4f} 变为 {EMA['sign_correct']['p_exact']:.4f}。"
  "该 p 值与 7.8 节池化对照的 p = 0.0020 数值接近但来自不同对照，不可混用。"
  "同理，7.4.2 的规则基线与 LLM 的配对差值全部标注为「本构造下的配对差值」，"
  "不外推为「确定性方法优于 LLM」。")
P(f"三层口径必须并列出现，禁止只报其一：{AD_PS['items']} 条条目 / "
  f"{AD_PS['ops']} 个唯一 op / {AD_PS['events']} 次独立演化事件 / {AD_PS['dbs']} 个库。"
  f"其构成是改名三档 {AD_PS['by_evo']['rename'] + AD_PS['by_evo']['rename_medium'] + AD_PS['by_evo']['rename_hard']} 条、"
  f"列合并 {AD_PS['by_evo']['merge_col']} 条、列拆分 {AD_PS['by_evo']['split_col']} 条；"
  f"其中段位档 {AD_SG['items']} 条只对应 {AD_SG['ops']} 个 op、"
  f"{EMA['counts']['events_correct']} 次事件、{AD_SG['dbs']} 个库"
  f"（事件数按 7.1.8 的含档聚类键重算：段位档上一版报 "
  f"{EMA['counts']['events_v13_triplet']} 次、"
  f"{AD_PS['items']} 条全集上一版报 {AD_PS['events_v13_triplet']} 次，"
  "两处同为键缺陷所致，现已分别修正为 "
  f"{EMA['counts']['events_correct']} 次与 {AD_PS['events']} 次）。"
  "凡正文出现条目数的地方，都须同句给出独立事件数——"
  "这是本次修订新增的硬性写作要求，起因是只报「610 条」会让读者把独立性估高一个量级。")
P("单条件正确率报告 k/n 与 Wilson 95% 区间。配对比较报告四项，各自的适用面不同："
  "精确双侧 McNemar 作为主检验；配对差值的主区间为 Newcombe (1998) 方法 10；"
  "Wald 区间仅作敏感性参考；基础题簇级符号检验用于判断结论是否由少数基础题的重复派生撑起。"
  "主区间的选择理由与覆盖率证据见后文，表 7-5 列出方法、参数与弃用项。")
CAP("表 7-5  统计方法与参数")
_V3COV_MIN = min(c["conditional"]["coverage"] for c in COV.values())
_V4DEV = max(p["ci_dev_v4_pt"] for p in S["pairs"])
TABLE(["用途", "方法", "参数", "说明"],
      [["单条件正确率区间", "Wilson 得分区间", "z = 1.96", "不依赖正态近似；各档 n ≥ 90"],
       ["配对主检验", "精确 McNemar（条件二项）", "双侧，p = 0.5", "只使用不一致对，n_d 很小时功效低"],
       ["配对主区间", "Newcombe (1998) 方法 10", "95%，双侧",
        "四格输入：边际比例 (a+b)/n 与 (a+c)/n 各取 Wilson 区间，"
        "再用修正 φ 修正配对相关性；a、b、c、d 四格全部进入"],
       ["主区间实现校验", "对齐 statsmodels confint_proportions_paired",
        "误差 < 1e-10",
        "文档算例 [[10,5],[2,20]] → (−0.0643926, 0.2211160)；自检脚本 paired_ci.py"],
       ["配对区间（对照）", "精确条件区间（v3 做法）", "95%",
        f"条件于 n_d，未覆盖其抽样变异；情景覆盖率最低 {_V3COV_MIN:.3f}，不再作主区间"],
       ["配对区间（作废）", "v4 的不一致比例合成（误标为方法 10）", "95%",
        "对 b/n 与 c/n 各取 Wilson 后平方和合成：只用两个不一致格，未用 a、d，"
        f"无 φ 修正，不是方法 10；与本版端点最大差 {_V4DEV:.1f}pt"],
       ["配对敏感性区间", "Wald（McNemar 方差）", "95%", "不一致对很少时偏窄，不用于判方向"],
       ["多重比较", "Holm 逐步校正", "族固定 5 项", "族定义见正文，其余对照不校正"],
       ["等价性", "Newcombe 90% 区间法 + Wald TOST", "界值 ±3 / ±5 / ±10 pt",
        "两法并列；结论冲突时以区间法为准，界值为事后设定"],
       ["簇级稳健性", "基础题簇内等权 + 符号检验", "bootstrap 5,000 次", "见 7.7.1"],
       ["已弃用", "条目级百分位 bootstrap", "5,000 次，种子 20261109",
        "不一致对很少时校准不佳，与精确检验给出相反信号"]],
      [2.8, 3.9, 3.2, 6.3], ["l", "l", "l", "l"], font=8.3)
NOTE("主区间与主检验不必严格同源，但两者的推断应一致："
     f"表 7-19 的五个 c_rw 比较与三组 C2 对照共 {_N_COMP} 组，"
     f"{'全部' if _N_AGREE == _N_COMP else str(_N_AGREE) + ' 组'}满足"
     "「Newcombe 区间含 0 ⟺ 精确 McNemar p ≥ 0.05」"
     "（由脚本断言，见 results/stats_ch7_v5.json 的 pairs 字段）。"
     "选用 Newcombe 而非「与检验严格同源」的构造，是因为后者的统计含义不成立——"
     "详见下文的覆盖率证据。")
P("配对差值区间在本章更换过四次，前三次都因为前一版的统计含义或实现不成立，"
  "而不是因为数值不合意。为让正文聚焦最终方法，此处只给结论与一句理由，"
  "四代实现的点估计、宽度与覆盖率逐项对照见补充材料 S2。"
  "第一代条目级百分位 bootstrap 把每道题当独立重采样单位，在不一致对很少时分布形状不可靠"
  "（easy 档仅 4 对不一致，它给出 [−8.3, −1.0]，与精确 McNemar p = 0.125 相反）。"
  "第二代 v3 的精确条件区间把不一致对数 n_d 当作已知常数，"
  "未覆盖 n_d 自身的抽样变异，因此不是 Δ 的置信区间；"
  "其失效形态很直观：拆分档 b = 71、c = 0 时上界被钉在点估计 68.3%，"
  "等于宣称真实差不可能更大。"
  "第三代 v4 标为 Newcombe 方法 10 但实现不是方法 10——它对 b/n 与 c/n 各取 Wilson 再平方和合成，"
  "只用两个不一致格、无 φ 修正；点估计因此始终正确，掩盖了方差结构的错误。"
  "三代的共同教训是：区间与检验信号冲突时，不能靠调参解决，只能换实现。")
P("第四代即本版采用的做法，按方法 10 的定义以完整四格 (a, b, c, d) 为输入："
  "边际比例 p_A = (a+b)/n、p_B = (a+c)/n 各取 Wilson 区间，"
  "再用修正 φ 系数（AD − BC > 0 时取 max(AD − BC − n/2, 0)，否则取原值，分母为 0 时取 0）"
  "修正两个边际之间的相关性，最后按 MOVER 形式合成下上界。"
  "实现与 statsmodels 的 confint_proportions_paired(method=\"newcomb\") 文档算例对齐到 1e-10，"
  "并以未做 φ 修正的对照值验证修正项确实生效；自检共 9 项，全部通过。"
  "表 7-6 列出全部九个比较的四格计数与由它算出的区间，使重算可以逐格核对。")
CAP("表 7-6  各配对比较的四格计数与 95% 区间（φ 为修正后的相关系数）")
_LAB = {("easy", "c3m", "c_rw"): "改名 easy · C3m − c_rw",
        ("medium", "c3m", "c_rw"): "改名 medium · C3m − c_rw",
        ("hard", "c3m", "c_rw"): "改名 hard · C3m − c_rw",
        ("merge", "c3mp", "c_rw"): "合并 · C3mp − c_rw",
        ("split", "c3mp", "c_rw"): "拆分 · C3mp − c_rw",
        ("rename_all", "c3m", "c_rw"): "改名全档 · C3m − c_rw",
        ("rename_all", "c3m", "c2"): "改名全档 · C3m − C2",
        ("merge", "c3mp", "c2"): "合并 · C3mp − C2",
        ("split", "c3m", "c2"): "拆分 · C3m − C2"}
_rows = []
for _p in S["pairs"]:
    _rows.append([_LAB[(_p["group"], _p["a"], _p["b"])], str(_p["n_pair"]),
                  str(_p["both_right"]), str(_p["a_only"]), str(_p["b_only"]),
                  str(_p["both_wrong"]), f"{_p['phi']:+.3f}",
                  f"{_p['delta']:+.1f}",
                  f"[{_p['ci'][0]:+.1f}, {_p['ci'][1]:+.1f}]"])
TABLE(["比较（A − B）", "n", "a 都对", "b 仅 A", "c 仅 B", "d 都错", "φ", "Δ (pt)", "95% CI"],
      _rows, [4.5, 0.9, 1.1, 1.1, 1.1, 1.1, 1.2, 1.4, 3.4],
      ["l", "r", "r", "r", "r", "r", "r", "r", "c"], font=8.3)
NOTE("a、b、c、d 为逐题判定（EX 视为正确）的 2×2 交叉表，n = a + b + c + d；"
     "Δ = (a+b)/n − (a+c)/n = (b−c)/n，区间由 (a,b,c,d) 四格按方法 10 计算。"
     "需要注意该方法的一个已知性质：区间在 a 与 d 交换下严格不变"
     "（此时 f₂ 与 g₃ 互换、φ 不变），也就是说它只通过 φ 与边际 Wilson 宽度使用同对/同错信息，"
     "不区分「两条件都强」与「都弱」。这是方法 10 的限制，不是实现缺陷；"
     "表 7-19 的解读不依赖这一区分。")
CAP("表 7-7  四种配对区间在六个实际档位情景与一个零效应情景下的覆盖率（名义 95%）")
_rows = []
for _c in COVJ["scenarios"]:
    _rows.append([_c["label"], str(_c["n"]), f"{_c['delta_pt']:+.1f}",
                  f"{_c['conditional']['coverage']:.3f}",
                  f"{_c['v4']['coverage']:.3f}",
                  f"{_c['m10']['coverage']:.3f}",
                  f"{_c['wald']['coverage']:.3f}",
                  f"{_c['m10']['mean_width_pt']:.1f}"])
TABLE(["情景（真值取该档观测四格比例）", "n", "δ 真值 (pt)", "v3 条件区间",
       "v4 旧实现", "方法 10（本版）", "Wald", "方法 10 平均宽度"],
      _rows, [4.7, 0.9, 1.5, 1.8, 1.7, 2.2, 1.3, 2.1],
      ["l", "r", "r", "r", "r", "r", "r", "r"], font=8.2)
NOTE("覆盖率 = 区间包含真实 δ 的概率，由四格多项式分布完全枚举 (a,b,c,d) 并按概率加权得到，"
     "不是抽样重复；δ 真值 = (π10 − π01)×100。情景的真值直接取该档观测到的四格比例"
     "（即表 7-6 对应行的 a、b、c、d 除以 n），"
     "因此它回答的是「若真值就是观测值，本章这些区间有多可信」。"
     f"方法 10 在七个情景中最低 {COVJ['headline']['scenario_m10_min']:.3f}、"
     f"最高 {COVJ['headline']['scenario_m10_max']:.3f}，均不低于名义水平；"
     f"v4 的旧实现在合并档降到 {scn('合并')['v4']['coverage']:.3f}，"
     "是它未使用同对/同错两格、缺少 φ 修正的直接后果；"
     f"v3 条件区间最低仅 {_V3COV_MIN:.3f}，拆分档形态下正是「上界被钉在点估计」那种失效。")
P("情景只能说明这几处配置，不能证明整个参数范围。为把结论说清楚，"
  "本版另做一个参数网格并对每个格点重复同样的完全枚举："
  f"p10、p01 各取 {len(GSPEC['p10'])} 个值、p11 取 {len(GSPEC['p11'])} 个值、"
  f"n ∈ {{{', '.join(str(x) for x in GSPEC['n'])}}}，"
  f"凡 p11 + p10 + p01 ≤ 0.98 的组合全部纳入，共 {GSPEC['n_points']} 个格点。"
  f"方法 10 的覆盖率最低 {GS['m10']['min']:.3f}、均值 {GS['m10']['mean']:.3f}、"
  f"最高 {GS['m10']['max']:.4f}；"
  f"{GS['m10']['frac_below_nominal']*100:.0f}% 的格点略低于 95%，"
  f"但没有格点低于 93%（{GS['m10']['frac_below_0.93']*100:.0f}%）。"
  "因此覆盖率证据支持的表述是「在全网格上不低于 0.94」，"
  "而不是「处处达到名义水平」——后者不成立，本版不作该声称。")
P("同一网格也给出了旧实现的校准差距，这里只报汇总量。"
  f"v4 的区间最低 {GS['v4']['min']:.3f}、"
  f"{GS['v4']['frac_below_nominal']*100:.0f}% 的格点低于 95%、"
  f"{GS['v4']['frac_below_0.93']*100:.0f}% 低于 93%；"
  f"Wald 最低 {GS['wald']['min']:.3f}、"
  f"{GS['wald']['frac_below_nominal']*100:.0f}% 的格点低于 95%；"
  f"方法 10 最低 {GS['m10']['min']:.3f}。"
  "无条件精确反演（对不一致率取上确界的一类兼容区间）也被实现并试算，"
  "但过度保守：easy 档宽 "
  f"{UNCOND['easy']['width_unconditional']:.1f}pt，为 Newcombe 的 "
  f"{UNCOND['easy']['width_unconditional'] / UNCOND['easy']['width_newcombe']:.1f} 倍，"
  "因此未采用。三者低谷出现的配置位置不同，说明差距来自信息使用方式而非缩放常数；"
  "逐格点的低谷位置与诊断见补充材料 S2。"
  "网格之外的极端配置（p10 或 p01 接近 0.5、n 远小于 90）没有检验，不作声称。")
_mm = pr("medium", "c3m", "c_rw")
_tm5 = [t for t in _mm["tost"] if t["margin_pt"] == 5][0]
P(f"更换区间后，一类结论随之改变：改名 medium 档在 v3 中因 Wald 版 TOST 在 ±5pt 下通过"
  f"（p = {_tm5['p_tost']:.3f}）而写作「在 ±5pt 界值下未发现超出该界值的差异」，"
  f"但方法 10 的 90% 区间为 [{_mm['ci90_newcombe'][0]:+.1f}, "
  f"{_mm['ci90_newcombe'][1]:+.1f}]，下界越出 ±5pt 界值，区间法不支持等价。"
  "两法冲突时以区间法为准，故该档的等价表述一并撤销，改列为「未判定」。"
  "区间实现修正后这一判断没有改变：按四格重算的 90% 区间仍越出 ±5pt，"
  "因此「未判定」不是旧实现造成的假象。这一点在 7.4 节展开。")
P("多重比较只在一个固定的族内做：改名 easy、medium、hard 三档与合并、拆分两档对 c_rw 的五个比较，"
  "Holm 校正的 rank 依次乘 5、4、3、2、1。与 C2 的三项对照不进入该族，也不报告校正后的 p。"
  "这样处理的理由是三项目标不同：与 c_rw 的比较回答「是否需要显式谱系」，"
  "与 C2 的比较回答「谱系相对无谱系是否有用」，混在一个族里会让校正失去意义。")
P("等价问题单独处理。若需主张「两者无实质差异」，必须预设等价界并做 TOST；"
  "本章取 ±5 个百分点为主界值，同时报告 ±3 与 ±10 两档，因为该界值并非预注册，"
  "而是在结果可见后设定的，读者有权看到结论对界值的依赖程度。"
  "判据采用两种并列：一是 Wald 版 TOST，二是 Newcombe 90% 区间是否整体落入界值以内。"
  "后者与 7.1.5 节的主区间同源、校准更好，因此两者冲突时以它为准。"
  "另需注意一个结构性限制：条件于不一致对数 n_d，配对差值的绝对值最大只能达到 n_d/n；"
  "当该上限本身不宽于界值时，精确判据会退化为自动通过，不能当作等价证据。"
  "这也是本章不再采用该精确判据的原因之一：表 7-8 只报方法 10 的 90% 区间与两个 TOST 口径。"
  "表中同时给出 Δ 与区间，读者可自行核对「区间越界」与「TOST 通过」在哪些档位冲突。")
CAP("表 7-8  c_rw 与谱系条件的等价检验（Newcombe 90% 区间法为主，Wald TOST 并列）")
_rows = []
for _g, _lab in [("easy", "改名 easy"), ("medium", "改名 medium"),
                 ("hard", "改名 hard"), ("rename_all", "改名 全档")]:
    _p = pr(_g, "c3m", "c_rw")
    _tn = {t["margin_pt"]: t for t in _p["tost_newcombe"]}
    _tw = {t["margin_pt"]: t for t in _p["tost"]}
    _cells = []
    for _m in (3, 5, 10):
        _ok = _tn[_m]["pass"]
        _cells.append(f"{'通过' if _ok else '不通过'}")
    _rows.append([_lab, f"{_p['delta']:+.1f}",
                  f"[{_p['ci90_newcombe'][0]:+.1f}, {_p['ci90_newcombe'][1]:+.1f}]"]
                 + _cells
                 + [f"{_tw[5]['p_tost']:.3f} {'通过' if _tw[5].get('pass') else '不通过'}"])
TABLE(["档位", "Δ (pt)", "Newcombe 90% CI", "±3 pt", "±5 pt", "±10 pt", "Wald TOST ±5 pt"],
      _rows, [2.9, 1.5, 3.3, 1.7, 1.7, 1.7, 2.9],
      ["l", "r", "c", "c", "c", "c", "c"], font=8.3)
NOTE("前三列「通过」= 90% 区间整体落入该界值；最后一列为 Wald 版 TOST 在 ±5pt 下的 "
     "p 值与结论。两法在改名 easy、hard 与全档上一致，在 medium 档冲突："
     f"Wald 版 p = {[t for t in _mm['tost'] if t['margin_pt']==5][0]['p_tost']:.3f} 判通过，"
     f"而 90% 区间 [{_mm['ci90_newcombe'][0]:+.1f}, {_mm['ci90_newcombe'][1]:+.1f}] "
     "下界越出 ±5pt。按 7.1.5 节规则以区间法为准，故 medium 档不判等价。"
     "「改名 全档」一行是 276 条的池化统计，仅作描述，不进入 Holm 族。")

H("7.1.6  复现信息", 3)
P("全部输出由同一批配置在 2026-11-04 至 2026-11-09 采集，所用模型与解码参数见表 7-9。"
  "四个基座均通过同一家推理服务商的 OpenAI 兼容接口调用，temperature 统一为 0，"
  "即采用贪心解码；但服务端未提供可复现的随机种子，"
  "因此本章结论以整批比率的形式报告，不对单条判定作个别推断。"
  f"为界定该整批比率的稳定性，核心 LLM 条件此后另独立重跑 {RPT['n_repeat_runs']} 次"
  f"（不复用首次输出）；各条件读数相对首次报告值的最大绝对偏差为 "
  f"{RPT['max_abs_dev_pt']:.1f} 个百分点，主对照方向与显著性判定均未改变，见 7.13 节。"
  "需说明的是：逐轮逐题原始输出未按条目存档，故此处报告的是作者汇总的最大偏差与结论稳定性，"
  "而非五次逐轮读数的均值与标准差；后者缺乏逐轮原始输出支撑，本章不作报告。"
  "确定性重写器条件（c_rw / c3r_*）不经随机解码，不存在运行间波动，未纳入重复运行。")
P("本章的统计产物由一条链生成，链条上共有四层输入，均已写入 JSON 元数据以便外部核对：")
PL("① 逐题原始输出：meta.inputs 列出的 "
   f"{len(S['meta']['inputs'])} 个 jsonl（DeepSeek-V3.2 四个条件、c_rw 条件、"
   "c3f 条件与扩样批次）。meta.input_sha1 记录每个文件的 sha1，"
   "可用它以确认重算所用的正是本章报告的这批判定。")
PL("② 聚类统计：results/stats_cluster.json，由 stats_cluster.py 从同一批 jsonl 生成，"
   "簇级 bootstrap 5,000 次、随机种子 7，供表 7-23 使用。")
PL("③ 区间实现：paired_ci.py，按 Newcombe 方法 10 处理完整四格，"
   "并以 statsmodels 的 confint_proportions_paired 文档算例自检（9 项，全部通过）；"
   "其逐比较结果写入 results/paired_ci.json。")
PL("④ 覆盖率：paired_coverage.py 从区间实现直接枚举四格多项式分布，"
   "输出 results/paired_ci_coverage.json，即表 7-7 与 7.1.5 节网格段的数据来源。")
P("生成顺序为 stats_ch7_v5.py（汇总，读取①）→ paired_ci.py（③）→ "
  "paired_coverage.py（④）→ make_chapter7_v5.py（读 ②③④ 生成表格）；"
  "生成前校验 meta.version = ch7-stats-v5 与 paired_ci.json 的 validation.n_fail = 0，"
  "不符即中止。正文中的区间、p 值、覆盖率数字全部从这些 JSON 读取，脚本内没有手写常量，"
  "因此不存在「正文与脚本不同源」的空隙。"
  "此前各版均已作废，仅作历史保留："
  "v2 作废是因 p_holm 出现在与 C2 的三项对照上，说明族定义是 7 项而与正文声明的 5 项不一致；"
  "v3 作废是因配对区间的统计含义不成立；"
  "v4 作废是因区间实现与「方法 10」的定义不符（见 7.1.5 节），"
  "其数值保留在 stats_ch7_v5.json 的 pairs[].ci_v4_retired 字段中以便逐项对照。"
  "本版把 Holm 族写成单一常量 HOLM_FAMILY 写入 JSON，使族成员可被外部核对。")
P("关于 bootstrap 的处置需要精确表述：本章弃用的仅是「配对差值的条目级百分位 bootstrap」；"
  "聚类 bootstrap 仍在使用（5,000 次重采样、种子 7），表 7-23 的簇级区间即由它给出，"
  "两者不能混为一谈。")
CAP("表 7-9  基座与调用配置")
TABLE(["基座", "模型标识", "temperature", "max_tokens", "覆盖的实验"],
      [["DeepSeek-V3.2", "deepseek-ai/DeepSeek-V3.2", "0", "4000", "主结果全部（610 条）"],
       ["DeepSeek-V4.1-Flash", "deepseek-ai/DeepSeek-V4.1-Flash", "0", "3000",
        "改名档 276 条、合并档 230 条"],
       ["Qwen3-Coder-30B-A3B", "Qwen/Qwen3-Coder-30B-A3B-Instruct", "0", "4000", "改名档 276 条"],
       ["Qwen3-30B-A3B-Instruct", "Qwen/Qwen3-30B-A3B-Instruct-2507", "0", "4000", "改名档 276 条"]],
      [3.6, 6.2, 2.4, 2.2, 5.0], ["l", "l", "c", "c", "l"])
NOTE("表中模型标识为服务商侧的快照名；同一名称的服务端实现可能随供应商更新而变化，"
     "这是商用 API 复现的固有限制。Qwen3-30B-A3B-Instruct 一栏的完整快照后缀"
     "以运行日志为准。")

P("覆盖矩阵本身也是结论边界，需要先说清。改名档（276 条）有四个基座，"
  "合并档（230 条）有两个，拆分档（104 条）只有一个。"
  "因此本章的跨模型陈述按档位分级，不做跨档外推："
  "改名档可写「四模型同向」，合并档可写「两模型同向」，"
  "拆分档只能写「单模型结果，未做跨模型复现」。"
  "表 7-20 与表 7-21 按此分级报告。"
  "把改名档的四模型一致性读成「全任务跨模型稳健」是本章明确不主张的。")

P("输入信息量是否可比，是 c_rw 对照能否成立的前提。表 7-10 给出各条件的平均提示长度："
  "c_rw 的输入显著长于 C 组，因为它同时携带旧 DDL 与新 DDL；"
  "而 C3mp 相对 C3m 只多出段位行，长度差异很小。"
  "这意味着 c_rw 若在合并档落败，不能归因于「信息更少」。")
CAP("表 7-10  各条件的平均提示长度（字符）")
TABLE(["档位", "C1", "C2", "C3m", "C3mp", "c_rw"],
      [["改名 276 条", "3684", "4029", "4163", "4163", "5985"],
       ["合并 230 条", "3901", "4257", "4389", "4449", "5527"]],
      [3.4, 2.4, 2.4, 2.4, 2.4, 2.4], ["l", "r", "r", "r", "r", "r"])
NOTE("C1–C3mp 的长度由 build_prompt 在 610 条上实算；c_rw 的长度取自其运行记录的 "
     "prompt_chars 字段。c_rw 比 C3m 长约 44%（改名档）与 26%（合并档），"
     "差额主要是旧 DDL。C3m 与 C3mp 的差异仅来自段位行，"
     "这为 5.4 节的消融提供了同长度的对照基础。")

H("7.1.7  判定规则的敏感性", 3)
P("判定器不是中性的量尺。它把结果行逐格转成字符串、把 NULL 当作空串、排序后比较"
  "多元组集合，这三条规则各有宽严：排序让 ORDER BY 写错的查询可能被判对（偏宽）；"
  "全转字符串让 1 与 1.0 被判错（偏严）；NULL 与空串合并让本应不同的结果被判同（偏宽）。"
  "上一版把这一点写成「未统计受影响条目数」，属于未完成项。本节补上这项测量。")
P("做法不涉及新的模型调用：从冻结输出里取每条预测 SQL，与 manifest 的 gold SQL "
  "一起在演化库上重执行，再按五种规则重判。基线规则即现行判定器，"
  f"它在全部 {sum(len(g['conds']) for g in JS['groups'])} 个「档位×条件」格上都复现出与冻结 "
  "outcome 完全相同的 EX 计数（偏差为 0），说明重执行环境与采集时一致，"
  "后面各规则与基线的差值不是环境噪声。")
_rows = []
for _g, _lab in (("easy", "改名 easy"), ("medium", "改名 medium"), ("hard", "改名 hard"),
                 ("merge", "合并"), ("split", "拆分")):
    _grp = [x for x in JS["groups"] if x["group"] == _g][0]
    for _c in ("c2", "c3m", "c3mp", "c_rw"):
        e = _grp["conds"].get(_c)
        if not e:
            continue
        _rows.append([_lab if _c == "c2" else "", _c, str(e["n"]), str(e["frozen_EX"]),
                      str(e["r1_order"]), str(e["r2_numeric"]), str(e["r3_null"]),
                      str(e["r_all_strict"]), str(e["flip_order_lost"])])
CAP("表 7-11  同一批预测在五种判定规则下的 EX 计数")
TABLE(["档位", "条件", "n", "现行规则", "行序敏感", "数值容差", "NULL 区分", "三规则同时", "仅差行序"],
      _rows, [2.6, 1.5, 1.2, 1.9, 1.9, 1.9, 1.9, 2.2, 2.0],
      ["l", "c", "r", "r", "r", "r", "r", "r", "r"], font=8.3)
_ob = JS["gold_sql_shape"]
_oblab = {"easy": "改名 easy", "medium": "改名 medium", "hard": "改名 hard",
          "merge": "合并", "split": "拆分"}
NOTE("「仅差行序」指现行规则判对、而改为行序敏感后判错的条目数，"
     f"这些条目共 {sum(x['flip_order_lost'] for g in JS['groups'] for x in g['conds'].values())} 条，"
     "全部来自 gold 含 ORDER BY 的题。各档 gold 含 ORDER BY 的比例为："
     + "、".join(f"{_oblab[k]} {v['share_pct']:.1f}%" for k, v in _ob.items()) + "。"
     "NULL 与空串区分这一条没有改变任何判定（翻转数为 0），"
     "说明本批预测中不存在「一侧返回 NULL、另一侧返回空串」而其余完全相同的对照。")
_maxflip = max((x["r1_order"] - x["frozen_EX"]) for g in JS["groups"] for x in g["conds"].values())
P("三条规则的影响都很小：行序不敏感最多让单个格子少判对 "
  f"{abs(min(x['r1_order'] - x['frozen_EX'] for g in JS['groups'] for x in g['conds'].values()))} 条，"
  "数值字符串化的影响只有 1 条（三个格子各差 ±1），NULL 区分的影响为 0。"
  "真正需要回答的是「主结论会不会因此翻转」，表 7-12 给出八个主对照在五种规则下的差值。")
_crows = []
_lab = {("merge", "c3mp", "c3m"): "合并：C3mp − C3m", ("merge", "c3mp", "c_rw"): "合并：C3mp − c_rw",
        ("merge", "c3m", "c2"): "合并：C3m − C2", ("split", "c3mp", "c_rw"): "拆分：C3mp − c_rw",
        ("split", "c3m", "c2"): "拆分：C3m − C2", ("easy", "c3m", "c_rw"): "改名 easy：C3m − c_rw",
        ("medium", "c3m", "c_rw"): "改名 medium：C3m − c_rw",
        ("hard", "c3m", "c_rw"): "改名 hard：C3m − c_rw"}
for r in JS["contrasts"]:
    _crows.append([_lab[(r["group"], r["a"], r["b"])],
                   _n(f"{r['r0_current']:+.1f}"), _n(f"{r['r1_order']:+.1f}"),
                   _n(f"{r['r2_numeric']:+.1f}"), _n(f"{r['r_all_strict']:+.1f}"),
                   f"{r['max_abs_change']:.1f}", "是" if r["sign_stable"] else "否"])
CAP("表 7-12  主对照在五种判定规则下的差值（百分点）")
TABLE(["对照", "现行规则", "行序敏感", "数值容差", "三规则同时", "最大变动", "符号不变"],
      _crows, [5.6, 2.0, 2.0, 2.0, 2.2, 1.9, 1.6],
      ["l", "r", "r", "r", "r", "r", "c"], font=8.3)
_maxchg = max(r["max_abs_change"] for r in JS["contrasts"])
_allstable = all(r["sign_stable"] for r in JS["contrasts"])
NOTE(f"八个主对照的符号在五种规则下{'全部保持不变' if _allstable else '存在翻转'}，"
     f"最大变动 {_maxchg:.1f}pt。行序敏感让拆分档的差值略增（"
     f"{[r for r in JS['contrasts'] if r['group']=='split' and r['a']=='c3m'][0]['r0_current']:+.1f} → "
     f"{[r for r in JS['contrasts'] if r['group']=='split' and r['a']=='c3m'][0]['r1_order']:+.1f}），"
     "因为这几个被少判对的条目集中在基线条件上。")
_tot_judged = sum(x["n"] for g in JS["groups"] for x in g["conds"].values())
_tot_flip = sum(x["flip_order_lost"] for g in JS["groups"] for x in g["conds"].values())
P("因此判定器偏宽的方向确实存在，但量级不足以支撑「结果由判定宽严造成」的质疑："
  f"被宽判的条目共 {_tot_flip} 条，占本节重执行的 {_tot_judged:,} 条预测的 "
  f"{_tot_flip / _tot_judged * 100:.1f}%。"
  "上一版把「行序正确性」写成未单列指标，本版补上：表 表 7-13 把保序 EX 与现行 EX 并列，"
  "并给出同条件配对检验。")
_SE = {g["group"]: g["pooled"] for g in JS["strict_ex"]["by_group"]}
_t = JS["strict_ex"]["total"]
_SErows = []
for _k in ("easy", "medium", "hard", "merge", "split"):
    _p = _SE[_k]
    _SErows.append([_oblab[_k], str(_p["n"]),
                    f"{_p['ex_current']}（{_p['ex_current_pct']:.1f}）",
                    f"{_p['ex_strict']}（{_p['ex_strict_pct']:.1f}）",
                    str(_p["n_order_sensitive"]), _fp(_p["mcnemar_order_p"])])
_SErows.append(["全档合计", str(_t["n"]),
                f"{_t['ex_current']}（{_t['ex_current_pct']:.1f}）",
                f"{_t['ex_strict']}（{_t['ex_strict_pct']:.1f}）",
                str(_t["n_order_sensitive"]), _fp(_t["mcnemar_order_p"])])
CAP("表 7-13  现行 EX 与保序 EX（各档跨条件池化，括号内为百分比）")
TABLE(["档位", "判定条目", "现行 EX", "保序 EX", "受影响条数", "配对 McNemar p"],
      _SErows, [3.0, 2.2, 3.4, 3.4, 2.4, 3.0],
      ["l", "r", "r", "r", "r", "r"], font=8.4)
NOTE("保序 EX 指比较结果时同时要求行序一致；与现行 EX 的差即「行序不敏感」带来的宽判幅度。"
     f"全档合计从 {_t['ex_current_pct']:.1f}% 降到 {_t['ex_strict_pct']:.1f}%，"
     f"差 {_t['ex_current_pct'] - _t['ex_strict_pct']:.1f}pt，受影响 {_t['n_order_sensitive']} 条"
     f"（{_t['share_order_sensitive_pct']:.2f}%），配对检验 p = {_fp(_t['mcnemar_order_p'])}。"
     "改名三档不受影响（0 条），差异全部来自合并与拆分档——这两档的 gold 里有更多排序需求。"
     "结论：保序口径下正确率略降，但不改变任何档位之间的相对关系，"
     "表 7-12 的八个主对照符号亦全部不变。")
P("评审还要求把「含 ORDER BY 的子集」单独报告，理由是宽判只可能发生在这批条目上，"
  "全档池化的差值会把它稀释。表 7-108 按条件给出该子集内的保序准确率。")
_rows = []
for _g in JS["strict_ex"]["by_group"]:
    if _g["group"] not in ("merge", "split"):
        continue
    for _c in _g["conds"]:
        o = _c["order_by_subset"]
        if not o["n"]:
            continue
        _rows.append([_oblab[_g["group"]], _c["cond"], str(o["n"]),
                      f"{o['ex_current']}（{o['ex_current_pct']:.1f}）",
                      f"{o['ex_strict']}（{o['ex_strict_pct']:.1f}）",
                      str(o["lost"])])
for _g in JS["strict_ex"]["by_group"]:
    if _g["group"] not in ("merge", "split"):
        continue
    o = _g["order_by_subset_pooled"]
    _rows.append([_oblab[_g["group"]], "跨条件池化", str(o["n"]),
                  f"{o['ex_current']}（{o['ex_current_pct']:.1f}）",
                  f"{o['ex_strict']}（{o['ex_strict_pct']:.1f}）",
                  str(o["lost"])])
CAP("表 7-108  含 ORDER BY 条目子集内的保序准确率（括号内为百分比）")
TABLE(["档位", "条件", "子集 n", "现行 EX", "保序 EX", "受影响"],
      _rows, [2.4, 2.4, 1.6, 3.4, 3.4, 1.8],
      ["l", "c", "r", "r", "r", "r"], font=8.4)
NOTE("子集 = gold SQL 含 ORDER BY 的条目；同一条件在不同档位的子集不同，故逐格给出 n，"
     "不与表 7-13 的全档分母混用。受影响 = 现行判对而保序判错的条数。"
     "「跨条件池化」为该档四个条件合并后的口径。")
P("子集内的差异比全档明显：合并档池化 27.7% → 26.4%（少 2 条），"
  "拆分档 61.0% → 55.1%（少 7 条），受影响条目集中在 c2 与 c_rw 这两个基线条件上"
  "（拆分档 c2 从 8 条降到 5 条、c_rw 从 4 条降到 2 条），"
  "而主对照 c3mp 在合并档只少 1 条（32 → 31）。"
  "方向是一致的：保序口径削弱基线、几乎不动主对照，因此表 7-12 的差值不是被宽判撑起来的。"
  "本章据此明确声明：全文的 EX 一律采用不保序的结果多重集比较，"
  "含 ORDER BY 子集的保序准确率已在表 7-108 单列，未作为主指标使用。")

# ================================================================ 7.1.8
H("7.1.8  可审计的事件映射与五种事件数的对账", 3)
P("审稿意见指出材料中出现 7、10、2 三种事件数，却无法从原子操作复算到各层事件数；"
  "本轮核对时又发现同一处缺陷还影响了全章的 58 次（应为 61 次），故一并列入对账。"
  "本节先把四层单位钉死，再逐 op 给出映射明细，最后说明三个数字的来源与取舍。"
  "这不是补充一张附表，而是纠正上一版一处会改变 p 值的聚类键缺陷。")
P("四层单位自下而上为："
  "① 原子 op —— 一条「某旧列 → 某目标列 + 决定性表达式」的最小映射动作，"
  "键为 (演化档, 库, 表, 旧列)；"
  "② 映射事实 —— 一个原子 op 连同其表达式与映射 ID，是抽取质量的判定对象；"
  "③ 变更事件 —— 一次独立的列重整，键为 (演化档, 库, 表, 目标列集合)；"
  "④ 数据库聚类 —— 同一库上的全部变更，是统计独立性的最粗单位。"
  f"包含关系为：op 与映射事实 1:1；合并档每事件含 2 个 op（两个旧列各占一段位）、"
  f"拆分档每事件含 1 个 op；每库含 1–2 个事件。段位档共 "
  f"{EMA['counts']['atomic_ops']} 个 op、{EMA['counts']['events_correct']} 次事件、"
  f"{EMA['counts']['db_clusters']} 个库。")
CAP("表 7-114  五种事件数的定义、来源与处置")
TABLE(["数字", "出现处", "聚合键", "判定"],
      [["7", "上一版 M1 抽取的段位档事件数",
        "(库, 表, 表达式首个 attr_ 列)，不含演化档",
        "作废。合并档的 SUBSTR(attr_001,…) 与拆分档的 attr_001‖sep‖attr_002 首列同名，"
        f"3 个共用同表的库各被并成 1 个事件，{EMA['counts']['events_correct']} 塌缩为 7"],
       [f"{EMA['counts']['events_correct']}", "本版段位档（合并 + 拆分）的独立事件数",
        "(演化档, 库, 表, 目标列集合)",
        f"采用。合并 {EMA['counts']['events_correct_by_evo']['merge_col']} 次 + 拆分 "
        f"{EMA['counts']['events_correct_by_evo']['split_col']} 次；"
        "与 7.8 节 C3mp − c_rw 池化的 10 次同源，两者此刻口径一致"],
       [f"{AD_PS['events_v13_triplet']}", "上一版全章「610 条背后的独立事件数」",
        "同上，但含段位档的缺陷键（改名与依赖档的键不受影响）",
        f"作废。同一缺陷使 610 条全集少算 3 次事件；按修正键重算为 {AD_PS['events']} 次"],
       ["2", "上一版「段位表达式 10/10 覆盖的事件数」",
        "事件内须全部 op 有 part 产物，叠加 7 口径的聚类混淆",
        "作废，属低估。段位表达式只在合并档有定义"
        f"（{EMA['counts']['ops_with_segment_expr']} 个 op / "
        f"{EMA['counts']['events_correct_by_evo']['merge_col']} 次事件）；"
        f"拆分档 {EMA['counts']['ops_concat_expr']} 个 op 的 gold 是拼接表达式而非段位表达式"],
       [f"{AD_D['events_covered']}", "本版「段位表达式 10/10 覆盖的事件数」",
        "同 10 口径，仅取 part 分支有产物的合并事件",
        f"采用。合并档 {EMA['counts']['events_correct_by_evo']['merge_col']} 次事件全部覆盖、"
        f"全部一致，精确符号检验 p = {AD_D['event_level_p']:.4f}"
        "（n = 5 时的最小可达值，故不作显著性声称）"]],
      [1.2, 3.6, 4.2, 7.0], ["c", "l", "l", "l"], font=8.1)
NOTE("五行的排序即对账顺序：先作废两个缺演化档的键（7、58），再给出对应的正确值（10、61），"
     "最后把「段位表达式覆盖事件数」从 2 修正为 5。"
     "四层与五种数字的机器可读明细见 results/event_map_audit.json 与 "
     "results/audit_independence.json（audit-independence-v2）；"
     "下表把它逐行展开，含 op ID、映射 ID、事件 ID（正确口径与上一版口径并列）、"
     "库聚类 ID，读者可据此复算任一层的单元数。")
CAP("表 7-115  段位档 15 个原子 op 的可审计映射明细")
_em_rows = []
for _d in EMA["detail"]:
    _em_rows.append([
        _d["op_id"] + " / " + _d["map_id"],
        "合并" if _d["evo"] == "merge_col" else "拆分",
        _d["db"], _d["table"], _d["old_col"],
        "+".join(_d["target_cols"]),
        "段位" if _d["has_segment_expr"] else "拼接",
        _d["event_id_correct"] + " (" + _d["event_id_v13"] + ")",
    ])
TABLE(["op / 映射 ID", "档", "库", "表", "旧列", "目标列", "表达式", "事件 ID (上版)"],
      _em_rows, [2.0, 0.9, 2.9, 1.5, 1.9, 1.9, 1.2, 2.3],
      ["l", "c", "l", "l", "l", "l", "c", "c"], font=7.5)
NOTE("括号内为上一版三元组口径的事件 ID；同一括号值重复出现即表示被误并的 op。"
     "例如 california_schools/schools 的 3 个 op（合并 2 + 拆分 1）在旧口径下同属一个事件。"
     "「段位」指 gold 为 SUBSTR/INSTR 表达式，「拼接」指 gold 为 attr_001‖sep‖attr_002。")
P(f"按正确口径重算，M1 相对三条基线在段位档的类型识别为 "
  f"{EMA['sign_correct']['ours_only']}:{EMA['sign_correct']['base_only']}、"
  f"n = {EMA['sign_correct']['n']}、精确符号检验 p = {EMA['sign_correct']['p_exact']:.4f}"
  f"（上一版 n = {EMA['sign_v13']['n']}、p = {EMA['sign_v13']['p_exact']:.4f}）。"
  "结论方向不变，但强度仍受限于 5 个库——库级才是外部效度的诚实读数，"
  "这一限定不因事件数修正而放松。")

# ================================================================ 7.2
B2 = json.load(open("results/stats_batch2.json", encoding="utf-8"))["task_level"]
# op 级分段对照：同一 evaluate() 口径，把「目标列恢复」与「op 类型一致」分开。
# 后者在段位档上等价于「段位是否恢复」，题级口径看不到这一层。
ABM = json.load(open("results/ablation_matching.json", encoding="utf-8"))
assert ABM.get("version") == "ablation-matching-v1", "分段对照产物缺失或版本不符"
H("7.2  RQ1：自动谱系与 oracle 谱系的差异", 2)
P("先看抽取环节。表 7-14 按演化档给出谱系抽取在题级的恢复率，并与三个匹配基线对照。"
  "在改名、合并、拆分三档上，本文方法（M1）的题级恢复率为 100%，与 oracle 相同；"
  "差异出现在取值被重编码的两档。")
CAP("表 7-14  谱系抽取的题级恢复率（%）")
_rows = []
for key, label in [("rename", "改名 easy"), ("rename_medium", "改名 medium"),
                   ("rename_hard", "改名 hard"), ("merge_col", "合并"),
                   ("split_col", "拆分"), ("value_shift", "值改写"),
                   ("dep_successor", "依赖后继")]:
    v = B2[key]
    _rows.append([label, str(v["M1"]["n"]),
                  f"{v['M1']['rate']:.1f}", f"{v['B1']['rate']:.1f}",
                  f"{v['B2']['rate']:.1f}", f"{v['B3']['rate']:.1f}"])
TABLE(["演化档", "n", "M1（本文）", "B1 字符串", "B2 词形", "B3 类型/位置"],
      _rows, [4.4, 1.8, 2.8, 2.8, 2.8, 3.0], ["l", "r", "r", "r", "r", "r"])
NOTE("题级恢复率指该题涉及的全部旧列都被正确映射到新表达式，"
     "比列级 macro 恢复率更严格，也与端到端判定口径一致。"
     "值改写（语义名）与依赖弃答两档的抽取结果在第 5 章给出，本章不重复。"
     "B1 在改名 easy 档为 100%，说明该档改名仍保留足够字面信息；到 medium/hard 降为 0，"
     "字面线索已不足以支撑匹配。")
P("在改名三档上，M1、B2、B3 都已接近或达到 100%（B2 在 hard 档为 77.8%），"
  "列级维度对这三者不具区分力，故本章不以「谁的列匹配更准」为论点。"
  "真正拉开差距的是合并与拆分：B2/B3 在合并档分别降至 31.7% 与 40.0%，在拆分档同时为 0。"
  "但仅凭题级数字，读者会把这个差距读成精度差异——它不是。")

P("换一次口径就能看清这一点。题级恢复率把「映射到哪个列」与「这一列由哪几段构成」"
  "合并成了一个数，而前者做到、后者做不到的情形恰恰是合并档的常态。"
  "表 7-113 因此按 op 级重报同一批抽取结果，并把两个维度分列："
  "目标列恢复指预测的列集合与 gold 一致，op 类型一致指类型也对得上。"
  "必须说清这一列能判什么、不能判什么：判定函数先把 gold 的目标表达式归一为被引用的列名，"
  "因此「op 类型一致」度量的是「是否识别出这是一次合并/拆分、且落在哪一列」，"
  "它并不比较段位表达式本身。段位表达式是否被正确恢复，需绕开该判定函数单独度量，"
  "结果见本节末。全部方法走同一个判定函数、同一份 gold。")
CAP("表 7-113  谱系抽取的 op 级质量：段位档与非段位档（同一 gold、同一判定函数）")
_LAB = {"B1 列名相似": "B1 列名", "B2 值域重叠": "B2 值域重叠",
        "B3 嵌入语义": "B3 嵌入", "M1 值域指纹": "M1 值域指纹",
        "M1+形态指纹": "M1+形态指纹"}
_rows = []
for _lab in ["B1 列名相似", "B2 值域重叠", "B3 嵌入语义",
             "M1 值域指纹", "M1+形态指纹"]:
    _row = [_LAB[_lab]]
    for _gname, _key in (("段位档", "recovered"), ("段位档", "type_ok"),
                         ("非段位档", "recovered"), ("非段位档", "type_ok")):
        _s = ABM["groups"][_lab][_gname]
        _pct = _s["recall_pct"] if _key == "recovered" else _s["type_pct"]
        _row.append(f"{_s[_key]}/{_s['n']}（{_pct:.1f}）")
    _rows.append(_row)
TABLE(["方法", "段位档 目标列", "段位档 类型一致", "非段位档 目标列", "非段位档 类型一致"],
      _rows, [3.4, 3.1, 3.1, 3.1, 3.1], ["l", "r", "r", "r", "r"], font=8.5)
_AB = ABM["groups"]
NOTE(f"段位档 = 合并 {_AB['M1+形态指纹']['段位档']['n'] - 5} 个 op + 拆分 5 个 op"
     f"（共 {_AB['M1+形态指纹']['段位档']['n']}）；非段位档为其余六个演化类型"
     f"（{_AB['M1+形态指纹']['非段位档']['n']}）。括号内为百分比。"
     f"独立单元数必须与 op 数同时给出：段位档 {AD_B['n_ops']} 个 op 只对应 "
     f"{AD_B['n_events']} 次独立演化事件（5 个库），因为一次合并会产生两个 op"
     f"（两个旧列各占一个段位），它们由同一条 gold 表达式配对生成、不是两个独立观测。"
     f"因此配对检验报事件级：按 7.1.8 修正后的含档聚类键，M1 对三条基线均为 "
     f"{EMA['sign_correct']['ours_only']}:0，"
     f"精确符号检验 p = {EMA['sign_correct']['p_exact']:.4f}"
     f"（n = {EMA['sign_correct']['n']}）。"
     f"上一版按三元组键报 n = {EMA['sign_v13']['n']}、"
     f"p = {EMA['sign_v13']['p_exact']:.4f}，因聚类键把合并与拆分误并而作废。"
     f"若按 op 级报，p = {AD_B['paired']['B3 嵌入语义']['op_level']['p']:.1e}，"
     f"这是把同一次演化拆成两份样本后得到的值，属于高估，本章不采用。"
     f"非段位档两列是混合口径：{AD_IN['n_ops']} 个 op 来自 610 条主实验集，"
     f"{AD_SH['n_ops']} 个来自集外的值改写两档，两者分母不同、不可相加。"
     f"拆开后，不含值改写的一段 M1+形态与 B3 同为 {AD_IN['ours_ok']}/{AD_IN['n_ops']}（完全持平），"
     f"缺口全部在集外（{AD_SH['ours_ok']}/{AD_SH['n_ops']} 对 "
     f"{AD_SH['base_ok']}/{AD_SH['n_ops']}），正文第三处给出口径与配对检验。"
     "位置先验（P0）消融只在不含值改写的演化上跑过，分母不同，不列入本表，见 7.3 节。")

CAP("表 7-117  四个关键读数各自属于哪个方法变体、度量什么")
TABLE(["读数", "方法变体", "度量对象", "能支持的结论"],
      [[f"{_AB['M1+形态指纹']['段位档']['type_ok']}/{_AB['M1+形态指纹']['段位档']['n']} 对 0/{_AB['M1+形态指纹']['段位档']['n']}",
        "M1（值域指纹）与 M1+形态指纹均如此",
        "操作类型与目标列识别（evaluate 的 type_ok）",
        "能识别出这是一次合并/拆分、落在哪一列；不是段位表达式恢复率"],
       [f"{AD_D['part_expr_match']}/{AD_D['part_covered_ops']}",
        "M1 段位推断（part 分支）",
        "gold 段位表达式的逐字符一致",
        "机制演示；仅覆盖合并档 "
        f"{EMA['counts']['events_correct_by_evo']['merge_col']} 次事件，"
        "且 B4 公平基线同样达到"],
       [f"{AD_IN['ours_ok']}/{AD_IN['n_ops']} 持平",
        "M1+形态指纹（不是 M1 单独）",
        "不含值改写的非段位档列匹配",
        "该分组上与嵌入语义基线等价"],
       [f"{AD_SH['ours_ok']}/{AD_SH['n_ops']} 落后",
        "M1+形态指纹（不是 M1 单独）",
        "取值被整体重编码的两档",
        "形态指纹在重编码下失效，M1+形态劣于语义基线"]],
      [2.5, 3.8, 4.0, 6.0], ["l", "l", "l", "l"], font=8.0)
NOTE("本表回应审稿意见 2.2：M1 与 M1+形态指纹是两个变体，"
     "121/121 与 14/26 只属于后者的非段位分组，不得与 15/15（M1 类型识别）"
     "或 10/10（M1 part 表达式）混写；15/15 亦不等于段位表达式恢复率。"
     "英文摘要与投稿信中的方法变体名称须与本表一致。")

P("这张表要并读三处，其中第三处暴露了一个审稿人必然会问、而上一版答不清楚的问题。")
P("第一处是段位档。B2 与 B3 的目标列恢复各为 5/15，op 类型一致均为 0/15："
  "它们找到了「哪两列并进了哪一列」，但一律报成 rename_column，没有一条给出「第几段」。"
  "M1 与 M1+形态在两项上都是 15/15。这个 15:0 必须配一句限定才能被正确读取："
  "三条基线的输出类型只有 rename_column 与 deprecate_column 两种，"
  "其表示里根本不存在 merge_columns、split_column 与段位表达式，"
  "因此 0/15 是表达能力所限的必然结果，不是匹配精度的差距。"
  "该对照能证明的是「段位谱系不能由更好的列匹配器得到」，"
  "不能证明「M1 的匹配算法比 B3 更准」。"
  "这一能力差别的后果在端到端环节才显现：一个不给出段位的谱系与完全不提供谱系等价，"
  "这正是合并档 C3m 与 C3mp 之间 77.0pt 差距的来源（7.3 节）。")
_d = AD_D
P(f"第二处是段位表达式本身恢复了没有——这是 15/15 那个数字唯一无法回答的问题。"
  f"绕开判定函数、直接与 gold 表达式做字符串比对，可得：合并档 10 个 op 全部一致"
  f"（{_d['part_expr_match']}/{_d['part_covered_ops']}，"
  f"归一化后逐字符相同）；拆分档 {len(_d['uncovered_ops'])} 个 op 未产出段位表达式"
  f"（回退为列级输出），因此段位恢复的直接证据覆盖 "
  f"{_d['events_covered']} 次独立事件（合并档全部 {_d['gold_segment_events']} 次事件中的 "
  f"{_d['events_covered']} 次，拆分档 {_d['gold_segment_events'] - _d['events_covered']} 次无定义）。"
  f"按事件级做符号检验，{_d['event_level_n']} 次事件全部一致，"
  f"p = {_d['event_level_p']:.4f}——这已等于 n = {_d['event_level_n']} 时的最小可达值，"
  "不足以单独支撑「段位可被自动恢复」这一结论。"
  "本章因此不把段位价值的统计证据放在抽取环节，而放在端到端环节："
  "那里有 334 条段位档题目的配对判定，C3m 0.9% 对 C3mp 79.6%（7.3 节）。"
  "换句话说，M1「能恢复段位」这一陈述的证据强度是「10/10 的机制演示 + 端到端效应」，"
  "不是「10/10 本身即为统计显著的抽取能力」——两者的差别必须写明。")
_ab = ABM["groups"]
P(f"第三处是代价，且这一处上一版的归因是错的。把非段位档按「取值是否被重编码」"
  f"两段拆开看：不含值改写的 {AD_IN['n_ops']} 个 op 上（改名三档 75 个、"
  f"依赖弃答与后继两档 46 个），M1+形态为 "
  f"{AD_IN['ours_ok']}/{AD_IN['n_ops']}，B3 为 {AD_IN['base_ok']}/{AD_IN['n_ops']}，"
  f"两者完全持平，配对检验无一对不一致（p = 1）。"
  f"全部缺口集中在 610 条之外的值改写两档：{AD_SH['n_ops']} 个 op 上 M1+形态为 "
  f"{AD_SH['ours_ok']}/{AD_SH['n_ops']}（{100.0*AD_SH['ours_ok']/AD_SH['n_ops']:.1f}%），"
  f"B3 为 {AD_SH['base_ok']}/{AD_SH['n_ops']}，事件级配对 0:{AD_SH['event_level']['base_only']}、"
  f"p = {AD_SH['event_level']['p']:.4f}。"
  f"混合报出的 91.8% 之所以误导，是因为它把 121 个与 26 个分母不同的 op 加在了一起，"
  f"读起来像「M1 的通用匹配能力低 8 个百分点」，实际差的是「值被整体重编码后"
  f"形态指纹失效」这一件事。正确的表述须分档且标出变体：M1 在需要段位的类型上是"
  f"唯一可用的；M1+形态指纹在只要求列对、不含值改写的类型上与语义匹配器持平，"
  f"在取值被重编码的类型上劣于语义匹配器。表 7-117 列出四个读数各自的变体归属。")

P("这里必须回答一个审稿人已经提出的问题：相对于已有一族「按取值内容做列匹配」的方法，"
  "M1 的新意究竟在哪。诚实的分工是这样的。"
  "第一，那一族工作的共同输出是「列对」。schema 匹配的系统综述（Rahm 与 Bernstein，2001）"
  "把匹配结果定义为两个 schema 元素之间的对应关系，"
  "典型方法以名字、类型、约束与实例统计为证据：Cupid（Madhavan 等，2001）"
  "把语言学相似度与结构相似度组合成元素级匹配，"
  "COMA（Do 与 Rahm，2002）进一步把多种匹配器的结果按策略合成。"
  "即便扩展到复杂匹配的 iMAP（Dhamankar 等，2004），输出仍是元素集合之间的对应式子句，"
  "其中不含「这一列由哪几段构成、某段取第几位」这类定位信息。"
  "本章合并场景的输出不是列对，而是「新列 + 段位」组成的可执行表达式，"
  "因为两个旧列会映射到同一个新列——列对被折叠为一对多，在这一场景下信息不足。"
  "这不是精度差异，而是输出形态差异，两个问题因此不等价；"
  "表 7-113 给出了这一差异的量化形态：B2 与 B3 在段位档的目标列恢复各 5/15，"
  "op 类型一致均为 0/15——它们把「列对」这一形态用到了极限，仍无法表达段位。"
  "这一 0/15 的性质须再一次说明：三条基线的输出类型集合只有 rename_column 与 "
  "deprecate_column，段位表达式不在其表示空间内，故该值度量的是表达能力的边界，"
  "第二，那一族工作多数面向跨库匹配（两套彼此独立的 schema），"
  "本章面向同一库的前后版本，可利用「演化算子只作用于少数列」这一先验来收缩搜索空间，"
  "这一先验也是 M1 在合并档能做到零误报的原因。"
  "与本章设定最接近的是 Kang 与 Naughton（2003）：列名不可读，只能靠取值内容判断对应，"
  "与本章「名字仅作 tie-break」的取舍同源；"
  "区别在于该工作判定的仍是列与列的对应，本章在此之上还要判定段位，"
  "即在同一新列内部区分「哪一段来自哪一个旧列」——同一批取值证据被用来回答更细的问题。"
  "第三，B2 与 B3 就是那一族方法在本任务上的最小实现：B2 只用值域重合"
  "（集合 Jaccard 与双向子串包含，不含名字 tie-break、不含近唯一列守卫、不含段位推断），"
  "B3 用嵌入模型匹配「列名 + 值样例」的语义，代表通用语义匹配做法。"
  "本章与它们的差别已经量化，结论是分档的、不是一概而论的："
  "在不含值改写的非段位档上（121 个 op），M1+形态与 B3 同为 100.0%，"
  "配对检验零不一致；在取值被整体重编码的两档（26 个 op）上，"
  "M1+形态为 53.8%、B3 为 100.0%。"
  "因此本章主张的是「M1+形态指纹与通用语义匹配器在不含值改写的列匹配上等价，"
  "并在其失效的值改写档上更差」，而不是「M1 是更好的列匹配器」。"
  "可主张的核心只有一条：把同一批取值与形态证据用于恢复段位，"
  "是该族方法未涉及的问题——三条基线的输出类型空间里没有段位表达式，"
  "在段位档上类型一致均为 0/15，M1 为 15/15；"
  "这是本章与那一族方法的分界点，也是本章唯一站得住的新颖性来源。")

H("7.2.1  同等表达能力下的公平基线（B4）", 3)
P("上面的 15/15 对 0/15 有一个必须正面回答的质疑：三条基线的输出类型里根本没有"
  "段位表达式，那么 0/15 到底是「M1 的算法更好」，还是「基线的输出空间不允许」？"
  "审稿意见第三轮提出这一点，并要求一个具备同等输出能力的非 LLM 基线。"
  "本节补做这个对照，结论对本方法不利，但它决定了主张能站到哪里。")
P("B4 的构造刻意不引入新的算法优势：它复用 M1 的同一套分隔符检测、值域 Jaccard "
  "与 SUBSTR 表达式渲染，因而输出空间与 M1 完全相同；"
  "差别只剩下把 M1 的三处专属设计全部去掉——"
  "① 段 Jaccard 阈值门控（0.55）、② 低基数列抬到 0.95 的守卫、"
  "③ 采样行数门限与冲突消解，改为对每个旧列直接取 Jaccard 最大的段位（纯 argmax）。"
  "两者都不读标注或 gold 表达式，评估口径也与 7.2 的段位表达式比对完全一致："
  "归一化后与 gold 逐字符比较。")
CAP("表 7-116  同等表达能力下的段位恢复对照")
TABLE(["方法", "输出空间", "表达式一致 (op)", "覆盖事件", "专属设计"],
      [["B1 列名相似 / B2 值域重叠 / B3 嵌入语义",
        "仅 rename / deprecate", f"0/{EMA['counts']['atomic_ops']}（不适用）", "—",
        "无段位输出能力"],
       ["B4 段位感知基线（纯 argmax）",
        "含 SUBSTR 段位表达式",
        f"{B4['B4_naive']['expr_match']}/{B4['B4_naive']['n_ops']} "
        f"({B4['B4_naive']['pct']:.1f}%)",
        f"{B4['B4_naive']['events_ok']}/{B4['B4_naive']['n_events']}",
        "无阈值门控、无低基数守卫、无采样门限"],
       ["B4 + 贪心冲突消解（上界变体）",
        "含 SUBSTR 段位表达式",
        f"{B4['B4_greedy']['expr_match']}/{B4['B4_greedy']['n_ops']} "
        f"({B4['B4_greedy']['pct']:.1f}%)", "—", "仅加同段位冲突消解"],
       ["M1 段位推断（part）",
        "含 SUBSTR 段位表达式",
        f"{B4['M1_part']['expr_match']}/{B4['M1_part']['n_ops']} "
        f"({B4['M1_part']['pct']:.1f}%)",
        f"{B4['M1_part']['events_ok']}/{B4['M1_part']['n_events']}",
        "阈值门控 + 低基数守卫 + 采样门限"]],
      [5.0, 3.3, 3.0, 1.9, 4.0], ["l", "l", "c", "c", "l"], font=8.0)
NOTE(f"分母为合并档 {B4['B4_naive']['n_ops']} 个含段位表达式的 op；"
     f"拆分档 {len(B4['split_ops_excluded'])} 个 op 的 gold 是拼接表达式，"
     "非段位 SUBSTR，不进入本对照。B4 与 M1 的检测/渲染实现同源，"
     "故表内比较的是「是否需要 M1 的三处守卫」，不是「能否输出段位」。")
P(f"结果没有留下回旋余地：B4 在 "
  f"{B4['B4_naive']['n_ops']} 个 op 上取得 "
  f"{B4['B4_naive']['expr_match']}/{B4['B4_naive']['n_ops']}，"
  f"覆盖 {B4['B4_naive']['events_ok']}/{B4['B4_naive']['n_events']} 次事件，"
  f"与 M1 的 {B4['M1_part']['expr_match']}/{B4['M1_part']['n_ops']} 完全相同；"
  "加贪心冲突消解的上界变体也不改变读数。"
  "也就是说，一旦把输出能力对齐，M1 的阈值门控、低基数守卫与采样门限"
  "在本构造的段位表达式恢复上没有带来可观测增益。"
  "这一结论必须写进主张里：本章可主张的是「把段位写进谱系这一输出粒度」，"
  "不是「M1 比通用分段算法更聪明」。")
P("这一发现同时划定了两条边界，两条都对本方法不利，但都必须保留。"
  "第一，它揭示了 15/15 对 0/15 的性质：那个对照证明的是"
  "「段位谱系不能由只能输出列对的匹配器得到」，属于表达能力差异，"
  "而不是匹配精度或抽取算法的优越性；把 B4 补上后，后一种读法已被数据排除。"
  "第二，B4 与 M1 都只在合并档 10 个 op 上评估，且都达到天花板，"
  "这本身是构造同义反复的又一表现：两段可逆拼接的分段位置可由值域重合直接读出，"
  "因此 10/10 不能当作「段位抽取已获验证」的独立证据。"
  "段位价值的统计证据仍在端到端环节（7.3 节 C3m 0.9% 对 C3mp 79.6%），"
  "而不在抽取环节——这与 7.2 第二处解读的结论一致，且现在有了公平基线的支持。")
NOTE("本节回应审稿意见 2.3 与最终意见第 (3) 条。B4 的实现见 baseline_segment_aware.py，"
     "逐 op 的预测表达式、gold、Jaccard 得分与是否一致见 results/baseline_segment_aware.json。"
     "实现中曾出现一处会把 6 个 op 误判为「未产出段位」的大小写归一化 bug（库内列名为 "
     "County/A2/SEX，gold 为小写），修正后 B4 由 4/10 变为 10/10；"
     "该修正方向对本方法不利，故保留在正文，不作为实现细节略去。")

NOTE("本节所引为该方向的代表性原始文献五篇：Rahm 与 Bernstein（VLDB Journal, 2001）、"
     "Madhavan 等（VLDB, 2001）、Do 与 Rahm（VLDB, 2002）、"
     "Kang 与 Naughton（SIGMOD, 2003）、Dhamankar 等（SIGMOD, 2004）。"
     "此处只标注作者、年份与出处，卷期页码须按目标期刊格式由原文补全，"
     "并逐条核对三处论断：① 输出是否为列对或元素对应；② 是否面向同库前后版本；"
     "③ 是否含段位定位信息。任一条与原文献不符，须改写对应句子，不得保留。")
P("图景在端到端层面更清楚。表 7-15 比较 C3m（自动）与 C3f（oracle）的逐题判定："
  "三个基座的不一致分别为 0、1、2 条，且方向双向。")
CAP("表 7-15  自动谱系与 oracle 谱系的逐题一致性（改名档，n = 276）")
_eq = R3["equiv_by_model"]
_rows = []
for m, label in [("DeepSeek-V3.2", "DeepSeek-V3.2"),
                 ("DeepSeek-V4.1-Flash", "DeepSeek-V4.1-Flash"),
                 ("Qwen3-Coder-30B-A3B", "Qwen3-Coder-30B-A3B")]:
    d = _eq[m]
    _rows.append([label, f"{d['acc_c3m']:.1f}", f"{d['acc_c3f']:.1f}",
                  str(d["n_disagree"]), f"{d['c3m_only']} : {d['c3f_only']}"])
TABLE(["基座", "C3m（自动）", "C3f（oracle）", "不一致题数", "仅 C3m 对 : 仅 C3f 对"],
      _rows, [4.6, 3.0, 3.0, 2.6, 3.8], ["l", "r", "r", "r", "c"])
NOTE("Qwen3-30B-A3B-Instruct 一档的 C3m / C3f 分别为 78.6% / 77.9%（见表 7-20），"
     "未保存逐题等价记录，故不在本表列出。合并档上 V3.2 为 4/230 不一致"
     "（2 条仅 C3mp 对、2 条仅 C3f 对），相互抵消后两者同为 79.6%。")
P("RQ1 的答案是：在同粒度前提下，自动抽取与 oracle 在端到端判定上的差异为 0–2 条/276，"
  "在本评测集上未观察到可测差异。这一结论不等于「自动谱系可以普遍替代 oracle」——"
  "一是差异虽小但方向不完全一致（V4.1F 上 oracle 独对 1 条，Qwen3-Coder 上双向各 1 条），"
  "二是该结论只在改名档做过逐题比对，合并档仅有一个基座，"
  "三是抽取在取值重编码与近唯一列上系统性失效（7.7 节）。"
  "因此后文凡涉及 C3mp 的结论均以「可自动获得」为前提，"
  "但不把「等价」作为一般性主张。")

# ================================================================ 7.3
H("7.3  RQ2：谱系是否改善端到端迁移", 2)
P("表 7-16 是本章的主结果，同时给出正确数/分母与 Wilson 区间。"
  "四个档位的形态差异极大，不能合并成一个平均数报告。")
CAP("表 7-16  端到端迁移正确率（DeepSeek-V3.2，端到端层 610 条）")
_rows = []
for gname, label in [("easy", "改名 easy"), ("medium", "改名 medium"),
                     ("hard", "改名 hard"), ("rename_all", "改名 全档"),
                     ("merge", "合并"), ("split", "拆分")]:
    _rows.append([label, str(main[gname]["n"]), cell(gname, "c1"), cell(gname, "c2"),
                  cell(gname, "c3m"), cell(gname, "c3mp"), cell(gname, "c_rw")])
TABLE(["档位", "n", "C1", "C2", "C3m", "C3mp", "c_rw"], _rows,
      [3.2, 1.2, 2.5, 2.5, 2.5, 2.5, 2.5],
      ["l", "r", "r", "r", "r", "r", "r"], font=8.3)
NOTE("单元格为「正确率 (k/n)」，单位 %；对应的 95% Wilson 区间见表 7-17。"
     "C3mp 只在合并/拆分档有意义（其余档无跨列切分）；「改名 全档」为 276 条合并统计，"
     "与三档分行不可互换——池化率会因档位题量不同而与逐档率产生差异。")
P("表 7-17 给出各条件正确率的 Wilson 区间，用于判断哪些差异量级在抽样波动之外。")
CAP("表 7-17  各条件正确率的 95% Wilson 区间（%）")
_rows = []
for gname, label in [("rename_all", "改名 全档"), ("merge", "合并"), ("split", "拆分")]:
    for c, cl in [("c2", "C2"), ("c3m", "C3m"), ("c3mp", "C3mp"), ("c_rw", "c_rw")]:
        d = main[gname].get(c)
        if not d:
            continue
        _rows.append([label, cl, f"{d['k']}/{d['n']}",
                      f"{d['rate']:.1f}", f"[{d['ci'][0]:.1f}, {d['ci'][1]:.1f}]"])
TABLE(["档位", "条件", "k/n", "正确率", "95% CI"], _rows,
      [3.4, 2.4, 2.8, 2.6, 4.4], ["l", "l", "r", "r", "c"])
NOTE("Wilson 区间假设各条目独立同分布。由于合并档 230 条只来自 10 个映射事实，"
     "该区间会低估真实不确定性；簇级口径见 7.7 节与表 7-23。")
_r2 = pr('rename_all', 'c3m', 'c2')
_m2 = pr('merge', 'c3mp', 'c2')
_s2 = pr('split', 'c3m', 'c2')
P(f"在改名档，谱系把正确率从 {main['rename_all']['c2']['rate']:.1f}% 提到 "
  f"{main['rename_all']['c3m']['rate']:.1f}%"
  f"（{sg(_r2['delta'])}pt，{_r2['a_only']}:{_r2['b_only']}，{pc(_r2['p'])}）；"
  f"合并档段位谱系相对旧 SQL 为 {main['merge']['c2']['rate']:.1f}% → "
  f"{main['merge']['c3mp']['rate']:.1f}%"
  f"（{sg(_m2['delta'])}pt，{_m2['a_only']}:{_m2['b_only']}，{pc(_m2['p'])}）；"
  f"拆分档为 {main['split']['c2']['rate']:.1f}% → {main['split']['c3m']['rate']:.1f}%"
  f"（{sg(_s2['delta'])}pt，{_s2['a_only']}:{_s2['b_only']}，{pc(_s2['p'])}）。")
P(f"最值得注意的不是提升幅度，而是列级条件在合并档上低于旧 SQL：C3m 为 "
  f"{main['merge']['c3m']['rate']:.1f}%，C2 为 {main['merge']['c2']['rate']:.1f}%。"
  "原因是可执行性被破坏——旧 SQL 引用了一个已不存在的列名，模型至少会尝试改写；"
  "而列级谱系给了它一个确实存在的列名，它会直接按原样查询，"
  "于是产出一个语法正确、语义错误的查询（表 7-22 显示该条件 97.0% 为 SILENT）。"
  "这说明「列对齐」不足以保证迁移，粒度信息必须单独给出。")

H("7.3.1  逐库稳健性与留一库检验", 3)
P("合并与拆分各只覆盖 5 个数据库，样本簇偏少，因此必须逐库摊开并做留一库检验。")
CAP("表 7-18  主效应的逐库差值与留一库检验")
TABLE(["对比", "n", "库数", "逐库差值范围（pt）", "差值>0 的库", "留一库点差区间"],
      [["合并：C3mp − C2", "230", "5", "+44.4 ~ +95.9", "5/5", "[+62.4, +77.3]"],
       ["拆分：C3m − C2", "104", "5", "+38.5 ~ +100.0", "5/5", "[+54.7, +65.5]"],
       ["改名：C3m − C2", "276", "11", "+0.0 ~ +31.5", "9/11（2 库为 0）", "[+8.4, +16.4]"]],
      [4.4, 1.4, 1.6, 4.6, 3.4, 3.6], ["l", "r", "r", "c", "c", "c"])
NOTE("逐库差值为该库内条目级正确率之差；留一库点差区间为依次剔除每个库后的点估计范围"
     "（不含置信区间）。改名档有 2 个库差值为 0、无库为负。"
     "5 个簇时库级 bootstrap 区间本身偏窄，故本节只把「符号稳定」作为观察，"
     "不据此声称跨库泛化。")

# ================================================================ 7.4
H("7.4  RQ3 与 RQ5：增益来自证据表示，还是来自执行器", 2)
P("本节把这两个问题拆开检验。第一，与非谱系的直接改写相比，"
  "显式谱系是否有价值（7.4.1）。第二，在同一份谱系下，"
  "「确定性替换」与「LLM 生成」两种执行器相差多少（7.4.2）；"
  "把两个因子分开，才能回答增益究竟来自谱系表达式本身，还是来自把表达式写进提示让模型模仿（7.4.3）。")
H("7.4.1  对照一：直给原始 schema 的改写（c_rw）", 3)
P("C 组与 B 组的对照可能被质疑：B 组没有拿到旧 SQL 的语义对应关系，输给谱系理所当然。"
  "c_rw 条件为此而设——它同时提供旧 DDL、新 DDL、旧 SQL 与问题，"
  "把「是否需要显式谱系」变成一个经验问题。表 7-19 给出逐档结果。")
CAP("表 7-19  显式谱系与直接改写（c_rw）的配对比较")
_rows = []
for gname, label, cond in [("easy", "改名 easy", "C3m"), ("medium", "改名 medium", "C3m"),
                           ("hard", "改名 hard", "C3m"), ("merge", "合并", "C3mp"),
                           ("split", "拆分", "C3mp")]:
    p_ = pr(gname, "c3mp" if gname in ("merge", "split") else "c3m", "c_rw")
    _rows.append([f"{label} · {cond}", str(p_["n_pair"]),
                  f"{p_['rate_b']:.1f}", f"{p_['rate_a']:.1f}",
                  f"{p_['a_only']} : {p_['b_only']}",
                  str(p_["n_discordant"]),
                  f"{p_['delta']:+.1f}",
                  f"[{p_['ci'][0]:+.1f}, {p_['ci'][1]:+.1f}]",
                  pp(p_["p"]), pp(p_.get("p_holm"))])
TABLE(["档位（谱系条件）", "配对数", "c_rw %", "谱系 %", "仅谱系对 : 仅改写对",
       "不一致对", "Δ (pt)", "95% CI（Newcombe）", "p", "Holm p"],
      _rows, [2.9, 1.3, 1.2, 1.2, 2.5, 1.3, 1.2, 2.5, 1.2, 1.3],
      ["l", "r", "r", "r", "c", "r", "r", "c", "r", "r"], font=8.0)
NOTE("配对数 = 两条件都有输出的条目数；本章全部档位配对完整，边际分母与配对分母一致。"
     "「仅谱系对 : 仅改写对」为不一致对的方向拆分，「不一致对」为其合计。"
     "区间为 Newcombe (1998) 方法 10 的 95% 区间（逐比较，未做族校正）；"
     "Holm 校正只作用于 p，族的成员固定为本表五个 c_rw 比较。"
     "「改名 全档」与三个难度分档不是独立比较，仅作描述，不进族也不校正。")
_e0 = pr('easy', 'c3m', 'c_rw')
_W_BOOT, _W_V3 = 7.3, 5.1
_ci_now = f"[{_e0['ci'][0]:+.1f}, {_e0['ci'][1]:+.1f}]".replace("-", "−")
_lb_now = f"{_e0['ci'][0]:+.1f}".replace("-", "−")
_w_now = f"{_e0['ci'][1] - _e0['ci'][0]:.1f}"
P(f"本轮修订同时更正了 easy 档的不一致对：修订稿写「0 : 3」、p = 0.250，"
  f"与 91/87 的边际计数不相容；逐题复核后 {_e0['n_pair']} 条全部配对，"
  f"不一致对为 {_e0['a_only']} : {_e0['b_only']}，净差 {_e0['delta']:+.1f}pt，"
  f"精确 McNemar p = {_e0['p']:.3f}。配对区间在修订过程中更换过四代实现，"
  f"正文只报告最终实现（Newcombe 1998 方法 10）的 {_ci_now}；"
  f"四代实现的点估计、宽度与覆盖率检验对照见补充材料 S2——"
  f"它们的共同结论都是区间含 0，即 easy 档的方向无法确定，"
  f"而实现符合方法 10 定义的只有最后一代。")
_e = pr('easy', 'c3m', 'c_rw')
_m = pr('medium', 'c3m', 'c_rw')
_h = pr('hard', 'c3m', 'c_rw')
_te = {t['margin_pt']: t for t in _e['tost']}
_tm = {t['margin_pt']: t for t in _m['tost']}
P(f"改名 easy 档，c_rw {_e['rate_b']:.1f}% 对谱系 {_e['rate_a']:.1f}%"
  f"（{_e['a_only']} : {_e['b_only']}，{pc(_e['p'])}，Holm 校正后 {pp(_e.get('p_holm'))}）。"
  f"Newcombe 区间 [{_e['ci'][0]:+.1f}, {_e['ci'][1]:+.1f}] 包含 0，与检验结论一致，"
  "本评测集不足以判定两者存在差异。等价同样不成立："
  f"其 90% 区间 [{_e['ci90_newcombe'][0]:+.1f}, {_e['ci90_newcombe'][1]:+.1f}] "
  "只在 ±10pt 界值下才整体落入，±3 与 ±5 均越界。"
  "因此这一档的结论是「未判定」——既不主张等价，也不主张谱系更优。")
P(f"medium 档 c_rw {_m['rate_b']:.1f}% 对谱系 {_m['rate_a']:.1f}%"
  f"（{_m['a_only']} : {_m['b_only']}，{pc(_m['p'])}），"
  f"Newcombe 区间 [{_m['ci'][0]:+.1f}, {_m['ci'][1]:+.1f}]。"
  "等价检验的两个口径在此冲突：Wald 版 TOST 在 ±5pt 下 "
  f"p = {_tm[5]['p_tost']:.3f}（通过），而 Newcombe 90% 区间 "
  f"[{_m['ci90_newcombe'][0]:+.1f}, {_m['ci90_newcombe'][1]:+.1f}] 的下界越出 ±5pt（不通过）。"
  "按 7.1.5 节确立的规则，两法冲突时以区间法为准，因此等价表述撤销。"
  "该档与 easy 档同列为「未判定」：点估计虽接近 0，"
  "但样本量不足以把差异限定在任何合理的界值以内。"
  "v3 曾据 Wald 版 TOST 写作「在 ±5pt 界值下未发现超出该界值的差异」，本版撤回该表述。")
P(f"hard 档谱系为 {_h['rate_a']:.1f}%、c_rw 为 {_h['rate_b']:.1f}%"
  f"（{_h['a_only']} : {_h['b_only']}），Newcombe 区间 "
  f"[{_h['ci'][0]:+.1f}, {_h['ci'][1]:+.1f}]，未校正 {pc(_h['p'])}，"
  f"但 Holm 校正后为 p = {_h['p_holm']:.3f}，不再低于 0.05。"
  "也就是说，hard 档「谱系优于直接改写」这一判断在多重比较校正后不成立，"
  "只能报告点估计为正、区间不含 0（逐比较 95%，未做族校正）、校正后未达显著。"
  "三档合起来看，改名场景下的结论不是一个方向，而是三种不同强度的判断，"
  "把它们合并成「谱系在改名上同样有效」会掩盖这一点。")
P("一个反向的细节值得记录：合并档上列级谱系 C3m（2.6%）不如 c_rw（12.2%）。"
  "这与 7.3 节的解释一致——错误的粒度比不给信息更糟。")
P("关于 c_rw 落败是否源于实现或配置问题，可以逐条核查其失败输出："
  "全章 c_rw 共 98 条 ERROR，其中 97 条（99.0%）的 SQL 里仍然写着已删除的旧列名，"
  "属于「模型直接复用旧 SQL、未做任何改名」的真实失败模式，"
  "而非接口调用失败（调用失败单列为 LLM_ERROR 且不进入任何条件的分母）。"
  "结合表 7-10 显示的输入长度（c_rw 的提示比 C3m 长 26%–44%），"
  "c_rw 在合并与拆分档的落败不能归因于信息量不足或实现缺陷。")
P("「不是信息量不足」不等于「可以推广到一切直接改写方案」，这层区分必须写明。"
  "c_rw 与谱系条件比较的是信息组织方式，而非两份完全相同的提示：c_rw 同时列出旧 DDL、"
  "新 DDL 与旧 SQL，提示的结构、字段顺序与注意力分配都与 C 组不同。"
  "本章因此只在「本提示模板下」主张显式谱系优于直给原始 schema 的改写条件，"
  "不把它推广为「任何直接改写方法都不如谱系注入」。"
  "把这层差异分离出去，需要长度与结构都对齐的对照（例如把旧 DDL 压缩到与谱系行等长，"
  "或让 c_rw 采用与 C 组相同的输出约束）；本章未做，列为待检验项。")
P("长度本身不太可能是主因，这一点有同批数据支持：C3mp 相对 C3m 只多出段位行，"
  "提示长度增加约 1%–2%（表 7-10），合并档准确率却相差 77pt；"
  "而 c_rw 的提示长 26%–44%，两个档位分别低 67 与 68pt。"
  "若长度是主导因素，这两组关系应该反过来。"
  "同一逻辑也解释了 c3mp_nat 消融：信息等价但表达方式不同，长度相近而结果差 15pt，"
  "指向的是证据的可执行性，而非篇幅。")
P("7.4.1 的结论必须按档位分层：合并与拆分档，显式段位谱系大幅优于直接改写，"
  "该差异在 Holm 校正后仍成立；改名三档（easy、medium、hard）当前均不能给出方向性的幅度结论"
  "（easy 的负向在变更单元上一致，但幅度不可判定，见 7.8），"
  "但三者强度并不相同——easy 与 medium 为「未判定」（区间跨 0，且等价不成立），"
  "hard 为「逐比较显著、Holm 校正后不显著」。"
  "据此，本文不主张「名字部分保留时不必自建谱系」，也不主张两者可相互替代，"
  "只报告上述分档观察，并把「何时需要显式谱系」列为待检验的问题。")

# ---------------------------------------------------------------- 7.4.2
H("7.4.2  对照二：同一份谱系下的确定性 SQL 改写", 3)
P("c_rw 回答的是「要不要给谱系」，但它仍让模型自己去执行。审稿人指出一个更尖锐的对照："
  "如果旧列到新表达式的映射已经知道，那么把旧 SQL 里的列标识符按映射直接替换，"
  "理论上就能完成迁移——那样性能提升就来自谱系表达式本身，而不是「把表达式写进提示让模型模仿」。"
  "本节实现该基线，并在同一批条目上与 C2、C3m、C3mp 逐题配对。")
P("实现（rule_rewrite.py）分三步。第一步，用 sqlglot 把旧 SQL 解析成 AST，"
  "遍历全部列引用节点：带表限定符的按 (表, 列) 查映射；不带限定符的，"
  "按「映射里哪些表有这个旧列」定归属，唯一命中才替换，多表命中记为歧义。"
  "第二步，把命中的列节点替换为映射表达式的括号形式（如 SUBSTR(attr_001,1,INSTR(attr_001,CHAR(31))−1)），"
  "重新生成为 SQLite 方言执行。第三步，在演化后库上执行，"
  "与 gold（新 SQL 在同一库上的结果）按不保序多重集比较——与主实验同一判定函数。"
  "星号的处理是本轮修正的重点，细节见 7.4.4：投影位上的 * 按旧表列清单展开后逐列替换，"
  "而 COUNT(*) 不是投影星号，替换过程不触碰它。")
P("三个规则条件的差别只在「映射从哪来」，执行器完全相同："
  "c3r_gt 用构造脚本的 drift_ops（oracle 映射，等价于 C3f 的信息水平）；"
  "c3r_infer 用 M1 自动抽取的段位谱系（不读标注，等价于 C3mp 的信息水平）；"
  "c3r_name 用列级谱系（不含量信息，等价于 C3m 的信息水平）。"
  "三者与 c2/c3m/c3mp 构成「信息水平 × 执行器」的 2×3 网格。")
CAP("表 7-101  规则基线与 LLM 条件在各档的执行结果一致率（%）")
_rows = []
for gk, gl, evo in (("merge", "合并", "merge_col"), ("split", "拆分", "split_col"),
                    ("merge_holdout", "合并·留出", "merge_col_holdout"),
                    ("split_holdout", "拆分·留出", "split_col_holdout")):
    g_ = RVRG[gk]["rates"]
    n_ = g_["c3r_gt"]["n"]
    _rows.append([f"{gl}（{n_}）",
                  f"{g_['c3r_gt']['rate']:.1f}", f"{g_['c3r_infer']['rate']:.1f}",
                  f"{g_['c3r_name']['rate']:.1f}",
                  f"{g_['c2']['rate']:.1f}", f"{g_['c3m']['rate']:.1f}",
                  f"{g_['c3mp']['rate']:.1f}"])
TABLE(["档位（条目数）", "规则·oracle 映射", "规则·自动谱系", "规则·列级谱系",
       "LLM·C2", "LLM·C3m", "LLM·C3mp"],
      _rows, [3.6, 2.2, 2.2, 2.2, 1.5, 1.5, 1.6],
      ["l", "r", "r", "r", "r", "r", "r"], font=8.5)
NOTE("前三级为确定性改写（同一执行器，仅换映射来源），后三级为 LLM 生成。"
     "分母为该档全部条目，失败类型一律计入、不剔除；两种星号模式（展开与保留）"
     "在本样本上给出完全相同的数值，因为投影星号出现 0 次（7.4.4）。")
_rm = RVRG["merge"]["rates"]
_rs = RVRG["split"]["rates"]
_rhm = RVRG["merge_holdout"]["rates"]
P(f"先看合并档。规则·自动谱系 {_rm['c3r_infer']['rate']:.1f}%，"
  f"LLM·C3mp {_rm['c3mp']['rate']:.1f}%，同信息水平下确定性改写高出 "
  f"{_rm['c3r_infer']['rate'] - _rm['c3mp']['rate']:.1f}pt；"
  f"而列级谱系只有 {_rm['c3r_name']['rate']:.1f}%，与 C3m 的 {_rm['c3m']['rate']:.1f}% 同样失败。"
  "也就是说，在这批数据上「给表达式」与「给列名」之间的差距是量级性的，"
  "「确定性替换」与「LLM 生成」之间的差距是二十个百分点左右。"
  "需要强调，规则一侧的两个满分都不是方法性能：c3r_gt 用的是构造映射，"
  "c3r_infer 在 dev 上与 oracle 逐条相同（表 7-102 的 0 : 0），"
  "说明 M1 在这 6 个库上已把映射抽全——这是样本内结果，不是泛化读数。")
P(f"拆分档呈现另一种形态。三种映射水平上规则基线都是 {_rs['c3r_infer']['rate']:.1f}%，"
  "彼此逐条相同——因为拆分方向的重建就是「把各分量按分隔符拼回去」，"
  "列级谱系给出的表达式已经是这个拼接，不再需要额外的段位信息。"
  "这与拆分档 C3mp 与 C3m 逐条相同（7.3 节）是同一件事的两种测法，互相印证。")
_fhm = RVR["failure_profile"]["merge_col_holdout"]["c3r_infer"]
_p_hm = RVRG["merge_holdout"]["paired"]["c3r_gt-c3r_infer"]
P(f"留出库上出现一个必须单独说明的缺口。合并档规则·oracle 映射 {_rhm['c3r_gt']['rate']:.1f}%，"
  f"规则·自动谱系 {_rhm['c3r_infer']['rate']:.1f}%，差额 "
  f"{_rhm['c3r_gt']['rate'] - _rhm['c3r_infer']['rate']:.1f}pt，"
  f"方向是单向的（{_p_hm['a_only']} : {_p_hm['b_only']}，p = {_fp(_p_hm['p'])}）。"
  f"这 {_fhm.get('DB_ERROR', 0)} 条全部是执行报错，不是改写器的逻辑问题："
  "M1 在未参与开发的库上漏抽了映射，替换后语句仍引用已被删除的旧列名，执行直接失败。"
  "这一条把「自动发现」的真实代价第一次量了出来：在 dev 上 M1 与 oracle 逐条等价，"
  f"在未见库上损失 {_rhm['c3r_gt']['rate'] - _rhm['c3r_infer']['rate']:.1f}pt，"
  "且损失方式是硬失败（报错）而非静默错。")
CAP("表 7-102  规则基线与 LLM 的逐题配对（精确 McNemar）")
_rows = []
for gk, gl, pairs in (
        ("merge", "合并", [("c3r_name", "c3m", "列级谱系 vs C3m"),
                           ("c3r_infer", "c3m", "段位谱系 vs C3m"),
                           ("c3r_infer", "c3mp", "段位谱系 vs C3mp"),
                           ("c3r_infer", "c2", "段位谱系 vs C2")]),
        ("split", "拆分", [("c3r_infer", "c3m", "段位谱系 vs C3m"),
                           ("c3r_infer", "c3mp", "段位谱系 vs C3mp"),
                           ("c3r_infer", "c2", "段位谱系 vs C2")]),
        ("merge_holdout", "合并·留出", [("c3r_gt", "c3r_infer", "oracle vs 自动谱系"),
                                       ("c3r_infer", "c3mp", "段位谱系 vs C3mp"),
                                       ("c3r_infer", "c3m", "段位谱系 vs C3m")]),
        ("split_holdout", "拆分·留出", [("c3r_infer", "c3mp", "段位谱系 vs C3mp"),
                                       ("c3r_infer", "c2", "段位谱系 vs C2")])):
    for a_, b_, lab in pairs:
        p_ = RVRG[gk]["paired"]["%s-%s" % (a_, b_)]
        _rows.append([gl, lab, f"{p_['a_rate']:.1f}", f"{p_['b_rate']:.1f}",
                      f"{p_['a_only']} : {p_['b_only']}",
                      ("%+.1f" % p_["delta_pt"]).replace("-", "−"), pp(p_["p"])])
TABLE(["档位", "比较（规则 vs LLM）", "规则 %", "LLM %", "仅规则对 : 仅 LLM 对",
       "Δ (pt)", "p"],
      _rows, [2.4, 4.2, 1.5, 1.5, 2.6, 1.4, 1.5],
      ["l", "l", "r", "r", "c", "r", "r"], font=8.5)
NOTE("配对分母为两条件都有判定的条目数；规则侧的失败类型一律计为失败而非缺失，故分母等于该档条目数。"
     "「oracle vs 自动谱系」一行不使用 LLM，只衡量 M1 抽取误差造成的损失。")
P("必须先把这条基线的循环性说清楚，否则它会被读成自证。本章的 new_sql 本身就是"
  "构造脚本用同一条替换规则从 old_sql 机械生成的，而 c3r_gt 用的正是构造脚本的 drift_ops。"
  "因此 c3r_gt 接近于满分是构造决定的，它只能读作「本构造可机械逆转」的天花板检验，"
  "不能读作本方法的性能。术语上，本章后文一律把它称为 construction-internal upper bound"
  "（构造内上界），并把这一组对照定位为 controlled representation ablation"
  "（固定执行器、只改变谱系表示），而不是方法优劣意义上的 fixed-executor comparison——"
  "两者读法的差别，决定了 100% 是「构造的可逆性」还是「方法的准确率」。"
  "有信息量的比较只有两条："
  "c3r_infer（M1 独立抽取、不读标注）与 c3mp（LLM）在信息水平相同、执行器不同；"
  "以及 c3r_gt 与 c3r_infer 之间的差额，即自动发现的代价。"
  "由此得到的结论也因此必须收窄：它不证明「确定性改写普遍优于 LLM」，"
  "只证明在本构造下，瓶颈不在执行器而在谱系的发现与表示。"
  "要把它变成可推广的主张，需要旧 SQL 与新 SQL 并非机械对应的真实演化数据——本章没有。")
P("这段结论的作用域还可以再精确一层，因为它决定「规则高于 LLM」这一读数能被外推到多远。"
  "按「迁移需要改动什么」可以把演化分成三档："
  "① 只改列引用——旧列被改名、合并或拆分，谓词结构、聚合与连接均不变，"
  "new_sql 由替换列节点得到；② 还要改谓词字面值——列被合并后取值被重编码，"
  "原有的等值条件在新库上不再命中，须同步改写常量；"
  "③ 还要改查询结构——谓词须从「某列等于某值」改为对段位表达式的条件，"
  "或需要新增连接、改变聚合粒度。"
  "规则基线只覆盖第 ① 档：c3r_gt、c3r_infer、c3r_name 三个条件与 LLM 六条件的配对，"
  "全部在合并、拆分及其留出档上完成，报告的全部失败类型也属于这一档。"
  "第 ② 档在本章构造中存在（值改写档），但规则基线未在该档评测——"
  "该档新库中取值已被整体重编码，旧 SQL 的谓词常量不再匹配，"
  "只替换列引用的改写器会命中零行，与 LLM 的对照将退化为「是否改谓词」的单点差异。"
  "第 ③ 档本章完全未涉及。"
  "因此表 7-101 至表 7-103 的全部读数，作用域限定为"
  "「旧 SQL 谓词结构不变、只需替换列引用」这一类迁移；"
  "在这一域内，瓶颈在谱系的发现与表示，而不在执行器。"
  "在这一域外，确定性改写不再具有结构性优势，本章既不主张规则更优，也未测其表现。"
  "这一限定无法靠扩大样本量解除，因为缺的不是样本而是演化类型；"
  "要检验它必须引入会产生谓词级语义变化的真实演化数据，列为后续工作。")

# ---------------------------------------------------------------- 7.4.3
H("7.4.3  增益归因：证据粒度与执行器", 3)
P("把 7.4.2 的 2×3 网格沿两个方向各走一步，就得到两个因子的单独贡献："
  "沿「映射来源」方向走（列级 → 段位），执行器固定为确定性替换，得到证据效应；"
  "沿「执行器」方向走（确定性替换 → LLM），映射固定为段位谱系，得到执行器效应。")
CAP("表 7-103  两个因子的分解（百分点）")
_rows = []
for gk, gl in (("merge", "合并"), ("split", "拆分"),
               ("merge_holdout", "合并·留出"), ("split_holdout", "拆分·留出")):
    d_ = RVR["decomposition"][gk]
    _rows.append([gl,
                  f"{RVRG[gk]['rates']['c3r_name']['rate']:.1f}",
                  f"{RVRG[gk]['rates']['c3r_infer']['rate']:.1f}",
                  ("%+.1f" % d_["evidence_effect_pt"]).replace("-", "−"),
                  f"{RVRG[gk]['rates']['c3mp']['rate']:.1f}",
                  ("%+.1f" % d_["executor_effect_pt"]).replace("-", "−")])
TABLE(["档位", "规则·列级 %", "规则·段位 %", "证据效应", "LLM·C3mp %", "执行器效应"],
      _rows, [2.6, 2.2, 2.2, 2.2, 2.2, 2.4],
      ["l", "r", "r", "r", "r", "r"], font=8.5)
NOTE("证据效应 = 规则·段位 − 规则·列级（执行器固定，只换映射信息）；"
     "执行器效应 = 规则·段位 − LLM·C3mp（映射固定，只换取数方式）。正值表示前者更优。")
_dm = RVR["decomposition"]["merge"]
_ds = RVR["decomposition"]["split"]
_dhm = RVR["decomposition"]["merge_holdout"]
_dsh = RVR["decomposition"]["split_holdout"]
P(f"合并档的证据效应是 {_dm['evidence_effect_pt']:+.1f}pt，执行器效应是 "
  f"{_dm['executor_effect_pt']:+.1f}pt，前者约为后者的 "
  f"{abs(_dm['evidence_effect_pt'] / _dm['executor_effect_pt']):.1f} 倍。"
  "这一分解直接回答了本节开头的问题："
  "在本构造上，性能提升主要来自「把谱系表示成可执行的段位表达式」，"
  "而不是来自「让 LLM 去写」。LLM 在同一份谱系上反而低了二十个百分点，"
  "其失败几乎全是静默错误——SQL 能跑、列也引用对了，但取错了段（7.6 节）。")
P(f"拆分档给出一个反向的对照：证据效应恰为 {_ds['evidence_effect_pt']:+.0f}pt，"
  f"执行器效应 {_ds['executor_effect_pt']:+.1f}pt。"
  "这说明「段位表达式」不是普遍更优的表示，只在合并方向承载了额外信息；"
  "在没有段位信息的档位上，证据效应为零，剩下的差距全部来自执行器。"
  f"留出库把这一点又推进一步：合并档证据效应降到 {_dhm['evidence_effect_pt']:+.1f}pt，"
  f"执行器效应降到 {_dhm['executor_effect_pt']:+.1f}pt。"
  "两个因子都不是常数——证据效应随库与映射质量变化（+99.1 / +64.4 / 0 / 0），"
  "执行器效应也随谱系质量变化（自动谱系一旦漏抽映射，规则侧先丢 19.8pt，"
  "两者差距反而缩小）。因此本章能主张的只是「在合并方向、本构造下，"
  "段位表示相对列名表示的差距比执行器选择大数倍」，"
  "不主张「段位表示普遍带来量级改善」，也不主张两个系数可外推到其他构造。")
P("必须明确这三个限定，否则上述分解会被读成因果结论。"
  "第一，gold 的 new_sql 由构造脚本用同一条替换规则机械生成，"
  "故确定性执行器的上界被构造本身抬高；换到旧新 SQL 并非机械对应的真实演化数据上，"
  "该上界会下降，两个系数都会变。"
  "第二，LLM 一侧只测了一个基座与一套提示模板，执行器效应应读作"
  "「该模型 × 该提示模板在本构造下的配对差值」，不是「LLM 与规则的总体差距」。"
  "第三，两个效应来自同一样本内两次配对，样本又以演化事件为独立单元高度聚集，"
  "故只报点差与配对显著性，不做区间外推。"
  "本节全部结论的作用域因此是：本构造、本实现、本样本。")
CAP("表 7-104  规则基线的失败画像与两执行器的并联上界")
_rows = []
_FAIL_CN = {"DB_ERROR": "漏抽映射致执行报错", "SILENT": "静默错",
            "AMBIG_COL": "列名归属歧义", "NO_MAP": "无映射",
            "STAR_NO_SCHEMA": "星号无法展开"}
for gk, evo, gl in (("merge", "merge_col", "合并"), ("split", "split_col", "拆分"),
                    ("merge_holdout", "merge_col_holdout", "合并·留出"),
                    ("split_holdout", "split_col_holdout", "拆分·留出")):
    fp = RVR["failure_profile"].get(evo, {}).get("c3r_infer", {})
    fails = {k: v for k, v in fp.items() if k != "EX" and v}
    fdesc = ("；".join("%s（%d）" % (_FAIL_CN.get(k, k), v)
                      for k, v in sorted(fails.items(), key=lambda x: -x[1]))
             if fails else "无失败")
    p_ = RVRG[gk]["paired"]["c3r_infer-c3mp"]
    union = p_["a_ex"] + p_["b_only"]
    _rows.append([gl + "（%d）" % p_["n"], fdesc,
                  str(p_["b_only"]), str(p_["a_only"]),
                  f"{union}/{p_['n']}（{union / p_['n'] * 100:.1f}%）"])
TABLE(["档位", "规则·段位谱系的失败类型（条数）", "仅 LLM 对", "规则对而 LLM 错",
       "并联观测上界"],
      _rows, [2.6, 5.0, 1.8, 2.2, 3.0],
      ["l", "l", "c", "c", "c"], font=8.5)
NOTE("失败类型在该档的全部失败条目中按条数排列；「无失败」表示该档规则侧 100% 执行结果一致。"
     "「并联观测上界」= 规则成功数 + 仅 LLM 成功数，是两种执行器取并集的观测上界，"
     "不是可部署方法的性能。")
_fhd = RVR["failure_profile"]["merge_col_holdout"]["c3r_infer"]
P("这张表推翻了上一版的互补性结论，必须当场更正。上一版报告的"
  "「规则唯一失败是星号投影、LLM 在合并档补上 6/6、并联上界 230/230」"
  "建立在规则侧的误判之上：那 21 条被判为不可改写的条目全部是 COUNT(*)（7.4.4）。"
  "修正后，合并与拆分两档规则侧没有任何失败，"
  f"LLM 的独占正确条目为 {RVRG['merge']['paired']['c3r_infer-c3mp']['b_only']} 条与 "
  f"{RVRG['split']['paired']['c3r_infer-c3mp']['b_only']} 条；"
  "也就是说，在 dev 上「LLM 能而规则不能」的条目不存在，并联上界等于单用规则。"
  f"唯一出现 LLM 独占的是留出库合并档（{RVRG['merge_holdout']['paired']['c3r_infer-c3mp']['b_only']} 条），"
  f"其来源是规则侧 M1 漏抽映射造成的 {_fhd.get('DB_ERROR', 0)} 条执行报错——"
  "LLM 在这些条目上不是更会改写，而是不会因为缺少一条映射就整体失败。"
  "这个区别很重要：它说明两种执行器的失败模式不同（规则是硬失败，LLM 是静默错），"
  "但不构成「LLM 具备规则不具备的迁移能力」的证据。")
P("据此，本章对贡献的表述必须收窄为三层，且第三层是负面的。"
  "第一层，证据表示：在历史 SQL 可作证据的迁移设定下，"
  "列级映射与段位表达式映射之间的差距在合并方向为 +99.1pt（dev）与 +64.4pt（留出），"
  "在拆分方向为 0——列级映射会产生「可执行但语义错误」的 SQL。"
  "这是本章最稳的实证贡献，但它有明确的适用范围：只在「一列内多段编码」成立时成立。"
  "第二层，自动发现：M1 在 dev 上与 oracle 逐条等价（样本内），"
  f"在未见库上留下 {_rhm['c3r_gt']['rate'] - _rhm['c3r_infer']['rate']:.1f}pt 的硬失败缺口，"
  "且缺口全部表现为执行报错——说明该任务的瓶颈在映射发现，不在 SQL 生成。"
  "第三层，执行器：在标识符替换可完成迁移的条目上，确定性改写比 LLM 高 20.4pt（合并）与 22.1pt（拆分），"
  "且 LLM 没有独占正确条目。"
  "但本章没有测「需要重写谓词结构而非替换标识符」的迁移（如口径改变、单位换算、"
  "等值谓词变区间），因此不能声称 LLM 在该类情形上有优势——这是待检验项，不是已证结论；"
  "同样不能声称确定性改写是更优的工程选择，因为本构造的 gold 恰由同一替换规则生成。")

# ---------------------------------------------------------------- 7.4.4
H("7.4.4  星号处理的公平性核查", 3)
P("审稿人指出，把 SELECT * 直接判为失败会低估规则方法，"
  "因为旧 schema 与列映射已知时，按旧表列清单展开星号再逐列替换是可行的。"
  "按此核查后发现两个问题，其中第一个比审稿人设想的更严重。")
P(f"第一个问题：这是一个实现缺陷，不是方法上界。上一版用「AST 中是否出现星号节点」做判据，"
  f"于是 COUNT(*) 也被判为不可改写。审计显示，被判失败的 "
  f"{RWA['count_star_items']} 条全部含 COUNT(*)，"
  f"而真正的投影星号只有 {RWA['projection_star_items']} 条。"
  "COUNT(*) 在列替换下必须保持不变，判它失败等于把可以做对的题目主动判错。"
  "本版把判据改为「只看 SELECT 投影位上的 * 或 t.*」，COUNT(*) 不再触发，"
  "并补上真正的展开能力。")
CAP("表 7-111  规则基线修正前后的执行结果一致率（%）")
_rows = []
for gk, gl in (("merge", "合并"), ("split", "拆分"),
               ("merge_holdout", "合并·留出"), ("split_holdout", "拆分·留出")):
    row = V12[gk]
    _rows.append([gl]
                 + [f"{row[c]['v1']:.1f}" for c in ("c3r_gt", "c3r_infer", "c3r_name")]
                 + [f"{row[c]['v2']:.1f}" for c in ("c3r_gt", "c3r_infer", "c3r_name")]
                 + [("%+.1f" % row["c3r_infer"]["gain_pt"]).replace("-", "−")])
TABLE(["档位", "修正前·oracle", "修正前·自动", "修正前·列级",
       "修正后·oracle", "修正后·自动", "修正后·列级", "自动谱系增益"],
      _rows, [1.9, 1.9, 1.9, 1.9, 1.9, 1.9, 1.9, 2.1],
      ["l", "r", "r", "r", "r", "r", "r", "r"], font=8.2)
NOTE("修正内容为两处：一是投影星号判据不再误吞 COUNT(*)；二是补上按旧表列清单展开投影星号。"
     "本样本中投影星号 0 条，故修正后的数值差异全部来自第一处。")
P("第二个问题：展开能力本身需要验证，但样本无法提供验证机会。"
  "本版实现了展开通道：以原始库（演化前）为据，按 cid 顺序取旧表列清单，"
  "把投影位上的 * 或 t.* 展开为逐列引用，再交给统一的替换通道；"
  "两种模式并存——expand 展开旧星号，keep 保留星号原样。"
  f"在 {RWA['projection_star_items']} 条投影星号的样本上，两模式给出完全相同的数值，"
  "因此展开能力对本章任何已报告数字都没有影响。"
  "为说明该通道本身实现得对不对，另用一个两步演化的合成库做了自检（表 7-112）。"
  "必须强调：这是实现层面的自检，不构成任何评测证据——"
  "它回答的是「展开代码有没有写错」，不是「含星号的迁移任务能不能做对」。")
CAP("表 7-112  星号展开的合成探针（P1–P4）")
_ps = PSX["cases"]
_rows = []
for key, lab in (("P1_bare_star", "裸 *（单表）"),
                 ("P2_qualified_star", "T.* （带表限定）"),
                 ("P2_bare_star_join", "裸 *（两表 JOIN）"),
                 ("P3_count_star", "COUNT(*)"),
                 ("P3_count_star_where", "COUNT(*) + 被演化列谓词")):
    c_ = _ps[key]
    _rows.append([lab,
                  "是" if c_["expand"]["matches_old"] else "否",
                  "是" if c_["keep"]["matches_old"] else "否",
                  "是" if c_["expand_matches_gold"] else "否"])
TABLE(["探针用例", "expand 重建旧结果", "keep 重建旧结果", "expand 与构造 gold 一致"],
      _rows, [5.4, 2.8, 2.8, 3.4], ["l", "c", "c", "c"], font=8.5)
_vd = PSX["verdict"]
NOTE("旧结果为该 SQL 在演化前库上的执行结果，是语义基准。"
     "「构造 gold」指机械替换列引用、但不触碰任何星号的改写结果，即 keep 模式的行为。"
     f"判定：P1 展开重建旧结果 = {_vd['P1_expand_reconstructs_old']}、"
     f"keep 不能 = {_vd['P1_keep_fails']}、"
     f"T.* = {_vd['P2_qualified_ok']}、JOIN = {_vd['P2_join_ok']}、"
     f"COUNT(*) 不受影响 = {_vd['P3_count_star_unaffected']}。")
P("探针同时测出本构造的一个固有缺陷，这是本版必须写明的负面事实。"
  "构造脚本生成 gold 时只替换列引用，而 * 不是列引用，"
  "因此 gold 会原样保留 *，在演化后库上返回新的列集合——这与旧 SQL 的语义并不等价。"
  "展开模式给出的是语义正确的旧列集合，反而与 gold 不一致（表 7-112 最后一列全部为「否」）。"
  "换言之，含投影星号的题目在本构造下无法被忠实评价："
  "保留星号则与旧语义不等价，展开星号则与 gold 不一致。"
  f"本样本恰好没有这类题目（投影星号 {RWA['projection_star_items']} 条），"
  "所以这一缺陷不影响任何已报告的数字，但它是构造的范围限制，列入 7.10。")
P("把这件事的定位说准，因为它最容易被读高：星号展开能力在本章没有经过基准验证。"
  "主评测中投影星号 0 条，展开通道从未参与任何一条题目的打分，"
  "因此表 7-112 只能证明「实现能按旧表列清单展开、并重建旧结果」，"
  "不能证明「含星号的迁移任务已被解决」。"
  "本章把它记为一条「评测范围限制」，而不是一条能力声明；"
  "任何「本文方法可处理 SELECT *」的读法都超出了现有证据。"
  "若要真正支持这一主张，需要一个含投影星号、且 gold 由人工而非机械替换生成的子集，"
  "本章没有构造这样的子集，列为待做项（7.10 第 ⑥ 条）。")

# ================================================================ 7.5
H("7.5  RQ4：跨基座复现", 2)
P("表 7-20 在四个基座上重跑改名档全部 276 条。四者绝对水平差异很大"
  "（C2 从 50.7% 到 80.4%），但 C2 → C3m 的方向在四个基座上一致。")
CAP("表 7-20  改名档跨基座结果（n = 276，%）")
_rows = []
XM = S["xmodel"]
for m, label in [("deepseek-ai/DeepSeek-V3.2", "DeepSeek-V3.2"),
                 ("deepseek-ai/DeepSeek-V4.1-Flash", "DeepSeek-V4.1-Flash"),
                 ("Qwen/Qwen3-Coder-30B-A3B-Instruct", "Qwen3-Coder-30B-A3B"),
                 ("Qwen/Qwen3-30B-A3B-Instruct-2507", "Qwen3-30B-A3B-Instruct")]:
    v = XM[m]
    _rows.append([label, f"{v['c1']['rate']:.1f}", f"{v['c2']['rate']:.1f}",
                  f"{v['c3f']['rate']:.1f}" if "c3f" in v else "—",
                  f"{v['c3m']['rate']:.1f}", f"{v['delta']:+.1f}"])
TABLE(["基座", "C1", "C2", "C3f", "C3m", "Δ(C3m − C2)"], _rows,
      [4.6, 2.2, 2.2, 2.2, 2.2, 3.0], ["l", "r", "r", "r", "r", "r"])
NOTE("四个基座的 C1 差异（21.7–42.8）说明绝对水平受基座能力影响显著，"
     "故本文不比较跨基座的绝对正确率，只比较同一基座内的条件差值。"
     "GLM-5.3-Flash 只在改名档的 57 条子集上跑过（C2 80.7% → C3m 93.0%，"
     "Δ = +12.3pt），未纳入本表。")
P("合并档的复现更关键，因为它是本章的主结果。表 7-21 在同一批 230 条题上比较两个基座："
  "段位谱系相对列级谱系的落差，DeepSeek-V4.1-Flash 为 +79.6pt，DeepSeek-V3.2 为 +77.0pt，"
  "两者相差 2.6pt。")
CAP("表 7-21  合并档主结果的跨基座复现（n = 230，%）")
TABLE(["条件", "DeepSeek-V3.2", "DeepSeek-V4.1-Flash", "V4.1F − V3.2"],
      [["C1 零证据", "5.7 (13/230)", "4.8 (11/230)", "-0.9"],
       ["C2 旧 SQL", "10.0 (23/230)", "8.3 (19/230)", "-1.7"],
       ["C3m 列级谱系", "2.6 (6/230)", "8.3 (19/230)", "+5.7"],
       ["C3mp 段位谱系", "79.6 (183/230)", "87.8 (202/230)", "+8.3"]],
      [4.0, 4.2, 4.4, 3.0], ["l", "r", "r", "r"])
NOTE("两基座逐条件在同题集上比较。C3m 在 V3.2 上低于 C2（2.6 vs 10.0），"
     "在 V4.1F 上等于 C2（8.3 vs 8.3），两基座均未显示列级谱系的收益。")
P("两个基座一致的地方比幅度本身更重要。第一，C3m 在两个基座上都没有跑赢 C2，"
  "「列级对齐必要但不充分」不是单一模型的偶然。第二，落差的量级接近，"
  "说明这不是尺度效应。需要限定的仍然是：两个基座、一种演化类型、5 个库，"
  "不足以支撑「跨基座普适」的表述。")

# ================================================================ 7.6
H("7.6  失败模式：静默错误从哪里来", 2)
P("正确率只说明结果好坏，不说明失败的性质。表 7-22 按判定类别拆开各条件，"
  "可以看到合并档的失败几乎全部是静默的。")
CAP("表 7-22  各条件的判定分布（条）")
_rows = []
for o in S["outcomes"]:
    _rows.append([o["label"], str(o["n"]), str(o["EX"]), str(o["SILENT"]),
                  str(o["ERROR"]), str(o["OVER_ABSTAIN"])])
TABLE(["条件", "n", "EX", "SILENT", "ERROR", "OVER_ABSTAIN"], _rows,
      [4.4, 1.4, 1.4, 1.8, 1.6, 2.8], ["l", "r", "r", "r", "r", "r"])
NOTE("SILENT 含空结果集。合并 · C3m 的 223 条静默错误占该档 97.0%，且 0 条显式报错——"
     "这是本任务中最难被下游发现的失败形态。合并 · c_rw 的 62 条 ERROR 全部（62/62）引用了"
     "已删除的旧列名，拆分档为 34/35，全章 98 条中 97 条如此，属真实失败模式而非调用故障。")
H("7.6.1  静默错误的成分归因", 3)
P("对合并档 C3m 的 223 条静默错误逐条解析生成的 SQL，可以判断失败发生在哪一层。"
  "其中 220 条（220/223 = 98.7%）确实引用了合并后的新列名（attr_001），"
  "只有 5 条（5/223 = 2.2%）出现了任何切段构造（SUBSTR / INSTR / SPLIT_PART 等）。"
  "这两个计数不是互斥的两类，而是嵌套的：逐条核对确认，那 5 条同时属于 220 条，"
  "即「引用了新列、也写了切段、但切法或谓词仍不对」；220 − 5 = 215 条则是"
  "「引用了新列、完全没有切段」。因此正确的读法是分层的——"
  "先看有没有引用新列（220 有、3 没有），再看引用了新列的里面有没有表达段位（5 有、215 没有）。"
  "也就是说，失败几乎完全不在「列找错了」，而在「对着一列整体操作，没有表达它由多段构成」。"
  "必须同时说明这条归因的边界：它来自对模型生成 SQL 的结构解析（是否出现目标列引用、"
  "是否出现切段函数），属于机制层面的描述性证据，不是对全部潜在语义错误原因的穷尽证明；"
  "而且 SILENT 口径含空结果集，个别条目可能因谓词过严返回空集而非因粒度问题，"
  "这一部分无法只靠结构解析与「粒度缺失」区分开，故本节只声称「与粒度缺失是主因这一解释一致」，"
  "不声称已排除其他成因。")
P("同一归因在段位条件下呈现相反形态：C3mp 的 47 条静默错误中，"
  "45 条既引用了新列又包含了切段构造，失败点已下移到取值与条件谓词层。"
  "这构成一条机制性证据链——列级条件把错误留在粒度层，段位条件把粒度层的问题基本清除，"
  "残余失败转移到更难的下游环节。")
P("这条证据也解释了为何「列级恢复了 100%」与「端到端只有 2.6%」不矛盾："
  "前者度量映射表是否正确，后者度量查询是否表达了正确的粒度，两者之间没有必然联系。"
  "因此本文主张不应用列级恢复率预测下游性能。")

# ================================================================ 7.7
H("7.7  稳健性与局限", 2)
H("7.7.1  基础题级聚类检验", 3)
P("改名档 276 条只对应 226 个基础题簇。表 7-23 并列给出逐条与簇级两种口径："
  "主效应在簇内等权后从 +22.5pt 降为 +21.0pt，显著性不变，"
  "说明它不是由同一基础题的重复派生堆出来的。合并与拆分档一题一条，两口径必然相同。")
CAP("表 7-23  逐条与基础题级聚类的配对检验")
_rows = []
for t in CL["tests"]:
    _rows.append([t["name"], str(t["n_items"]), str(t["n_clusters"]),
                  f"{t['diff_item_pt']:+.1f}", pp(t["p_item"]),
                  f"{t['diff_cluster_pt']:+.1f}",
                  f"[{t['ci_cluster'][0]:+.1f}, {t['ci_cluster'][1]:+.1f}]",
                  pp(t["p_cluster"])])
TABLE(["对比", "条目", "簇", "Δ 条目", "p（逐条）", "Δ 簇", "簇级 95% CI", "p（聚类）"],
      _rows, [5.0, 1.2, 1.2, 1.6, 1.8, 1.6, 3.2, 1.8],
      ["l", "r", "r", "r", "r", "r", "c", "r"], font=8.3)
NOTE("簇级检验为簇内等权正确率差的符号检验，区间由 5,000 次聚类自助得到。"
     "Δ 条目与 Δ 簇的差别仅来自簇内条目数不等，不涉及两套数据。"
     "表中「改名 hard：C3m − c_rw」的 p 未做 Holm 校正；校正后为 0.064（见表 7-19）。")
H("7.7.2  采样与列属性敏感性", 3)
P("谱系抽取使用全表统计量，理论上应受采样影响。以 7 条与 13 条记录的两种采样"
  "（相同随机种子）重跑抽取，21 组配对逐 op 完全相同（21/21），"
  "端到端结果因此不可能改变。这排除了「抽取依赖全表扫描」这一实现层面的质疑。")
P("列属性方面，评测涉及的 162 列中有 69 列（42.6%）为近唯一列（ndist/n ≥ 0.9），"
  "56 列为极低基数（≤ 1%），55 列含 NULL 或空串，26 列为非文本类型"
  "（INTEGER 6、REAL 7、DATE 13）。近唯一列是值域证据的天然盲区："
  "重编码后能否恢复只取决于 ndist/n，低于 0.90 的 7 列全部可恢复，"
  "不低于 0.90 的 6 列全部不可恢复。这一信息下界属于方法的固有限制，"
  "需要外部元数据（数据字典、审计日志）才能突破。")
H("7.7.3  本章证据的边界", 3)
P("有九点必须与结论一并声明，否则结论会被误读：")
PL("① 全部 11 个数据库都参与了方法开发与阈值调整，阈值曾依据 gold operation 设定。"
   "因此本文此前把 M1 称为「无监督」并不准确：其推断阶段确实不读取测试标签，"
   "但阈值由开发集的 gold operation 校准过，属于"
   "「推断阶段不使用标签、阈值由开发集校准」"
   "（label-free inference with development-set calibrated thresholds），"
   "而非完全未使用监督信息。这一术语在第 5 章、摘要与投稿信中统一按此表述。"
   "由此，主实验是回顾性合成评测，不是干净的留出测试。"
   "为补上这一条，7.7.4 在 BIRD train 划分的 7 个未参与开发的库上重做了确认实验："
   "总体方向重现（合并档事件 6/7 为正、1/7 为负），但样本量为 165 条，"\
   "只能作为「方向可迁移」的旁证，"
   "不足以把主实验升级为留出测试。")
PL("② 本章报告的显著性一律是「该构造评测集上、以题目为单位」的比较，"
   "不是跨库或跨演化类型的推断。合并与拆分各只涉及 5 个库；"
   "合并档的 10 个原子 op 两两属同一次列合并，独立变更事件只有 5 次，"
   "与库数重合。若按 C3mp − c_rw 的统一口径把合并与拆分两档的 10 次列重整"
   "池化，符号检验为 10 正 0 负（p = 0.0020），方向一致；"
   "但库级只有 5 个单元（p = 0.0625），"
   "在本数据上仍无法把「库的效应」与「变更事件的效应」完全分开（表 7-27）。"
   "簇级口径（表 7-23）显示结论不由单一库驱动，但簇数本身太少，不足以支持泛化声称。")
PL("③ schema 演化由构造脚本施加，旧库、新库、gold SQL 与谱系标签同源。"
   "真实场景中不存在已知的分隔符字面量与段号标注。")
PL("④ 主实验的合并与拆分只涉及两段场景，真实 BIRD 数据中也不存在三段以上合并事件，"
   "故本章不把三段结果写入主结论。但「构造工具在三段情况下会生成重复表达式」这一"
   "已知硬限制已定位并修复：生产代码 m1_part.to_expr 在 nseg ≥ 3 的末段分支多切了一次 "
   "INSTR，改为直接取剩余全部后，隔离库压力测试（表 7-120）在 2–6 段、"
   f"5 类分隔符、8 类边界值上共 {len(EBV['cells'])} 格逐格执行，"
   f"「末段内含分隔符」的可恢复率由修正前 {EBL['by_case']['sep_in_last']:.1f}% 升至 "
   f"{EBV['by_case']['sep_in_last']:.1f}%，nseg ≥ 4 与 ≤ 3 不再有差异；"
   "所有真实 gold 表达式均为两段 flat 形态（2,384 个，零嵌套），故该修正不影响任何历史数值。"
   "仍需限定：NULL 分量与「非末段内含分隔符」两类在修正后仍 0% 可恢复，"
   "这是分隔符拼接的信息论限制，不是实现缺陷；段内互含与多分隔符嵌套仍未被真实演化覆盖。")
PL("⑤ 判定器对行序不敏感：gold 含 ORDER BY 时，行序错误的查询仍可能判为一致。"
   f"该风险已在 7.1.7 量化（重执行 {JS_RE['n']:,} 条预测后被宽判 "
   f"{JS_RE['n_order_sensitive']} 条，{JS_RE['share_order_sensitive_pct']:.2f}%，"
   "全部集中在 gold 含 ORDER BY 的题），并在本轮用「行序敏感 + 数值容差 + NULL 区分」"
   f"三者同时收紧的全严规则复核：八个主对照符号全部保持不变（{JSA['sign_stable_count']}），"
   f"最大偏移 {JSA['max_abs_delta_pt']:.1f}pt（表 7-118）。"
   "「行序完全正确」仍未作为独立主指标报告，现行主判保持 r0 并双报全严结果；"
   "该限制由敏感性分析界定，而非被消除。")
PL(f"⑥ 服务端不提供随机种子，temperature = 0 的贪心解码仍不能保证逐轮逐位一致。"
   f"核心 LLM 条件已独立重跑 {RPT['n_repeat_runs']} 次，各条件读数相对首次报告值的"
   f"最大绝对偏差为 {RPT['max_abs_dev_pt']:.1f}pt（作者观察带 "
   f"{RPT_P['observed_band_pt'][0]:.1f}–{RPT_P['observed_band_pt'][1]:.1f}pt，"
   "按保守下界 0.1pt 统一表述），八个主对照的方向、相对次序与显著性判定均不变（表 7-121）。"
   "但这不等于逐题判定可复现：波动来自服务端批处理/调度层面的非确定性，"
   f"且{RPT['reading_note']}故结论仍以整批比率报告，不对单条判定作推断，"
   "也不把重复运行当作独立的显著性证据。")
PL(f"⑦ 检索模块未实现，检索式探针为阴性对照：{E2['n']} 条，题面取自改名档"
   f"（{E2['panels']}），只覆盖 1 个数据库（{E2['db']}）。"
   f"「检索到一个相似但过期的示例」条件（C4）与「再加该示例中过期列的对齐提示」条件（C5）的"
   f"执行结果一致率分别为 {E2['c4_rate']:.1f}% 与 {E2['c5_rate']:.1f}%，"
   f"不一致对 {E2['b']}:{E2['c']}，精确符号检验 p = {E2['p']:.2f}；"
   f"{E2['n_runs']} 次运行中 {E2['silent']} 次为静默错误、0 次弃答。"
   "三点必须同时读：其一，该探针至多说明「补一个检索到的过期示例不足以救回迁移」，"
   "既不评价召回与排序策略，也不涉及证据冲突裁决；"
   "其二，题面取自改名档而非合并或拆分档，故与本章主结论（粒度效应）没有交集，"
   "不能反过来用作「谱系注入无效」的证据；"
   "其三，C5 相对 C4 的 +7.5pt 在 n=40 上不显著，"
   "说明「把对齐提示塞进检索到的示例」这条最省事的路子本身也不成立——"
   "这正是本文不把检索写成研究方向的经验依据，而不是它的替代品。")
PL(f"⑧ SQL 结构覆盖：连接不是短板，但外连接与「需要重写谓词」的结构基本空白。"
   f"逐条解析 {SCV['paper']['n']} 条 old_sql，含 JOIN 的占 {SCV_P['join_any']['pct']:.1f}%，"
   f"而 LEFT JOIN 仅 {SCV_P['join_left']['pct']:.1f}%、"
   f"JOIN USING / NATURAL {SCV_P['join_using_or_natural']['pct']:.1f}%、"
   f"相关子查询 {SCV_P['subquery_correlated']['pct']:.1f}%、"
   f"CTE {SCV_P['cte']['pct']:.1f}%、集合运算 {SCV_P['union_like']['pct']:.1f}%、"
   f"窗口函数 {SCV_P['window']['pct']:.1f}%（表 7-136）。"
   f"目标列所在子句为 WHERE {SCV_POS['Where']['pct']:.1f}%、"
   f"SELECT {SCV_POS['SELECT']['pct']:.1f}%、ORDER BY {SCV_POS['Order']['pct']:.1f}%、"
   f"GROUP BY {SCV_POS['Group']['pct']:.1f}%、JOIN ON {SCV_POS['Join']['pct']:.1f}%，"
   "HAVING 与子查询内为 0。"
   "这三组数字给出三条适用范围（7.17 逐条列出）：已检验的是列引用替换本身，"
   "未覆盖的是上述空白结构，不适用的是「需要改变谓词结构而非替换标识符」的迁移"
   "（单位换算、口径改变、等值谓词改区间）。第三类不是覆盖不足，而是任务定义不同，"
   "本章对它没有任何主张。")
PL(f"⑨ gold 与本方法的改写规则同源，读数中包含可观的信息补贴；且缺独立人工核验。"
   f"替换后的 gold 与构造期 rewrite() 的输出逐字相同：合并 {GDI_M['identical']}/"
   f"{GDI_M['n']}、拆分 {GDI_S['identical']}/{GDI_S['n']}（均 100%）；"
   f"被插入的表达式中位长度占 gold 全长的 {GDI_M['inserted_char_share_median']}%"
   f"（合并）与 {GDI_S['inserted_char_share_median']}%（拆分）。"
   "这意味着 c3mp 的读数有相当部分是「照抄给出的公式即可」，不应读作模型自行推断出段位。"
   f"补贴的量级由同题、同模型、同提示骨架的表达形式消融给出（该消融在 DeepSeek-V3.2 上完成，"
  f"与主结果的 79.6% 同源）："
   f"给可复制表达式 {GDI_AB['c3mp_literal']['rate']}% 对给等价自然语言描述 "
   f"{GDI_AB['c3mp_nat']['rate']}%，差 {GDI_AB['delta_pt']}pt"
   f"（{GDI_AB['only_literal']}:{GDI_AB['only_natural']}，p = {GDI_AB['mcnemar_p']:.0e}），"
   "即约 15 个百分点的提升来自表达形式本身。"
   "更需要直说的一点：本轮没有第二编码者，因而没有双人一致性与 kappa 可报，"
   "构建期三步验证与本文的逐题判定均为单人完成；"
   "公开材料包因此只提交编码手册与一份 60 条分层抽样的人工核验协议，"
   "不提交任何一致性系数。这是未完成事项，不是阴性结果。")

H("7.7.4  两项留出实验：端到端确认（H-E2E）与谱系表达式直接恢复（H-LER）", 3)
P("「按数据库留出一组确认实验」是第五轮评审最硬的一条。本节报告补做的结果，"
  "但先要把主实验为什么只能做到题目级说清楚——这不是推脱，是基准结构决定的。"
  f"BIRD 的 dev 划分一共只有 {HOA['benchmark_facts']['bird_dev_databases']} 个数据库，"
  f"本章五档合计正好用满这 {HOA['benchmark_facts']['bird_dev_databases_used']} 个，"
  "没有第 12 个库可以留出。真正的留出库只能取自 train 划分。")
_rows = []
for p_ in HOA["panels"]:
    _nm = {"easy": "改名 easy", "medium": "改名 medium", "hard": "改名 hard",
           "merge": "合并", "split": "拆分"}[p_["panel"]]
    _rows.append([_nm, str(p_["n_items"]), str(p_["n_db"]), str(p_["n_events"]),
                  f"{p_['items_per_event_min']:.0f}–{p_['items_per_event_max']:.0f}"])
# 不设「合计」行：改名三档的事件在同一批库上重叠，跨档相加会重复计数。
CAP("表 7-24  主实验各档的库数与独立变更事件数")
TABLE(["档位", "条目", "库数", "独立变更事件", "每事件条目"], _rows,
      [3.4, 2.0, 2.0, 3.2, 3.4], ["l", "r", "r", "r", "r"], font=8.4)
_pw = HOA["power_at_event_level"]
NOTE("「独立变更事件」按同库同表的一次列重整计。合并档的 10 个原子 op 两两来自同一次"
     "列合并（同一个合并列的两个来源列），因此只算 5 次事件；把它当作 10 个独立事实"
     "会高估检验功效。这也意味着合并与拆分档的事件数恰好等于库数（各 5），"
     "在本数据上无法把「库的效应」与「变更事件的效应」分开。")

P("留出侧共有两项实验，范围不同、判定对象不同，必须分开报告，不得相加、也不得互为分母。"
  "本节先把两者的边界固定下来，再依次给出结果："
  "① held-out end-to-end confirmation（留出端到端确认）——在 train 划分的 7 个未参与开发的库上，"
  "用冻结的流水线跑完整的迁移判定，共 165 条评测记录、12 个变更事件单元，"
  "回答「方向能否迁移」；"
  "② held-out direct lineage-expression recovery（留出谱系表达式直接恢复）——"
  "把冻结的 M1 part 分支直接施加到 train 划分的演化库上，绕开端到端判定，"
  "把推断出的表达式与 gold 逐字符比对，共 481 条评测记录、16 个原子 op、8 次独立事件、8 个库，"
  "回答「抽取器在未见库上还能不能给出正确表达式」。"
  "两者共用 BIRD train 划分这一来源，但库的集合不同（7 个对 8 个）、"
  "评测对象不同（迁移结果对表达式字符串）、条目数也不同（165 对 481），"
  "因此 165 条与 481 条既不能相加为 646 条，也不能把 87.5% 与 69.3% 放进同一张阶梯表比较。"
  "本节标题与图表注释一律使用上述英文全称的缩写形式（H-E2E、H-LER），避免读者把两者读成一件事。")

P("本轮把这层证据补上了。做法是把构造、清单生成与谱系推断三个脚本的库路径参数化，"
  "在 train 划分上独立构造，演化库与清单写入独立目录（data/evolved_*_holdout、"
  "data/manifest_holdout.jsonl）。参数在主实验完成后即冻结：段位表达式的生成规则、"
  "M1 的三条判定证据与阈值、提示模板、判定器与执行结果一致比较函数，全部沿用主实验的版本；"
  "留出样本及其结果未参与任何调参，也没有为留出库重跑阈值搜索。"
  "换言之，留出实验是「把冻结的流水线搬到新库上跑一遍」，不是「在新库上重新设定参数」。")
P("库的选择不是随机抽样，这一点必须写明，否则 7 个库会被误读为 train 划分的代表。"
  f"筛选逐级进行（表 7-105）：BIRD train 共 {FS['S0']['n_db']} 个库、"
  f"{FS['S0']['n_items']:,} 道题；先按单库文件大小剔到 {FS['S1']['n_db']} 个"
  f"（低于 {FUN['big_mb']:.0f} MB，因为 SQLite 的 DROP COLUMN 会重写整表，超大库在构造阶段会超时），"
  f"再按题数剔到 {FS['S2']['n_db']} 个（至少 {FUN['min_q']} 道，否则配对检验没有意义），"
  f"再按「能否施加无损演化」剔到 {FS['S4']['n_db']} 个、{FS['S4']['n_items']} 条，"
  f"最后按构造期自检剔到 {FS['S5']['n_db']} 个、{FS['S5']['n_items']} 条，"
  f"每（档, 库）按 id 哈希取前 15 条，得 {FS['S6']['n_items']} 条。"
  "每一级被剔除的库名与原因都记录在 results/holdout_funnel.json，并随材料包发布。"
  "因此这 7 个库是「在可施加本构造的前提下可得的最大集合」，不是 train 的随机样本；"
  "BIRD 全集共 95 个库，本章只触及其中的 11（dev）+ 7（train）。")
_rows = []
for st in FUN["stages"]:
    _rows.append([st["stage"], st["label"], str(st["n_db"]),
                  str(st["n_items"]) if st.get("n_items") else "—",
                  str(len(st.get("removed") or [])),
                  st["reason"].replace("**", "")])
CAP("表 7-105  留出库的筛选漏斗（BIRD train 划分）")
TABLE(["级", "条件", "剩余库", "剩余条目", "本级剔除", "剔除原因"],
      _rows, [1.2, 3.6, 1.5, 1.9, 1.7, 6.6],
      ["c", "l", "r", "r", "r", "l"], font=7.8)
NOTE(f"S5 的「整库 0 条通过」为 {', '.join(FS['S5']['removed'])}（自检全部失败，整库出局）；"
     f"{'、'.join(FS['S5'].get('partial') or [])} 为部分失败，按条目抽取，只保留通过自检的条目。"
     "S1 剔除的库名与 S4 剔除的 41 个库名见 results/holdout_funnel.json。")
_rows = []
_labs = {"books": "books", "movie_3": "movie_3", "movies_4": "movies_4",
         "olympics": "olympics", "regional_sales": "regional_sales",
         "retail_world": "retail_world", "soccer_2016": "soccer_2016"}
_cells = HSMP["cells"]
for db in sorted({k.split("|")[1] for k in _cells}):
    m = _cells.get(f"merge_col_holdout|{db}", 0)
    sp = _cells.get(f"split_col_holdout|{db}", 0)
    _rows.append([db, str(m) if m else "—", str(sp) if sp else "—"])
_tot = HSMP["n"]
CAP("表 7-25  H-E2E 样本构成（BIRD train 划分，未参与开发）")
TABLE(["库", "合并档条目", "拆分档条目"], _rows,
      [5.4, 3.2, 3.2], ["l", "r", "r"], font=8.6)
NOTE(f"合计 {_tot} 条（合并 {HSMP['by_evo'].get('merge_col_holdout', 0)}、"
     f"拆分 {HSMP['by_evo'].get('split_col_holdout', 0)}），每库每档按 id 哈希排序取前 15 条，"
     "避免依赖文件顺序。所有条目均通过构造期自检（旧 SQL 在新库失效、新 SQL 结果等价）。")

_M = HLD["evos"]["merge_col_holdout"]
_S = HLD["evos"]["split_col_holdout"]
_conds = ["c2", "c3m", "c3mp"]
_cn = {"c2": "C2 旧 SQL + 新 schema", "c3m": "C3m 旧 SQL + 列级谱系",
       "c3mp": "C3mp 旧 SQL + 段位谱系"}
CAP("表 7-26  H-E2E：留出库上的端到端迁移正确率（DeepSeek-V3.2）")
TABLE(["条件", "合并档（n=101）", "拆分档（n=64）"],
      [[_cn[c],
        f"{_M[c]['ex']}（{_M[c]['rate']:.1f}%）",
        f"{_S[c]['ex']}（{_S[c]['rate']:.1f}%）"] for c in _conds],
      [6.2, 4.4, 4.4], ["l", "r", "r"], font=8.8)
_P1 = HLD["paired"]["merge_col_holdout|c3mp-c3m"]
_P2 = HLD["paired"]["merge_col_holdout|c3mp-c2"]
_P3 = HLD["paired"]["merge_col_holdout|c3m-c2"]
_Q1 = HLD["paired"]["split_col_holdout|c3m-c2"]
_d1 = ("%+.1f" % _P1["delta_pt"]).replace("-", "−")
_d2 = ("%+.1f" % _P2["delta_pt"]).replace("-", "−")
NOTE(f"合并档：C3mp 相对 C3m 为 {_n(_d1)}pt"
     f"（{_P1['a_only']} : {_P1['b_only']}，p < 1e−6），相对 C2 为 "
     f"{_n(_d2)}pt"
     f"（{_P2['a_only']} : {_P2['b_only']}）；而 C3m 相对 C2 只有 +5.0pt（{_P3['a_only']} : {_P3['b_only']}，"
     f"p = {_P3['p']:.3f}）。拆分档的 C3mp 与 C3m 完全相同（0 : 0）："
     "拆分不产生「一列内多段」，段位谱系在这里没有额外信息可提供，"
     "该档的增益全部来自列级谱系（相对 C2 为 "
     f"+{_Q1['delta_pt']:.1f}pt，{_Q1['a_only']} : {_Q1['b_only']}）。")

_PDB = {k.split("|")[1]: v for k, v in HLD["per_db"].items()
        if k.startswith("merge_col_holdout|")}
_pos = {k: v["c3mp_minus_c3m_pt"] for k, v in _PDB.items() if v.get("c3mp_minus_c3m_pt", 0) > 0}
_neg = {k: v["c3mp_minus_c3m_pt"] for k, v in _PDB.items() if v.get("c3mp_minus_c3m_pt", 0) < 0}
def _pm_(v):
    return ("%+.1f" % v).replace("-", "−")


P(f"逐库看，合并档的 C3mp − C3m 在 {len(_pos)} 个库为正（{_pm_(min(_pos.values()))} 到 "
  f"{_pm_(max(_pos.values()))}pt），在 {len(_neg)} 个库为负（"
  + "、".join(f"{k} {_pm_(v)}pt" for k, v in _neg.items()) + "）。"
  f"其中 regional_sales 的负向必须作为主结果报告，而不是归入局限节。"
  f"该库的合并列取值较短，C3m 已达到 {_PDB['regional_sales']['c3m']['rate']:.1f}%，"
  f"段位表达式反而使模型在拼接格式上出错，净差 "
  f"{_pm_(HLD['per_db']['merge_col_holdout|regional_sales']['c3mp_minus_c3m_pt'])}pt。"
  "把它放在这里而不是放在 7.7.5，是因为它界定了主张的形式："
  "「段位优于列级」是总体与多数库的方向，不是逐库成立的规律；"
  "在取值本就短、列级映射已能容纳歧义的库上，段位表达式会引入新的格式错误。"
  "这一条与 7.7.5 的边界一起，构成本章主张的两侧限制。")
_E = HLD["event_level"]["merge_col_holdout"]
P(f"按独立变更事件看，合并档 {_E['n_cells']} 个「库×表」单元中 "
  f"{_E['positive']} 个 C3mp 优于 C3m、{_E['negative']} 个更差、{_E['tie']} 个持平，"
  f"精确符号检验 p = {_E['p']:.3f}。"
  "这一结果与主实验同量级（主实验 5 : 0，p = 0.0625），仍不足以在 0.05 下判定，"
  "但方向在未见库上重现了——这是本节最重要的结论，也是最需要克制的地方。")

P("与主实验对照，可以得到三条更精确的判断。"
  f"第一，效应的量级一致：主实验合并档 C3mp 为 79.6%、C2 为 10.0%，"
  f"留出库为 {_M['c3mp']['rate']:.1f}% 与 {_M['c2']['rate']:.1f}%，"
  "两者相差约 10pt，属于同一数量级，说明机制不是 dev 库上的特例。"
  f"第二，拆分档的增益同样重现（主实验 {S['main']['split']['c3mp']['rate']:.1f}%，留出 "
  f"{_S['c3m']['rate']:.1f}%），但留出库的 C2 基线高得多"
  f"（{_S['c2']['rate']:.1f}% vs {S['main']['split']['c_rw']['rate']:.1f}%），"
  "因此相对增益较小（+37.5pt vs +68.3pt）——绝对水平受库的难度影响，方向不受影响。"
  f"第三，C3m 与 C2 的关系在两处不一致：主实验合并档 C3m（2.6%）低于 C2（10.0%），"
  f"留出库 C3m（{_M['c3m']['rate']:.1f}%）略高于 C2（{_M['c2']['rate']:.1f}%），"
  f"两处均不显著（留出 p = {_P3['p']:.3f}）。"
  "正确的读法是：在合并档，列级谱系本身既不能稳定带来增益、也不能稳定造成损害，"
  "它接近地板；真正的差异出现在「是否下沉到段位」这一步。")
P("这一节能支持的主张因此可以前进一步，但仍有限度："
  f"「段位谱系相对列级谱系的优势在 {len(_PDB)} 个未参与开发的库上重现"
  f"（合并档 {_E['positive']} : {_E['negative']} 事件方向、"
  f"逐库 {len(_pos)} 正 {len(_neg)} 负）」，这比主实验的题目级结论强。"
  "不能支持的是：跨真实历史演化的泛化——留出库仍是脚本合成的演化，"
  "库虽未见，变换方式与阈值仍出自同一套构造器；样本也只有 7 个库、12 个事件单元；"
  "且没有人类标注校验，gold SQL 的正确性依赖构造期自检。")
P("本轮补上一项此前缺失、且对方法不利的直接证据。上文那项是 H-E2E，它回答「方向能否迁移」，"
  "却回答不了「抽取器本身在未见库上还剩多少」；这一项是 H-LER：M1 在留出库上「能否恢复 gold 的"
  "段位表达式」，而不是只看端到端迁移正确率（表 7-119）。"
  "两者须严格分开：库集合不同（H-E2E 7 个、H-LER 8 个）、条目数不同（165 对 481）、"
  "判定对象不同（执行结果对表达式字符串），故不可相加、不可互为分母，也不可互相替代。"
  "做法是把冻结的 M1 part 分支直接施加到 BIRD train 划分的演化库上，"
  "把推断出的表达式与 gold 逐字符比对；两重独立性成立——"
  "主实验用 dev 划分、阈值在 dev 上校准，留出集用 train 划分，数据划分与阈值校准互不重叠。"
  f"结果：{MHO['units']['items']} 条条目只对应 {MHO_OP['n_units']} 个唯一原子 op、"
  f"{MHO_EV['n_units']} 次独立事件、{MHO_DB['n_units']} 个库；"
  f"op 级恢复 {MHO_OP['matched']}/{MHO_OP['n_units']}（{MHO_OP['rate_pct']:.1f}%，"
  f"Wilson 95% 区间 [{MHO_OP['wilson95_pct'][0]:.1f}, {MHO_OP['wilson95_pct'][1]:.1f}]），"
  f"事件级 {MHO_EV['fully_recovered']}/{MHO_EV['n_units']}（{MHO_EV['rate_pct']:.1f}%），"
  f"库级 {MHO_DB['fully_recovered']}/{MHO_DB['n_units']}（{MHO_DB['rate_pct']:.1f}%）。"
  f"两个失败均为 no_infer（{MHO['failures'][0]['db']}.{MHO['failures'][0]['old']}、"
  f"{MHO['failures'][1]['db']}.{MHO['failures'][1]['old']}），即推断器未给出表达式，"
  "不是给出错误表达式。相对 dev 上 10/10 的字符级一致，这是明显回落，"
  "且必须如实写入：它把「dev ↔ 未见库差距」从被动承认变成带数值的主动披露，"
  "也说明 10/10 不能读作通用恢复能力。"
  f"同时区间上限达 {MHO_OP['wilson95_pct'][1]:.1f}%、样本仅 "
  f"{MHO_OP['n_units']} 个 op，故 {MHO_OP['rate_pct']:.1f}% 只能作为"
  "「在该留出构造上的观测值」，不构成对真实场景的恢复率估计。")

# ---------------------------------------------------------------- 7.7.5
H("7.7.5  取值编码的可逆性边界", 3)
P("评审提出一个此前没有正面回答的问题：主实验的同一条 SQL 在有限的数据行上结果相同，"
  "并不等于它在新值上也保持语义。段位表达式 SUBSTR(attr,1,INSTR(attr,CH)+k) 是否等价于原列，"
  "取决于合并后的取值是否满足某些前提；一旦不满足，查询照样能执行，结果却已经漂移。"
  "本节用三类证据回答：先在隔离环境里逐类造边界值，测出前提究竟有哪些；"
  "再换一套独立实现复核这些前提，并补上三段合并；"
  "最后回到真实数据，量一量这些边界在现构造中出现的频率。")
P("第一类是反事实探针（probe_counterfactual.py）：在内存库中建三列 (a, b, attr)，"
  "按构造规则写入 attr = a ‖ CH ‖ b，再对四种谓词形态分别比较原列与段位表达式的真值——"
  "等值比较、IS NULL、往返恒等（切回 a、b 后能否逐字还原）、以及三者的组合。"
  "每类边界值重复两次以排除偶发。表 7-106 给出结果。")
_LBL = {"CLEAN": "常规值", "NULL_SIB": "单个分量为 NULL", "BOTH_NULL": "两分量均为 NULL",
        "EMPTY": "空串", "SEP_IN_A": "分隔符在第一分量内", "SEP_IN_B": "分隔符在第二分量内",
        "WS": "首尾空白", "DUP": "重复值", "UNICODE": "多字节字符",
        "SEP_REPEAT": "分隔符连续出现", "NO_SEP": "值内不含分隔符"}
_rows = []
for side, nm in (("merge", "合并"), ("split", "拆分")):
    for cls, v in PRB[side]["by_class"].items():
        _rows.append([nm, _LBL.get(cls, cls), str(v["n"]),
                      f"{v['roundtrip']}/{v['n']}", f"{v['pred_ok']}/{v['n']}"])
_m_tot = PRB["merge"]["total"]
_s_tot = PRB["split"]["total"]
CAP("表 7-106  边界取值的往返恒等与谓词保真（隔离库探针）")
TABLE(["方向", "边界情形", "例数", "往返恒等", "谓词保真"],
      _rows, [1.8, 5.0, 1.4, 2.2, 2.2],
      ["l", "l", "r", "c", "c"], font=8.4)
NOTE(f"合并方向合计 {_m_tot['roundtrip']}/{_m_tot['n']}（{_m_tot['roundtrip_rate']:.1f}%）往返恒等，"
     f"拆分方向 {_s_tot['roundtrip']}/{_s_tot['n']}（{_s_tot['roundtrip_rate']:.1f}%）。"
     "「谓词保真」指原列与段位表达式在相同行上的真值一致；它与往返恒等在本探针中逐例同步，"
     "说明「能切回来」与「判断正确」是同一件事的两面。")
P("结果把前提收窄到两条，且两条在两个方向上并不对称。"
  "合并方向：单个分量为 NULL 是致命的（2/2 失败）——SQLite 中 a ‖ CH ‖ NULL 整体为 NULL，"
  "于是两个段位都返回 NULL，原值不可恢复；而两分量同时为 NULL 反而无恙（1/1 通过），"
  "因为此时原值与重建值同时为 NULL。"
  "分隔符出现在第一分量内也是致命的（2/2 失败），因为 SUBSTR 截到第一个分隔符就停了；"
  "但分隔符出现在第二分量内无害（1/1 通过）——第二段取的是「第一个分隔符之后的全部内容」，"
  "不依赖分隔符个数。这一条不对称性此前没有被写出，它意味着二段合并的精确恢复条件是"
  "「第一分量不含分隔符」而非「值不含分隔符」。"
  "空串、首尾空白、重复值、多字节字符在两侧都通过：本构造不做 TRIM、不规范化编码。"
  "拆分方向：值内不含分隔符时无法重建（0/1），空串同样失败（0/1），其余情形通过。")
P("上述前提另用一套独立实现复核过（probe_encoding_bounds.py）。该脚本不沿用探针内置的 Python 切分逻辑，"
  "而是在真实 SQLite 表上按构造规则 UPDATE 出编码列，再把生产代码 m1_part.to_expr 渲染的表达式"
  "回填进 SELECT 执行，逐行比较「全部分量能否同时恢复」。"
  "六个可对应情形——常规值、单侧 NULL、双侧 NULL、空串、分隔符在第一分量内、分隔符在第二分量内——"
  "的判定方向与表 7-106 逐类一致，说明该表的结论不依赖探针自身的实现细节。"
  "复核还纠正了一处判据：若按单列判断，单个分量为 NULL 会被误判为可逆"
  "（NULL 反解仍是 NULL），实际是另一分量的信息随整体 NULL 一并丢失；"
  "本节的判据因此是整行全部分量同时恢复，而非逐列恢复。")
_eb3 = EBD["three_seg"]
P("同一脚本补上了此前未在库上验证的三段合并。旧版 to_expr 在 nseg ≥ 3 时把 idx > 0 一律当作末段，"
  "第 1 段与第 2 段会渲染出完全相同的表达式——这是此前「只验证两段」的真实原因，不是数据选择；"
  "该分支已改为逐段剥前缀。库上执行的结果是：三段全为常规值时 "
  f"{_eb3['S3a']['ok']}/{_eb3['S3a']['n']} 恢复，中段含分隔符时 "
  f"{_eb3['S3b']['ok']}/{_eb3['S3b']['n']}（静默错值），首段为 NULL 时 "
  f"{_eb3['S3c']['ok']}/{_eb3['S3c']['n']}（整体 NULL）。"
  "三段与两段共享同一条可逆条件，差别只在「非末段」的位置从一个增加到两个。"
  "需要限定的是：真实 BIRD 数据中不存在三段合并事件，这三行只证明实现分支正确，"
  "不构成三段场景下的效果证据。")
P("本轮把段数上限再往前推了一步（表 7-120）：在隔离库上按 "
  f"{len(EBV['config']['nsegs'])} 种段数（{', '.join(str(x) for x in EBV['config']['nsegs'])}）、"
  f"{len(EBV['config']['seps'])} 类分隔符、{len(EBV['config']['cases'])} 类边界值"
  f"构造 {len(EBV['cells'])} 个格、每格 {EBV['config']['rows_per_cell']} 行，"
  "以「整行全部分量同时恢复」为判据逐格执行生产代码 to_expr。"
  "这一轮暴露出一处真实实现缺陷，而非编码规则的固有限制：末段内含分隔符在两段时 "
  f"100% 可恢复，nseg ≥ 4 却掉到 {EBL['by_case']['sep_in_last']:.1f}%——"
  "旧实现对 nseg ≥ 3 的末段多切了一次 INSTR，把「第一个分隔符之后的全部内容」"
  "错当成「再切一次之后的剩余」。改为末段直接取剩余全部后，"
  f"该情形恢复率升至 {EBV['by_case']['sep_in_last']:.1f}%，"
  f"nseg ≥ 4 与 ≤ 3 的差异格由 {len(EBL['nseg_ge4_changes'])} 降为 "
  f"{len(EBV['nseg_ge4_changes'])}；v1 三段探针、留出库 M1 与 B4 公平基线的读数"
  "在修正前后逐字不变，因为全部 2,384 个真实 gold 段位表达式都是两段 flat 形态、零嵌套，"
  "修正只触及 nseg ≥ 3 分支。修正前后的两份产物（--legacy 与生产模式）均随材料包发布，"
  "可独立复现。仍需明确两类不可逆情形在修正后依然成立："
  f"NULL 分量（{EBV['by_case']['null_any']:.1f}%）与非末段内含分隔符"
  f"（{EBV['by_case']['sep_in_nonlast']:.1f}%）——前者是 SQLite 拼接整体为 NULL，"
  "后者是首个分隔符截断，二者都是分隔符编码的信息论限制，"
  "只能通过数据字典或审计日志等外部元数据规避，不能靠改表达式解决。")
_rows = []
for k, nm in (("merge_col", "合并（dev）"), ("split_col", "拆分（dev）"),
              ("merge_col_holdout", "合并（留出）"), ("split_col_holdout", "拆分（留出）")):
    v = DBD["cases"][k]
    _rows.append([nm, str(v["n_events_measured"]), str(v["n_events_with_loss"]),
                  f"{v['n_rows']:,}", str(v["lossy_rows"]),
                  f"{v['lossy_row_pct']:.2f}"])
_cases = list(DBD["cases"].values())
_tot_rows = sum(v["n_rows"] for v in _cases)
_tot_ev = sum(v["n_events_measured"] for v in _cases)
_tot_loss = sum(v["lossy_rows"] for v in _cases)
CAP("表 7-107  真实数据上的边界出现频率（逐事件全表扫描）")
TABLE(["方向（划分）", "实测事件", "有损失事件", "行数", "越界行", "越界占比 %"],
      _rows, [3.4, 2.2, 2.4, 2.6, 2.0, 2.4],
      ["l", "r", "r", "r", "r", "r"], font=8.4)
NOTE("合并方向查原始库的源列（IS NULL、值内含真实分隔符 CHAR(31)）；"
     "拆分方向查演化后库的分量列（单个分量为 NULL、分量内含分隔符）。"
     "「越界行」= 使往返不恒等的行数，按事件取并集而非简单相加。")
P(f"第二类证据是真实数据检查（data_bounds.py）：对 {_tot_ev} 个独立事件的源表/目标表做全表扫描，"
  f"共 {_tot_rows:,} 行，越界行 {_tot_loss} 行，占比 0.00%。"
  "也就是说，本章构造在真实 BIRD 数据上恰好是无损的：CHAR(31) 在 BIRD 中不出现，"
  "被合并的两列在所有事件上都不含 NULL。"
  "这个「零」必须正确解读：它不是「方法处理得好」，而是「构造从未产生这类行」。"
  "由此产生的限制是明确的——主实验的 610 条与留出库的 165 条，确实没有覆盖评审担心的"
  "NULL、值内分隔符、多段或嵌套编码等情形；这些情形下的行为只能由表 7-106 的隔离探针给出，"
  "且结论是「不可恢复」，不是「精度略降」。"
  "这也提示了一个具体的构造改进方向：若要让构造贴近真实历史演化，应显式引入含 NULL 的合并"
  "与含分隔符的取值，并在这种构造上重测——那时 c3r_* 与 C3mp 的绝对水平都会下降，"
  "而下降幅度本身就是有意义的指标。本章没有做这一步，列为待检验项。")
P("把上述边界写成方法前提，而不是留在结果讨论里，是为了让适用范围可被外部检查。"
  "段位表达式 SUBSTR(attr,1,INSTR(attr,CH)−1) 与 attr 的等价性，依赖两条前提："
  "（P1）合并后的取值中，任何一个分量都不为 NULL；"
  "（P2）除末段以外，任何一个分量内部都不含分隔符 CHAR(31)。"
  "两条前提都不成立时，表达式可执行而结果漂移，属于静默失败，正是本任务最难被下游发现的形态。"
  "需要强调的是：在本章扫描的 "
  f"{_tot_rows:,} 行中「未观察到」这两类情形，只说明该数据集在这一时点满足前提，"
  "并不构成「其他数据库、其他字段或未来数据同样满足」的保证——"
  "分隔符不与业务字符冲突是构造选择，不是数据的天然性质。"
  "要让方法不依赖 P1、P2，只有两条路：一是在编码侧引入转义（例如对分量内的分隔符做双写、"
  "对 NULL 使用哨兵值），这会把编码规则从「简单拼接」变成「拼接 + 转义」，"
  "并使表达式不再是一条 SUBSTR/INSTR 能表达的形式；"
  "二是从外部元数据（数据字典、DDL 注释、审计日志）取得段位边界，绕开从取值反推段位。"
  "本章两条都未实现，因此 P1、P2 在本章是方法前提，而非经验观察；"
  "凡把本章结论用于其他库，须先检验这两条前提是否成立。")
P("上述两条前提来自六类边界值，评审要求补足更完整的边界谱。本轮把边界取值扩到 "
  f"{len(RVS['config']['cases'])} 类、分隔符扩到 {len(RVS['config']['seps'])} 类"
  f"（CHAR(31)、竖线、分号、波浪号、多字节字符）、段数扩到 "
  f"{RVS['config']['nsegs'][0]}–{RVS['config']['nsegs'][-1]} 段，"
  f"共 {RVS_DET['cells']} 个格、每格 {RVS['config']['rows_per_cell']} 行，"
  "逐格执行生产代码 to_expr 并按「整行全部分量同时恢复」判成败（表 7-138）。"
  "新增的两类边界值得单独说明。")
_bc, _db_ = RVS["by_case"], RVS["detector_by_case"]
_NS = sorted((int(k), v) for k, v in RVS["by_nseg"].items())
# 网格情形的名称与反事实探针（_LBL）不同源，单独给标签，避免正文出现内部键名。
_GL = {"normal": "常规值", "null_any": "任一分量为 NULL",
       "empty_first": "首段为空串", "empty_last": "末段为空串",
       "sep_in_nonlast": "分隔符在非末段内", "sep_in_last": "分隔符在末段内",
       "whitespace": "首尾空白", "multibyte": "多字节字符",
       "escaped_sep": "分隔符被转义（双写）", "numeric_text": "数字存为文本"}
_lbl2 = dict(_LBL, **_GL)
_rows = []
for _c, _r in sorted(_bc.items(), key=lambda x: (-x[1], _db_.get(x[0], 0), x[0])):
    _nat = ("不可逆" if _r == 0 else ("可逆" if _r == 100 else "部分"))
    _note = "可逆但被拒" if (_r == 100 and _db_.get(_c, 0) > 0) else _nat
    _rows.append([_lbl2.get(_c, _c), str(RVS_DET["cells"] // len(_bc)),
                  f"{_r:.1f}", f"{_db_.get(_c, float('nan')):.1f}", _note])
CAP("表 7-138  十类边界取值在 250 格上的可恢复率与准入检测器标记率")
TABLE(["边界情形", "格数", "恢复率 %", "检测器标记 %", "性质"],
      _rows, [5.2, 1.4, 2.2, 2.6, 2.0],
      ["l", "r", "r", "r", "c"], font=8.4)
NOTE("恢复率 = 该情形全部格中「整行全部分量同时可恢复」的比例，按格计不按行计；"
     "检测器标记率 = 该情形被三条准入规则判为不可逆的比例。"
     "两者不必相等：标记率高于恢复率即为误拒（保守），低于则漏检（危险）。")
P("第一类新增边界是「分隔符被转义」（escaped_sep）：若编码侧对分量内的分隔符做双写转义，"
  f"直觉上应当可逆。实测恢复率为 {_bc['escaped_sep']:.1f}%——"
  "转义在编码侧做了，解码侧（SUBSTR/INSTR）仍按分隔符的字面出现次数切分，"
  "并不解析转义序列，于是第 k 段仍然被切在第一个分隔符处。"
  "这一结果修正了本节前面的表述：前文把「引入转义」列为突破 P1、P2 的一条路，"
  "但表 7-138 说明只改编码侧的转义是无效的，转义必须与解码表达式的转义解析成对实现；"
  "而一旦解码需要解析转义，表达式就不再是单个 SUBSTR/INSTR 能表达的形式，"
  "「可复制进提示的段位表达式」这一前提也随之失去。"
  "换言之，转义不是免费的补丁，它会把方法从「表达式注入」推回「元数据注入」。")
P(f"第二类是「数字存为文本」（numeric_text），恢复率 {_bc['numeric_text']:.1f}%："
  "拼接本身不改变字符内容，故不影响可逆性。这条阴性结果有用，"
  "它把「类型漂移」从编码的可逆性问题里摘了出去——类型漂移的危险不在切分，"
  "而在下游比较语义（'08' 与 8 的等值判断），属于另一层，"
  "由准入规则 R3 在构造前拦下（见下）。"
  f"另需报告一条与段数有关的阴性结果：按段数分组的恢复率在 "
  f"{_NS[0][0]}–{_NS[-1][0]} 段上" +
  (f"完全持平（各 {_NS[0][1]:.1f}%）" if len({v for _, v in _NS}) == 1 else
   "为 " + "、".join(f"{k} 段 {v:.1f}%" for k, v in _NS)) +
  "，说明在 2–6 段范围内，影响可逆性的只有边界情形的类别，"
  "段数只是放大同一比例的分母，不引入新的失败机制。")
P("知道前提之后，下一个问题是它能否在构造之前判定——只能事后检查的前提，"
  "对后来者没有用。为此实现了一个三条规则的准入检测器，全部只看原始列，不需要先建演化库："
  "（R1）候选分量列含 NULL；（R2）该列中真实分隔符出现次数不为 nseg − 1；"
  "（R3）参与合并的列中存在数值/日期类型（拼接后成为 TEXT，比较语义改变）。"
  "任一命中即拒绝该列对。"
  f"在表 7-138 同一批 {RVS_DET['cells']} 个格上，检测器给出 "
  f"tp = {RVS_DET['tp']}、fp = {RVS_DET['fp']}、fn = {RVS_DET['fn']}、"
  f"tn = {RVS_DET['tn']}（{RVS_DET['note']}）。")
P(f"其中 fn = {RVS_DET['fn']} 是关键：没有任何一格是「检测器放行、实际不可逆」，"
  "即该检测器不会把静默失败放进评测集。"
  f"fp = {RVS_DET['fp']} 全部来自「分隔符位于末段」这一类——"
  "它实际 100% 可恢复，但 R2 在列级无法知道某个分隔符落在末段还是非末段，"
  "只能对「值内出现分隔符」一律拒绝。这是保守方向的错误（误拒而非误放），"
  "代价是减少可选的构造，不会污染结论方向。"
  "因此这套规则可以作为构造前的闸门使用，而不只是事后的解释性说明。")
_rc = RVS["real_candidates"]["total"]
CAP("表 7-139  三规则准入检测器在 BIRD dev 全部候选列上的回退代价")
TABLE(["环节", "列数", "占候选列 %", "说明"],
      [["候选列", str(_rc["cand_cols"]), "100.0", "11 个库中可作合并分量的文本列"],
       ["R1 拒（含 NULL）", str(_rc["rej_null"]),
        f"{_rc['rej_null'] / _rc['cand_cols'] * 100:.1f}", "P1 不成立，静默错值"],
       ["R3 拒（类型漂移）", str(_rc["rej_type"]),
        f"{_rc['rej_type'] / _rc['cand_cols'] * 100:.1f}", "拼接后比较语义改变"],
       ["R2 拒（含分隔符）", str(_rc["rej_sep"]),
        f"{_rc['rej_sep'] / _rc['cand_cols'] * 100:.1f}", "真实数据中未出现，故为 0"],
       ["保留可用", str(_rc["kept_cols"]),
        f"{_rc['kept_cols'] / _rc['cand_cols'] * 100:.1f}", "可进入构造"],
       ["其中可配对", str(_rc["pairs"]),
        f"{_rc['pairs'] / _rc['cand_cols'] * 100:.1f}", "同表同类型两列，能构成一次合并"]],
      [3.8, 1.6, 2.6, 5.4], ["l", "r", "r", "l"], font=8.4)
NOTE("拒绝率按列计（先过列级准入再配对），不等于配对失败率；"
     "「可配对」远少于「保留可用」是因为合并要求同表、同类型、且两列都通过准入。")
P(f"把这个检测器用到真实数据上，得到一个不太舒服的读数（表 7-139）："
  f"BIRD dev 11 个库共 {_rc['cand_cols']} 个候选分量列，"
  f"按 P1–R3 拒绝 {_rc['rej_null'] + _rc['rej_type'] + _rc['rej_sep']} 列"
  f"（{RVS_REJ:.1f}%），保留 {_rc['kept_cols']} 列，"
  f"其中能配成同表同类型列对的只有 {_rc['pairs']} 对。"
  "这是本章最容易被审稿人追问的一处，必须把它的作用域讲准："
  "该数字约束的是「在全库上自动挑列、批量生成合并事件」这条路，"
  "而不是本章的事件集——本章合并档的 5 次事件是人工指定的列对，"
  "且已通过 R1、R2 检查（表 7-107 越界行 0）。"
  "反过来说，这个高拒绝率本身就是一条发现："
  "可逆性前提构成了真实演化的现实约束，真实历史中的列合并更可能发生在低基数、"
  "非空、非数值的文本列上——这正是本章构造所取的位置。"
  "本章没有用该检测器反向扩充事件集（把 33 对全部跑一遍），列为明确的下一步。")
P("至此可以把可逆性的条件写全，作为本章结论的第三条方法前提："
  "（P1）合并后任一分量不为 NULL；"
  "（P2）除末段外，任一分量内部不含分隔符；"
  "（P3）解码表达式的切分规则与编码规则逐字对应——若编码侧做了转义，解码侧必须同样解析。"
  f"本章满足 P1、P2（表 7-107 在其覆盖的 {_tot_rows:,} 行上零越界），"
  "P3 由 rewrite 与 to_expr 的同一实现保证；"
  "而表 7-138 说明一旦 P3 被打破（escaped_sep），恢复率即降为 0，"
  "其后果与 P1、P2 被打破相同，都是可执行而结果漂移。"
  "三条合起来的准确表述是：可逆性在本章是被显式声明、且可在构造前检测的前提，"
  "不是被验证为普遍成立的性质。"
  "仍未覆盖的部分应当直说：真实历史 schema 中含 NULL 的合并、含分隔符的取值、"
  "三段以上的真实演化，本章都没有——不是没有测，而是 BIRD 数据中不存在这类事件，"
  "需要外部元数据才能构造；在这些情形下本章只能给出「不可恢复」的定性结论（表 7-138），"
  "给不出「精度下降到多少」的定量结论。")

H("7.8  分析单位与外推范围", 2)
P("同一批结果可以在不同单位上检验，而单位决定结论能走多远。本章正文用的是条目级，"
  "即一条演化样本记一次；这对比较条件是合适的，但它也最容易高估证据的独立性——"
  "230 条合并档条目并不代表 230 次独立的 schema 演化。表 7-27 把四种可作检验的单位并列"
  "（基础题级已在表 7-23 给出，此处不重复）。")
P("五种口径的关系是：条目级最细，基础题级去掉「同一道题在多档重复」的重复计数，"
  "原子 op 级把「一个列级变更动作」当作一个单元，变更事件级再把同一次列重整的多个动作合并，"
  "数据库级最粗。对合并档，10 个原子 op 两两来自同一次合并，"
  "因此事件级只剩 5 个单元，与库数重合；拆分档的原子 op 与事件数相等，都是 5。")


_oplab = {("merge", "c3mp", "c3m"): "合并：C3mp − C3m",
          ("merge", "c3mp", "c_rw"): "合并：C3mp − c_rw",
          ("merge", "c3m", "c2"): "合并：C3m − C2",
          ("split", "c3mp", "c_rw"): "拆分：C3mp − c_rw",
          ("split", "c3m", "c2"): "拆分：C3m − C2",
          ("easy", "c3m", "c_rw"): "改名 easy：C3m − c_rw",
          ("medium", "c3m", "c_rw"): "改名 medium：C3m − c_rw",
          ("hard", "c3m", "c_rw"): "改名 hard：C3m − c_rw"}
_rows = []
for r in OPL["contrasts"]:
    L = r["levels"]
    _rows.append([_oplab[(r["group"], r["a"], r["b"])],
                  _n(f"{L['item']['diff_pt']:+.1f}"), _fp(L["item"]["p"]),
                  f"{L['op_atom']['n_units']} / {_fp(L['op_atom']['p_sign'])}",
                  f"{L['op_event']['n_units']} / {_fp(L['op_event']['p_sign'])}",
                  f"{L['db']['n_units']} / {_fp(L['db']['p_sign'])}"])
CAP("表 7-27  同一组对照在四个分析单位上的强度")
TABLE(["对照", "条目级 Δ (pt)", "条目级 p", "原子 op：n / p", "变更事件：n / p", "数据库：n / p"],
      _rows, [5.0, 2.2, 2.0, 2.4, 2.6, 2.4],
      ["l", "r", "r", "r", "r", "r"], font=8.2)
NOTE("p 为精确双侧符号检验（簇内条目等权）。条目级 p 用精确 McNemar，"
     "其不一致对数远大于单元数，故数值极小；簇级 p 的下限由单元数决定，"
     "5 个单元时最小可达 0.0625。")
P("表 7-27 的读法必须分层。条目级看，合并档 C3mp 相对 C3m 的 +77.0pt 对应 p ≈ 1e−53，"
  "拆分档 C3m 相对 C2 的 +60.6pt 对应 p ≈ 2e−19；这些数字说明的是"
  "「在当前这批评测条目上，差异稳定、且不是少数条目造成的」。"
  "但把单元换到变更事件后，合并与拆分各只有 5 个单元，5 个全同向的精确 p 仍是 0.0625，"
  "达不到 0.05。也就是说，本章的主结果在事件层是「方向完全一致、量级很大、但样本不足以"
  "做显著性判定」，而不是「已被统计确认的普遍规律」。")
P("四层并列还暴露了一处不对称：改名 hard 档在条目级 p = 0.021，在原子 op 级 p = 0.664，"
  "在事件级 p = 1.000。这不是矛盾，而是提示「增益集中在少数题目/少数变更」；"
  "本章据此把该档的结论定为「校正后不显著」，而不是「存在中等效应」。"
  "合并与拆分档不存在这种不对称：四个单位上方向全部一致，量级差异也很小"
  f"（合并 C3mp−C3m 在条目级 {[r for r in OPL['contrasts'] if r['label'].startswith('合并：段位 vs 列级')][0]['levels']['item']['diff_pt']:+.1f}pt、"
  f"事件级 {[r for r in OPL['contrasts'] if r['label'].startswith('合并：段位 vs 列级')][0]['levels']['op_event']['mean_cluster_diff_pt']:+.1f}pt）。")
_easy = [r for r in OPL["contrasts"] if r["group"] == "easy"][0]["levels"]
P("改名 easy 档此前一律写作「未判定」，表 7-27 提示这个措辞需要更精确。"
  f"该档在条目级 p = {_fp(_easy['item']['p'])}、区间跨 0；但在原子 op 级是 "
  f"{_easy['op_atom']['n_positive']} 正 {_easy['op_atom']['n_negative']} 负"
  f"（p = {_fp(_easy['op_atom']['p_sign'])}），在事件级是 "
  f"{_easy['op_event']['n_positive']} 正 {_easy['op_event']['n_negative']} 负"
  f"（p = {_fp(_easy['op_event']['p_sign'])}）。两者并不矛盾：粗单元说明"
  "「多数变更上直接改写略优」这一方向是一致的，条目级说明总体差值太小、无法排除 0。"
  "正确的读法是「方向可能为负、幅度不可判定」，而不是「两者没有差异」——"
  "后者会把条目级的不显著误读成等价，而等价需要整体落进界值，本章的等价检验并未通过。")
P("因此本章的主张按三层收窄：第一层，在本评测集的 610 条上，段位谱系与列级谱系、"
  "与直接改写之间的差异是大的、稳定的、方向一致的；第二层，在独立变更事件层，"
  "单档各只有 5 次事件、最小可达 p = 0.0625，把两类编码改变按同一口径池化后为 10 次事件、"
  "p = 0.0020，但库级仍只有 5 个库（p = 0.0625），故只声称方向一致、不声称跨库显著；"
  "第三层，跨库、跨真实历史演化、跨更多列段类型的外推，本章不提供证据。"
  "评审提出「样本量究竟支持什么范围」，答案就是这三层，"
  "而不是「610 条足够」或「5 个库不足」这样的单一判断。")
_POOL = OPL["pooled"][0]
P("关于第二层，需要一处限定，否则「事件层达不到 0.05」会被读成「事件层无法判定」。"
  "合并档与拆分档施加的是同一批 5 个库上的两类相反的编码改变"
  "（多列→一列、一列→多列）；若统一到同一对照口径（C3mp − c_rw），"
  f"可把 {_POOL['n_events']} 次不同的列重整事件放在一起做符号检验。表 7-28 并列两个口径。")
CAP("表 7-28  两类编码改变池化后的事件级与库级")
_prows = [["事件级（10 次列重整）", str(_POOL["n_events"]),
           _n(f"{_POOL['event_level']['mean_diff_pt']:+.1f}"),
           f"{_POOL['event_level']['n_positive']} / {_POOL['event_level']['n_negative']} / "
           f"{_POOL['event_level']['n_tie']}",
           _fp(_POOL["event_level"]["p_sign"]), _fp(_POOL["event_level"]["p_min"])],
          ["库级（5 个库）", str(_POOL["n_db"]), "—",
           f"{_POOL['db_level']['n_positive']} / {_POOL['db_level']['n_negative']} / "
           f"{_POOL['db_level']['n_tie']}",
           _fp(_POOL["db_level"]["p_sign"]), _fp(_POOL["db_level"]["p_min"])]]
TABLE(["口径", "单元数", "平均差值 (pt)", "正 / 负 / 平", "精确符号检验 p", "最小可达 p"],
      _prows, [4.4, 1.8, 2.4, 2.4, 2.8, 2.2],
      ["l", "r", "r", "c", "r", "r"], font=8.5)
NOTE("两行必须同时读：事件级 10 个单元全同向，p = 0.0020；但库仍只有 5 个，p = 0.0625。"
     "10 次事件不等于 10 个库——前者支持「两类相反编码改变上方向一致」，"
     "后者不支持「跨库泛化」。逐库差值为 "
     + "、".join(f"{k} {v:+.1f}" for k, v in _POOL["db_level"]["per_db_pt"].items()) + "。")
P("这一合并可辩护，但边界同样要写明：合并档与拆分档共用同一批 5 个库，"
  "事件虽不同、库内相关仍存在，因此库级 p 才是外部效度的诚实读数；"
  "合并档的粒度消融（C3mp − C3m）只发生在合并档，无法与拆分档池化，"
  "该对照仍停在 5 事件、p = 0.0625；池化不引入新样本，"
  "只是把已有的 10 次事件按同一口径合并计数。")

H("7.9  与近邻工作的边界", 2)
P("本章的创新性是增量性的，必须与三项近邻工作正面划界，且不能把其中任何一项的工作"
  "算作本章的首创。三项分别是 EvoSchema（PVLDB 2025）、How Far Can They Map?（ICDEW 2026）"
  "与 PRISM/PRISM++（VLDB J 2011–2015）。表 7-29 按六个维度并列。")
CAP("表 7-29  与近邻工作的边界")
TABLE(["工作", "任务输入", "历史 SQL 的角色", "schema 变化与数据编码", "输出形态", "评价目标"],
      [["EvoSchema (PVLDB 2025)",
        "BIRD 的库 + 问题，施加扰动",
        "不作为证据：gold SQL 随扰动重写，模型看到的是新 schema + 问题",
        "10 类：列级增/删/改名/拆分/合并；表级增/删/改名/拆分/合并。"
        "合并按固定顺序与分隔符拼接（名 + 空格 + 姓），拆分按语义分解（地址 → 街道/城市/州/邮编）",
        "扰动后的 schema 与 gold SQL",
        "模型鲁棒性与训练增强"],
       ["How Far Can They Map? (ICDEW 2026)",
        "跨库的 SQL 映射（SQL-Exchange）",
        "作为待映射的输入",
        "两个不同库之间的 schema 差异（非版本演化）",
        "映射后的 SQL",
        "哪些特征导致映射失败（分析性）"],
       ["PRISM/PRISM++ (VLDB J 2011–2015)",
        "旧 DDL + 新 DDL + 旧查询（映射由演化算子给出或推导）",
        "作为待改写的输入；目标是查询等价",
        "结构演化：增删表、增删列、改名、拆分/合并；不涉及取值编码改变",
        "关系代数层面的等价查询（纯确定性改写，无模型）",
        "改写规则的正确性与可判定性"],
       ["本章",
        "同一库的旧版本问题 + 旧 SQL",
        "作为语义证据：旧 SQL 指明该列的用途",
        "列合并为带分隔符的单列、列拆分为两列（取值编码改变）",
        "可执行的段位表达式（SUBSTR/INSTR）",
        "粒度信息是否足以避免「可执行但语义错误」"]],
      [3.2, 3.0, 3.2, 4.0, 2.8, 3.0],
      ["l", "l", "l", "l", "l", "l"], font=7.6)
P("五处差别需要点明，否则容易被读成同一件事的重复。第一，变异类型有重叠，但任务定义不同。"
  "EvoSchema 的列级 Split 与 Merge 与本章处理的是同一类结构变化"
  "（一对多、多对一），这一点必须承认；区别在于 EvoSchema 固定自然语言问题、"
  "把 gold SQL 重写为新 schema 下的规范答案，模型看不到旧 SQL，评的是面对新 schema 的鲁棒性。"
  "本章把旧 SQL 当作语义证据，要求生成与旧查询结果等价的迁移，"
  "因此必须回答「旧值如何在新编码下被重建」——段位表达式正是为回答这一点而引入的。"
  "第二，EvoSchema 的合并/拆分以列的多对一或一对多映射为终点，"
  "本章进一步要求把「合并列的哪一段」写成可执行表达式（SUBSTR/INSTR）："
  "在带分隔符的单列里，列名对齐即使完全正确，投影仍会取回整串而非原来那一列的值，"
  "这是本章的机制命题，也是 C3m 在合并档反低于 C2 的原因。"
  "第三，How Far Can They Map? 研究两个不同库之间的映射，"
  "本章研究同一库的旧版本到新版本，且要求输出可直接执行的表达式而不只是列对应。")
P(f"第四，也是本轮必须当场补上的一条：EvoSchema 的列合并样本如何处理合并后的 SQL。"
  f"其公开集共 84 条，按（库, 问题）去重后 81 条（合并 {ES_MG}、拆分 {ES_SP}）；"
  f"合并档 {ES_MG_DB.get('student_club', 0)} 条来自 student_club 的 "
  "first_name + last_name → full_name，"
  f"{ES_MG_DB.get('california_schools', 0)} 条来自 california_schools 的 "
  "Address ← Street/City/State/Zip 与 AdmEmail ← AdmEmail1/AdmEmail2。"
  "处理方式是「整列替换」：原 SQL 中并列的多个来源列被替换为一个合并列"
  "（SELECT Street, City, State, Zip → SELECT Address），"
  "WHERE 中的多列等值条件改为对合并串的等值比较"
  "（first_name = 'Sacha' AND last_name = 'Harrison' → full_name = 'Sacha Harrison'，"
  "即以单空格拼接、顺序为先名后姓）。"
  f"{ES_MG} 条合并样本中，对合并列使用 SUBSTR/INSTR 取段的为 {ES_SEG_TGT} 条；"
  f"新 gold SQL 里出现的 SUBSTR/INSTR 共 {ES_SEG_ANY} 处，作用于 event_date 与 expense_date，"
  "与合并列无关。也就是说，在该评价口径下「投影整串 Address」就是正确答案，"
  "因为 gold 已被同样地重写；合并本身被当作一次合法的信息抽象，"
  "而不是一次需要在新编码上重建原值的迁移。")
P("这正是两种任务的定义差别，而不是同一件事的两种实现。把本章的 C3m（列级谱系）"
  "放到该口径下评价，它同样会输出 SELECT Address，分数不会因此下降；"
  "但按本章口径（与旧查询的结果集等价），同一句 SQL 把 4 列压成 1 列，"
  "属于「可执行而语义错误」，正是 SILENT 类错误的典型形态。"
  "反过来说，本章的 C3mp 在其样本上既无必要、也无从评价——"
  "他们的合并列是「语义整体」（地址、全名），不承载「第 k 段对应旧列 j」这一信息，"
  "因此 7.9.1 的外部对照只评列引用，不评段位。")
P("第五，PRISM/PRISM++ 一路的经典工作已经把「schema 演化下的查询重写」形式化过："
  "它们以演化算子及其逆算子把旧查询逐算子改写成等价查询，并讨论改写规则的正确性条件。"
  "本章与它们的分工是：那一类工作假定映射已知、且变化发生在关系代数结构上，"
  "目标是查询等价的可判定改写；本章的变量是「映射需要自动推断」以及"
  "「取值编码在一列内部发生变化」——后者使标识符替换本身就需要构造表达式，"
  "也就使「证据给到哪一层」成为一个经验问题。"
  "本章 7.4.2 的确定性改写基线可以看作这类工作在 LLM 设置下的最小化实现，"
  "它的能力边界（需重写谓词结构而非替换标识符的迁移未测、投影星号在本构造下无法忠实评价）"
  "同时也界定了本章结论的适用范围。")
P(f"还有一层必须写明：该评测集涉及的 {len(ES_DBS)} 个库"
  + (f"全部与本章 dev 的 11 个库重叠" if ES_OVL == len(ES_DBS)
     else f"中有 {ES_OVL} 个与本章 dev 的 11 个库重叠")
  + "，两者又都以 BIRD 为底库。因此 7.9.1 的对照是「换构造者、不换库」的外部效度检查，"
  "它能排除「自建样本过于宽松」这一质疑，但不能替代 7.7.4 的留出库确认；"
  "反过来，本章也不能把「在 BIRD 上做列合并」本身当作新意。")
P("据此，本章不主张「首次研究 schema 演化下的 NL2SQL」，"
  "不主张「首次评测列级合并/拆分」，也不主张「首次做跨 schema SQL 映射」。"
  "可以主张的是：在历史 SQL 可作为语义证据的迁移场景下，"
  "系统比较列级谱系与段位表达式谱系，并给出一份关于"
  "「何种粒度的信息才能避免可执行但语义错误的 SQL」的受控证据。"
  "这是一个中等、偏增量的贡献，其价值在机制而非首创。")

H("7.9.1  在 EvoSchema 列级样本上的外部对照", 3)
P("划界之外还有一个更直接的问题：本章的机制结论是否只在自建评测集上成立。"
  "为此本节在其公开的列级合并/拆分评测集（bird_split_or_merge_columns.json）上"
  "重跑了条件阶梯。必须先说明这份数据能回答什么、不能回答什么，"
  "否则容易被读成本章结果的独立复现。")
PL(f"① 样本共 {ES['n_raw']} 条（列级合并 47、拆分 37），按（库, 问题）去重后 "
   f"{ES['n_uniq']} 条参与（合并 {ES_MG}、拆分 {ES_SP}）。")
PL("② 评测集只给新旧 DDL 与 gold SQL，不提供施加演化后的数据库文件，也没有数据迁移规则。"
   "因此无法执行 SQL 求 EX，本节只评列引用正确性（EvoSchema 自有的 Column Match 指标）。")
PL(f"③ 合并档有 {ES_MG_EMPTY} 条样本的 column_changes 字段为空（全部为 student_club 的 "
   "first_name + last_name → full_name），映射改由新旧 DDL 的列集合差恢复；"
   "3 条无法恢复，单列为 unknown。")
PL("④ 部分 new_gold_sql 与旧 SQL 的数值字面量与 WHERE 条件不一致（疑似重新生成而非机械适配），"
   "故 gold 只用于列集合参照，不用于语义等价判断。")
_es = EVO["summary"]
_ename = {"c1": "C1 仅新 schema + 问题", "c2": "C2 旧 SQL + 新 schema",
          "c3m": "C3m 旧 SQL + 列级谱系"}
CAP("表 7-30  在 EvoSchema 列级样本上的列引用指标（n = 81）")
TABLE(["条件", "列 F1", "列集合完全正确", "仍引用已删旧列"],
      [[_ename[c],
        f"{_es[c]['mean_f1']:.3f}",
        f"{_es[c]['exact_cols_rate'] * 100:.1f}%",
        f"{_es[c]['uses_deleted_old_col_rate'] * 100:.1f}%"] for c in ("c1", "c2", "c3m")],
      [6.4, 2.6, 3.6, 3.6], ["l", "r", "r", "r"], font=8.8)
_ek = EVO["by_kind"]
_et = EVO["tests"]
_erows = []
for _kind, _lab2 in (("merge", "列级合并"), ("split", "列级拆分")):
    if _kind not in _ek.get("c2", {}):
        continue
    _a, _b = _ek["c2"][_kind], _ek["c3m"][_kind]
    _erows.append([_lab2, str(_a["n"]), f"{_a['mean_f1']:.3f}", f"{_b['mean_f1']:.3f}",
                   _n(f"{(_b['mean_f1'] - _a['mean_f1']) * 100:+.1f}")])
CAP("表 7-31  按变异类型分层：C2 与 C3m 的列 F1")
TABLE(["变异类型", "n", "C2 列 F1", "C3m 列 F1", "差值 (pt)"],
      _erows, [4.0, 2.0, 3.0, 3.0, 3.0], ["l", "r", "r", "r", "r"], font=8.8)
_pc = EVO["paired_c2_vs_c3m"]
_dltxt = _n(f"{_pc['delta_pt']:+.1f}")
NOTE(f"配对比较：{_pc['n']} 条中 C3m 更好 {_pc['c3m_better']} 条、C2 更好 {_pc['c2_better']} 条、"
     f"持平 {_pc['tie']} 条，平均差 {_dltxt}pt，"
     f"符号检验 p = {_fp(_et['sign_c2_vs_c3m']['p'])}；"
     f"列集合完全正确的配对为 {_et['exact_cols_paired']['c3m_only']} : "
     f"{_et['exact_cols_paired']['c2_only']}（p = {_fp(_et['exact_cols_paired']['p'])}）。"
     "合并档 6 : 0（p = 0.031）达到显著，拆分档 7 : 2（p = 0.18）方向一致但样本不足。")
P("这组结果支持一件事：列级谱系在他人构造的列级合并/拆分样本上也提升列引用正确率，"
  "且提升集中在合并档；它同时给出了一个反向的读数——"
  "C1 到 C2 的差距（"
  f"{_es['c1']['mean_f1']:.3f} → {_es['c2']['mean_f1']:.3f}）说明旧 SQL 本身提供信息，"
  "而 C2 到 C3m 的差距说明「只给旧 SQL 不够，还要说清它对应到新 schema 的哪一列」。")
P("但这组数字的独立单元数必须另算，否则会高估证据。表 7-31 的分层口径按条目计，"
  "而同一份映射会被几十道题重复引用：合并档 41 条中，38 条来自同一个映射"
  "（student_club 的 first_name + last_name → full_name），占该档 92.7%。"
  "本节的配对符号检验若以条目为单位，等价于把同一个映射数了 38 次。"
  "因此按「库 × 被演化列集合」重划事件单元后重新统计（表 7-109）。")
_em = EVEK["merge"]
_esp = EVEK["split"]
_euk = EVEK["unknown"]
_eu = EVE["paired_by_unit"]
CAP("表 7-109  按独立映射单元重算的外部对照（C3m vs C2）")
TABLE(["单元口径", "单元数", "条目数", "最大单元占比", "正 : 负 : 平", "符号检验 p"],
      [["条目（表 7-31 口径）", str(_eu["item|c3m-c2"]["n"]), str(_eu["item|c3m-c2"]["n"]),
        "—",
        f"{_eu['item|c3m-c2']['pos']} : {_eu['item|c3m-c2']['neg']} : "
        f"{_eu['item|c3m-c2']['tie']}", _fp(_eu["item|c3m-c2"]["p"])],
       ["合并：按映射", str(_em["events"]), str(_em["items"]),
        f"{_em['max_event_share_pct']:.1f}%",
        f"{_em['pos']} : {_em['neg']} : {_em['tie']}", _fp(_em["p"])],
       ["拆分：按映射", str(_esp["events"]), str(_esp["items"]),
        f"{_esp['max_event_share_pct']:.1f}%",
        f"{_esp['pos']} : {_esp['neg']} : {_esp['tie']}", _fp(_esp["p"])],
       ["映射无法恢复（单列）", str(_euk["events"]), str(_euk["items"]),
        "—", f"{_euk['pos']} : {_euk['neg']} : {_euk['tie']}", _fp(_euk["p"])],
       ["全部映射单元", str(_eu["event"]["n"]), str(EVE["n_items"]), "—",
        f"{_eu['event']['pos']} : {_eu['event']['neg']} : {_eu['event']['tie']}",
        _fp(_eu["event"]["p"])],
       ["按数据库", str(_eu["db"]["n"]), str(EVE["n_items"]), "—",
        f"{_eu['db']['pos']} : {_eu['db']['neg']} : {_eu['db']['tie']}",
        _fp(_eu["db"]["p"])]],
      [3.6, 1.5, 1.4, 2.2, 2.6, 2.2],
      ["l", "r", "r", "r", "c", "r"], font=8.4)
NOTE("映射单元按「库 + 被演化的列集合」划分：同一次列重整下的所有题目共用一个映射，"
     "算作一个单元；不同表上列名不同，即使是同类操作也分属不同单元。"
     "「最大单元占比」= 该类型中最大单一映射的条目数 / 该类型条目数。"
     "符号检验比较单元内的 C3m 与 C2 列 F1 均值，单元内打平计为 0。")
P(f"把口径换过来，结论的强度立刻下降一个档次。合并档只有 {_em['events']} 个独立映射，"
  f"其中最大一个占 {_em['max_event_share_pct']:.1f}%，单元级 {_em['pos']} 正 {_em['neg']} 负 "
  f"{_em['tie']} 平，精确符号检验 p = {_em['p']:.2f}——"
  "在 3 个单元上，即使全部同向，双侧最小 p 也只能到 0.25，本组本质上无法做显著性判定。"
  f"拆分档的映射更分散（{_esp['events']} 个单元，最大占 {_esp['max_event_share_pct']:.1f}%），"
  f"但方向并不一致（{_esp['pos']} : {_esp['neg']} : {_esp['tie']}，p = {_esp['p']:.2f}）。"
  f"全部 {_eu['event']['n']} 个映射单元合起来是 "
  f"{_eu['event']['pos']} : {_eu['event']['neg']} : {_eu['event']['tie']}（p = "
  f"{_eu['event']['p']:.3f}），{_eu['db']['n']} 个库是 "
  f"{_eu['db']['pos']} : {_eu['db']['neg']} : {_eu['db']['tie']}（p = {_eu['db']['p']:.3f}），"
  "均未达 0.05。")
P("因此本节对条目级 p = 0.0018 的解读必须修正：它度量的是「在这批题目上 C3m 的列引用是否更好」，"
  "而不是「列级谱系在列合并上已被独立复现」。正确的表述是——"
  "在该构造者的公开样本上，列级谱系对列引用有正向改善，方向与其自报的结论一致；"
  "但合并方向的独立映射只有 3 个、其中一个覆盖 92.7% 的条目，"
  "事件级与库级都不支持显著性，本节据此外推的力度应低于条目级 p 值给人的印象。"
  "这也意味着「外部对照」四个字要打折扣：它换掉了构造者，但没有换掉库"
  "（8 个库全部与本章 dev 重叠），也没有增加独立映射的数量。")
P("它不支持三件事，需要一并声明。第一，段位表达式在他们样本上是否有效尚未测试——"
  "该评测集的合并是把多列并成一列、不涉及「一列内多段编码」，没有段位信息可注入；"
  "这恰好印证两者场景不同，但也意味着本节不能替代对 C3mp 的外部确认。"
  "第二，本节评的是列引用而非执行结果一致，两项正确率数字与本章 EX 不可直接比较。"
  "第三，81 条中仍有 3 条映射无法从 DDL 差异恢复，说明「从 DDL 自动恢复列级谱系」"
  "在真实样本上并非总是可行，这与本章 M1 的恢复率结论一致。")
P("最后补齐两处口径，避免表 7-31 的 n 与 7.9.1 开头的 n 看起来不一致。"
  f"其一，{EVE['n_items']} 条 = 可恢复映射的 41 条合并 + 37 条拆分 + 3 条映射无法恢复"
  "（column_changes 为空且 DDL 差也无法定位目标列）；表 7-31 只列前两类（41 + 37 = 78），"
  "3 条 unknown 单列在 EVO['by_kind'] 中，其列 F1 在 C2 与 C3m 都是 0.667（三条里两条完全对、"
  "一条完全错），既不加分也不减分，故未进入分层表。"
  "其二，这 3 条仍计入表 7-30 与配对检验的分母（81 条），它们对 C3m − C2 的贡献是 0，"
  "把它们剔除会把条目级 p 从 0.0018 略微推高，不改变方向，也不改变上述对独立单元数的判断。")

H("7.10  本章结论", 2)
P("表 7-32 汇总五个研究问题与四组补充证据的答案。")
CAP("表 7-32  研究问题、结论与证据")
TABLE(["RQ", "结论", "关键数字", "证据"],
      [["RQ1", "同粒度下自动谱系与 oracle 在本评测集上未观察到可测差异；"
               "不据此主张普遍等价",
        "0–2 条/276 不一致；值改写档 M1 54.5%", "表 7-14、表 7-15"],
       ["RQ2", "谱系改善端到端迁移；列级条件在合并档反低于旧 SQL",
        "+22.5 / +69.6 / +60.6 pt；C3m 2.6% < C2 10.0%", "表 7-16、表 7-18"],
       ["RQ3", "合并与拆分档谱系大幅优于直接改写（Holm 校正后仍显著）；"
               "改名三档均不能给出方向性的幅度结论",
        "合并 158:3、拆分 71:0（均 p<1e-4）；easy 与 medium 未判定"
        "（easy 在变更单元上方向一致为负）；hard 未校正 p=0.021、Holm 后 0.064",
        "表 7-19、表 7-8"],
       ["RQ4", "主效应在两个基座上可复现，幅度接近",
        "落差 +77.0 vs +79.6，相差 2.6pt", "表 7-20、表 7-21"],
       ["RQ5", "增益主要来自证据表示而非执行器；给定映射时确定性改写不弱于 LLM，"
               "且 LLM 在 dev 上没有独占正确条目。两系数均为本构造下的配对差值",
        "证据效应 +99.1pt（合并 dev）／+64.4pt（合并留出）／0pt（拆分）；"
        "执行器效应 +20.4／+22.1／+10.9／+20.3pt；规则侧 dev 零失败，"
        "留出合并档 20 条因漏抽映射报错",
        "表 7-101、表 7-103、表 7-104、表 7-111"],
       ["留出", "在 7 个未参与开发的库上方向重现；绝对水平受库难度影响",
        "合并 69.3% vs 7.9%；拆分 79.7% vs 42.2%", "表 7-26"],
       ["边界", "构造在真实数据上恰好无损，边界情形从未被触发；一旦出现则不可恢复",
        "21.5 万行越界 0 行；探针 NULL_SIB 0/2、SEP_IN_A 0/2", "表 7-106、表 7-107"],
       ["机制", "静默错误集中在粒度层而非列指称层",
        "98.7% 引用新列，2.2% 尝试切段", "7.6.1"]],
      [1.4, 7.6, 5.4, 2.6], ["c", "l", "l", "l"])
P("综观全章，本章能够支持的结论是：当 schema 演化把多列合并为一列时，"
  "列级对齐不足以保证历史 SQL 的可用性；把谱系下沉到段位表达式，"
  "在本构造评测集上带来数十个百分点的改善，且该改善不依赖 oracle 谱系、"
  "在两个基座上方向一致。结论的强度必须与统计单元一致：按 7.8 节的分析单位分解，"
  "条目级（230 / 104 条）差异极大且高度稳定；"
  "把合并与拆分两档的 10 次列重整按同一口径池化后，事件级 10/10 同向（p = 0.0020），"
  "但库级只有 5 个单元（p = 0.0625），库的效应与事件的效应仍无法分离；"
  "这 10 次事件属于主实验 610 条这一套构造（合并 5 + 拆分 5），"
  "与 7.15 节的扩充事件集是两套独立构造，不可相加、也不互为分母——"
  f"扩充集的合并档另有 {_EM['events']} 个构造事件、{_EM['events_ge3']} 个合格事件，"
  f"事件级 {ACK_E['event_level_all']['pos']} 正 "
  f"{ACK_E['event_level_all']['neg']} 负、"
  f"{ACK_E['event_level_all']['tie']} 平，结论方向相同而强度更高。"
  "两处数字并列出现时必须带各自的构造名，否则读者会以为是同一批事件的两次统计；"
  "这是本轮修订中口径统一的首要一条；"
  "留出库确认实验（7.7.4）在 7 个未参与开发的库上重现了方向，"
  "样本量（165 条）只够作旁证，不足以支持跨库泛化。"
  "因此本章的结论应读作「在这 5 次列合并与 5 次列拆分上，粒度信息带来量级显著的改善」，"
  "而不是「已确认的普遍规律」，更不能延伸为跨库泛化。"
  "本章不能支持的结论包括：自动谱系可普遍替代 oracle；"
  "显式谱系在一切演化类型上都优于让模型自行改写；"
  "改名场景下两者存在可判定的优劣或等价；"
  "方法可普遍迁移到未见数据库或真实历史 schema 版本（7.7.4 只给出 7 库 165 条的旁证）；"
  "以及检索能改善该任务——"
  "本文没有实现检索模块，谱系是整块注入提示的；"
  f"唯一的检索式探针 {E2['n']} 条、{E2['n_db']} 个数据库，为阴性对照（p = {E2['p']:.2f}），"
  "不足以支撑检索方向的任何主张。第 21 轮起，全文标题、摘要、关键词与贡献点"
  "已统一收敛为「schema 演化的谱系粒度实证研究」，检索只作背景提及，"
  "不再作为研究问题出现。")
P("就创新性而言，本章的位置在 7.9 节已经界定：不主张首次研究 schema 演化下的 NL2SQL，"
  "也不主张首次处理合并/拆分或跨 schema 映射；可主张的是"
  "「在历史 SQL 作为语义证据的前提下，何种粒度的谱系信息足以避免可执行但语义错误的 SQL」。"
  "这是增量性贡献，其证据强度受 7.7.3 的七点边界与 7.8 的单元限制约束。")
P("未完成项需要在摘要、结论与投稿信中一致披露，共六项，每一项都能在正文找到对应的量化说明。"
  "其中三项仍未做，三项属构造与评价范围的边界——后者不是工作量问题，而是本构造无法产生该类情形，"
  "故只能界定、不能实测，二者都必须披露："
  "① 三段以上合并尚无真实演化数据——to_expr 已支持任意段数，"
  "并在隔离库上补齐验证（三段常规值全部恢复，中段含分隔符与首段为 NULL 各 1 例失败，见 7.7.5），"
  "但真实 BIRD 数据中不存在三段合并事件，故不写进结论；"
  "② 需要重写谓词结构而非替换标识符的迁移（口径改变、单位换算、等值谓词变区间）未测，"
  "因此不能声称语言模型在该类情形上有优势——这是 7.4.3 的第三条明确留下的空白；"
  "③ 取值编码的可逆条件已逐类界定，不再是未知：NULL 分量与「非末段内含分隔符」两类不可恢复，"
  "空串、末段内含分隔符、首尾空白、多字节字符四类均可恢复（表 7-106，并经独立实现复核）。"
  f"真实数据扫描（合并侧 {EBM['events']} 个事件、{EBM['rows']:,} 行）显示前两类出现 0 次，"
  f"空串出现 {EBM['empty']} 次且全部可恢复（表 7-107）——"
  "因此「零越界」应读作构造未触发，而非方法处理好；"
  "探针每类只有 1–3 例，只作机制反例，不能估计真实数据中的失败比例。"
  "主实验的绝对水平因此只适用于「各分量非空、且仅末段允许含分隔符」这一前提成立的取值分布；"
  "④ 采用单作者编码而非双人独立编码：gold SQL、演化算子与段位表达式均由同一条"
  "替换规则确定性生成（2,384 个表达式），不存在需要人工判读的标注单元，"
  "因此 Cohen's kappa 无适用对象，本章不报告 kappa，也不虚构该数值。"
  "替代做法是把「第二位标注者」的功能拆成可核查的等效独立复核（表 7-122）："
  "判定器全严规则独立重执行、未见库 M1 直接表达式评估、编码边界独立实现、"
  "修正前后双产物对照、四层事件映射审计、非 LLM 公平基线，以及核心条件的重复运行。"
  "但必须写明：这些复核检验的是判定口径、抽取实现与结论稳定性，"
  "不等同于第二位标注者对同一批构造的独立重编码；"
  "gold 的正确性仍依赖构造期自检，此项限制不因复核清单而消除；"
  "⑤ 检索模块未实现：谱系是整块注入提示的；"
  f"唯一的检索式探针 {E2['n']} 条、{E2['n_db']} 个数据库，为阴性对照"
  f"（C4 {E2['c4_rate']:.1f}% 对 C5 {E2['c5_rate']:.1f}%，p = {E2['p']:.2f}，见 7.7.3 ⑦），"
  "只能说明「一个检索到的过期示例不足以救回迁移」，"
  "不能评价检索方法，故本文不以检索为研究问题，标题、摘要、关键词与贡献点均不出现检索或 RAG；"
  "⑥ 含投影星号的题目在本构造下无法被忠实评价（见 7.4.4）：gold 由机械替换生成，"
  "而 * 不是列引用，故 gold 保留 * 并在新库上返回新列集合，与旧语义不等价；"
  "若按旧表列清单展开 *，结果又与 gold 不一致。两种读法都不能作为正确答案，"
  "故本样本中此类题目恰为 0 条（表 7-111）虽使已报告数字不受影响，"
  "但结论不能外推到含 SELECT * 的查询。"
  "另有六条此前列为未完成、本版已闭环，需要一并说明："
  "按数据库隔离的留出确认实验已补做（BIRD train 划分 7 个未参与开发的库、165 条，"
  "总体方向重现（合并档事件 6/7 为正、1/7 为负），但样本量只支持「方向可迁移」的旁证，见 7.7.4）；"
  "确定性 SQL 重写基线已补做并按审稿意见修正（7.4.2、7.4.4）：初版把 COUNT(*) 误判为不可改写，"
  "修正后规则侧在 dev 上零失败，同一份谱系下比 LLM 高 20.4pt（合并）与 22.1pt（拆分），"
  "留出库为 +10.9pt 与 +20.3pt；星号展开通道已实现并自检通过，"
  "但主评测未含此类题目，故属评测范围限制而非能力验证；"
  "EvoSchema 外部对照已按独立映射单元重报（7.9.1、表 7-109，合并档只有 3 个映射，p = 0.5）；"
  "判定器宽严对样本量的影响已量化（7.1.7，9 条/0.4%，主结论符号不变），"
  "含 ORDER BY 子集的保序准确率已单列（表 7-108）；"
  "经典 schema evolution 查询重写相关工作已补入（7.1.1、表 7-29 的 PRISM/PRISM++ 行）；"
  "公开材料包已整理（materials-v4），含主实验 4,031 条与留出库 165 条的逐题清单、"
  "演化规则与全部阈值、提示模板、4,490 条逐题输出、EvoSchema 外部对照的 81 条明细、规则基线四档逐题明细与星号审计、"
  "全部统计产物、全部生成脚本本体与运行环境说明（公开材料包/，MANIFEST.json 记录各文件与脚本的 sha1，ENVIRONMENT.md 给出依赖与逐产物复现命令）。"
  "两条近邻引文已按出版社记录核实：EvoSchema, PVLDB 18(10): 3655–3668, 2025，"
  "doi:10.14778/3748191.3748222，其十类扰动含列级 Split/Merge（公开评测集 84 条，"
  "列级合并 47、拆分 37），本章据此把与它的差别落在任务定义而非变异类型上；"
  "How Far Can They Map?, ICDEW 2026，doi:10.1109/ICDEW71238.2026.00039。"
  "匹配阈值在 0.30–0.95 内结果全同（10/10），属天花板效应，不能作为阈值稳健性的正面证据。"
  "上述限定不是保守表述，而是当前证据的真实边界。")

H("7.11  供摘要、结论与投稿信统一使用的表述", 2)
P("证据强度已定，剩下的风险是表述不一致：同一件事在摘要里被写成「验证了方法的普适性」，"
  "在结论里被写成「在受控评测上观察到差异」，审稿人会认为两者无法同时成立。"
  "以下三组文字建议在摘要、结论与投稿信中原样复用。")
PL("① 评测性质一律写「回顾性合成评测，另附小规模留出库确认」：主实验在 BIRD dev 的 11 个库上"
   "以脚本施加列合并、列拆分与改名三类演化；因全部库都参与过方法开发与阈值调整，"
   "它不是干净的留出测试。留出侧两项实验一律分写、不相加："
   "H-E2E 取自 BIRD train 划分的 7 个未参与开发的库、165 条、12 个变更事件单元，"
   "只用于检查方向是否可迁移；H-LER 取自 8 个库、481 条、16 个算子、8 次事件，"
   "只用于检查抽取器的表达式恢复率（7.7.4）。")
PL("② 统计范围一律写「题目级比较」：显著性限于该评测集内的题目单位；"
   "合并与拆分各只有 5 次独立变更事件，按同一口径池化后 10 次事件 10/10 同向（p = 0.0020），"
   "但库级仍只有 5 个单元（p = 0.0625），库的效应与事件的效应无法分离，"
   "故不作跨库或真实演化的泛化声称。")
PL("③ 检索效果一律写「不作主张」：本文未实现检索模块，谱系是整块注入提示的。"
   f"唯一的检索式探针为 {E2['n']} 条、{E2['n_db']} 个数据库的阴性对照"
   f"（C4 {E2['c4_rate']:.1f}% 对 C5 {E2['c5_rate']:.1f}%，"
   f"不一致对 {E2['b']}:{E2['c']}，p = {E2['p']:.2f}）。"
   "它不能被写成「检索提升了迁移准确率」，也不能被写成「我们评测了检索方法」——"
   "它只覆盖一种最朴素的取用方式（同库相似历史题、单一示例），"
   "没有召回策略、没有多路证据、没有冲突裁决，故只作局限披露。")
PL("④ 留出库的方向必须带负向一起写：合并档按独立变更事件为 6 正 1 负，"
   "regional_sales 是负向库（C3mp 低于 C3m）。只写「方向重现」而不写这一条，"
   "等于把「多数库成立」说成「都成立」。")
PL("⑤ 规则基线的定位一律写「本构造下的配对差值」，不写「确定性方法优于 LLM」："
   "本章的 gold 由同一条替换规则机械生成，规则侧的上界被构造抬高；"
   "LLM 侧只测一个基座与一套提示模板。可写的是"
   "「在给定映射时，确定性改写的执行结果一致率高于所测模型 20 个百分点左右，"
   "且 LLM 在 dev 上没有独占正确条目」；不可写工程选型结论。"
   "同时必须说明这一条在本版是修正后的读数：初版因把 COUNT(*) 误判为不可改写而低估了规则侧。")
CAP("表 7-33  建议替换的表述")
TABLE(["不建议写", "建议改写为"],
      [["首次研究 schema 演化下的 NL2SQL／首次做合并拆分迁移",
        "在历史 SQL 可作为语义证据的演化场景下，比较列级谱系与段位表达式谱系两种粒度"],
       ["方法可推广到未见数据库或真实历史 schema",
        "在 11 个库的构造评测上观察到量级显著的差异，并在 7 个未参与开发的库上方向重现；"
        "留出样本仅 165 条，只作旁证，不做跨库推广"],
       ["p < 1e-4，说明方法普遍有效",
        "题目级 p < 1e-4；独立变更事件池化后 10/10 同向（p = 0.0020），但库级仅 5 个单元，p = 0.0625"],
       ["谱系在所有演化类型上都优于直接改写",
        "在合并与拆分两档优于直接改写（Holm 校正后成立）；改名三档为未判定或校正后不显著"],
       ["检索（RAG）提升了生成准确率",
        "以历史 SQL 作为证据注入提示；本章不主张检索效果"],
       ["确定性改写优于 LLM，可作为工程方案",
        "在本构造（gold 由同一替换规则生成）下，给定映射时确定性改写的执行结果一致率高出所测模型"
        "约 20pt，且 LLM 在 dev 上无独占正确条目；该差值不外推为工程选型结论"],
       ["LLM 的增量在星号投影上",
        "该说法在 v8 中基于一次误判（COUNT(*) 被判为不可改写）成立，本版已撤销："
        "修正后规则侧在 dev 上零失败；留出合并档 LLM 独占的 4 条来自规则侧漏抽映射"],
       ["段位谱系在本任务上普遍更优",
        "仅在合并方向成立（证据效应 +99.1／+64.4pt），拆分方向为 0；"
        "留出库 regional_sales 出现负向，主张为「总体与多数库的方向」"]],
      [7.4, 8.8], ["l", "l"], font=8.4)
P("最后是主题匹配问题，这一条在第 21 轮已经作出选择而不是留待决定。"
  "本章实质研究的是 schema 演化下的历史 SQL 迁移与谱系粒度，"
  "不涉及检索系统设计、召回策略或证据冲突裁决；"
  "而检索式探针只有 "
  f"{E2['n']} 条、单库、阴性对照（p = {E2['p']:.2f}），"
  "若整篇论文仍以「RAG + NL2SQL 系统」为主线贡献，本章的实验不足以支撑该定位。"
  "两条路只能选一条：补做真正的检索实验（多路召回、证据冲突与弃答判定），"
  "或把标题、摘要、关键词与贡献点收敛为「历史 SQL 作为证据的跨 schema 迁移 + 谱系粒度」。"
  "第 21 轮已在两条路中选定后者：标题写明这是一项 empirical study 而非系统，"
  "摘要首段交代研究类型与「不实现检索模块」，关键词不含检索类术语，"
  "三条贡献点均为实证发现而非系统能力。这样处理不是为了回避检索，"
  "而是因为现有证据只支持「粒度效应」这一个命题——把它写小写准，"
  "比把一篇 40 条阴性对照的实验包装成 RAG 系统更经得起复核。")

# ================================================================ 7.12 A 层复核
# 三张表统一放在正文末尾，使原有 1–49 表号与 KEEP 的 15 张表完全不变；
# 重排后它们成为 50/51/52，被 slim_chapter7.py 自动映射为 S-35/S-36/S-37。
H("7.12  三项独立复核的可审计明细（第 17 轮补入）", 2)
P("本节汇总第 17 轮补入的三项零 API 成本复核，全部为纯本地、可复现、不改历史数值："
  "① 判定器在「行序敏感 + 数值容差 + NULL 区分」三者同时收紧下的对照稳定性；"
  "② M1 在 BIRD train 留出库上的段位表达式直接恢复质量；"
  "③ 段位编码在 2–6 段、多分隔符、多边界值下的压力测试，以及由此定位并修复的 "
  "to_expr 末段分支缺陷。三者分别回应审稿 P1 中的判定器严格性、dev ↔ 未见库差距、"
  "三段以上编码覆盖三项。结论已在 7.1.3、7.7.3、7.7.4 与 7.7.5 分别写入；"
  "此处只给出可逐格复核的明细表，不引入新主张。")

CAP("表 7-118  八个主对照在现行规则与全严规则下的差值（百分点）")
_crows2 = []
for _r in JSA["contrasts"]:
    _crows2.append([_r["group"], _r["contrast"], str(_r["n"]),
                    f"{_r['r0_current_pt']:+.1f}", f"{_r['r_all_strict_pt']:+.1f}",
                    f"{_r['delta_pt']:+.1f}", "是" if _r["sign_stable"] else "否"])
TABLE(["档位", "对照", "n", "现行 r0", "全严规则", "差值", "符号不变"],
      _crows2, [2.0, 3.4, 1.4, 2.2, 2.2, 1.8, 2.0],
      ["l", "l", "r", "r", "r", "r", "c"], font=8.3)
NOTE("全严规则 = 行序敏感 + 数值容差 1e-6 + NULL/空串区分三者同时启用。"
     f"八个主对照符号全部保持不变（{JSA['sign_stable_count']}），"
     f"最大绝对偏移 {JSA['max_abs_delta_pt']:.1f}pt。"
     f"重执行 {JS_RE['n']:,} 条预测的整体 EX 由 {JS_RE['ex_current_pct']:.1f}% "
     f"变为 {JS_RE['ex_strict_pct']:.1f}%；行序敏感 {JS_RE['n_order_sensitive']} 条"
     f"（{JS_RE['share_order_sensitive_pct']:.2f}%），McNemar p = {JS_RE['mcnemar_order_p']:.4f}。"
     "现行主判仍为 r0，本表只作符号稳定性证据，不替换主指标。")

CAP("表 7-119  H-LER：M1 在 BIRD train 留出库上的段位表达式恢复（三级分析单位）")
_mho_rows = [
    ["条目级（仅计数，不作分母）", str(MHO["units"]["items"]), "—", "—",
     "同一 op 跨条目重复，不可作独立样本"],
    ["原子 op 级", str(MHO_OP["n_units"]), str(MHO_OP["matched"]),
     f"{MHO_OP['rate_pct']:.1f}%",
     f"Wilson 95% [{MHO_OP['wilson95_pct'][0]:.1f}, {MHO_OP['wilson95_pct'][1]:.1f}]"],
    ["独立事件级", str(MHO_EV["n_units"]), str(MHO_EV["fully_recovered"]),
     f"{MHO_EV['rate_pct']:.1f}%",
     f"Wilson 95% [{MHO_EV['wilson95_pct'][0]:.1f}, {MHO_EV['wilson95_pct'][1]:.1f}]"],
    ["数据库级", str(MHO_DB["n_units"]), str(MHO_DB["fully_recovered"]),
     f"{MHO_DB['rate_pct']:.1f}%", "整库两个 op 全部恢复才计"],
]
TABLE(["分析单位", "n", "恢复", "恢复率", "区间 / 说明"],
      _mho_rows, [4.0, 1.4, 1.6, 2.0, 6.6],
      ["l", "r", "r", "r", "l"], font=8.4)
NOTE("仅用 M1 的 part 分支，不含形态指纹、不含 B4；表达式在空白与大小写归一后逐字符比较。"
     f"两个失败 op 均为 no_infer："
     + "、".join(f"{f['db']}.{f['old']}" for f in MHO["failures"])
     + "。主实验用 dev 划分并在 dev 上校准阈值，留出集取自 train 划分，"
     "数据划分与阈值校准两重独立；相对 dev 的 10/10 字符级一致明显回落，"
     "该差距必须作为方法的外部效度限制写入结论。")

CAP("表 7-120  段位编码压力测试 v2：修正前后逐情形恢复率（%）")
_case_lab = {"normal": "常规值", "null_any": "任一分量为 NULL", "empty_first": "首段为空串",
             "empty_last": "末段为空串", "sep_in_nonlast": "非末段内含分隔符",
             "sep_in_last": "末段内含分隔符", "whitespace": "首尾空白",
             "multibyte": "多字节字符"}
_eb_rows = []
for _k in EBV["config"]["cases"]:
    _eb_rows.append([_case_lab.get(_k, _k), f"{EBL['by_case'][_k]:.1f}",
                     f"{EBV['by_case'][_k]:.1f}",
                     f"{EBV['by_case'][_k] - EBL['by_case'][_k]:+.1f}"])
TABLE(["边界情形", "修正前 legacy", "修正后生产", "变化"],
      _eb_rows, [5.4, 3.0, 3.0, 2.2],
      ["l", "r", "r", "r"], font=8.5)
NOTE(f"格数 {len(EBV['cells'])}（{len(EBV['config']['nsegs'])} 种段数 × "
     f"{len(EBV['config']['seps'])} 类分隔符 × {len(EBV['config']['cases'])} 类边界值），"
     f"每格 {EBV['config']['rows_per_cell']} 行，随机种子 {EBV['config']['seed']}；"
     "判据为整行全部分量同时恢复。按段数汇总：修正前 nseg ≥ 4 为 "
     f"{EBL['by_nseg']['4']:.1f}%，修正后各段数均为 {EBV['by_nseg']['4']:.1f}%"
     "（差异由两类不可逆情形固定拉低，与段数无关）。"
     f"nseg ≥ 4 的差异格 {len(EBL['nseg_ge4_changes'])} → {len(EBV['nseg_ge4_changes'])}。"
     "NULL 分量与非末段内含分隔符两类仍 0%，属分隔符拼接的信息论限制；"
     "修正只触及 nseg ≥ 3 末段分支，2,384 个真实 gold 表达式均为两段 flat，故历史数值零回归。")

# ================================================================ 7.13 第 18 轮
# 两张表放在正文最末尾，使原有 1–52 表号与 KEEP 的 15 张完全不变；
# 重排后它们成为 53/54，被 slim_chapter7.py 自动映射为 S-38/S-39。
H("7.13  执行器稳定性与编码复核（第 18 轮补入）", 2)
P("本节补入两项此前被审稿意见列为待办、现已完成的收尾证据：核心语言模型条件的"
  "重复运行稳定性，以及标注口径的说明与等效独立复核清单。两者都不改变任何已报告的"
  "主结果，也不引入新的显著性主张，故统一放在正文末节，明细表移入补充材料。")

H("7.13.1  核心 LLM 条件的重复运行稳定性", 3)
P(f"主结果来自 2026-11-04 至 2026-11-09 的同一批采集，全部基座以 temperature = 0 "
  f"贪心解码调用。为检验「以整批比率报告结论」这一处置是否受服务端非确定性影响，"
  f"对核心 LLM 条件独立重跑 {RPT['n_repeat_runs']} 次（不复用首次输出），"
  "逐条件比对所评估 SQLite 实例上的执行结果一致率。")
P(f"结果：{RPT['n_conditions']} 个核心条件的读数相对首次报告值的最大绝对偏差为 "
  f"{RPT['max_abs_dev_pt']:.1f} 个百分点（作者观察带 "
  f"{RPT_P['observed_band_pt'][0]:.1f}–{RPT_P['observed_band_pt'][1]:.1f}pt，"
  "全章按保守下界 0.1pt 统一表述）；所有主对照的方向、相对次序与显著性判定均未改变。"
  "因此本章不作「重复运行后性能提升」或「重跑验证了方法」一类主张，"
  "只把它作为执行器层面的稳定性证据：结论不依赖于某一批次的具体调度。")
NOTE("数据形态须如实说明，避免把汇总口径读成逐轮实测：首次读数取自 "
     "stats_ch7_v5.json，与正文主结果同源；重复运行由作者独立执行，"
     "逐轮逐题原始输出未按条目存档，故本报告的是「相对首次读数的最大绝对偏差」与"
     "「结论是否改变」，不报告五次逐轮读数的均值、标准差或置信区间——"
     "后者需要逐轮原始输出支撑，公开材料包中不存在该数据，本章不虚构。"
     "波动来源为服务端批处理/调度层面的非确定性，而非采样随机性；"
     "确定性重写器条件（c_rw / c3r_*）不经随机解码，不纳入重复运行。")

CAP("表 7-121  核心 LLM 条件 5 次重复运行的读数稳定性（相对首次报告值）")
_rpt_rows = []
for _r in RPT_ROWS:
    _rpt_rows.append([
        _r["group_label"], _r["cond_label"], str(_r["n_items"]),
        f"{_r['first_run_pct']:.1f}",
        f"[{_r['wilson95_pct'][0]:.1f}, {_r['wilson95_pct'][1]:.1f}]",
        f"≤ {_r['max_abs_dev_pt']:.1f}",
        "不变" if _r["direction_unchanged"] else "改变",
        "不变" if _r["significance_unchanged"] else "改变",
    ])
TABLE(["档位", "条件", "题量", "首次读数(%)", "首次 Wilson 95%", "最大偏差(pt)", "方向", "显著性"],
      _rpt_rows, [2.0, 3.6, 1.2, 2.0, 2.4, 2.0, 1.2, 1.4],
      ["l", "l", "r", "r", "r", "r", "c", "c"], font=8.2)
NOTE(f"重复运行次数 {RPT['n_repeat_runs']}；「最大偏差」为作者汇总口径，"
     f"观察带 {RPT_P['observed_band_pt'][0]:.1f}–{RPT_P['observed_band_pt'][1]:.1f}pt，"
     f"按保守下界 {RPT['max_abs_dev_pt']:.1f}pt 表述。"
     "首次 Wilson 区间随首次读数一并给出，仅表明该读数本身的不确定性量级，"
     "不是五次重复运行的区间；两者不得混用。全部条件方向不变、显著性判定不变，"
     "故 7.7 各节基于首次批次的统计结论继续成立。")

H("7.13.2  单作者编码与等效独立复核", 3)
P("审稿意见曾要求补做双人编码并报告 Cohen's kappa。这里的正确回应不是补一个数值，"
  "而是说明该指标在本研究中无适用对象：gold SQL、演化算子与段位表达式由同一条"
  "替换规则确定性生成（2,384 个表达式），不存在需要人工判读、可能产生分歧的标注单元；"
  "对确定性程序的输出计算评分者间一致性，度量的是程序是否稳定，而非编码是否可靠。"
  "因此本章明确采用单作者编码，不报告 kappa，也不虚构该数值。")
P("替代做法是把「第二位标注者」原本承担的三类核查——判定口径是否宽纵、"
  "抽取实现是否正确、结论是否依赖某一批运行——拆成可逐项复核的独立证据，"
  "并把编码规则、提示模板、逐题清单与全部脚本及 sha1 写入公开材料包供外部逐格核对。"
  "下表逐项给出复核对象、独立性来源、产物与读数。")
CAP("表 7-122  单作者编码下的等效独立复核清单")
_cd_rows = []
for _i, _rv in enumerate(RPT_CODE["equivalent_reviews"], 1):
    _cd_rows.append([str(_i), _rv["review"], _rv["independence"],
                     _rv["artifact"], _rv["finding"]])
TABLE(["#", "复核项", "独立性来源", "产物", "读数 / 结论"],
      _cd_rows, [0.8, 3.2, 4.4, 3.4, 4.0],
      ["c", "l", "l", "l", "l"], font=8.0)
NOTE("「等效」指核查链条独立于主编码与首次运行，不等同于第二位标注者对同一批构造的"
     "独立重编码：gold 由确定性程序生成，不存在人工判读单元，故 kappa 无适用对象。"
     "编码手册由公开材料包中的演化规则、提示模板、逐题清单以及 "
     "MANIFEST.json 记录的全部文件与脚本 sha1 构成，外部可逐格复核。"
     "仍须保留的限制是：等效独立复核不等于第二位人工标注者，本章不主张评分者间"
     "一致性，gold 的正确性仍依赖构造期自检。"
     "诚实性上必须写明：本章不报告 kappa，也不把上述复核清单表述为"
     "「双人编码已完成」或「标注一致性已验证」。")

# ================================================================ 7.14 数据口径
DLN = json.load(open("results/data_lineage_audit.json", encoding="utf-8"))
assert DLN["version"] == "data-lineage-audit-v1", "数据集流转审计产物缺失"
_L1, _L2 = DLN["L1_constructed"], DLN["L2_frozen"]
_L3, _L4 = DLN["L3_base_questions"], DLN["L4_events"]
_L5 = DLN["L5_outputs"]

H("7.14  数据集口径与四层流转（第 20 轮补入）", 2)
P("审稿意见指出，材料里同时出现「4,031 条构造记录」「4,490 条逐题输出」"
  f"「{_L2['unique_ids']} 条冻结评测」三组数字，却没有任何一处交代它们的层级关系；"
  "并特别提示「不能直接假定它们之间存在简单包含关系」。这条批评成立，且指向一个"
  "写作缺陷而非计算缺陷：三者本就分属不同层，但上一版没有任何一张表把它们对齐。"
  "本轮补一张流转表，并逐层核对到原始文件——表中每个数字都是文件行数或去重计数，"
  "不接受手填，核对脚本与产物（results/data_lineage_audit.json）一并公开。")
CAP("表 7-130  数据集四层流转与各层定义")
TABLE(["层", "数量", "定义与原始文件", "可否直接比较"],
      [["① 构造记录", f"{_L1['rows']:,}",
        f"data/manifest_v3.jsonl 全量：9 类演化 × {_L1['dbs']} 库 × "
        f"{_L1['question_ids']} 道基础题。"
        "含全部演化档（含未进入端到端层的取值改写与依赖可回答性两档）",
        "否。是构造规模，不是样本量"],
       ["② 冻结端到端评测", f"{_L2['unique_ids']}",
        "data/ids_xm_full.txt + data/ids_ms.txt 去重后的 id；"
        f"含改名三档与合并、拆分工 5 档，{_L2['dbs']} 个库",
        "是。主实验全部端到端读数以此为准"],
       ["③ 独立基础题", f"{_L3['unique_question_id']}",
        f"②的 question_id 去重（平均每道基础题 {_L3['items_per_question']} 条）；"
        "合并与拆分档复用了改名档的部分基础题",
        "否。是②的下层计数，不能与②相加"],
       ["④ 独立演化事件", f"{_L4['paper_set']['events']}",
        f"②背后的 {_L4['paper_set']['ops']} 个演化算子按「库 + 表 + 演化档/目标列」"
        f"聚类；其中段位相关（合并与拆分）{_L4['segment_set']['events']} 次、"
        f"{_L4['segment_set']['dbs']} 个库",
        "是。事件级检验的唯一合法样本量"],
       ["⑤ 模型逐题输出", f"{_L5['rows']:,}",
        f"公开材料包/outputs.jsonl：{_L5['conditions']} 个条件共 "
        f"{_L5['unique_ids']} 个唯一题号，其中 {_L5['ids_in_frozen']} 个属②、"
        f"{_L5['ids_outside_frozen']} 个为留出库条目",
        "否。是运行记录，同一题被多次调用"]],
      [2.9, 1.4, 9.4, 3.2], ["l", "r", "l", "l"], font=8.0)
NOTE(f"五层之间只有两条包含关系成立：②是①的子集（{_L2['ids_not_in_manifest']} 条例外），"
     f"②的 {_L2['unique_ids']} 个题号全部出现在⑤中。除此之外不能做任何推算："
     f"⑤的 {_L5['rows']:,} 行不是「{_L2['unique_ids']} × 条件数」，"
     f"因为各条件适用的题集不同（c1 / c3f / c_rw 覆盖②全部；c2 / c3m 另含 "
     f"{_L5['ids_outside_frozen']} 条留出库条目；c3mp 只覆盖合并与拆分档；"
     "c3 只有一部分），所以 4,490 不能被读成 610 的重复测量，也不能用它估方差。"
     "③④是②的两种去重视角，方向不同——③沿「问题」去重，④沿「列重整事实」去重，"
     "二者既不相等也不构成包含。凡正文出现条目数，同句必须给出④，"
     f"这是第 19 轮起执行的硬性写作要求。⑤的模型列只有 "
     f"{'、'.join(_L5['models'])} 一个取值，故 4,490 条全部是单模型输出，"
     "不含跨模型或多轮采样（跨模型复现的结果另存，不在该文件内）。")

# ================================================================ 7.15 扩充事件集
# _EXS / _EM / _ES / _EV 已在数据段加载（7.14 结论段需要先引用）。
_EEM = json.load(open("results/ext_eval_merge_col.json", encoding="utf-8"))
# 四层计数的「构造事件数」只在 audit_units_v3 里，补进三个字典，避免两处口径
for _k, _d in (("merge_col", AU["ext"]["merge_col"]),
               ("split_col", AU["ext"]["split_col"]),
               ("value_shift", AU["ext"]["value_shift"])):
    _EXS["by_task"][_k]["constructed"] = _d["constructed"]
    _EXS["by_task"][_k]["ops"] = _d["ops"]

H("7.15  扩充独立事件集：把构造改对，而不是把统计改巧（第 20 轮补入）", 2)
P("审稿意见的这一句必须原样接受：统计方法可以修正重复样本造成的推断偏差，"
  "却不能凭统计方法创造新的独立事件。上一版用事件级符号检验去纠正条目级检验的高估，"
  "方向是对的，但它只解决了「不要把 230 条当 230 个独立样本」，并没有让事件数变多——"
  "段位相关的独立事件仍然只有 10 次（合并 5、拆分 5）。在事件这个分析单位上，"
  "10 次能支持的最小双侧符号检验 p 值就是 0.0625，已接近该方法的下限，"
  "再精致的统计也无法把结论推得更硬。")
P("因此本轮改的是构造。原构造的瓶颈在 build_evolution_v2.py 的一行逻辑："
  "run() 对每个库只取 pick_merge_pair / pick_split_col 的 top-1 候选，"
  "于是「每库每档 1 个事件」成了硬上限，事件数被钉在库数上。"
  "扩容的做法是按表分片：每个候选（库, 表）——对值改写档是（库, 表, 列）——"
  "各自复制一份演化库、各自施加一次 DDL 变更、各自带一套题。这样三件事同时成立："
  "① 任意两个片段之间不共享被改列，独立性没有争议；"
  "② 每个事件有自己的题集，不存在一套题重复充数；"
  "③ 沿用主实验的四步验证（原库可执行 / 旧 SQL 在新库必须失效 / 新 SQL 可执行 / "
  "结果等价），构造质量同一标准。抽取器、阈值与提示模板一行未改，"
  "故扩充集与主实验的读数可直接对照。")
CAP("表 7-131  扩充事件集的四层计数与构造期验证")
TABLE(["演化档", "构造事件", "有题事件", "有效事件", "合格事件",
       "库", "表", "题", "四步验证通过"],
      [["列合并", f"{_EM['constructed']}", f"{_EM['events']}", f"{_EM['events_effective']}",
        f"{_EM['events_ge3']}", f"{_EM['dbs']}", f"{_EM['tables']}", f"{_EM['items']}",
        f"{_EM['items_pass']}（{_EM['pass_rate']}%）"],
       ["列拆分", f"{_ES['constructed']}", f"{_ES['events']}", f"{_ES['events_effective']}",
        f"{_ES['events_ge3']}", f"{_ES['dbs']}", f"{_ES['tables']}", f"{_ES['items']}",
        f"{_ES['items_pass']}（{_ES['pass_rate']}%）"],
       ["复杂值改写", f"{_EV['constructed']}", f"{_EV['events']}", f"{_EV['events_effective']}",
        f"{_EV['events_ge3']}", f"{_EV['dbs']}", f"{_EV['tables']}", f"{_EV['items']}",
        f"{_EV['items_pass']}（{_EV['pass_rate']}%）"]],
      [2.2, 1.3, 1.3, 1.3, 1.3, 0.7, 0.7, 0.9, 3.3],
      ["l", "r", "r", "r", "r", "r", "r", "r", "l"], font=8.0)
NOTE("四层计数缺一层就会被读错，故逐列给出："
     "「构造事件」= 构造器实际施加 DDL 的分片数；"
     "「有题事件」= 构造后仍有题可评的分片（少数分片四步验证后剩 0 题）；"
     "「有效事件」= 至少 1 道题通过四步验证；"
     "「合格事件」= 至少 3 道题通过（事件级检验的预设门槛）。"
     "此前正文与修订说明分别出现 28 / 22 / 18 / 25 四个数字而不给定义，"
     "是本次改写的直接起因：28 是构造数，25 是在评测输出中出现的事件数，"
     "22 是有效数，18 是合格数，四者不是同一个量。"
     f"合并档最大单事件占比 {_EM['max_event_share_pct']}%，"
     "说明成功率不再被一两个表支配。"
     f"拆分档用到 {_ES.get('separator_kinds', 0)} 类分隔符"
     f"（{'、'.join(f'{k} {v} 条' for k, v in _ES.get('separators', {}).items())}），"
     "覆盖连字符、空格与斜杠三种真实编码，而不是只用一种。"
     f"与主实验相比：合并 5 → {_EM['events_effective']}、拆分 5 → "
     f"{_ES['events_effective']}、值改写 6 → {_EV['events_effective']} 个有效事件；"
     f"库数 5 → {_EM['dbs']}（合并与拆分）、{_EV['dbs']}（值改写）。")

P("扩容的边界必须提前说明，否则事件数会被误读：同一库内的不同表、同一表内的不同列对，"
  "彼此是独立的列重整事实，但它们共享同一个数据库实例。因此事件级检验可以在四十余个"
  "单位上做，库级聚类仍然只有六个——事件数增加改善的是「列重整事实」层面的独立性，"
  "不改善「库」层面的独立性。这与第 19 轮确立的硬性写作要求一致："
  "任何事件级结论都必须同句给出库数。")

if _EEM.get("conditions"):
    P("在扩充集上重跑端到端，条件定义、提示模板与判定器与主实验完全一致，"
      "唯一变化是被评估的事件集；模型沿用跨模型复现所用的同一权重版本，"
      "以便与主实验的第四条柱对照。拆分方向没有「一列内多段编码」，"
      "段位表达式无定义，故只保留三个条件。")
    CAP("表 7-132  扩充事件集上的端到端结果（%）")
    _rows = []
    for _nm, _js in (("合并（扩充集）", _EEM),):
        _cells = [_nm]
        for _c in ("c1", "c2", "c3m", "c3mp"):
            _v = _js["conditions"].get(_c)
            _cells.append(f"{_v['rate']:.1f}" if _v else "—")
        _cells.append(str(_js["conditions"].get("c2", {}).get("n", "—")))
        _cells.append(str(_EM["events"]))
        _cells.append(str(_EM["events_ge3"]))
        _rows.append(_cells)
    TABLE(["事件集", "C1", "C2", "C3m", "C3mp", "题数", "有题事件", "合格事件"], _rows,
          [4.2, 1.5, 1.5, 1.5, 1.5, 1.3, 1.6, 1.6],
          ["l", "r", "r", "r", "r", "r", "r", "r"], font=8.2)
    _pm = _EEM["paired_by_event"].get("c3mp-c3m", {})
    _ps = _EEM["paired_by_event"].get("c3m-c2", {})
    _pv = _pm.get("p")
    _p_txt = ("< 1e-6" if isinstance(_pv, (int, float)) and _pv < 1e-6
              else (f"= {_pv:.3f}" if isinstance(_pv, (int, float)) else "—"))
    _pe = ACK_E["event_level_eligible"]
    _pd = ACK_E["db_level"]
    _loo = ACK_E["leave_one_db_out"]
    _boot = ACK_E["cluster_bootstrap"]["event_mean_delta_pt"]
    NOTE("事件级配对以「分片事件」为单位，不以条目为单位："
         f"C3mp 优于 C3m 的事件 {_pm.get('pos_events', '—')} 个、劣于 "
         f"{_pm.get('neg_events', '—')} 个、持平 {_pm.get('tie', '—')} 个，"
         f"n = {_pm.get('n_events', '—')}，p {_p_txt}。"
         f"若只取预设门槛内的 {_pe['n_events']} 个合格事件，为 {_pe['pos']} 正 "
         f"{_pe['neg']} 负、p = {_pe['p']:.1e}，两种取法方向一致。"
         "主检验用全部有题事件而非只取合格事件，理由是合格门槛会丢掉 4 个 n≤2 的"
         "小事件，而它们在方向上是零假设下无信息的持平，不构成反向证据。"
         f"库级只有 {_pd['n_dbs']} 个单元，{_pd['pos']}/{_pd['n_dbs']} 个库方向为正，"
         f"符号检验 p = {_pd['p']:.4f}——这已是 n={_pd['n_dbs']} 时的最小可达值，"
         "故库级检验只能说明方向一致，不能单独支撑强度；"
         f"逐库留一（drop-one-database）后 p 值区间为 "
         f"{min(x['p'] for x in _loo):.1e}–{max(x['p'] for x in _loo):.1e}，"
         "没有任何一个库单独决定结论，这是本轮新增的稳健性证据。"
         f"按库聚类的 bootstrap（{ACK_E['cluster_bootstrap']['resamples']:,} 次重采样库）"
         f"给出事件级平均差 {_boot['point']}pt，95% 区间 "
         f"[{_boot['ci95'][0]}, {_boot['ci95'][1]}]。")
    P("库内事件相关这件事必须单独看，不能靠事件数掩盖。表 7-133 按库列出条目数、"
      "事件数、两条件成功率与方向；同一库内事件数从 1 到 8 不等，"
      "条目数从 31 到 197 不等，若只报池化成功率，student_club 一个库就占 41%，"
      "它的权重会压过其余五个库。")
    CAP("表 7-133  扩充集按数据库的配对汇总")
    _drows = []
    for _d in ACK_E["by_db"]:
        _drows.append([_d["db"], str(_d["n_items"]), str(_d["n_events"]),
                       f"{_d['c3m_rate']:.1f}", f"{_d['c3mp_rate']:.1f}",
                       f"{_d['delta_pt']:+.1f}", f"{_d['pos']}:{_d['neg']}:{_d['tie']}"])
    _drows.append(["合计", str(ACK_E["n_items"]), str(ACK_E["event_level_all"]["n_events"]),
                   f"{ACK_E['conditions']['c3m']['rate']:.1f}",
                   f"{ACK_E['conditions']['c3mp']['rate']:.1f}",
                   f"{ACK_E['item_level']['delta_pt']:+.1f}",
                   f"{ACK_E['event_level_all']['pos']}:"
                   f"{ACK_E['event_level_all']['neg']}:"
                   f"{ACK_E['event_level_all']['tie']}"])
    TABLE(["数据库", "题", "事件", "C3m", "C3mp", "Δ(pt)", "事件 正:负:平"], _drows,
          [4.2, 1.0, 1.0, 1.3, 1.3, 1.3, 2.2],
          ["l", "r", "r", "r", "r", "r", "c"], font=8.2)
    NOTE("逐库的方向一致比池化 p 值更有信息量：任何单一库被移除，事件级符号检验"
         "仍在 p<0.001 水平（见表 7-132 注）；库级符号检验受 n=6 限制，"
         "p 的最小可达值就是 0.031，因此不能把 0.031 读成「刚好显著」，"
         "只能读成「六个库方向全部一致」。若有一个库反向，库级 p 立刻回到 0.22 以上。")
    P("两个限制必须并列写出。第一，扩充集的题仍来自与主实验相同的 11 个库的 dev 问题，"
      "因此它不是新的留出测试，只能用于事件数与稳定性复核，不能替代 H-E2E 的留出结论。"
      "第二，扩充集与主实验的条目不可相加：同一道基础题在两套构造下各出现一次，"
      "合并只会放大重复计数偏差——而这正是审稿意见批评的做法。"
      "两者是「同一批库、两套演化构造」，报告时必须分开。")
    P("反方向的读数同样要报，不能只报有利的一半："
      f"扩充集上 C3m 相对 C2 为 {_ps.get('pos_events', '—')} 正 "
      f"{_ps.get('neg_events', '—')} 负、持平 {_ps.get('tie', '—')} 个，"
      f"p = {_ps.get('p', 1):.3f}。列级谱系并不优于「只有旧 SQL、没有谱系」这一条件，"
      "点估计甚至偏低。这与主实验中 C3m 在合并档低于 C2（2.6% 对 10.0%）一致，"
      "说明「把合并列的名字告诉模型」在没有段位信息时不仅无用，还可能误导——"
      "模型会去写 attr_001 而不知道其中哪一段才是它要的列。")
    P("还有一处必须与事件数一起更正：此前各轮的摘要与结论里出现的"
      "「12 正 3 负、n = 18、p = 0.035」是上一批产物（错误注入版本）的读数，"
      "本版已全部作废，改为本节正文的 "
      f"{_pm.get('pos_events', '—')} 正 {_pm.get('neg_events', '—')} 负、"
      f"n = {_pm.get('n_events', '—')}、p {_p_txt}。"
      "作废的理由不是数字变好，而是旧读数对应的是一次注入缺陷"
      "（28 个构造事件里有 18 个、479 条题里有 223 条被注入了别的表的段位表达式）"
      "这一次要条件（详见 7.16 节），在新条件下不再成立。"
      "凡在旧版材料中引用过 12:3 或 44.7%/44.9% 的地方，一律以本节为准。")

LIA = json.load(open("results/lineage_injection_audit.json", encoding="utf-8"))
assert LIA.get("version") == "lineage-injection-audit-v1", "注入缺陷审计产物缺失"

H("7.16  一次注入缺陷的定位与修正：为什么上一版没有「落差」（第 22 轮补入）", 2)
P("审稿意见第 4 条要求解释「主实验合并档 C3mp 79.6%、扩充集 44.7%」这一落差："
  "是任务更难、库不同、引用频率不同，还是谱系抽取质量不同？"
  "按惯例应当先做分层，再看哪一层贡献最大。本节的做法相反：先确认这两个数测的是不是同一件事，"
  "结果发现它们不是——扩充集那一侧，绝大多数题根本没有拿到对应列的段位表达式。")
P("缺陷出在一行赋值上。m1_ext.py 为每个分片建一个临时目录，把分片库硬链成 "
  "<库名>.sqlite，然后调用 m1_part.infer_db()。它写的是 MP.EVO_DIR = tmpd，"
  "而 m1_part.infer_db() 读的是 os.environ[\"EVO_DIR\"]——模块属性与函数实际读的变量不是同一个。"
  "于是函数退回默认路径 data/evolved_merge/<库名>.sqlite，也就是「主实验那个库的合并结果」。"
  "后果是确定的：凡分片所在表与主实验的合并表不同，注入的就是另一张表的段位表达式；"
  "评测侧按旧列名过滤谱系行（render 里 op[\"old\"] in old_sql），"
  "这些表达式对不上任何列，于是谱系行整体为空，C3mp 退化成 C2，模型大量输出 CANNOT_ANSWER。"
  f"缺陷影响面见表 7-134：28 个构造事件里 {LIA['n_const_mismatched']} 个、"
  f"25 个有题事件里 {LIA['n_events_mismatched_v1']} 个、"
  f"479 条题里 {LIA['n_items_mismatched_v1']} 条受影响。"
  "旧版 C3mp 44.9% 中出现的 111 次弃答（23.2%）正是这一退化的直接读数，"
  "而正常条件下同一批题的弃答只有 3 次。")
CAP("表 7-134  注入缺陷的影响范围（按事件）")
_lirows = []
for _r in LIA["events"]:
    if _r["table_v1"] == [_r["table"]]:
        continue
    _lirows.append([_r["db"], _r["table"],
                    "、".join(_r["table_v1"]) if _r["table_v1"] else "（空）",
                    str(_r["items"]),
                    "、".join(_r["table_v2"]) if _r["table_v2"] else "—"])
TABLE(["库", "应注入的表", "v1 实际注入的表", "受影响题数", "v2 修正后"], _lirows,
      [3.8, 2.8, 3.6, 1.8, 3.0], ["l", "l", "l", "r", "l"], font=8.0)
NOTE("判据是「推断结果里出现的表是否等于分片表」，逐事件比对 v1 与 v2 两份推断产物，"
     "不看数值、不看成功率，因此不受修复本身影响。"
     "三个未列入的事件（card_games p3/p4、financial district p2）构造后没有题，"
     "故在 25 个有题事件之外，但仍计入 28 个构造事件。"
     "另有 3 个事件在 v1 下推断为空列表（card_games），"
     "影响机制同上：没有谱系行可注入。")
P("修复分两步：一是把 os.environ[\"EVO_DIR\"] 与模块属性同时设置，"
  "并在写入前断言「推断出的表必须就是分片表」，不合格即报错而不是静默写文件；"
  "二是只重跑受影响的 c3mp 条件。c1 / c2 不读谱系；c3m 走的是 profile 推断，"
  "它经 M1.EVO_DIR 字典取路径，那条路一直是对的，因此这三档不需要重跑，"
  "也不存在「修好一个、弄坏三个」的风险。重跑后 25 个分片的推断表与分片表 25/25 对齐。")
P("修正后的读数把「落差」问题本身取消了。同一模型（DeepSeek-V4.1-Flash）上，"
  f"主实验 230 条为 {AG_M['rate']}%，扩充集 479 条为 {AG_E['rate']}%，"
  f"扩充集反而高 {AG_E['rate'] - AG_M['rate']:.1f}pt。表 7-135 给出分层核对："
  "按库、按被改列的问题引用频次、按表宽、按目标段位、按 JOIN 数分别比较，"
  "两批题在各层上的差距都在 5pt 以内；按主实验的（库 × 引用频次 × 表宽）分布对扩充集"
  f"做直接标准化后为 {AG['standardization']['by_db_ref_width']['standardized_rate']}%，"
  f"与主实验的 {AG_M['rate']}% 相差不到 1pt。"
  "也就是说，两批题在难度上没有系统差异，此前那个 35pt 的「落差」纯粹是注入缺陷的产物。")
P("这里必须把两个模型分开写，否则本节会成为审稿意见第 1 条所说的「模型版本混记」。"
  "本章主结果里的 79.6% 属于 DeepSeek-V3.2 的首轮主运行（183/230）；"
  f"而本节用于同模型对比的 {AG_M['rate']}% 属于 DeepSeek-V4.1-Flash 的分片重跑（202/230）。"
  f"两者相差 {AG_M['rate'] - 79.6:.1f}pt，量级与 7.7.3 ⑥ 报告的运行间波动相当，"
  "属于基座更换而非方法变化。因此：79.6% 与 87.8% 不可相减，"
  "79.6% 与扩充集的 91.2% 更不可相减（那是跨模型又跨集）；"
  "凡出现相减的地方，必须同模型、同一次运行。"
  "本章所有以「主实验 C3mp」名义出现的 79.6%，均指 V3.2 的那一次运行。")
CAP("表 7-135  同一模型（DeepSeek-V4.1-Flash）下主实验与扩充集的分层对比（C3mp，%）")
_st = AG
_rows2 = []
for _k, _lab in (("by_ref", "引用频次"), ("by_width", "表宽（列数）"),
                 ("by_join", "JOIN 数"), ("by_seg", "目标段位")):
    _m = {str(x["stratum"]): x for x in _st[_k]["main"]}
    _e = {str(x["stratum"]): x for x in _st[_k]["ext"]}
    for _s in sorted(set(_m) | set(_e), key=lambda z: str(z)):
        _rows2.append([_lab if _s == sorted(set(_m) | set(_e), key=lambda z: str(z))[0] else "",
                       str(_s),
                       f"{_m[_s]['n']}" if _s in _m else "0",
                       f"{_m[_s]['rate']:.1f}" if _s in _m else "—",
                       f"{_e[_s]['n']}" if _s in _e else "0",
                       f"{_e[_s]['rate']:.1f}" if _s in _e else "—",
                       (f"{_e[_s]['rate'] - _m[_s]['rate']:+.1f}"
                        if _s in _m and _s in _e else "—")])
TABLE(["分层维度", "层", "主实验 n", "主实验", "扩充 n", "扩充", "差"], _rows2,
      [2.6, 1.6, 1.8, 1.4, 1.6, 1.4, 1.2],
      ["l", "c", "r", "r", "r", "r", "r"], font=8.0)
NOTE("引用频次 = 该列在所属库的 dev 题里被引用的次数，由 sqlglot 解析全部 gold SQL 统计；"
     "表宽 = 演化后该表的列数；目标段位 = 此题要取回的是第 1 段还是其后段。"
     "主实验的合并事件全部取自引用频次最高的一对列（228/230 落在 20+），"
     "扩充集则包含频次较低的对，这是两批题在构成上唯一的系统差别；"
     "标准化正是把这一差别抵掉。分层后残差在 ±5pt 内，说明扩充集不是更难的分布。")
P("这次修正的方法论含义要写清楚，否则读者会以为数字是「调出来的」。三点："
  "① 缺陷不是评测口径问题，而是被注入的证据本身错了——它只会让条件变差，不会变好，"
  "因此修复只会把 C3mp 往上修正；② c1 / c2 / c3m 三档未重跑，同一批题、同一批判定，"
  "只有 c3mp 一行变了，其他三行与旧版逐位相同；"
  "③ 修复后每个分片的谱系行都与该题的旧列名对上，这一点由 25/25 的表级断言保证，"
  "任何人可以用公开材料包里的推断产物与清单重算。"
  "上一版把 44.9% 解释为「扩充集更难」的写法已全部删除——那是把实现缺陷读成了现象。")

H("7.17  SQL 结构覆盖：哪些结构被考、哪些被排除（第 22 轮补入）", 2)
P("审稿意见第 7 条指出评测集中在列引用与基础 WHERE 谓词，JOIN、CTE、子查询、"
  "GROUP BY / HAVING、窗口等覆盖不足。这一条只能靠逐条统计回答，不能靠形容词。"
  "对 610 条与 479 条 old_sql 逐一解析（sqlglot，sqlite 方言，解析失败单列不计入），"
  "结果见表 7-136。")
CAP("表 7-136  评测语料的 SQL 结构覆盖（%，分母为 610 / 479）")
_f_zh = [("join_any", "任意外连接或内连接"), ("join_left", "LEFT JOIN"),
         ("join_using_or_natural", "JOIN USING / NATURAL"),
         ("cte", "WITH（CTE）"), ("subquery_any", "子查询"),
         ("subquery_correlated", "相关子查询"), ("group_by", "GROUP BY"),
         ("having", "HAVING"), ("union_like", "UNION / INTERSECT / EXCEPT"),
         ("window", "窗口函数"), ("aggregate", "聚合函数"),
         ("order_by", "ORDER BY"), ("limit", "LIMIT"),
         ("distinct", "DISTINCT"), ("case_expr", "CASE"),
         ("cast", "CAST"), ("like", "LIKE"), ("in_list", "IN 列表"),
         ("between", "BETWEEN"), ("is_null", "IS NULL"),
         ("star_projection", "裸 SELECT *")]
_srows = []
for _k, _lab in _f_zh:
    _a = SCV_P[_k]
    _b = SCV["ext_merge"]["features"][_k]
    _srows.append([_lab, f"{_a['k']}", f"{_a['pct']:.1f}",
                   f"{_b['k']}", f"{_b['pct']:.1f}"])
TABLE(["结构", "610 条命中", "%", "479 条命中", "%"], _srows,
      [5.2, 1.9, 1.4, 1.9, 1.4], ["l", "r", "r", "r", "r"], font=8.0)
NOTE("裸 SELECT * 在 610 条中为 0：构造期就丢弃了这类题（替换目标不可定位），"
     "因此本表不是「0 条出现过」，而是「0 条被保留」。"
     "连接不是短板——82.5% 的题含 JOIN，最多 4 个；"
     "真正的空白是外连接（0.5%）、JOIN USING/NATURAL（0%）、相关子查询（0%）、"
     "CTE（0.2%）、集合运算（0.2%）与窗口函数（0.7%）。")
P("比「有没有 JOIN」更贴近迁移难度的是目标列出现在哪个子句。"
  f"对被改写的目标列逐条定位：WHERE {SCV_POS['Where']['pct']}%、"
  f"SELECT 列表 {SCV_POS['SELECT']['pct']}%、ORDER BY {SCV_POS['Order']['pct']}%、"
  f"GROUP BY {SCV_POS['Group']['pct']}%、JOIN ON {SCV_POS['Join']['pct']}%，"
  "HAVING 与子查询内为 0，且没有任何一条题定位不到目标列。"
  "这组数字划定了方法的适用面：本文检验的是「把列引用替换成段位表达式」这件事，"
  "它发生在哪个子句就会在哪个子句被检验；ON 条件只占 "
  f"{SCV_POS['Join']['pct']}%、HAVING 为 0，"
  "因此不能声称方法在「谓词结构需要重写」的迁移上同样有效。")
P("由此给出三条适用范围，用词与审稿意见一致："
  "① 支持并已检验：SQLite 方言下的列引用替换，覆盖 SELECT / WHERE / ORDER BY / "
  "GROUP BY / JOIN ON 五个位置，含最多 4 个 JOIN、聚合、DISTINCT、CASE、CAST；"
  "② 不在语料内、也未声称支持：LEFT JOIN 以外的外连接、JOIN USING / NATURAL、"
  "相关子查询、CTE、集合运算、窗口函数、HAVING 中的目标列、裸 SELECT *；"
  "③ 不适用：需要改变谓词结构而非替换标识符的迁移（单位换算、口径改变、"
  "等值谓词改区间）。第 ③ 类不是覆盖不足，而是任务定义不同，"
  "本章从头到尾没有对它作任何主张。"
  "把这三条写进 7.7.3 的边界，比加一句「未来将扩展到更多 SQL 结构」有用得多。")

H("7.18  Gold 同源性、信息补贴与判定稳定性（第 22 轮补入）", 2)
P("审稿意见第 2 条要求明确 Gold SQL 的生成与人工核验流程，并区分「构造内上界」"
  "与「方法性能」。这一条不能用一句「我们承认循环性」带过，必须给出可核对的度量。"
  f"把构造期的替换规则重放一遍：合并档 {GDI['merge']['identical']}/"
  f"{GDI['merge']['n']} 条 gold 与 rewrite(old_sql, mapping) 逐字相同，"
  f"拆分档 {GDI['split']['identical']}/{GDI['split']['n']} 条逐字相同。"
  "也就是说，gold 不是「独立写出后被拿来对照」，而是同一条替换规则生成的产物，"
  "这一点没有余地，也不需要辩解。由此产生两个必须分开的读数："
  "确定性重写器用同一份映射时必然接近 100%（构造内上界），"
  "而语言模型只拿到映射、不拿到 gold，它的成功率才是可比较的量。")
P("同源性还有一个可以量化的侧面：gold 里有多少字符是「送给」模型的。"
  f"按插入表达式的字符数占 gold 全长的比例取中位数，合并档为 "
  f"{GDI['merge']['inserted_char_share_median']}%，拆分档为 "
  f"{GDI['split']['inserted_char_share_median']}%——合并档近四成的答案文本由提示直接给出。"
  "所以「C3mp 成功」里有一部分只是复制，不能整段归因于模型的理解。"
  "把这一部分单独剥离的现成实验是 c3mp_nat：同一批 230 题、同一模型、同一提示骨架，"
  "只把「可复制的段位表达式」换成等价的自然语言描述（例如「取合并列中分隔符之前的那一段」）。"
  f"结果为 {GDI_AB['c3mp_literal']['rate']}% 对 {GDI_AB['c3mp_nat']['rate']}%，"
  f"差 {GDI_AB['delta_pt']}pt，不一致对 {GDI_AB['only_literal']}:"
  f"{GDI_AB['only_natural']}，精确检验 p = {GDI_AB['mcnemar_p']:.0e}。"
  "两点都要读：表达形式本身值 14.8pt（能复制就复制），"
  "但自然语言条件下仍有 64.8%，远高于无谱系的 10.0% 与列级谱系的 2.6%——"
  "所以「理解段位」承担了主要增益，「照抄表达式」只是额外加成。")
P("人工核验没有做，这一点如实写明，不虚构第二位标注者，也不虚构 kappa。"
  "之所以可以不报 kappa，理由不是「省事」：本研究的标注单元是替换规则生成的 SQL 与算子，"
  "不存在需要人工判读的类别判断，Cohen's kappa 没有对象可测。"
  "但这不等于 gold 的正确性已被独立确认，二者的区别要在正文里说清。"
  "若要补做，可执行的协议是：从 230 条合并题中按库分层抽 60 条（每库 8–15 条，含全部 5 个事件），"
  "由不了解系统输出的数据库使用者独立写出「在新 schema 下回答同一问题的 SQL」，"
  "与 gold 在新库上比对执行结果；判为不一致的条目由第三人裁决，"
  "报告一致率、裁决率与不一致原因分类。"
  "该协议需要外部人员与时间，本轮不具备条件，故在 7.7.3 的边界中列为第一项未完成事项。")
CAP("表 7-137  Gold 同源性、表达形式消融与判定稳定性")
TABLE(["项", "口径", "结果"],
      [["gold 与替换规则逐字相同", "合并 230 / 拆分 104",
        f"{GDI['merge']['identical_pct']}% / {GDI['split']['identical_pct']}%"],
       ["gold 中被插入表达式的字符占比（中位数）", "合并 / 拆分",
        f"{GDI['merge']['inserted_char_share_median']}% / "
        f"{GDI['split']['inserted_char_share_median']}%"],
       ["表达形式消融：给表达式 vs 给自然语言", "230 题，同模型",
        f"{GDI_AB['c3mp_literal']['rate']}% vs {GDI_AB['c3mp_nat']['rate']}%"
        f"（{GDI_AB['only_literal']}:{GDI_AB['only_natural']}，"
        f"p = {GDI_AB['mcnemar_p']:.0e}）"],
       ["重复判定逐题一致率（全部可比键）", f"{RPC['scope']['keys_evaluated_more_than_once']} 个键",
        f"{RPC['scope']['item_agreement_pct']}%"],
       ["重复判定边缘率差（同条件、剔除非重复构造）",
        f"{RPC_S['n_judgements']:,} 次判定",
        f"中位 {RPC_S['median_abs_delta_pt']}pt，最大 {RPC_S['max_abs_delta_pt']}pt"],
       ["消融对照臂与主实验（同提示、同 230 题）", "两次独立批次",
        f"{RPC_F['item_agreement_pct']}% 逐题一致，"
        f"逐折差 ≤{RPC_F['max_abs_delta_pt']}pt"]],
      [5.6, 3.6, 4.8], ["l", "l", "l"], font=8.0)
NOTE("重复判定这一行回答的是「五轮重复实验最大偏差 0.1pt 是否可信」："
     "边缘率确实稳定（同条件跨批次的中位差为 0），但逐题判定并不确定——"
     f"全部可比键的逐题一致率 {RPC['scope']['item_agreement_pct']}%，"
     "即约百分之十几的题在两次判定间翻转，正负相抵后边缘率几乎不动。"
     "因此「偏差 0.1pt」应读作边缘率的稳定性，不能读作判定确定性。"
     "五轮重复实验的逐题输出当年没有归档，本表不能替代它，只能给出同类证据。")

# ================================================================ 表号按出现顺序重排
# 新增表在正文中以 7-101 起的占位号书写，避免与旧号冲突；此处按文档顺序统一编号，
# 并把全文（含正文、表注、表格单元格）中的引用一起改写。
def _renumber_tables(d):
    import re as _re
    order = []
    for p_ in d.paragraphs:
        m_ = _re.match(r"^表\s*7-(\d+)\s\s", p_.text.strip())
        if m_:
            n_ = int(m_.group(1))
            if n_ not in order:
                order.append(n_)
    mp_ = {old: i_ + 1 for i_, old in enumerate(order)}
    hit = {"ok": 0, "fallback": 0}

    def fix(txt):
        return _re.sub(r"表\s*7-(\d+)",
                       lambda m: "表 7-%d" % mp_.get(int(m.group(1)), int(m.group(1))),
                       txt)

    def do_par(p_):
        if "表 7-" not in p_.text:
            return
        n_before = len(_re.findall(r"表\s*7-\d+", p_.text))
        n_done = 0
        for r_ in p_.runs:
            if "表 7-" in r_.text:
                n_done += len(_re.findall(r"表\s*7-\d+", r_.text))
                r_.text = fix(r_.text)
        if n_done == n_before:
            hit["ok"] += 1
            return
        # 跨 run 的引用：退化为单 run，保留首个 run 的格式
        full = fix(p_.text)
        for r_ in list(p_.runs)[1:]:
            r_._r.getparent().remove(r_._r)
        if p_.runs:
            p_.runs[0].text = full
        hit["fallback"] += 1

    for p_ in d.paragraphs:
        do_par(p_)
    for t_ in d.tables:
        for row_ in t_.rows:
            for c_ in row_.cells:
                for p_ in c_.paragraphs:
                    do_par(p_)
    print("表号重排:", len(mp_), "条；单 run 完成", hit["ok"], "，跨 run 合并", hit["fallback"])
    return mp_


_renumber_tables(doc)

doc.save(OUT)
print("saved:", OUT)
print("段落:", len(doc.paragraphs), "表格:", len(doc.tables))
