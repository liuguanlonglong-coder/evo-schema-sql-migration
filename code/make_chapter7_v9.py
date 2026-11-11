# -*- coding: utf-8 -*-
"""第 7 章实验结果 修订版（v9）。

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
OUT = sys.argv[1] if len(sys.argv) > 1 else "投稿版_第7章_实验结果_v9.docx"

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
  "不涉及新的列匹配算法，也不以检索效果为贡献——真检索实验仅 40 题且为阴性对照，"
  "只在 7.7 节作为局限报告。与三项近邻工作的六维对比见 7.9 节表 7-29。")
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
  "而非替换标识符的迁移，本章未测；二是投影星号，本构造无法忠实评价（7.4.4）。")

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
     "凡以条目为单位的检验都需同时报告簇数，7.7 节给出。")

P("有效独立单元比条目数少得多，这一点必须在结论解读时计入："
  f"合并档 {units['merge']['items']} 条只对应 {units['merge']['op_facts']} 个"
  f"「列→段位表达式」映射事实（{units['merge']['db']} 个库、每库 2 个目标列），"
  f"平均每事实被 {units['merge']['items_per_op']} 条题重复使用；"
  f"拆分档 {units['split']['items']} 条只对应 {units['split']['op_facts']} 个 op，"
  f"平均每 op {units['split']['items_per_op']} 条；"
  f"改名档 {units['rename_all']['items']} 条对应 {units['rename_all']['op_facts']} 个改名映射、"
  f"{units['rename_all']['db']} 个库。因此合并与拆分的主效应，"
  "其证据强度由 5 个库和 10 个／5 个独立事实决定，而不是由 230／104 条题决定。")

H("7.1.5  统计口径", 3)
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
  "因此本章结论以整批比率的形式报告，不对单条判定作个别推断。")
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

# ================================================================ 7.2
B2 = json.load(open("results/stats_batch2.json", encoding="utf-8"))["task_level"]
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
  "真正拉开差距的是合并与拆分：B2/B3 在合并档分别降至 31.7% 与 40.0%，在拆分档同时为 0，"
  "因为它们无法表达「一列由多段构成」或「一列拆成多列」这一变化。")

P("这里必须回答一个审稿人已经提出的问题：相对于已有一族「按取值内容做列匹配」的方法，"
  "M1 的新意究竟在哪。诚实的分工是这样的。"
  "第一，那一族方法（以值域重叠、取值分布或样本值相似度为证据的实例级 schema 匹配）"
  "解决的是「两个 schema 之间哪一列对应哪一列」，输出是列对；"
  "本章的合并场景输出不是列对，而是「新列 + 段位」组成的可执行表达式，"
  "因为两个旧列会映射到同一个新列——列对形式在这一场景下信息不足，"
  "这不是精度差异，而是输出形态差异，问题因此不等价。"
  "第二，那一族方法多数面向跨库匹配（两套独立 schema），"
  "本章面向同一库的前后版本，可利用「演化算子只作用于少数列」这一先验来收缩搜索空间，"
  "这一先验也是 M1 在合并档能做到零误报的原因。"
  "第三，本章的 B1–B3 是名字与类型层面的朴素对照，不是那一族方法的实现，"
  "因此本章不主张「M1 在列匹配精度上优于实例级匹配方法」——这个比较本章没有做，"
  "也不该用 B1–B3 的结果代替。可主张的是：把值域与形态证据用于"
  "「恢复段位」而非「匹配列名」，是本章与那一族方法的分界点。")
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
CAP("表 7-101  规则基线与 LLM 条件在各档的执行等价率（%）")
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
  "不能读作本方法的性能。有信息量的比较只有两条："
  "c3r_infer（M1 独立抽取、不读标注）与 c3mp（LLM）在信息水平相同、执行器不同；"
  "以及 c3r_gt 与 c3r_infer 之间的差额，即自动发现的代价。"
  "由此得到的结论也因此必须收窄：它不证明「确定性改写普遍优于 LLM」，"
  "只证明在本构造下，瓶颈不在执行器而在谱系的发现与表示。"
  "要把它变成可推广的主张，需要旧 SQL 与新 SQL 并非机械对应的真实演化数据——本章没有。")

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
NOTE("失败类型在该档的全部失败条目中按条数排列；「无失败」表示该档规则侧 100% 执行等价。"
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
CAP("表 7-111  规则基线修正前后的执行等价率（%）")
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
  "为避免「写了但从未跑过」的质疑，另用一个两步演化的合成库做了探针（表 7-112）。")
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
  "其中 220 条（98.7%）确实引用了合并后的新列名，只有 5 条（2.2%）出现了任何切段构造"
  "（SUBSTR / INSTR / SPLIT_PART 等）。也就是说，失败几乎完全不在「列找错了」，"
  "而在「对着一列整体操作，没有表达它由多段构成」。")
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
P("有七点必须与结论一并声明，否则结论会被误读：")
PL("① 全部 11 个数据库都参与了方法开发与阈值调整，阈值曾依据 gold operation 设定。"
   "因此主实验是回顾性合成评测，不是干净的留出测试。"
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
PL("④ 合并与拆分只验证了两段场景，未覆盖三段以上、段内互含与多分隔符嵌套；"
   "构造工具在三段情况下会生成重复表达式，属已知硬限制。")
PL("⑤ 判定器对行序不敏感：gold 含 ORDER BY 时，行序错误的查询仍可能判为一致。"
   "该风险已在 7.1.7 量化：重执行 2,164 条预测后，被宽判的条目共 9 条（0.4%），"
   "全部集中在 gold 含 ORDER BY 的题；八个主对照的符号在五种判定规则下均不改变。"
   "但「行序完全正确」仍未被作为独立指标报告，这一项未闭环。")
PL("⑥ 所有条件均为单次运行，服务端不提供随机种子。同批重跑已观察到个别判定翻转，"
   "故结论以整批比率报告，不对单条判定作推断。")
PL("⑦ 真检索实验仅 40 题，且为阴性对照。本文没有实现检索模块，"
   "谱系是整块注入提示的，因此本章不对检索效果作任何主张。")

H("7.7.4  留出库确认实验", 3)
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

P("本轮把这层证据补上了。做法是把构造、清单生成与谱系推断三个脚本的库路径参数化，"
  "在 train 划分上独立构造，演化库与清单写入独立目录（data/evolved_*_holdout、"
  "data/manifest_holdout.jsonl）。参数在主实验完成后即冻结：段位表达式的生成规则、"
  "M1 的三条判定证据与阈值、提示模板、判定器与执行等价比较函数，全部沿用主实验的版本；"
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
CAP("表 7-25  留出库实验的样本构成（BIRD train 划分，未参与开发）")
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
CAP("表 7-26  留出库上的端到端迁移正确率（DeepSeek-V3.2）")
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

# ---------------------------------------------------------------- 7.7.5
H("7.7.5  取值编码的可逆性边界", 3)
P("评审提出一个此前没有正面回答的问题：主实验的同一条 SQL 在有限的数据行上结果相同，"
  "并不等于它在新值上也保持语义。段位表达式 SUBSTR(attr,1,INSTR(attr,CH)+k) 是否等价于原列，"
  "取决于合并后的取值是否满足某些前提；一旦不满足，查询照样能执行，结果却已经漂移。"
  "本节用两类证据回答：先在隔离环境里逐类造边界值，测出前提究竟有哪些；"
  "再回到真实数据，量一量这些边界在现构造中出现的频率。")
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
  "第二，本节评的是列引用而非执行等价，正确率数字与本章 EX 不可直接比较。"
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
  "留出库确认实验（7.7.4）在 7 个未参与开发的库上重现了方向，"
  "样本量（165 条）只够作旁证，不足以支持跨库泛化。"
  "因此本章的结论应读作「在这 5 次列合并与 5 次列拆分上，粒度信息带来量级显著的改善」，"
  "而不是「已确认的普遍规律」，更不能延伸为跨库泛化。"
  "本章不能支持的结论包括：自动谱系可普遍替代 oracle；"
  "显式谱系在一切演化类型上都优于让模型自行改写；"
  "改名场景下两者存在可判定的优劣或等价；"
  "方法可普遍迁移到未见数据库或真实历史 schema 版本（7.7.4 只给出 7 库 165 条的旁证）；"
  "以及检索能改善该任务——"
  "本章没有实现检索模块，谱系是整块注入提示的，若整篇论文仍以检索创新为主线，"
  "则与本章提供的证据不匹配，标题、摘要、关键词与贡献点须相应收敛。")
P("就创新性而言，本章的位置在 7.9 节已经界定：不主张首次研究 schema 演化下的 NL2SQL，"
  "也不主张首次处理合并/拆分或跨 schema 映射；可主张的是"
  "「在历史 SQL 作为语义证据的前提下，何种粒度的谱系信息足以避免可执行但语义错误的 SQL」。"
  "这是增量性贡献，其证据强度受 7.7.3 的七点边界与 7.8 的单元限制约束。")
P("未完成项需要在摘要、结论与投稿信中一致披露，共六项，每一项都能在正文找到对应的量化说明。其中五项是尚未做的实验，"
  "第六项是本构造自身的评价边界，二者都必须披露："
  "① 三段以上合并与分隔符嵌套——当前只在两段场景上验证；to_expr 已支持任意段数，"
  "但三段情形尚未在任何真实库上跑过，因此不写进结论；"
  "② 需要重写谓词结构而非替换标识符的迁移（口径改变、单位换算、等值谓词变区间）未测，"
  "因此不能声称语言模型在该类情形上有优势——这是 7.4.3 的第三条明确留下的空白；"
  "③ 构造未覆盖 NULL 与值内分隔符：真实数据扫描显示这两类行占比 0.00%（7.7.5、表 7-107），"
  "隔离探针显示一旦出现即不可恢复（表 7-106）；探针每类只有 1–3 个例子，"
  "只作机制反例，不能用来估计真实数据中的失败比例。主实验的绝对水平因此只适用于"
  "「两分量非空且不含分隔符」这一前提成立的取值分布；"
  "④ 无人工独立复核：编码由单作者完成，gold SQL 的正确性依赖构造期自检，"
  "本章未做双人编码，也未报告 kappa，附编码手册供复核；"
  "⑤ 检索实验未做：本章未实现检索模块，谱系是整块注入提示的，真检索实验仅 40 题且为阴性对照；"
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
  "留出库为 +10.9pt 与 +20.3pt；星号展开能力已实现并经合成探针验证；"
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
   "它不是干净的留出测试。留出确认取自 BIRD train 划分的 7 个未参与开发的库、165 条，"
   "只用于检查方向是否可迁移。")
PL("② 统计范围一律写「题目级比较」：显著性限于该评测集内的题目单位；"
   "合并与拆分各只有 5 次独立变更事件，按同一口径池化后 10 次事件 10/10 同向（p = 0.0020），"
   "但库级仍只有 5 个单元（p = 0.0625），库的效应与事件的效应无法分离，"
   "故不作跨库或真实演化的泛化声称。")
PL("③ 检索效果一律写「不作主张」：本章未实现检索模块，谱系是整块注入提示的，"
   "真检索实验仅 40 题且为阴性对照。")
PL("④ 留出库的方向必须带负向一起写：合并档按独立变更事件为 6 正 1 负，"
   "regional_sales 是负向库（C3mp 低于 C3m）。只写「方向重现」而不写这一条，"
   "等于把「多数库成立」说成「都成立」。")
PL("⑤ 规则基线的定位一律写「本构造下的配对差值」，不写「确定性方法优于 LLM」："
   "本章的 gold 由同一条替换规则机械生成，规则侧的上界被构造抬高；"
   "LLM 侧只测一个基座与一套提示模板。可写的是"
   "「在给定映射时，确定性改写的执行等价率高于所测模型 20 个百分点左右，"
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
        "在本构造（gold 由同一替换规则生成）下，给定映射时确定性改写的执行等价率高出所测模型"
        "约 20pt，且 LLM 在 dev 上无独占正确条目；该差值不外推为工程选型结论"],
       ["LLM 的增量在星号投影上",
        "该说法在 v8 中基于一次误判（COUNT(*) 被判为不可改写）成立，本版已撤销："
        "修正后规则侧在 dev 上零失败；留出合并档 LLM 独占的 4 条来自规则侧漏抽映射"],
       ["段位谱系在本任务上普遍更优",
        "仅在合并方向成立（证据效应 +99.1／+64.4pt），拆分方向为 0；"
        "留出库 regional_sales 出现负向，主张为「总体与多数库的方向」"]],
      [7.4, 8.8], ["l", "l"], font=8.4)
P("最后是主题匹配问题。本章实质研究的是 schema 演化下的历史 SQL 迁移与谱系粒度，"
  "不涉及检索系统设计、召回策略或证据冲突裁决。若整篇论文仍以「RAG + NL2SQL」为"
  "主线贡献，则本章的实验不足以支撑该定位，二者只能选一：要么补充真正的检索实验"
  "（多路召回、证据冲突与弃答判定），要么把标题、摘要、关键词与贡献点收敛为"
  "「历史 SQL 作为证据的跨 schema 迁移 + 谱系粒度」，检索只作为背景提及。"
  "本章建议后者——它与现有证据一致，也不需要新的模型调用。")

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
