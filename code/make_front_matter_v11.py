# -*- coding: utf-8 -*-
"""生成《摘要_标题_贡献点_结论_v11.docx》。

回应第九轮评审的五条意见与三项优先事项，全部落在章外文字层：

1. 摘要紧邻规则 100% 处加「在本文构造中／构造内上界」，不能只写在正文与修订说明；
2. M1 与实例级 schema matching 的边界补具体文献（文献条目在章内 7.3，本文件给核对清单）；
3. 规则基线的结论限定到当前构造——摘要、贡献点、结论、投稿信四处口径一致；
4. 星号探针定位为评测范围限制，不作能力证据（章内已改，本文件在限定条款里同步）；
5. 统一目标期刊为 Applied Intelligence，删除 JMS 口径；修正「约 400 字」与实测 576 字
   的内部不一致（英文亦由「约 230 词」更正为实测值）。

相对 v9 新增两节：投稿信（可直接投递）与投稿前待核清单（文献五篇 + AI 三项强制要求）。
"""
import json
import os
import re

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

BASE = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE)
OUT = "摘要_标题_贡献点_结论_v11.docx"

EN, CN, CN_H = "Times New Roman", "宋体", "黑体"

# 归档状态由 archive_info.json 单点驱动（与 prepare_archive.py 共用同一文件）。
# 不存在 = 未归档，投稿信与表 9-2/9-3 保持「将在发表前归档」；
# status=deposited = 已发布，自动切为完成时态并写入真实 DOI。
# 这样 DAS、投稿信、表 9-2、表 9-3 四处不会因手工改稿而分叉。
AI = {}
if os.path.exists("archive_info.json"):
    AI = json.load(open("archive_info.json", encoding="utf-8"))
ARCH_DONE = AI.get("status") == "deposited"
ARCH_DOI = AI.get("doi", "[repository DOI]")
ARCH_REPO = AI.get("repo_name", "[REPOSITORY_NAME]")
# 预留（reserved）阶段：DOI 已可写进待上传的 CITATION.cff / DAS.md 等文件，
# 但 record 尚未 Publish，稿件里的 DAS 与投稿信必须仍是「将归档」口径。
# 故稿件正文专用的 DOI 只在 deposited 时才换成真实值。
MANUSCRIPT_DOI = ARCH_DOI if ARCH_DONE else "[repository DOI]"

S = json.load(open("results/stats_ch7_v5.json", encoding="utf-8"))["main"]
RVR = json.load(open("results/rule_vs_llm.json", encoding="utf-8"))
G = RVR["groups"]
D = RVR["decomposition"]
HLD = json.load(open("results/holdout_stats.json", encoding="utf-8"))
EVO = json.load(open("results/evoschema_events.json", encoding="utf-8"))
DBD = json.load(open("results/data_bounds.json", encoding="utf-8"))
PRB = json.load(open("results/probe_counterfactual.json", encoding="utf-8"))
RWA = json.load(open("results/rule_rewrite.json", encoding="utf-8"))["audit"]
ABM = json.load(open("results/ablation_matching.json", encoding="utf-8"))
assert RVR["version"] == "rule-vs-llm-v2"
assert ABM["version"] == "ablation-matching-v1", "分段对照产物缺失（ablation_matching.py）"
_AB = ABM["groups"]
# 段位档：M1 的类型恢复（n = 15）——仍用于「识别出合并/拆分」这一口径
_EB = _AB["M1+形态指纹"]["段位档"]        # n = 15，type_ok = 15

# 统计独立性审计：条目 / op / 事件三层口径，与段位表达式的真实恢复率
ADI = json.load(open("results/audit_independence.json", encoding="utf-8"))
assert ADI["version"] == "audit-independence-v2", "独立性审计产物缺失（audit_independence.py，需 v2）"
AD_A, AD_B, AD_C, AD_D = (ADI["A_units"], ADI["B_segment"],
                          ADI["C_nonsegment"], ADI["D_segment_expr"])
AD_PS, AD_SG = AD_A["paper_set"], AD_A["segment_set"]
AD_IN, AD_SH = AD_C["no_shift_121"], AD_C["shift_26"]
_ABEV = AD_B["paired"]["B3 嵌入语义"]["event_level"]   # 上一版三元组事件级（n = 7，作废）
_ABOP = AD_B["paired"]["B3 嵌入语义"]["op_level"]      # op 级（高估，仅作对照）
# 第 14 轮：含演化档的正确事件聚类键（10 次）与段位感知公平基线
EMA = json.load(open("results/event_map_audit.json", encoding="utf-8"))
assert EMA["version"] == "event-map-audit-v1"
B4 = json.load(open("results/baseline_segment_aware.json", encoding="utf-8"))
assert B4["version"] == "segment-aware-baseline-v1"
# 第 17 轮：A 层三项复核（数值一律从产物读取，不手写常量）
MHO = json.load(open("results/m1_holdout_eval.json", encoding="utf-8"))
assert MHO["version"] == "m1-holdout-eval-v1"
JSA = json.load(open("results/judge_strict_audit.json", encoding="utf-8"))
assert JSA["version"] == "judge-strict-audit-v1"
EBV = json.load(open("results/encoding_bounds_v2.json", encoding="utf-8"))
assert EBV["version"] == "encoding-bounds-v2"
EBL = json.load(open("results/encoding_bounds_v2_legacy.json", encoding="utf-8"))
assert EBL["config"]["mode"] == "legacy_prefix"
MHO_OP = MHO["op_level"]
MHO_EV = MHO["event_level"]
MHO_DB = MHO["db_level"]
# 第 18 轮：5 次重复运行汇总 + 单作者编码/等效独立复核
RPT = json.load(open("results/llm_repeat_runs.json", encoding="utf-8"))
assert RPT["version"] == "llm-repeat-runs-v1", "重复运行产物缺失或版本不符"
RPT_P = RPT["provenance"]
# 章外统一改报正确口径：事件数 10、p = 0.0020
EVN = EMA["sign_correct"]["n"]          # 10
EVP = EMA["sign_correct"]["p_exact"]    # 0.001953
SGEV = EMA["counts"]["events_correct"]  # 10（段位档事件）

# 第 21 轮：定位修订所需的两组数字（第零节起即要用，故提前加载）
# ① 扩充合并档的端到端配对（事件级符号检验；数字全部从产物读取）
_XEM = json.load(open("results/ext_eval_merge_col.json", encoding="utf-8"))
_XPM = _XEM["paired_by_event"]["c3mp-c3m"]
# ② 检索式探针 E2：40 条、单库、阴性对照。数字从逐条输出重算，不手写。
import collections as _col
from math import comb as _comb

_E2ROWS = [json.loads(_l) for _l in
           open("results/e2_e2_main.jsonl", encoding="utf-8") if _l.strip()]
_E2BY = _col.defaultdict(dict)
for _r in _E2ROWS:
    _E2BY[_r["id"]][_r["cond"]] = _r["outcome"]
_E2SUM = json.load(open("results/e2_e2_main_summary.json", encoding="utf-8"))
_E2 = {
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
assert _E2["n"] == _E2SUM["n_samples"] == 40, "E2 逐条数与 summary 不一致"
assert _E2["c4_ex"] == _E2SUM["by_cond"]["c4"]["EX"]
assert _E2["c5_ex"] == _E2SUM["by_cond"]["c5"]["EX"]
_E2N = _E2["b"] + _E2["c"]
_E2["p"] = (1.0 if _E2N == 0 else
            min(1.0, 2 * sum(_comb(_E2N, _k)
                             for _k in range(0, min(_E2["b"], _E2["c"]) + 1)) / 2 ** _E2N))
_E2["c4_rate"] = 100.0 * _E2["c4_ex"] / _E2["n"]
_E2["c5_rate"] = 100.0 * _E2["c5_ex"] / _E2["n"]

# 第 22 轮：扩充集的四层计数与事件级/库级检验（与第 7 章同一产物、同一口径）
AUF = json.load(open("results/audit_units_v3.json", encoding="utf-8"))
assert AUF["version"] == "audit-units-v3"
AX_M = AUF["ext"]["merge_col"]
ACF = json.load(open("results/audit_cluster_v3.json", encoding="utf-8"))
assert ACF["version"] == "audit-cluster-v3"
ACE_ALL, ACE_ELI = ACF["ext_merge"]["event_level_all"], ACF["ext_merge"]["event_level_eligible"]
ACE_DB = ACF["ext_merge"]["db_level"]

# ---- 补充材料表号反查 -------------------------------------------------------
# 主文瘦身后，S 号由「未被主文保留的表按文档顺序」决定，任何中间插表都会整体平移。
# 因此这里不写死 S 号，改为按表题关键词反查；反查失败则直接报错，绝不静默用旧号。
_KW = {"strict": "全严规则", "enc": "段位编码压力测试", "hler": "H-LER",
       "rep": "重复运行的读数稳定性", "single": "单作者编码",
       "strat": "分层对比", "advice": "建议替换的表述"}
_SMAP = {}
if os.path.exists("补充材料_第7章_v11.docx"):
    _SD = Document("补充材料_第7章_v11.docx")
    for _p_ in _SD.paragraphs:
        _m_ = re.match(r"^表\s*S-(\d+)\s\s(.+)$", _p_.text.strip())
        if not _m_:
            continue
        for _k_, _kw_ in _KW.items():
            if _kw_ in _m_.group(2):
                _SMAP.setdefault(_k_, int(_m_.group(1)))
assert len(_SMAP) == len(_KW), ("补充材料表号反查失败，先运行 slim_chapter7.py："
                               f"{sorted(set(_KW) - set(_SMAP))}")
SN = lambda k: f"S-{_SMAP[k]}"
_NSUPP = max(_SMAP.values())


_n = lambda k: S[k]["n"]
N_ALL = _n("rename_all") + _n("merge") + _n("split")
_mg, _sp = S["merge"], S["split"]
_rm, _rs = G["merge"]["rates"], G["split"]["rates"]
_rhm, _rsh = G["merge_holdout"]["rates"], G["split_holdout"]["rates"]
_ev = HLD["event_level"]["merge_col_holdout"]
_gap_ho = _rhm["c3r_gt"]["rate"] - _rhm["c3r_infer"]["rate"]
_d_mg = _rm["c3r_infer"]["rate"] - _rm["c3mp"]["rate"]
_d_sp = _rs["c3r_infer"]["rate"] - _rs["c3mp"]["rate"]
_d_hm = _rhm["c3r_infer"]["rate"] - _rhm["c3mp"]["rate"]
P_mg = G["merge"]["paired"]["c3r_infer-c3mp"]
P_sp = G["split"]["paired"]["c3r_infer-c3mp"]
P_hm = G["merge_holdout"]["paired"]["c3r_infer-c3mp"]

doc = Document()
sec = doc.sections[0]
sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
sec.left_margin = sec.right_margin = Cm(2.4)
sec.top_margin = sec.bottom_margin = Cm(2.4)
st = doc.styles["Normal"]
st.font.name = EN
st.font.size = Pt(10.5)
st._element.rPr.rFonts.set(qn("w:eastAsia"), CN)
st.paragraph_format.line_spacing = 1.34
st.paragraph_format.space_after = Pt(4)


def _run(p, txt, bold=False, size=10.5, font=CN):
    r = p.add_run(txt)
    r.bold = bold
    r.font.size = Pt(size)
    r.font.name = EN
    r._element.rPr.rFonts.set(qn("w:eastAsia"), font)
    return r


def emit(p, txt, size=10.5):
    for j, part in enumerate(txt.split("**")):
        if part:
            _run(p, part, bold=(j % 2 == 1), size=size)


def H(txt, lvl=1):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(15 if lvl == 1 else 10)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.keep_with_next = True
    _run(p, txt, bold=True, size={1: 15, 2: 12, 3: 11}[lvl], font=CN_H)
    return p


def P(txt, indent=True, size=10.5):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    if indent:
        p.paragraph_format.first_line_indent = Pt(21)
    emit(p, txt, size)
    return p


def PB(txt, size=10.5):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_after = Pt(7)
    emit(p, txt, size)
    return p


def PL(txt):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Pt(19)
    p.paragraph_format.space_after = Pt(2)
    emit(p, txt, 10.3)
    return p


def CAP(txt):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(9)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.keep_with_next = True
    _run(p, txt, bold=True, size=9)
    return p


def NOTE(txt):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(8)
    emit(p, txt, 8.5)
    return p


def TABLE(head, rows, widths, aligns, font=8.6):
    t = doc.add_table(rows=1, cols=len(head))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    for i, cell in enumerate(t.rows[0].cells):
        cell.width = Cm(widths[i])
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        _run(cell.paragraphs[0], str(head[i]), bold=True, size=font)
    for row in rows:
        cs = t.add_row().cells
        for i, v in enumerate(row):
            cs[i].width = Cm(widths[i])
            p_ = cs[i].paragraphs[0]
            p_.alignment = {"l": WD_ALIGN_PARAGRAPH.LEFT, "c": WD_ALIGN_PARAGRAPH.CENTER,
                            "r": WD_ALIGN_PARAGRAPH.RIGHT}[aligns[i]]
            p_.paragraph_format.space_after = Pt(1)
            emit(p_, str(v), font)
    return t


def _m(v, nd=1):
    return f"{v:+.{nd}f}".replace("-", "−")


# ===================================================================== 封面
p_ = doc.add_paragraph()
p_.alignment = WD_ALIGN_PARAGRAPH.CENTER
_run(p_, "摘要、标题、贡献点、结论与投稿信（v11）", bold=True, size=17, font=CN_H)
p_ = doc.add_paragraph()
p_.alignment = WD_ALIGN_PARAGRAPH.CENTER
_run(p_, "第十轮：主张重摆为「谱系粒度决定成败」，并补 M1 相对匹配基线的分段对照",
     size=10, font=CN)
NOTE("目标期刊：**Applied Intelligence**（Springer，ISSN 0924-669X）。"
     "本文件所有字数均为脚本实测值，不是估计值；"
     "末尾附「数字 → 出处」自检表与「投稿前待核清单」，两处都需逐行核对后再粘贴。"
     "正文证据见《投稿版_第7章_实验结果_v11.docx》。"
     "本版的核心变化不是加数字，而是换主句：全文第一位的结论从「我们构造基准并比较两种执行器」"
     "改为「同一执行器下，谱系的粒度决定迁移成败」。"
     "第 21 轮另加第零节：把**研究类型**写进标题（An Empirical Study of …），"
     "并逐处清理 RAG 与「NL2SQL 系统」的遗留口径——"
     "因为现有检索证据只有 40 条、单库、阴性对照，撑不起一个系统类主张。")

# ===================================================================== 零
H("零、定位：这是一项实证研究，不是一个系统")
P("第 21 轮只做一件事：把研究类型写在最显眼的位置，并清掉与它不一致的旧口径。"
  "起因是外部意见指出，本项目的检索证据只有 40 道题、属于阴性对照，且没有实现完整检索模块；"
  "在这种情况下若把 RAG 放进标题、摘要或主要贡献，审稿人一旦翻到实验部分就会判定"
  "「研究问题与实验内容不一致」——这与数值是否正确无关，属于定位错误。")
P("因此本文的类型声明改为三句话：① 标题以 An Empirical Study of 开头，"
  "声明产出是证据而非系统；② 摘要首段交代研究类型与「不实现检索模块」；"
  "③ 三条贡献点全部写成实证发现，不出现「我们提出／构建了某系统」。"
  "下列两张表分别是「本文主张什么」与「旧口径在哪些位置泄漏」，两表都必须与正文一致。")
CAP("表 0-1  本文主张与不主张的六件事")
TABLE(["主张", "是否成立", "依据", "处置"],
      [["提出 RAG-NL2SQL 系统", "否",
        f"无检索模块，谱系整块注入提示；唯一检索式探针 {_E2['n']} 条"
        f"（改名档、单库 {_E2['db']}）、p = {_E2['p']:.2f}",
        "标题、摘要、关键词、贡献点一律不出现"],
       ["提出端到端 NL2SQL 生成方法", "否",
        "生成侧只用单一基座与一套提示模板，未训练、未调优、无多模型集成",
        "不主张方法创新"],
       ["给出谱系抽取器 M1", "有限",
        f"段位档 {AD_SG['ops']}/{AD_SG['ops']} 对基线 0/{AD_SG['ops']}；"
        f"字符级表达 10/10 与段位感知基线持平；值改写两档落后 "
        f"{AD_SH['ours_ok']}/{AD_SH['n_ops']} 对 {AD_SH['base_ok']}/{AD_SH['n_ops']}",
        "只写「段位专用抽取器」，不写「更好的匹配器」"],
       ["给出粒度效应的受控证据", "是（主贡献）",
        f"同一执行器下 {_rm['c3r_name']['rate']:.1f}% 对 {_rm['c3r_infer']['rate']:.1f}%；"
        f"扩充构造后事件级 {_XPM['pos_events']} 正 {_XPM['neg_events']} 负"
        f"（n={_XPM['n_events']}，p={_XPM['p']:.3f}）",
        "置于贡献点首位，且限于合并方向"],
       ["给出确定性重写基线", "是（佐证）",
        "gold 由同一条替换规则生成，其 100.0% 是构造内上界",
        "不作工程选型结论"],
       ["评测了一种检索方法", "否",
        f"{_E2['n']} 条、{_E2['n_db']} 个数据库的阴性对照，无召回策略、无多路证据、无冲突裁决",
        "只在 7.7.3 作局限披露，不进摘要与贡献点"]],
      [3.0, 1.6, 7.2, 4.4], ["l", "c", "l", "l"], font=8.2)
NOTE("第四行是唯一可以放在标题里的主张；其余五行要么是佐证，要么必须消失。"
     "第三行的「有限」不是谦辞：M1 只在段位这一件事上不可替代，"
     "在不需要段位的 121 个算子上与嵌入基线同为 100%，在值改写两档明显更弱。")
CAP("表 0-2  旧口径的七个泄漏点与现状")
TABLE(["位置", "旧口径", "现状"],
      [["标题", "曾以问句「谱系要细到哪一层」作主标题，未声明研究类型",
        "改为 An Empirical Study of Schema Lineage Granularity for Historical SQL "
        "Migration under Schema Evolution"],
       ["英文摘要", "未交代研究类型；「不实现检索模块」只在末段以否定句出现一次",
        "首段写明 empirical study 与 no retrieval module，末段保留原否定句作呼应"],
       ["中文摘要", "同上", "首段写明「一项实证研究」与「不实现检索模块」"],
       ["关键词", "早期版本含「检索增强生成」", "5 个关键词全部指称正文实存内容，检索类词列入删除候选"],
       ["贡献点", "以「我们构建／提出」开头，读者会期待一个系统",
        "三条均改写为实证发现；另立「定位」行负向声明不主张系统"],
       ["结论", "未交代研究类型", "首句写明「作为一项实证研究」"],
       ["投稿信", "主题句为 an applied AI study，未说明不做系统",
        "改为 empirical study，并在第二段加一句 no system 声明"]],
      [1.8, 6.0, 8.4], ["l", "l", "l"], font=8.2)
NOTE("第 7 章内部的四处相关表述（7.1.1、7.7.3 ⑦、7.10、7.11 ③）已同轮改写："
     "原句是「若整篇论文仍以 RAG 为主线，二者只能选一」的待决口吻，"
     "现改为已决断，并把检索式探针写成 "
     f"{_E2['n']} 条 / {_E2['n_db']} 库 / 阴性对照（p = {_E2['p']:.2f}）的完整口径。"
     "保留一个待决问句在正文里，等于把定位问题的处置权交还给审稿人。")

# ===================================================================== 一
H("一、第十轮要解决的三件事")
P("这一轮不是补数字，而是换主句。第九轮之后实验已经够用，但主张的位置错了："
  "摘要首句仍在讲「比较两种执行器」，而数据里最有说服力的对比是"
  "「同一执行器、两种谱系粒度」。三件事里的第一件来自自查，"
  "它同时推翻了我上一轮的一条审稿意见。")
TABLE(["问题", "现状", "本轮改法", "不做的后果"],
      [["M1 相对匹配基线",
        "三条基线（B1 名字、B2 值域重叠、B3 嵌入）早已实现并跑过，题级恢复率此前已有报告；"
        "但题级口径把「映射到哪个列」与「这一列由哪几段构成」合成了一个数",
        "按 op 级重报，拆成「目标列恢复」与「op 类型一致」两列，另立一张 op 级质量表；"
        f"三条基线在段位档均为 0/{AD_SG['ops']}，M1 为 {AD_SG['ops']}/{AD_SG['ops']}，"
        f"事件级 p = {EVP:.4f}（{EVN} 次独立事件，按含演化档的聚类键）",
        "B2 在合并档的 31.7% 会被读成「精度不足」而非「能力缺失」；"
        "而这一读法会让 C3m 与 C3mp 之间 77.0pt 的差距显得无法解释"],
       ["主张位置",
        "首句为「比较两种执行器」，规则 100% 承担全部卖点",
        "首句改为「同一执行器下，谱系的粒度决定迁移成败」；"
        "规则 100% 降为佐证，且必须与列级 0.9% 成对出现",
        "构造循环性一旦被看穿，全文最强的那个数字同时失效——"
        "而它本来只是同义反复"],
       ["分析单位",
        "条目、op、变更事件、库、映射单元五种单位混用，每处都有说明，但没有一条选择规则",
        "7.1.5 写明固定规则：描述性按条目，推断按最粗可用单位，条目级显著仅为必要条件",
        "审稿人会问「主结论以哪个单位为准」，并怀疑按显著性择优报告"]],
      [2.7, 5.8, 5.2, 4.0], ["l", "l", "l", "l"], font=8.3)
NOTE("第三件事的规则不是新解释，而是把已经发生过的一次降级写成通则："
     "EvoSchema 外部对照的条目级 p = 0.0018 被降为映射单元级 p = 0.5，正文据此改写了结论。"
     "规则若是事后为了自洽而定的，读者能看出来；这里能指出它第一次生效的位置，所以站得住。")

# ===================================================================== 二
H("二、标题：三案与取舍（第 21 轮改为以研究类型开头）")
P("定位改变后，标题的第一个成分必须是研究类型而不是研究对象。"
  "理由很直接：标题若写对象（RAG、NL2SQL、检索增强），读者会去实验部分找那个系统，"
  "找不到就判「声称与内容不符」；标题若写研究类型（An Empirical Study of），"
  "读者的预期变成「你分离出了哪个变量、证据有多强」，而这恰好是本文能交付的东西。")
PL("**A（推荐，投稿用）** An Empirical Study of Schema Lineage Granularity for "
   "Historical SQL Migration under Schema Evolution")
PL("　　中文对照：schema 演化下历史 SQL 迁移的谱系粒度实证研究")
PL("**B** How Fine Must Lineage Be? Granularity Effects in Migrating Historical SQL "
   "under Schema Evolution")
PL("**C** One Executor, Two Granularities: Segment-Expression Lineage for Historical SQL")
NOTE("A 的四个成分各对应正文一处，没有装饰词："
     "**An Empirical Study of** 声明研究类型（第零节表 0-1 的第六行因此不成立）；"
     "**Schema Lineage Granularity** 是唯一自变量（7.4、7.5）；"
     "**Historical SQL Migration** 是任务，刻意不写 NL2SQL 生成（7.1.1 已划清与 Text-to-SQL 的界）；"
     "**under Schema Evolution** 是前提条件而非贡献（7.2 的构造）。"
     "B 是上一轮的推荐，本轮降为备选：问句标题会让审稿人期待「你打算怎么做到」，"
     "而本文既不实现检索也不实现生成系统，问句容易读成承诺未兑现；"
     "C 最能体现「控制了执行器」，但它并列了两个成分、且仍带问句色彩。"
     "三案均不含 RAG、NL2SQL 与医疗字样——"
     "副标题「——以电子病历结构化字段为例」不建议加，因为本研究没有医疗数据；"
     "英文主标题 15 词，在 Applied Intelligence 的常见长度内。"
     "若改用 B 或 C，投稿信第一段引号内与摘要首句须同步替换，三处不能只改一处。")

# ===================================================================== 三
H("三、英文摘要（实测 __ENWORDS__ 词，Applied Intelligence 要求 150–250 词）")
_XEXT = json.load(open("results/ext_events_summary.json", encoding="utf-8"))["by_task"]

_EN_ABS = [
    ("We report an empirical study of schema lineage granularity for historical SQL "
     "migration under schema evolution. When columns are merged into one delimited "
     "column, column-level lineage returns the whole string: the migration executes but "
     "is semantically wrong. We build no retrieval module and no NL-to-SQL system; "
     "lineage enters the prompt whole."),
    ("On 11 BIRD databases we build a 610-item ladder (90 operations, 61 events) under "
     "retrospective synthetic evolution; all informed development, so the main study is "
     "not held out."),
    ("With the executor fixed, merged items reach 0.9% execution-result agreement with "
     "column-level and 100.0% with segment-level lineage, a property of the representation "
     "rather than the executor. A prompted model reaches 79.6%; the rewriter's 100.0% is "
     "construction-internal, its gold SQL using the same substitution. Splits gain "
     "nothing (77.9%)."),
    ("M1 recovers operation type and target column on 15/15 segment-bearing operations "
     "versus 0/15 for matchers unable to emit segment expressions: expressiveness, not "
     "accuracy. Character-level agreement (10/10, five events) matches a segment-aware "
     "baseline, so no algorithmic gain is claimed; M1 loses under value recoding "
     "(53.8% vs 100%)."),
    (f"Scope is the evaluated SQLite subset and three operators. Held-out confirmation "
     f"covers 165 items on 7 databases: 6 events recur, 1 reverses; a recovery study "
     f"(481 items, 8 databases) reaches 14/16. Expanding merge to "
     f"{AX_M['constructed']} constructed events is event-level significant "
     f"({ACE_ELI['pos']}/{ACE_ELI['n_events']} qualified, p<1e-6); 5 events cannot. "
     f"A {_E2['n']}-item retrieval probe is a negative control (p={_E2['p']:.2f})."),
]
for _seg in _EN_ABS:
    PB(_seg)
_EN_N = sum(len(re.findall(r"[A-Za-z0-9][A-Za-z0-9%\.\-–']*", s)) for s in _EN_ABS)
NOTE("__ENNOTE__")


# ===================================================================== 四
H("四、中文摘要（实测 __CNCHARS__ 字）")
_CN_ABS = [
    ("本文是一项关于谱系粒度的实证研究，不是系统论文：不实现检索模块，也不实现端到端 NL2SQL 生成，"
     "谱系整块注入提示，故不主张任何系统能力。schema 演化会使历史 SQL 失效："
     "当多列被合并为带分隔符的单列时，列级映射会把整串返回，迁移结果可执行却语义错误。"
     "要问的核心问题不是用哪个模型重写，而是谱系必须细到哪一层。"),
    ("我们在 11 个 BIRD 数据库上构造 610 条受控样本（改名 276、列合并 230、列拆分 104），"
     "条件阶梯从「仅新 schema」逐级升到指明段位位置的可执行表达式谱系（SUBSTR/INSTR）。"
     "执行器固定不变：由确定性重写器沿谱系替换列引用；"
     "另一侧比较 M1（推断阶段不使用标签、阈值由开发集校准）与名字、值域重叠、嵌入三条匹配基线。"),
    (f"在同一执行器下，合并档在所评估 SQLite 实例上的执行结果一致率为列级谱系 {_rm['c3r_name']['rate']:.1f}%、"
     f"段位表达式谱系 {_rm['c3r_infer']['rate']:.1f}%——差距来自谱系本身，而非执行器。"
     f"同一份段位谱系交给语言模型时为 {_mg['c3mp']['rate']:.1f}%；"
     f"重写器的 {_rm['c3r_infer']['rate']:.1f}% 是构造内上界"
     f"（gold SQL 由同一条替换规则机械生成），不作为独立测得的性能。"
     f"拆分方向不受益于段位信息（{_sp['c3mp']['rate']:.1f}%）。"),
    (f"在需要段位的 {_EB['n']} 个算子上 M1 全部识别出「这是一次合并或拆分」"
     f"（{_EB['type_ok']}/{_EB['n']}），三条基线均为 0/{_EB['n']}；"
     f"但基线无法表达段位，故这是表达能力而非精度差异，"
     f"按 {EVN} 次独立事件检验为 p = {EVP:.4f}"
     f"（按算子计为 {_ABOP['p']:.1e}，属高估）。"
     f"在不含值改写的非段位档上 M1 与嵌入基线持平（各 {AD_IN['n_ops']}/{AD_IN['n_ops']}），"
     f"仅在值改写两档落后（{AD_SH['ours_ok']}/{AD_SH['n_ops']} 对 "
     f"{AD_SH['base_ok']}/{AD_SH['n_ops']}）。"
     "故 M1 是段位专用抽取器，在通用列匹配上与语义方法等价，在值改写档上更弱。"),
    (f"适用范围必须与结果同时给出。评测性质：主实验是回顾性合成评测——"
     "数据库内容取自 BIRD，但改名、合并与拆分由受控脚本施加，且 11 个库都参与过方法开发"
     "与阈值校准，故它不是干净的留出测试；SQL 范围限于所评估的 SQLite 子集与这三类演化算子，"
     "CTE、复杂子查询、JOIN USING/NATURAL JOIN、聚合与 HAVING、投影星号、"
     "位置 ORDER BY、显式类型转换及其他方言均未纳入。"
     f"单元口径：{AD_PS['items']} 条背后只有 {AD_PS['ops']} 个演化算子、"
     f"{AD_PS['events']} 次独立演化事件，段位档 {AD_SG['ops']} 个算子更只对应 "
     f"{SGEV} 次事件，故合并与拆分两档一律以事件为单位复核。"
     "这一功效瓶颈只能靠构造解决：按表分片重建后，合并档构造事件由 5 次扩至 "
     f"{AX_M['constructed']} 次（有题 {_XEXT['merge_col']['events']}、有效 "
     f"{_XEXT['merge_col']['events_effective']}、合格 "
     f"{_XEXT['merge_col']['events_ge3']}），拆分档 {_XEXT['split_col']['events']} 次"
     f"（合格 {_XEXT['split_col']['events_ge3']}），值改写档 "
     f"{_XEXT['value_shift']['events']} 次（有效 "
     f"{_XEXT['value_shift']['events_effective']}）；四层口径不可相加。"
     "在扩充集合并档上重跑端到端后，段位谱系优于列级谱系的合格事件为 "
     f"{ACE_ELI['pos']} 正 {ACE_ELI['neg']} 负（n={ACE_ELI['n_events']}），"
     f"有题事件 {ACE_ALL['pos']} 正 {ACE_ALL['neg']} 负另 {ACE_ALL['tie']} 持平，"
     f"库级 {ACE_DB['pos']}/{ACE_DB['n_dbs']}；这是事件层面首次取得显著性——"
     "10 次事件时最小可达 p 为 0.0625，无论方向多稳定都不可能越过 0.05 线。"
     "编码边界已逐类测定：空值分量与非末段内含分隔符不可恢复，"
     "空串、末段内含分隔符、首尾空白、多字节字符四类可恢复；"
     "21.5 万行真实数据扫描中前两类出现 0 次，故零越界应读作构造未触发。"
     "留出侧的两项实验范围不同、不可相加：端到端留出确认 165 条、7 个库、"
     "12 个变更事件单元，合并档方向 6 正 1 负；谱系表达式直接恢复 481 条、"
     f"{MHO_OP['n_units']} 个算子、{MHO_EV['n_units']} 次事件、{MHO_DB['n_units']} 个库，"
     f"M1 为 {MHO_OP['matched']}/{MHO_OP['n_units']}（{MHO_OP['rate_pct']:.1f}%），"
     "故两者既不合并报告，也不互为分母；本文未实现检索模块。"
     f"三项独立复核进一步界定适用范围：全严判定规则下 {JSA['sign_stable_count']} 个主对照符号不变"
     f"（最大偏移 {JSA['max_abs_delta_pt']:.1f}pt）；"
     f"2–6 段编码压力测试定位并修复了末段表达式缺陷（末段含分隔符可恢复率 "
     f"{EBL['by_case']['sep_in_last']:.1f}% → {EBV['by_case']['sep_in_last']:.1f}%），"
     f"全部真实 gold 表达式均为两段，历史数值零回归"
     f"（详见补充材料表 {SN('strict')} 至 {SN('enc')}）。"
     f"核心语言模型条件另独立重跑 {RPT['n_repeat_runs']} 次，各条件读数相对首次报告值的"
     f"最大偏差为 {RPT['max_abs_dev_pt']:.1f} 个百分点，主对照方向与显著性判定均不变；"
     "标注采用单作者编码并附等效独立复核清单，gold 由确定性程序生成、无人工判读单元，"
     f"故不报告 kappa（详见补充材料表 {SN('rep')} 至 {SN('single')}）。"
     f"定位上还需交代一句：本文不涉及检索系统。唯一的检索式探针为 {_E2['n']} 条、"
     f"{_E2['n_db']} 个数据库的阴性对照（C4 {_E2['c4_rate']:.1f}% 对 C5 {_E2['c5_rate']:.1f}%，"
     f"不一致对 {_E2['b']}:{_E2['c']}，p = {_E2['p']:.2f}），"
     "只能说明「补一个检索到的过期示例不足以救回迁移」，不能评价召回、排序或多路证据，"
     "故全文不以检索为研究问题，标题与关键词亦不含检索类术语。"),
]
for _seg in _CN_ABS:
    PB(_seg)
_CN_N = sum(len(re.sub(r"\s", "", s)) for s in _CN_ABS)
NOTE("__CNNOTE__")


# ===================================================================== 五
H("五、关键词与贡献点")
P("关键词 5 个，全部指称正文真实存在的内容：")
PL("schema 演化；历史 SQL 迁移；表达式级谱系；确定性查询重写；执行结果一致率")
NOTE("删除候选：RAG、检索增强生成、NL2SQL 端到端、提示工程——这些词所指的模块在本研究中不存在。"
     "Applied Intelligence 要求 4–6 个关键词，此处 5 个在区间内。"
     "第六个关键词曾考虑补 empirical study。最终不加——研究类型已由标题首位承载，"
     "重复一次等于占掉一个本可用于检索性的词位；若编辑或审稿人要求补足 6 个，"
     "按推荐顺序补「controlled evaluation」，仍不补任何检索类词。"
     "检索类词在本版是硬禁区：正文 7.7.3 的检索式探针只有 "
     f"{_E2['n']} 条、{_E2['n_db']} 个数据库、p = {_E2['p']:.2f}，"
     "写进关键词等于把阴性对照升格为研究主题。"
     "第五个关键词原本写作「执行等价」，本轮改为「执行结果一致率」："
     "正文的操作性指标是 execution-result agreement，"
     "若关键词写「等价」，等于在检索层面声称本文做了形式化语义等价证明——"
     "而 7.1.3 已明确那是未做的。关键词、摘要、正文三处的措辞必须同源。")

P("贡献点四条：第一条是新增的定位声明，用负向写法把「不主张什么」先钉住；"
  "其余三条重写，因为主句变了，粒度证据升为首位、规则基线退为佐证、第二条改为带边界的正面贡献。")
TABLE(["", "上一版的写法", "本版写法"],
      [["定位",
        "未声明研究类型；三条贡献均以「我们构建／提出」开头，读者会期待一个系统",
        "本文是实证研究，不是系统论文：不提出 RAG-NL2SQL 系统、不实现检索模块、"
        "不做端到端生成方法；谱系整块注入提示，"
        f"唯一检索式探针 {_E2['n']} 条、{_E2['n_db']} 个数据库、阴性对照"
        f"（p = {_E2['p']:.2f}），只作局限披露，不进贡献"],
       ["贡献一",
        "给出粒度命题的受控证据：合并档由 "
        f"{_mg['c3m']['rate']:.1f}% 升至 {_mg['c3mp']['rate']:.1f}%（证据效应 "
        f"{_m(D['merge']['evidence_effect_pt'])}pt）",
        "以受控实证设计把「谱系粒度」分离为独立变量（不提出系统）：固定同一执行器，"
        f"合并档列级谱系 {_rm['c3r_name']['rate']:.1f}%、段位表达式 "
        f"{_rm['c3r_infer']['rate']:.1f}%。两侧执行器相同，差距只能归于谱系表示——"
        "这是全文唯一不受构造循环性影响的结论。拆分方向效应为零，故命题限于「一列内多段编码」；"
        "并把该效应落到失败形态上：C3m 的 223 条静默错误里 220 条仍引用合并后的新列、"
        "仅 5 条含切段构造，而 C3mp 的 47 条静默错误里 45 条已含切段构造、失败点下移到取值与谓词层——"
        "粒度改变的不是「找不找得到列」，而是错误留在哪一层」；"
        "为摆脱事件数不足的功效限制，按表分片把合并档事件由 5 扩到 "
        f"{_XEXT['merge_col']['events_effective']} 个有效事件，同一方向在事件层面达到 "
        f"{_XPM['pos_events']} 正 {_XPM['neg_events']} 负（n={_XPM['n_events']}，"
        f"p={_XPM['p']:.3f}），是本章唯一一次事件级显著性"],
       ["贡献二",
        "把误差来源定位到「谱系能否被正确发现」，未见库留下 "
        f"{_gap_ho:.1f}pt 缺口",
        "给出唯一能产出段位表达式的抽取器 M1（推断阶段不使用标签、阈值由开发集校准）：在需要段位的 "
        f"{AD_SG['ops']} 个 op 上全部识别出合并或拆分（三条基线均为 0/{AD_SG['ops']}，"
        f"按 {EVN} 次独立事件检验 p = {EVP:.4f}），且对需要段位的 "
        f"{AD_D['part_covered_ops']} 个合并 op 的段位表达式 {AD_D['part_expr_match']}/"
        f"{AD_D['part_covered_ops']} 与 gold 一致（基线输出空间不含段位表达式）；"
        f"其边界——不含值改写时与嵌入基线持平（各 {AD_IN['n_ops']}/{AD_IN['n_ops']}），"
        f"仅值改写两档落后（{AD_SH['ours_ok']}/{AD_SH['n_ops']} 对 "
        f"{AD_SH['base_ok']}/{AD_SH['n_ops']}），故 M1 是段位专用抽取器，"
        f"不是更好的通用匹配器。未见库上自动谱系另有 {_gap_ho:.1f}pt 的硬失败缺口"],
       ["贡献三",
        f"给出确定性重写这条强基线：高出所测模型 {_d_mg:.1f}pt（合并）",
        "量化执行器效应并说明其构造内性质：同一段位谱系下规则改写高出所测模型 "
        f"{_d_mg:.1f}pt（合并）与 {_d_sp:.1f}pt（拆分），但 gold SQL 由同一条替换规则生成，"
        "该差值只衡量本构造下的执行器差异，且只覆盖谓词结构不变、仅需替换列引用的迁移；"
        f"表示效应（{_m(D['merge']['evidence_effect_pt'])}pt）约为它的 "
        f"{abs(D['merge']['evidence_effect_pt'] / D['merge']['executor_effect_pt']):.1f} 倍"],
       ["作用域", "证据来自两段可逆拼接的合成编码；未覆盖空值与值内分隔符",
        "同上；另加：留出确认 7 库 165 条且含一个负向库（regional_sales）；"
        "未实现检索；投影星号题在本构造下无法被忠实评价"]],
      [1.4, 5.0, 9.8], ["c", "l", "l"], font=8.3)
NOTE("贡献二的写法在本轮又改了一次，而且是往更不利的方向改。v10 只说「把误差定位到谱系发现」，"
     "读者无法判断 M1 本身值多少；v11 给出 15/15 对 0/15 的对照。"
     "但本轮核对判定函数后确认：那个 15/15 度量的是「识别出这是合并或拆分」，"
     "因为判定前 gold 的目标表达式已被归一为纯列名，段位字符串从未进入比较。"
     "因此同一段必须再补一句——段位表达式本身的字符级一致率为 10/10，"
     f"覆盖合并档全部 {EMA['counts']['events_correct_by_evo']['merge_col']} 次独立事件，"
     "该事件级符号检验 p = 0.0625，已等于 n = 5 时的最小可达值，"
     "故不能作为「抽取能力显著」的证据；段位价值的统计证据不在抽取环节，而在端到端环节。"
     "同理，上一版写的「通用匹配上不如嵌入基线的 91.8%」也是错的："
     "拆开后 610 条主实验集内的 121 个算子两者持平于 100%，"
     "缺口 12 个全部来自集外的值改写两档。"
     "两处都改成了更弱、但经得起复算的表述——这比保留一个好看的数字重要。")

P("相关工作与本任务的边界，用一张四列表比三句话更清楚。这张表建议直接贴进引言相关工作末尾，"
  "四个维度分别回答「输入是什么」「输出是什么」「用不用历史 SQL」「能不能表达列内段位」。"
  "第四列是本任务与前三条工作线的真正分界：不是他们做得不够准，而是他们的输出空间里"
  "没有「一段取自哪一位」这种对象。填表时不写「已有方法都无法处理段位表达式」这类总括句——"
  "那是一个超出检索范围的断言；只写各条工作线的输出形态，把差别留在第四列。")
CAP("表 5-1  三条相关工作线与本任务的边界（建议贴入引言相关工作末尾）")
TABLE(["工作线", "输入", "输出", "用历史 SQL", "表达列内段位"],
      [["schema 匹配（Rahm 与 Bernstein 2001；Cupid 2001；COMA 2002；Kang 与 Naughton 2003）",
        "两套 schema 及其实例统计",
        "元素之间的对应关系，多为列对；复杂匹配扩展到子句",
        "否", "否"],
       ["schema 演化的查询重写（PRISM/PRISM++ 一类）",
        "旧查询 + 已知的演化算子",
        "适配新 schema 的查询；替换粒度为整列",
        "是", "否"],
       ["Text-to-SQL 的 schema 扰动鲁棒性（EvoSchema 等）",
        "自然语言问题 + 当前 schema",
        "新生成的查询；题目从自然语言出发",
        "否", "否"],
       ["本任务（谱系粒度）",
        "旧 SQL + 演化后 schema + 自动抽取的谱系证据",
        "迁移后的查询；证据可为列级或段位表达式",
        "是", "是（SUBSTR/INSTR 形式的段位表达式）"]],
      [4.6, 3.4, 4.4, 1.9, 3.4], ["l", "l", "l", "c", "l"], font=8.2)
NOTE("英文版对应段落（可直接用于 Introduction 末尾）："
     "Prior schema-matching work primarily estimates correspondences between schema elements, "
     "while schema-evolution rewriting focuses on adapting queries to changed schemas. "
     "This study isolates a narrower question: when an old column is embedded as a segment of "
     "a merged column, does column-level lineage retain enough information to migrate "
     "historical SQL?")
NOTE("两点提醒。其一，第四列写「否」不是贬低，而是陈述输出空间的形状——"
     "列级匹配把两个旧列映射到同一个新列时会折叠成一对多，信息在这一步就已经丢失；"
     "其二，表内每一项都须与引言正文的引文一一对应，若某处引文最终被替换，"
     "本表对应行也要同步改写，不能只改正文。")


# ===================================================================== 六
H("六、结论段")
PB(f"作为一项实证研究，本文的证据支持一个窄命题：当 schema 演化把多列合并为带分隔符的一列时，"
   f"谱系粒度是迁移成败的决定因素。固定同一个确定性改写器，"
   f"列级谱系只到 {_rm['c3r_name']['rate']:.1f}%，段位表达式谱系到 "
   f"{_rm['c3r_infer']['rate']:.1f}%；把同一份段位谱系交给语言模型则到 "
   f"{_mg['c3mp']['rate']:.1f}%。三条路径共用一套谱系与一套判定，"
   f"唯一变化的是表示与执行器，而表示的效应（{_m(D['merge']['evidence_effect_pt'])}pt）"
   f"约为执行器效应（{_m(D['merge']['executor_effect_pt'])}pt）的 "
   f"{abs(D['merge']['evidence_effect_pt'] / D['merge']['executor_effect_pt']):.1f} 倍。")
PB(f"由此得到三条对后续工作更有用的读数。其一，该任务的瓶颈在谱系发现："
   f"自动抽取在开发库上与构造映射逐条等价，在未见库上留下 {_gap_ho:.1f}pt 的硬失败缺口，"
   f"且缺口全部表现为执行报错而非静默错。"
   f"其二，确定性重写是一条必须报告的强基线，但它的 "
   f"{_rm['c3r_infer']['rate']:.1f}% 是构造内上界——gold 由同一替换规则生成——"
   f"因此只能用于说明「给定映射后执行器不是瓶颈」，不能用于工程选型。"
   f"方法一侧的定位同样需要收窄，且分三档而不是一概而论：M1 在需要段位的 "
   f"{_EB['n']} 个 op 上全部识别出合并或拆分、三条基线为 0（事件级 p = {EVP:.4f}，"
   f"原因之一是基线的输出空间不含段位表达式）；在不含值改写的非段位档上"
   f"它与嵌入基线持平（各 {AD_IN['n_ops']}/{AD_IN['n_ops']}）；"
   f"其三，粒度命题的统计强度不再依赖十个事件：按表分片重建构造后，"
   f"合并档有 {_XEXT['merge_col']['events_effective']} 个有效事件、"
   f"拆分档 {_XEXT['split_col']['events_effective']} 个、值改写档 "
   f"{_XEXT['value_shift']['events_effective']} 个，"
   f"端到端上段位谱系优于列级谱系的事件为 {_XPM['pos_events']} 正 "
   f"{_XPM['neg_events']} 负（p={_XPM['p']:.3f}）——"
   f"此前的「不显著」是功效不足，而非效应不存在；"
   f"只在值改写两档落后（{AD_SH['ours_ok']}/{AD_SH['n_ops']} 对 "
   f"{AD_SH['base_ok']}/{AD_SH['n_ops']}）。"
   f"它是段位专用抽取器，不是更好的通用匹配器，也不是更差的通用匹配器。")
PB("这些结论的外推受五重限制，均已在正文量化：样本为 11 个库上的合成演化，"
   f"{AD_PS['events']} 次独立演化事件（段位档 {SGEV} 次）；"
   "留出确认仅 7 个库、165 条且存在一个负向库"
   "（regional_sales），只能支持「总体方向可迁移」；编码为两段可逆拼接，"
   "边界情形已在隔离库逐类测定（空值分量与非末段内含分隔符不可恢复，其余四类可恢复），"
   "而真实数据扫描中这两类致命情形出现 0 次，故不覆盖此类取值分布；"
   "含投影星号的查询在本构造下无法被忠实评价，而样本中恰好没有此类条目；"
   "规则基线与语言模型的对照只覆盖「谓词结构不变、仅需替换列引用」的迁移，"
   "需要改写谓词字面值或查询结构的迁移未纳入评测；"
   "三项独立复核还给出一条对方法不利的量化边界：M1 在未见库上的段位表达式恢复仅 "
   f"{MHO_OP['matched']}/{MHO_OP['n_units']}（{MHO_OP['rate_pct']:.1f}%，"
   f"Wilson 上限 {MHO_OP['wilson95_pct'][1]:.1f}%），明显低于开发集的 10/10，"
   f"故不能把开发集读数外推为通用恢复能力（补充材料表 {SN('hler')}）。"
   f"核心语言模型条件独立重跑 {RPT['n_repeat_runs']} 次后，各条件读数相对首次报告值的"
   f"最大偏差为 {RPT['max_abs_dev_pt']:.1f} 个百分点，方向与显著性结论不变；"
   "逐轮逐题原始输出未按条目存档，故只报告汇总偏差与结论稳定性，"
   f"不报告缺乏数据支撑的均值与标准差（补充材料表 {SN('rep')}）。"
   "标注采用单作者编码，并以判定器全严复核、未见库独立评估、编码边界独立实现等构成"
   "等效独立复核；gold 由确定性程序生成、无人工判读单元，故不报告 kappa，"
   f"也不把复核清单表述为双人编码已完成（补充材料表 {SN('single')}）。"
   "本文因此不主张跨库泛化，不主张已解决真实历史 schema 演化，"
   "也不主张确定性方法在工程上优于语言模型。")
PB("最后交代研究类型，因为它决定上面这些结论该被怎样读。本文是一项实证研究，不是系统论文："
   "不提出 RAG-NL2SQL 系统，不实现检索模块，也不提出端到端生成方法——谱系整块注入提示，"
   f"唯一的检索式探针为 {_E2['n']} 条、{_E2['n_db']} 个数据库的阴性对照"
   f"（C4 {_E2['c4_rate']:.1f}% 对 C5 {_E2['c5_rate']:.1f}%，p = {_E2['p']:.2f}）。"
   "因此本文的贡献是「把一个变量分离出来并给出效应量」，而不是「造出一个能用的系统」；"
   "若要支撑后者，至少需要多路召回、证据冲突裁决与弃答判定三部分，本文一部都没有做。"
   "把这条边界写进结论而不是留给读者发现，是本文对自身证据强度的一次校准。")


# ===================================================================== 七
H("七、投稿信（Applied Intelligence 版，可直接投递）")
P("这封信的写法与常规不同：第二段先给主命题，第三段主动交代规则基线更强、"
  "以及自家抽取器在通用匹配上不占优。评审已明确指出坦诚披露是本文可信度的来源；"
  "若把这些留给审稿人自己发现，整封信的可信度都会被折价。")
PB("Dear Editor-in-Chief,", size=10.5)
PB("We submit for consideration in Applied Intelligence the manuscript entitled "
   "“An Empirical Study of Schema Lineage Granularity for Historical SQL Migration "
   "under Schema Evolution”.", size=10.5)
PB("Schema evolution invalidates historical SQL. This is an empirical study, not a system "
   "paper: we propose no retrieval module and no end-to-end NL-to-SQL system, and lineage "
   "is injected into the prompt as a whole; the one retrieval-style probe we ran is a "
   "negative control (40 items, a single database, p = 0.25) and is reported only as a "
   "limitation. Our study isolates a variable that prior work has not separated: how fine "
   "the lineage must be. Holding the executor fixed, "
   "execution-result agreement on merged columns is 0.9% with column-level lineage and 100.0% "
   "with an executable expression that pins segment positions; a prompted model given the "
   "same lineage reaches 79.6%. Because both sides run through the same rewriter, the gap "
   "can be attributed to the lineage representation rather than to the generator. The "
   "effect is absent for splits, which bounds the claim. We keep two held-out studies "
   "apart: an end-to-end confirmation over 165 items on seven unseen databases, where the "
   "merge direction recurs in six of seven independent change events and we report the "
   "single negative case in the main results; and a separate direct lineage-expression "
   "recovery study over 481 items from eight databases, where M1 recovers 14 of 16 "
   "operations. The two sets differ in databases, items and target of judgement, so we "
   "never add or pool them. Scope is the evaluated SQLite query subset and the rename, "
   "merge and split operators; CTE, JOIN USING, HAVING, asterisk projection, ordinal "
   "ORDER BY, explicit casts and other dialects are untested.", size=10.5)
PB("Because a reviewer may reasonably ask whether a language model is needed at all, the "
   "manuscript also reports a deterministic rewriter built on the same lineage. Given "
   "identical lineage it outperforms the tested model by 20.4 points on merged columns and "
   "22.1 on splits. We state plainly that the gold SQL in our construction is generated by "
   "the same substitution rule the rewriter applies, so its 100.0% is a construction-internal "
   "upper bound and not an engineering recommendation. We also bound the comparison domain: "
   "every rewriter-versus-model reading holds only for migrations whose predicates survive "
   "unchanged and which require identifier substitution alone; migrations that force predicate "
   "or query-structure rewriting are outside our construction and we claim nothing about them. "
   "We likewise report where our own extractor is weak, and we distinguish its variants. "
   "M1 recovers the correct operation type and target column on all 15 segment-bearing "
   "operations, against 0/15 for three schema-matching baselines; this is a measure of output "
   "expressiveness, not accuracy, because those baselines cannot emit segment expressions. "
   "We re-test it at the level of ten independent change events (p = 0.0020) after correcting "
   "an event-clustering key that had merged merge- and split-operations sharing a target column. "
   "Direct character-level agreement with gold segment expressions is a separate, narrower "
   "reading (10/10, five events), and a segment-aware non-LLM baseline given the same output "
   "space reaches the same 10/10; we therefore claim no algorithmic advantage for the extractor. "
   "On the 121 operations without value recoding M1+shape ties the embedding matcher at 100%, "
   "and loses only on the 26 recoded ones, 53.8% against 100%. The contribution we claim is "
   "diagnostic: what matters is the granularity of the lineage, not the generator or the "
   "matching algorithm. We repeated the core model conditions five times independently; "
   "every condition deviated from the first run by at most 0.1 percentage points, with all "
   "contrast directions and significance decisions unchanged. Per-run per-item outputs were "
   "not archived item-by-item, so we report the aggregate maximum deviation and conclusion "
   "stability rather than an unsupported mean or standard deviation. Coding was performed by "
   "one author; because the gold SQL and drift operators are generated deterministically by a "
   "single substitution rule, there are no human-annotated judgement units and Cohen's kappa "
   "is not applicable. Instead we provide an auditable set of equivalent independent reviews, "
   "and we do not claim that double coding or inter-rater agreement was established.", size=10.5)
if ARCH_DONE:
    _cover_avail = (
        "All scripts, per-item outputs, statistical routines and the environment "
        "specification are archived in the " + ARCH_REPO + " repository and are "
        "publicly available under a persistent identifier, DOI " + ARCH_DOI + "; "
        "the Data Availability Statement gives the same link. ")
else:
    _cover_avail = (
        "All scripts, per-item outputs, statistical routines and the environment "
        "specification have been assembled into a reproducible package, which we will deposit in "
        "a public repository with a persistent identifier prior to publication; the Data "
        "Availability Statement will then give that link, and we have not yet completed the "
        "deposit. ")
PB("The work fits the journal's scope as an empirical study with a controlled evaluation "
   "protocol. " + _cover_avail +
   "The manuscript is original, is not under consideration elsewhere, and all "
   "authors have approved the submission. We declare no competing interests.", size=10.5)
PB("Sincerely,", size=10.5)
PB("（通讯作者姓名、单位、邮箱、ORCID）", size=10.5)
NOTE("第二段是这封信的实质内容，第三段是主动交代，两段都不能剪。"
     "**不要在信中写「LLM 提升了迁移能力」**——这与本文件第三节的摘要口径冲突，"
     "而编辑会同时看到两处。"
     "第四段在上一版写的是「a released artefact」「a Data Availability Statement is included」，"
     "但同批修订说明又承认 DAS 与 GitHub/Zenodo 归档尚未完成——这是审稿报告 7.3 点名的"
     "既成事实矛盾，必须修正。现改为「已整包、将在发表前存入公开仓库并申请持久标识符、"
     "归档尚未完成」，不把未完成的发布写成 released。"
     "一旦归档真的完成并取得 DOI，本段改由 archive_info.json 自动切为完成时并在 DAS 中给出真实链接"
     "（运行 python fill_doi.py 一次同步六处，不要手改）；"
     "在此之前不得使用 provided / included / released 等词描述公开状态（见第九节）。"
     "标题已换成本轮推荐的 A 案；若最终改用其它案，第一段引号内须同步替换。")


# ===================================================================== 八
H("八、必须与摘要同进同出的七条限定")
P("以下七条，在摘要、结论、投稿信与贡献点中只要缺一条，就会重新出现第一节表里的不一致。"
  "前四条沿用上一轮，第五条是主张重摆后新增的，第六条是比较域限定，"
  "第七条是本轮新增的定位限定——它约束的不是某个数字，而是整篇的研究类型。")
PL(f"**① 编码可逆性**：写「两段、可逆、非空、不含分隔符的取值分布」，不写「一般的列合并」。"
   f"真实数据侧的证据是 21.5 万行扫描、0 行越界，且四类边界已在隔离库逐类测定"
   f"（空值分量与非末段内含分隔符不可恢复，空串、末段内含分隔符等可恢复）——"
   f"这恰好说明构造没有对这类型数据做压力测试，而不是说明数据干净。"
   f"第 17 轮补入的 2–6 段压力测试进一步把前者坐实：修正 to_expr 末段缺陷后，"
   f"末段含分隔符的可恢复率由 {EBL['by_case']['sep_in_last']:.1f}% 升至 "
   f"{EBV['by_case']['sep_in_last']:.1f}%，但 NULL 分量与非末段含分隔符两类仍为 0%——"
   f"限定语不因实现修复而放宽（补充材料表 {SN('enc')}）。")
PL(f"**② 构造内上界**：规则 100% 必须与「gold 由同一条替换规则机械生成」同句出现，"
   f"或紧邻一句内出现。这一条对摘要、贡献三、结论、投稿信四处同时生效。"
   f"本轮主张重摆后，它从「主结果」降为「佐证」，但限定语一个字都不能少——"
   f"降级的是它的位置，不是它的性质。")
PL(f"**③ 单元与方向**：留出库一律写「{_ev['positive']} 正 {_ev['negative']} 负」，"
   f"不写「方向完全重现」；EvoSchema 对照一律按映射单元写（合并档仅 3 个映射、"
   f"最大一个占 92.7%，单元级 p = 0.5），不用条目级 p = 0.0018 替代事件级证据。"
   f"本轮新增一条硬性要求：凡出现「610 条」「230 条」等条目数，"
   f"必须同句或紧邻给出独立单元数（合并档 10 个原子 op、5 次列重整事件；拆分档 5 个），"
   f"否则读者会把条目数读成样本量——这是审稿意见第一条的成因，"
   f"而它只需要半句话就能避免。")
PL(f"**④ 星号与评测范围**：一律写「投影星号题在本构造下无法被忠实评价，样本中恰为 0 条」，"
   f"不写「已支持 SELECT *」「经探针验证了星号能力」。"
   f"合成探针只说明实现没写错，不构成评测证据。")
PL(f"**⑤ M1 的定位必须分三档**：写「段位档 {AD_SG['ops']}/{AD_SG['ops']} 识别出合并或拆分"
   f"对基线 0/{AD_SG['ops']}，事件级 p = {EVP:.4f}」时，必须同时给出后两档——"
   f"不含值改写时与嵌入基线持平（各 {AD_IN['n_ops']}/{AD_IN['n_ops']}），"
   f"值改写两档落后（{AD_SH['ours_ok']}/{AD_SH['n_ops']} 对 {AD_SH['base_ok']}/{AD_SH['n_ops']}）。"
   f"单写第一档会被读成「M1 更好」，而事实是它就只解决段位这一件事。"
   f"第 17 轮再加一条外部效度读数：同一 part 分支在 BIRD train 留出库上只恢复 "
   f"{MHO_OP['matched']}/{MHO_OP['n_units']} 个 op（{MHO_OP['rate_pct']:.1f}%），"
   f"明显低于 dev 的 10/10，故凡出现 10/10 必须同句说明它是开发集读数、"
   f"不构成通用恢复能力（补充材料表 {SN('hler')}）。"
   f"这不是谦逊问题：审稿人自己复算时若发现作者只报强项、略去持平与落后两档，"
   f"全部正向数字都要重新打折。")
PL(f"**⑥ 比较域**：凡出现「规则改写高出所测模型 X pt」，必须同句或紧邻写明"
   f"该差值只在「旧 SQL 谓词结构不变、仅需替换列引用」的迁移上成立；"
   f"需要改写谓词字面值（口径改变、单位换算）或改写查询结构（新增连接、改变聚合粒度）"
   f"的迁移本章未测，也不主张规则占优。"
   f"这一条与第 ②条不是同一件事：②限定「这个数怎么来的」，⑥限定「这个数能比到哪」。"
   f"缺了⑥，二十个百分点的对照会被读成「确定性方法普遍优于语言模型」，"
   f"而这恰好是本章数据支持不了、且最容易被真实演化数据推翻的一句。")
PL(f"**⑦ 研究类型**：凡出现「我们提出／构建了一个 RAG-NL2SQL 系统」「检索增强的迁移框架」一类句子，"
   f"一律删。全文只有一处可以提检索，即 7.7.3 的阴性对照，且必须带 "
   f"{_E2['n']} 条 / {_E2['n_db']} 个数据库 / p = {_E2['p']:.2f} 三个限定。"
   f"标题以 An Empirical Study of 开头，摘要首段交代不实现检索模块，"
   f"三条贡献点均为实证发现——这三处只要有一处回到系统口径，"
   f"定位就从「证据不足」变成「声称与内容不符」，后者是无法靠补充实验挽回的。"
   f"这一条与第⑥条的区别：⑥限的是「这个比较能走多远」，⑦限的是「这篇是什么类型的论文」。")
NOTE("第四条的起因是第九轮评审的措辞「不能暗示星号能力已经通过基准验证」——"
     "这类暗示往往不是刻意写的，而是「能力已实现并经探针验证」这种顺手的句子造成的。"
     "第五条的起因是本轮自查：三条匹配基线早已存在且已报告，"
     "但题级口径把「类型全错」写成了「精度不足」，等于用一个更软的说法掩盖了更硬的不利事实。"
     "两条的共同点是，问题都不在数字，而在数字旁边缺少的那半句。"
     "第六条的性质不同：数字没错，来源也交代了，缺的是「这个比较能走多远」的边界。"
     "这类缺口不会引发数值争议，只会让结论在没有证据的地方继续外推——"
     "而外推部分一旦被真实演化数据否掉，读者会连带怀疑数字本身。"
     "第七条来自外部意见而非自查，性质又不同一层：它一个数字都不改，"
     "却决定这些数字被当作什么来读。标题写系统，读者就按系统的标准审；"
     "标题写实证研究，读者才按「变量是否被分离、证据是否够」来审——"
     "而本文能交付的恰好是后者。")

# ===================================================================== 九
H("九、投稿前待核清单")
P("以下三项是硬门槛，缺一项会被退回补充材料或直接质疑。")

CAP("表 9-1  M1 边界段所引五篇文献（作者与出处已核对，须按原文补卷期页码）")
TABLE(["文献", "在本章中的作用", "须核对的三处论断"],
      [["Rahm 与 Bernstein, VLDB Journal, 2001（schema 匹配系统综述）",
        "确立该方向的输出形态为「元素之间的对应关系」", "输出是否为对应关系／映射"],
       ["Madhavan, Bernstein 与 Rahm, VLDB, 2001（Cupid）",
        "组合语言学与结构相似度的元素级匹配代表", "是否输出元素级匹配"],
       ["Do 与 Rahm, VLDB, 2002（COMA）",
        "多匹配器结果合成的代表", "是否以合成策略为核心"],
       ["Kang 与 Naughton, SIGMOD, 2003",
        "设定最接近：列名不可读、只能靠取值内容判断对应", "① 是否用取值内容判列对；"
        "② 是否输出列对"],
       ["Dhamankar 等, SIGMOD, 2004（iMAP）",
        "支持复杂匹配，但输出仍为元素集合间的对应式子句", "① 是否支持非 1:1 匹配；"
        "② 是否含段位定位信息"]],
      [6.6, 5.4, 4.2], ["l", "l", "l"], font=8.2)
NOTE("**本表不含页码与卷期，这是刻意的。** 我只标注作者、年份与出处，"
     "其余字段必须从出版社页面或原文补全——若依二手来源填写，一旦卷期页码有误，"
     "整段文献支撑都会被视为未核。补全后请顺带核对章内 7.3 的三处论断："
     "① 输出是否为列对或元素对应；② 是否面向同库前后版本；③ 是否含段位定位信息。"
     "任一条与原文献不符，须改写对应句子，不得保留。")

if ARCH_DONE:
    _das_status = "已完成：材料包已公开归档，DOI " + ARCH_DOI + "（投稿信已按完成时表述）"
    _das_action = ("无需再调；DAS 与表 9-3 已给出同一 DOI。"
                   "注意 Zenodo 发布后文件与 DOI 不可修改，如需改文件须发新版本并同步更新稿件 DOI")
    _code_status = "同左，含 ENVIRONMENT.md 与最短复现路径，链接已在 DAS 给出"
    _code_action = "归档完成后在 Data Availability 中给出链接（已完成）"
    _t93_cap = "表 9-3  Data Availability Statement（已归档，可直接写入投稿稿）"
    _t93_note = ("下方声明中的 DOI 与 archive_info.json、CITATION.cff、DAS.md、投稿信第四段一致，"
                 "由 fill_doi.py 一次同步；请勿单独手改其中一处。")
else:
    _das_status = "材料包已整理，但公开归档与 DOI 尚未完成（投稿信已按未完成表述）"
    _das_action = ("投稿前完成 Zenodo/GitHub 归档并取得 DOI，把下方模板 [repository DOI] 换成真实链接，"
                   "再把投稿信第四段改回完成时；在此之前不得写 released / included")
    _code_status = "同左，含 ENVIRONMENT.md 与最短复现路径"
    _code_action = "归档完成后在 Data Availability 中给出链接"
    _t93_cap = "表 9-3  Data Availability Statement 模板（归档完成前不得作为已发布声明使用）"
    _t93_note = ("下方模板中的 [repository DOI] 是占位符，当前稿件尚未完成归档，"
                 "故这段不能原样写进投稿稿，也不能在投稿信里说 DAS 已 included。"
                 "待 Zenodo/GitHub 归档取得真实 DOI 后再启用，并与投稿信第四段同时改口径。")

CAP("表 9-2  Applied Intelligence 的三项强制或高概率要求")
TABLE(["要求", "状态", "动作"],
      [["Data Availability Statement（原创研究强制）", _das_status, _das_action],
       ["代码可用性（非强制但必被问）", _code_status, _code_action],
       ["摘要 150–250 词、关键词 4–6 个",
        "__ABSLENOK__", "已按实测值对齐，无需再调"]],
      [5.0, 6.0, 5.2], ["l", "l", "l"], font=8.2)
NOTE("另需注意：该刊不在中科院预警名单（2020–2026 各版均未列入），"
     "但 2025 版分区表中小类「计算机:人工智能」为 4 区（大类计算机科学 3 区）。"
     "若毕业或考核按小类认定，需先与本单位确认口径，这不是稿件层面的问题。")

CAP(_t93_cap)
NOTE(_t93_note)
PB("The datasets analysed are derived from the BIRD benchmark, which is publicly "
   "available. The evolution operators, item-level outputs, and all scripts required to "
   "reproduce every reported number are archived at " + MANUSCRIPT_DOI + ". The environment "
   "specification, including interpreter and dependency versions and the shortest "
   "reproduction path, is included in the archive. No restricted or proprietary data were "
   "used.", size=10.5)

# ===================================================================== 十
H("十、数字 → 出处自检表")
P("粘贴前逐行核对。任一行对不上，说明章内章外已经开始分叉。"
  "本表的表号不再是手写常量：每次生成时从第 7 章 docx 按表题反查，"
  "因此章内增删表格触发重排后，这里会自动跟着变。")
_TBL = {}


def _count_caps(f):
    """数一份 docx 里的表题数；章外文档的表数表述不写死。"""
    import re as _re
    from docx import Document as _Doc
    if not os.path.exists(f):
        raise SystemExit(f"缺少 {f}；先运行 make_chapter7_v11.py 再运行 slim_chapter7.py")
    return sum(1 for p0 in _Doc(f).paragraphs
               if _re.match(r"^表\s*(?:7|S)-\d+\s\s", p0.text.strip()))


_NSUPP = _count_caps("补充材料_第7章_v11.docx")
_NFULL = _count_caps("第7章_v11_完整版_49表.docx")


def _T(keyword):
    """按表题关键词反查表号。上一版把表号写死，在两次重排后全部失效。

    正文瘦身后表格分布在两份文档：主文 表 7-1..7-15、补充材料 表 S-1..S-37
    （表数随入稿动态增长，第 17 轮由 34 增至 37）。
    这里同时反查，命中补充材料时在出处后标注，避免自检表与主文分叉。
    """
    import re as _re
    from docx import Document as _Doc
    if not _TBL:
        for f in ("投稿版_第7章_实验结果_v11.docx", "补充材料_第7章_v11.docx"):
            if not os.path.exists(f):
                raise SystemExit(
                    f"缺少 {f}；先运行 make_chapter7_v11.py 再运行 slim_chapter7.py")
            d0 = _Doc(f)
            for p0 in d0.paragraphs:
                m0 = _re.match(r"^表\s*(7|S)-(\d+)\s\s(.+)$", p0.text.strip())
                if m0:
                    _TBL[m0.group(3)] = (m0.group(1), int(m0.group(2)))
    for _t, (_pref, _n) in _TBL.items():
        if keyword in _t:
            return (f"表 S-{_n}（补充材料）" if _pref == "S"
                    else f"表 7-{_n}")
    raise SystemExit(f"[自检表] 未找到表题：{keyword}")


TABLE(["数字／表述", "取值", "出处"],
      [["主实验样本", f"{N_ALL} 条 = 改名 {_n('rename_all')} + 合并 {_n('merge')} + "
        f"拆分 {_n('split')}", "stats_ch7_v5.json main"],
       ["合并档阶梯", f"C2 {_mg['c2']['rate']:.1f} / C3m {_mg['c3m']['rate']:.1f} / "
        f"C3mp {_mg['c3mp']['rate']:.1f}", _T("端到端迁移正确率")],
       ["拆分档阶梯", f"C2 {_sp['c2']['rate']:.1f} / C3m {_sp['c3m']['rate']:.1f} / "
        f"C3mp {_sp['c3mp']['rate']:.1f}", _T("端到端迁移正确率")],
       ["同一执行器两种粒度", f"列级 {_rm['c3r_name']['rate']:.1f} vs 段位 "
        f"{_rm['c3r_infer']['rate']:.1f}（合并）",
        _T("规则基线与 LLM 条件在各档")],
       ["M1 与三条基线", f"段位档 {AD_SG['ops']}/{AD_SG['ops']} vs 0/{AD_SG['ops']}"
        f"（事件级 p = {EVP:.4f}）；不含值改写持平 {AD_IN['ours_ok']}/"
        f"{AD_IN['n_ops']} vs {AD_IN['base_ok']}/{AD_IN['n_ops']}；值改写 "
        f"{AD_SH['ours_ok']}/{AD_SH['n_ops']} vs {AD_SH['base_ok']}/{AD_SH['n_ops']}",
        _T("谱系抽取的 op 级质量")],
       ["规则·自动谱系", f"合并 {_rm['c3r_infer']['rate']:.1f} / 拆分 "
        f"{_rs['c3r_infer']['rate']:.1f} / 留出合并 {_rhm['c3r_infer']['rate']:.1f} / "
        f"留出拆分 {_rsh['c3r_infer']['rate']:.1f}", _T("规则基线与 LLM 条件在各档")],
       ["规则 − LLM", f"+{_d_mg:.1f}（{P_mg['a_only']} : {P_mg['b_only']}）／"
        f"+{_d_sp:.1f}（{P_sp['a_only']} : {P_sp['b_only']}）／"
        f"+{_d_hm:.1f}（{P_hm['a_only']} : {P_hm['b_only']}）",
        _T("规则基线与 LLM 的逐题配对")],
       ["两因子分解", f"合并 证据 {_m(D['merge']['evidence_effect_pt'])} / 执行器 "
        f"{_m(D['merge']['executor_effect_pt'])}；拆分 证据 "
        f"{_m(D['split']['evidence_effect_pt'])}", _T("两个因子的分解")],
       ["规则失败画像", f"dev 零失败；留出合并漏抽映射致执行报错；"
        f"COUNT(*) 修正 {RWA['count_star_items']} 条", _T("规则基线的失败画像")],
       ["留出库", f"7 库 165 条；事件 {_ev['positive']} 正 {_ev['negative']} 负；"
        f"合并 {_rhm['c3mp']['rate']:.1f} vs 列级 {_rhm['c3m']['rate']:.1f}",
        _T("留出库上的端到端迁移正确率")],
       ["自动谱系缺口", f"oracle − 自动 = {_gap_ho:.1f}pt（留出合并）",
        _T("规则基线与 LLM 条件在各档")],
       ["编码边界", "21.5 万行扫描、0 行越界；探针 NULL_SIB 0/2、SEP_IN_A 0/2",
        _T("真实数据上的边界出现频率")],
       ["星号探针", f"投影星号 {RWA['projection_star_items']} 条；P1–P4 自检通过",
        _T("星号展开的合成探针")],
       ["EvoSchema 单元级", "合并 3 映射、最大占 92.7%、p = 0.5",
        _T("按独立映射单元重算的外部对照")],
       ["检索式探针（E2）", f"{_E2['n']} 条 / {_E2['n_db']} 库 / "
        f"C4 {_E2['c4_rate']:.1f}% 对 C5 {_E2['c5_rate']:.1f}% / "
        f"不一致对 {_E2['b']}:{_E2['c']} / p = {_E2['p']:.2f}",
        "e2_e2_main.jsonl（逐条重算，与 summary 交叉校验）"]],
      [2.5, 9.4, 3.6], ["l", "l", "l"], font=8.2)
NOTE("表号由 _T() 按表题从两份 docx 反查（主文 表 7-1 至 7-15、"
     f"补充材料 表 S-1 至 S-{_NSUPP}），"
     "因此章内增删或搬移表格触发重排后，这里会自动跟着变。"
     "这一改动针对的是一个实际出过的错误：v9 的自检表把「合并档阶梯」手写成了另两张表的旧编号，"
     "而表格在两次重排后已经换了含义——正文与表题由脚本同步改写，手写的表号却不会。"
     "正文瘦身后 _T() 相应扩为双文档查询；自检表的作用是防止章内外分叉，"
     "它自己先分叉过一次，这条教训值得写进流程。")

# ===================================================================== 十一
H("十一、本文件不做的事")
PL("不改第 7 章的数值与判定口径；章内证据见《投稿版_第7章_实验结果_v11.docx》"
   f"（主文 15 表）与《补充材料_第7章_v11.docx》（S-1 至 S-{_NSUPP}），"
   f"完整 {_NFULL} 表素材库见《第7章_v11_完整版_49表.docx》"
   "（文件名沿用瘦身时的历史约定，内容随入稿增长）。")
PL("不再保留 JMS 口径。目标期刊已确定为 Applied Intelligence，"
   "该刊不要求 PRISMA/SANRA、不设字数上限、不需要循证检索式，"
   "因此这些内容一律不进摘要、贡献点与投稿信；"
   "原 JMS 版所需的 5,500 词裁剪方案同时失效——第 7 章不必砍到 4,000 字。")
PL("不代拟审稿回复信。回复信应逐条引用本文件第五、六、八节的句子，"
   "而不是重新组织措辞——两处措辞一旦不同，审稿人会认为作者自己也没统一口径。")
PL("不虚构文献条目。第九节表 9-1 只给作者、年份与出处，"
   "卷期页码留空待补，这是刻意的取舍。")
PL("不补检索实验去换取「RAG」这个标题词。若真要以 RAG 为题，至少需要多路召回、"
   "证据冲突裁决与弃答判定三部分，那是另一篇论文的工作量；"
   "把 40 条阴性对照包装成检索研究，代价是审稿人翻到实验部分当场识破，"
   "且此后所有数值都会按「作者是否夸大」重新打折。")

# 回填实测字数
_ENNOTE = (f"实测 {_EN_N} 词，落在本轮设定的 220–235 词目标区间内（刊方区间为 150–250 词）。"
           "压到这一区间的原因不是刊方要求，而是投稿系统的计数口径与 Word 不同："
           "连字符、缩写、百分号的处理各有出入，贴着上限写等于把风险留给自己，"
           "故主动留出余量。"
           "第 21 轮的改写位置有两处：首句由「Schema evolution breaks historical SQL」"
           "改为「We report an empirical study of…」，把研究类型放在第一位；"
           "同段末补「We build no retrieval module and no NL-to-SQL system」；"
           "末段补「A 40-item retrieval probe is a negative control（p=0.25）」。"
           "为腾出这三句，删去的是「The question is not which model rewrites the query」"
           "这一设问句，以及 16 operations、ten events could not 等可从正文补齐的细节。"
           "三处限定语仍然在："
           "「construction-internal…the same substitution」是构造循环性的限定；"
           "「expressiveness, not accuracy」与「no algorithmic gain is claimed」"
           "两句是 M1 定位的两半，缺任一半 15/15 都会变成单向宣传；"
           "末段两项留出实验分列（165 条 / 7 库 与 481 条 / 8 库），不得合并成一句。")
_CNNOTE = (f"实测 {_CN_N} 字（含标点）。Applied Intelligence 只约束英文摘要 150–250 词，"
           "中文摘要不受该刊字数限制，仅供中文版本或作者自用，故此处只报实测值、不设目标值。"
           "第 21 轮在首段加了研究类型声明（实证研究、不实现检索模块、不实现端到端生成），"
           "末段加了检索式探针的完整口径（40 条 / 1 库 / p = 0.25）——"
           "这两处不是补充说明，而是把「本文是什么」提前到读者形成预期的位置。"
           "第三段与第四段仍是这一版的主张落点：第三段给出「同一执行器、两种粒度」，"
           f"第四段给出 M1 的段位档 15/15 对 0/15（事件级 p = {EVP:.4f}）、不含值改写时与嵌入"
           "基线持平、仅值改写两档落后——三档并陈，缺一档都会被读成单向宣传。"
           "两段顺序不可互换：先立粒度这个主命题，再谈方法定位；"
           "若把 M1 放在前面，读者会以为这是一篇「提出新抽取器」的论文。")
_ABSLENOK = (f"英文实测 {_EN_N} 词（刊方区间 150–250；第 21 轮改写研究类型与检索探针口径后"
             f"为 {_EN_N} 词，仍在上限内留有余量）、关键词 5 个，均在要求区间内")
_MAP = {"__ENWORDS__": str(_EN_N), "__CNCHARS__": str(_CN_N),
        "__ENNOTE__": _ENNOTE, "__CNNOTE__": _CNNOTE, "__ABSLENOK__": _ABSLENOK}


def _fill(paras):
    for _p in paras:
        for _r in _p.runs:
            for _k, _v in _MAP.items():
                if _k in _r.text:
                    _r.text = _r.text.replace(_k, _v)


_fill(doc.paragraphs)
for _t in doc.tables:
    for _row in _t.rows:
        for _cell in _row.cells:
            _fill(_cell.paragraphs)

doc.save(OUT)
print("saved:", OUT)
print(f"英文摘要 {_EN_N} 词 / 中文摘要 {_CN_N} 字")
print("段落:", len(doc.paragraphs), "表格:", len(doc.tables))
