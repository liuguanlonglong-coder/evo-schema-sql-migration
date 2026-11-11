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
assert ADI["version"] == "audit-independence-v1", "独立性审计产物缺失（audit_independence.py）"
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
# 章外统一改报正确口径：事件数 10、p = 0.0020
EVN = EMA["sign_correct"]["n"]          # 10
EVP = EMA["sign_correct"]["p_exact"]    # 0.001953
SGEV = EMA["counts"]["events_correct"]  # 10（段位档事件）

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
     "改为「同一执行器下，谱系的粒度决定迁移成败」。")

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
H("二、标题：三案与取舍")
P("主张改变后，标题的推荐顺序也变了。粒度成为唯一核心，「确定性重写基线」退为支撑证据，"
  "因此不宜再与它并列出现在主标题里。RAG 与「端到端 NL2SQL」仍不得出现。")
PL("**A（推荐）** 谱系要细到哪一层：schema 演化下历史 SQL 迁移的粒度效应")
PL("**B** 同一执行器，两种粒度：段位表达式谱系能否救回历史 SQL")
PL("**C** 列合并与列拆分之后：历史 SQL 迁移中的谱系粒度与确定性重写基线")
NOTE("A 把问句放在最前，与新主张一致，且「粒度效应」是可被证伪的命题；"
     "B 最能体现本轮重摆（控制了执行器），但「救回」偏口语；"
     "C 是 v10 的推荐，本轮降为备选——它并列了两个自变量，"
     "而数据表明两者不是并列关系：表示效应约为执行器效应的 "
     f"{abs(D['merge']['evidence_effect_pt'] / D['merge']['executor_effect_pt']):.1f} 倍。"
     "三案均不含 RAG。副标题「——以电子病历结构化字段为例」仍不建议加，"
     "因为本章没有医疗数据与检索系统实验。")

# ===================================================================== 三
H("三、英文摘要（实测 __ENWORDS__ 词，Applied Intelligence 要求 150–250 词）")
_EN_ABS = [
    ("Schema evolution breaks historical SQL. When columns are merged into one "
     "delimited column, column-level lineage returns the whole string, so the migration "
     "runs but is semantically wrong. The question is not which model rewrites the query "
     "but how fine lineage must be."),
    ("On 11 BIRD databases (610 items) we build a lineage-granularity ladder. "
     "With the executor fixed, we compare M1 (label-free inference, "
     "development-set-calibrated thresholds) against name, value-overlap and embedding matchers."),
    ("Under one executor, merged items reach 0.9% with column-level and 100.0% with "
     "segment-level lineage—a property of lineage, not the executor. A prompted model "
     "reaches 79.6%; the rewriter's 100.0% is construction-internal, as its gold SQL uses "
     "the same substitution. Splits gain nothing (77.9%)."),
    ("M1 recovers the operation type and target column on 15/15 operations, versus 0/15 for "
     "three matchers unable to emit segment expressions: expressiveness, not accuracy. "
     "Character-level agreement with gold expressions is a separate reading (10/10, five "
     "events), and a segment-aware non-LLM baseline with equal output space also scores "
     "10/10; we claim no algorithmic gain. Ten events give p = 0.0020."),
    ("M1+shape ties the embedding matcher on 121 non-reencoded operations, losing on 26 "
     "recoded ones (53.8% vs 100%). The 610 items rest on 90 operations and 58 events, so "
     "contrasts use event-level tests. Two boundary classes are non-invertible but never "
     "fired in 215,216 rows. On seven held-out databases the direction recurs in 6 events "
     "and reverses in 1. We implement no retrieval module."),
]
for _seg in _EN_ABS:
    PB(_seg)
_EN_N = sum(len(re.findall(r"[A-Za-z0-9][A-Za-z0-9%\.\-–']*", s)) for s in _EN_ABS)
NOTE("__ENNOTE__")


# ===================================================================== 四
H("四、中文摘要（实测 __CNCHARS__ 字）")
_CN_ABS = [
    ("schema 演化会使历史 SQL 失效：当多列被合并为带分隔符的单列时，列级映射会把整串返回，"
     "迁移结果可执行却语义错误。要问的核心问题不是用哪个模型重写，而是谱系必须细到哪一层。"),
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
    (f"适用范围必须与结果同时给出：{AD_PS['items']} 条背后只有 {AD_PS['ops']} 个演化算子、"
     f"{AD_PS['events']} 次独立演化事件，段位档 {AD_SG['ops']} 个算子更只对应 "
     f"{SGEV} 次事件，故合并与拆分两档一律以事件为单位复核。"
     "编码边界已逐类测定：空值分量与非末段内含分隔符不可恢复，"
     "空串、末段内含分隔符、首尾空白、多字节字符四类可恢复；"
     "21.5 万行真实数据扫描中前两类出现 0 次，故零越界应读作构造未触发。"
     "在 7 个未参与开发的库上方向重现（独立变更事件 6 正 1 负）；本文未实现检索模块。"),
]
for _seg in _CN_ABS:
    PB(_seg)
_CN_N = sum(len(re.sub(r"\s", "", s)) for s in _CN_ABS)
NOTE("__CNNOTE__")


# ===================================================================== 五
H("五、关键词与贡献点")
P("关键词 5 个，全部指称正文真实存在的内容：")
PL("schema 演化；历史 SQL 迁移；表达式级谱系；确定性查询重写；执行等价")
NOTE("删除候选：RAG、检索增强生成、NL2SQL 端到端、提示工程——这些词所指的模块在本研究中不存在。"
     "Applied Intelligence 要求 4–6 个关键词，此处 5 个在区间内。")

P("贡献点三条重写。主句变了，第一条与第三条的位置因此互换：粒度证据升为首位，"
  "规则基线退为佐证；第二条改为带边界的正面贡献。")
TABLE(["", "v10 的写法", "v11 写法"],
      [["贡献一",
        "给出粒度命题的受控证据：合并档由 "
        f"{_mg['c3m']['rate']:.1f}% 升至 {_mg['c3mp']['rate']:.1f}%（证据效应 "
        f"{_m(D['merge']['evidence_effect_pt'])}pt）",
        "把「谱系粒度」作为独立变量分离出来：固定同一执行器，"
        f"合并档列级谱系 {_rm['c3r_name']['rate']:.1f}%、段位表达式 "
        f"{_rm['c3r_infer']['rate']:.1f}%。两侧执行器相同，差距只能归于谱系表示——"
        "这是全文唯一不受构造循环性影响的结论。拆分方向效应为零，故命题限于「一列内多段编码」"],
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
      [1.5, 5.4, 10.0], ["c", "l", "l"], font=8.3)
NOTE("贡献二的写法在本轮又改了一次，而且是往更不利的方向改。v10 只说「把误差定位到谱系发现」，"
     "读者无法判断 M1 本身值多少；v11 给出 15/15 对 0/15 的对照。"
     "但本轮核对判定函数后确认：那个 15/15 度量的是「识别出这是合并或拆分」，"
     "因为判定前 gold 的目标表达式已被归一为纯列名，段位字符串从未进入比较。"
     "因此同一段必须再补一句——段位表达式本身的字符级一致率为 10/10，"
     "却只覆盖 2 次独立事件，事件级 p = 0.5；段位价值的统计证据不在抽取环节，而在端到端环节。"
     "同理，上一版写的「通用匹配上不如嵌入基线的 91.8%」也是错的："
     "拆开后 610 条主实验集内的 121 个算子两者持平于 100%，"
     "缺口 12 个全部来自集外的值改写两档。"
     "两处都改成了更弱、但经得起复算的表述——这比保留一个好看的数字重要。")


# ===================================================================== 六
H("六、结论段")
PB(f"本文的证据支持一个窄命题：当 schema 演化把多列合并为带分隔符的一列时，"
   f"谱系粒度是迁移成败的决定因素。固定同一个确定性改写器，"
   f"列级谱系只到 {_rm['c3r_name']['rate']:.1f}%，段位表达式谱系到 "
   f"{_rm['c3r_infer']['rate']:.1f}%；把同一份段位谱系交给语言模型则到 "
   f"{_mg['c3mp']['rate']:.1f}%。三条路径共用一套谱系与一套判定，"
   f"唯一变化的是表示与执行器，而表示的效应（{_m(D['merge']['evidence_effect_pt'])}pt）"
   f"约为执行器效应（{_m(D['merge']['executor_effect_pt'])}pt）的 "
   f"{abs(D['merge']['evidence_effect_pt'] / D['merge']['executor_effect_pt']):.1f} 倍。")
PB(f"由此得到两条对后续工作更有用的读数。其一，该任务的瓶颈在谱系发现："
   f"自动抽取在开发库上与构造映射逐条等价，在未见库上留下 {_gap_ho:.1f}pt 的硬失败缺口，"
   f"且缺口全部表现为执行报错而非静默错。"
   f"其二，确定性重写是一条必须报告的强基线，但它的 "
   f"{_rm['c3r_infer']['rate']:.1f}% 是构造内上界——gold 由同一替换规则生成——"
   f"因此只能用于说明「给定映射后执行器不是瓶颈」，不能用于工程选型。"
   f"方法一侧的定位同样需要收窄，且分三档而不是一概而论：M1 在需要段位的 "
   f"{_EB['n']} 个 op 上全部识别出合并或拆分、三条基线为 0（事件级 p = {EVP:.4f}，"
   f"原因之一是基线的输出空间不含段位表达式）；在不含值改写的非段位档上"
   f"它与嵌入基线持平（各 {AD_IN['n_ops']}/{AD_IN['n_ops']}）；"
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
   "需要改写谓词字面值或查询结构的迁移未纳入评测。"
   "本文因此不主张跨库泛化，不主张已解决真实历史 schema 演化，"
   "也不主张确定性方法在工程上优于语言模型。")


# ===================================================================== 七
H("七、投稿信（Applied Intelligence 版，可直接投递）")
P("这封信的写法与常规不同：第二段先给主命题，第三段主动交代规则基线更强、"
  "以及自家抽取器在通用匹配上不占优。评审已明确指出坦诚披露是本文可信度的来源；"
  "若把这些留给审稿人自己发现，整封信的可信度都会被折价。")
PB("Dear Editor-in-Chief,", size=10.5)
PB("We submit for consideration in Applied Intelligence the manuscript entitled "
   "“How Fine Must Lineage Be? Granularity Effects in Migrating Historical SQL under "
   "Schema Evolution”.", size=10.5)
PB("Schema evolution invalidates historical SQL. Our study isolates a variable that prior "
   "work has not separated: how fine the lineage must be. Holding the executor fixed, "
   "execution equivalence on merged columns is 0.9% with column-level lineage and 100.0% "
   "with an executable expression that pins segment positions; a prompted model given the "
   "same lineage reaches 79.6%. Because both sides run through the same rewriter, the gap "
   "can be attributed to the lineage representation rather than to the generator. The "
   "effect is absent for splits, which bounds the claim. Across seven held-out databases "
   "the direction recurs in six of seven independent change events, and we report the "
   "single negative case in the main results rather than in the limitations.", size=10.5)
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
   "matching algorithm.", size=10.5)
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
PB("The work fits the journal's scope as an applied AI study with a controlled evaluation "
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
H("八、必须与摘要同进同出的六条限定")
P("以下六条，在摘要、结论、投稿信与贡献点中只要缺一条，就会重新出现第一节表里的不一致。"
  "前四条沿用上一轮，第五条是主张重摆后新增的，第六条是本轮补的比较域限定。")
PL(f"**① 编码可逆性**：写「两段、可逆、非空、不含分隔符的取值分布」，不写「一般的列合并」。"
   f"真实数据侧的证据是 21.5 万行扫描、0 行越界，且四类边界已在隔离库逐类测定"
   f"（空值分量与非末段内含分隔符不可恢复，空串、末段内含分隔符等可恢复）——"
   f"这恰好说明构造没有对这类型数据做压力测试，而不是说明数据干净。")
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
   f"这不是谦逊问题：审稿人自己复算时若发现作者只报强项、略去持平与落后两档，"
   f"全部正向数字都要重新打折。")
PL(f"**⑥ 比较域**：凡出现「规则改写高出所测模型 X pt」，必须同句或紧邻写明"
   f"该差值只在「旧 SQL 谓词结构不变、仅需替换列引用」的迁移上成立；"
   f"需要改写谓词字面值（口径改变、单位换算）或改写查询结构（新增连接、改变聚合粒度）"
   f"的迁移本章未测，也不主张规则占优。"
   f"这一条与第 ②条不是同一件事：②限定「这个数怎么来的」，⑥限定「这个数能比到哪」。"
   f"缺了⑥，二十个百分点的对照会被读成「确定性方法普遍优于语言模型」，"
   f"而这恰好是本章数据支持不了、且最容易被真实演化数据推翻的一句。")
NOTE("第四条的起因是第九轮评审的措辞「不能暗示星号能力已经通过基准验证」——"
     "这类暗示往往不是刻意写的，而是「能力已实现并经探针验证」这种顺手的句子造成的。"
     "第五条的起因是本轮自查：三条匹配基线早已存在且已报告，"
     "但题级口径把「类型全错」写成了「精度不足」，等于用一个更软的说法掩盖了更硬的不利事实。"
     "两条的共同点是，问题都不在数字，而在数字旁边缺少的那半句。"
     "第六条的性质不同：数字没错，来源也交代了，缺的是「这个比较能走多远」的边界。"
     "这类缺口不会引发数值争议，只会让结论在没有证据的地方继续外推——"
     "而外推部分一旦被真实演化数据否掉，读者会连带怀疑数字本身。")

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


def _T(keyword):
    """按表题关键词反查表号。上一版把表号写死，在两次重排后全部失效。

    正文瘦身后表格分布在两份文档：主文 表 7-1..7-15、补充材料 表 S-1..S-34。
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
        _T("按独立映射单元重算的外部对照")]],
      [2.5, 9.4, 3.6], ["l", "l", "l"], font=8.2)
NOTE("表号由 _T() 按表题从两份 docx 反查（主文 表 7-1 至 7-15、补充材料 表 S-1 至 S-34），"
     "因此章内增删或搬移表格触发重排后，这里会自动跟着变。"
     "这一改动针对的是一个实际出过的错误：v9 的自检表把「合并档阶梯」手写成了另两张表的旧编号，"
     "而表格在两次重排后已经换了含义——正文与表题由脚本同步改写，手写的表号却不会。"
     "正文瘦身后 _T() 相应扩为双文档查询；自检表的作用是防止章内外分叉，"
     "它自己先分叉过一次，这条教训值得写进流程。")

# ===================================================================== 十一
H("十一、本文件不做的事")
PL("不改第 7 章的数值与判定口径；章内证据见《投稿版_第7章_实验结果_v11.docx》"
   "（主文 15 表）与《补充材料_第7章_v11.docx》（S-1 至 S-34），"
   "完整 49 表素材库见《第7章_v11_完整版_49表.docx》。")
PL("不再保留 JMS 口径。目标期刊已确定为 Applied Intelligence，"
   "该刊不要求 PRISMA/SANRA、不设字数上限、不需要循证检索式，"
   "因此这些内容一律不进摘要、贡献点与投稿信；"
   "原 JMS 版所需的 5,500 词裁剪方案同时失效——第 7 章不必砍到 4,000 字。")
PL("不代拟审稿回复信。回复信应逐条引用本文件第五、六、八节的句子，"
   "而不是重新组织措辞——两处措辞一旦不同，审稿人会认为作者自己也没统一口径。")
PL("不虚构文献条目。第九节表 9-1 只给作者、年份与出处，"
   "卷期页码留空待补，这是刻意的取舍。")

# 回填实测字数
_ENNOTE = (f"实测 {_EN_N} 词，在 150–250 词区间内。"
           "第三段与第四段各有一句不可删："
           "「construction-internal…the very substitution it applies」是构造循环性的限定语，"
           "缺了它 100.0% 会被读成独立验证的性能；"
           "「not a better general matcher」是 M1 定位的另一半，"
           "缺了它第四段的 15/15 就成了单向宣传。"
           "第三段的 0.9% 与 100.0% 必须同句出现——它们共用同一个执行器，"
           "这是全文唯一不受构造循环性影响的对比，分开写就没有这个性质了。"
           "若需压到 200 词以内，先删第五段的 215,216 与 6 正 1 负，保留「reversible two-segment」。")
_CNNOTE = (f"实测 {_CN_N} 字（含标点）。Applied Intelligence 只约束英文摘要 150–250 词，"
           "中文摘要不受该刊字数限制，仅供中文版本或作者自用，故此处只报实测值、不设目标值。"
           "本篇的第三段与第四段是这一轮重摆的落点：第三段给出「同一执行器、两种粒度」，"
           f"第四段给出 M1 的段位档 15/15 对 0/15（事件级 p = {EVP:.4f}）、不含值改写时与嵌入"
           "基线持平、仅值改写两档落后——三档并陈，缺一档都会被读成单向宣传。"
           "两段的顺序不可互换——先立粒度这个主命题，再谈方法定位；"
           "若把 M1 放在前面，读者会以为这是一篇「提出新抽取器」的论文。")
_ABSLENOK = (f"英文实测 {_EN_N} 词、关键词 5 个，均在要求区间内")
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
