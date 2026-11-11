# -*- coding: utf-8 -*-
"""生成《摘要_标题_贡献点_结论_v10.docx》。

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
OUT = "摘要_标题_贡献点_结论_v10.docx"

EN, CN, CN_H = "Times New Roman", "宋体", "黑体"

S = json.load(open("results/stats_ch7_v5.json", encoding="utf-8"))["main"]
RVR = json.load(open("results/rule_vs_llm.json", encoding="utf-8"))
G = RVR["groups"]
D = RVR["decomposition"]
HLD = json.load(open("results/holdout_stats.json", encoding="utf-8"))
EVO = json.load(open("results/evoschema_events.json", encoding="utf-8"))
DBD = json.load(open("results/data_bounds.json", encoding="utf-8"))
PRB = json.load(open("results/probe_counterfactual.json", encoding="utf-8"))
RWA = json.load(open("results/rule_rewrite.json", encoding="utf-8"))["audit"]
assert RVR["version"] == "rule-vs-llm-v2"

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
_run(p_, "摘要、标题、贡献点、结论与投稿信（v10）", bold=True, size=17, font=CN_H)
p_ = doc.add_paragraph()
p_.alignment = WD_ALIGN_PARAGRAPH.CENTER
_run(p_, "回应第九轮评审：规则结果标注构造内边界、补齐文献、统一 Applied Intelligence 口径",
     size=10, font=CN)
NOTE("目标期刊：**Applied Intelligence**（Springer，ISSN 0924-669X）。"
     "本文件所有字数均为脚本实测值，不是估计值；"
     "末尾附「数字 → 出处」自检表与「投稿前待核清单」，两处都需逐行核对后再粘贴。"
     "正文证据见《投稿版_第7章_实验结果_v10.docx》。")

# ===================================================================== 一
H("一、第九轮要解决的四件事")
P("这一轮评审没有要求新实验，四条意见全部落在「同一个数字被两种口径描述」上。"
  "其中前三条与摘要层有关，第四条是文件内部的前后矛盾。")
TABLE(["问题", "现状", "本轮改法", "不做的后果"],
      [["规则 100% 的性质", "摘要写「达到 100.0%」，"
        "「gold 由同一条替换规则机械生成」只在正文与修订说明",
        "紧邻该结果加「在本文构造中」，并展开一句说明它是构造内上界",
        "读者会把构造内自洽读成独立验证的系统性能"],
       ["M1 与已有方法的边界", "只写「那一族方法」这一类别，无具体文献",
        "补五篇原始文献并逐条对照输出形态（章内 7.3），本文件给核对清单",
        "新颖性无法评估，审稿人只能按「增量」判"],
       ["规则基线的适用范围", "章内已收窄，摘要、贡献点、投稿信仍偏强",
        "四处统一为「本构造下的配对差值」，投稿信主动交代",
        "章内外分叉，比数值错误更难解释"],
       ["字数标注", "英文标「约 230 词」实测 240；中文标「约 400 字」实测 576",
        "全部改为实测值，并说明中文摘要不受该刊字数约束",
        "审稿人核对时会怀疑其他数字同样未核"]],
      [2.9, 5.6, 5.0, 4.2], ["l", "l", "l", "l"], font=8.3)
NOTE("四条的共同点是：数字本身没错，错的是数字旁边的限定语。"
     "第九轮评审之所以把它们列进「优先完成」，是因为摘要是第一份被读的文件——"
     "限定语缺在摘要里，等于全文都没有。")

# ===================================================================== 二
H("二、标题：三案与取舍")
P("标题要让两个真实的自变量出现在最前面：谱系粒度、给定映射后的确定性重写。"
  "RAG 与「端到端 NL2SQL」都不能出现在标题中，因为它们所指的模块在本研究中不存在。")
PL("**A（推荐）** 列合并与列拆分之后：历史 SQL 迁移中的谱系粒度与确定性重写基线")
PL("**B** 谱系要细到哪一层：schema 演化下历史 SQL 迁移的粒度效应")
PL("**C** 段位表达式谱系：schema 演化下历史 SQL 迁移的可执行证据表示")
NOTE("A 同时写出自变量与强基线，与正文一一对应；B 简洁但丢掉规则基线这条最重要的对照；"
     "C 强调表示形式，容易被读成「提出一种新表示」而非「检验粒度是否够用」。"
     "三案均不含 RAG。若需保留应用领域，只可加副标题「——以电子病历结构化字段为例」，"
     "但第九轮评审已指出本章没有医疗数据与检索系统实验，"
     "故该副标题只能用于强调任务来源，不能当作应用贡献。")

# ===================================================================== 三
H("三、英文摘要（实测 __ENWORDS__ 词，Applied Intelligence 要求 150–250 词）")
_EN_ABS = [
    ("Schema evolution breaks historical SQL. When several columns are merged into one "
     "delimited column, a column-level mapping returns the whole string rather than the "
     "original value; a migration is executable yet semantically wrong. We ask how fine a "
     "lineage must be, and whether a deterministic rewriter given that lineage matches a "
     "prompted language model."),
    ("We construct a controlled benchmark on 11 BIRD databases (610 items). Conditions form "
     "an information ladder up to expression-level lineage specifying segment positions "
     "(SUBSTR/INSTR). A deterministic rewriter substitutes column references through the "
     "same lineage, with three mapping sources: gold, inferred, column-level."),
    ("For merged columns, execution equivalence is 10.0% with the new schema only, 2.6% with "
     "column-level lineage, and 79.6% with segment-level lineage. Given the same lineage the "
     "rewriter reaches 100.0% within this construction: the gold SQL is generated by the "
     "very substitution rule the rewriter applies, so this is a construction-internal upper "
     "bound rather than independently measured performance. It still stands 20.4 points "
     "above the model (47:0), and no item is solved by the model alone. Splits gain nothing "
     "from segment information (77.9%), yet the rewriter gains 22.1 points."),
    ("On seven held-out databases (165 items) the direction recurs (6 events positive, "
     "1 negative), with the rewriter at 80.2–100.0%. Scope is explicit: the encoding is a "
     "reversible two-segment concatenation, and a full scan of 215,216 rows never triggers "
     "null components or in-value separators, so these results do not characterize such "
     "data. No retrieval module is implemented; we make no retrieval claim."),
]
for _seg in _EN_ABS:
    PB(_seg)
_EN_N = sum(len(re.findall(r"[A-Za-z0-9][A-Za-z0-9%\.\-–']*", s)) for s in _EN_ABS)
NOTE("__ENNOTE__")

# ===================================================================== 四
H("四、中文摘要（实测 __CNCHARS__ 字）")
_CN_ABS = [
    ("schema 演化会使历史 SQL 失效：当多列被合并为带分隔符的单列时，列级映射会把整串返回，"
     "迁移结果可执行却语义错误。本文检验两个问题：谱系信息需要细到哪一层；"
     "以及在给定该谱系时，确定性重写是否不弱于把谱系写进提示的语言模型。"),
    ("我们在 11 个 BIRD 数据库上构造 610 条受控样本（改名 276、列合并 230、列拆分 104），"
     "条件阶梯从「仅新 schema」逐级升到给出段位位置的可执行表达式谱系（SUBSTR/INSTR）。"
     "在同一谱系下另设确定性重写器，只做列引用替换，并按三种映射来源运行："
     "构造映射、自动抽取谱系、列级谱系。"),
    (f"列合并档的执行等价率随信息水平上升：仅新 schema {_mg['c2']['rate']:.1f}%、"
     f"列级谱系 {_mg['c3m']['rate']:.1f}%、段位表达式谱系 {_mg['c3mp']['rate']:.1f}%。"
     f"给定同一份段位谱系，确定性重写在本文构造内达到 {_rm['c3r_infer']['rate']:.1f}%——"
     f"此处 gold SQL 由同一条替换规则机械生成，故该值应读作构造内上界，"
     f"而非独立测得的系统性能；即便如此，它仍高出所测模型 {_d_mg:.1f} 个百分点"
     f"（{P_mg['a_only']} : {P_mg['b_only']}），"
     f"且在开发库上没有一条是模型独有做对。列拆分档不受益于段位信息"
     f"（{_sp['c3mp']['rate']:.1f}% 与列级相同），但确定性重写仍高出 {_d_sp:.1f} 个百分点。"
     f"在 7 个未参与开发的库（165 条）上方向重现（独立变更事件 {_ev['positive']} 正 "
     f"{_ev['negative']} 负），确定性重写处于 {_rhm['c3r_infer']['rate']:.1f}–"
     f"{_rsh['c3r_infer']['rate']:.1f}% 之间。"),
    ("适用范围必须与结果同时给出：本文的编码是可逆的两段拼接，"
     "对 21.5 万行真实数据的全表扫描未触发空值分量与值内分隔符，"
     "因此结论不覆盖含空值或分隔符冲突的取值分布；"
     "本文亦未实现检索模块，不对检索效果作任何主张。"),
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

P("贡献点三条，第三条是负面的。第九轮之后，三条都要带上限定语。")
TABLE(["", "v8 的写法", "v10 写法（限定语已内嵌）"],
      [["贡献一", "提出 RAG 驱动的 NL2SQL 迁移框架，在医疗库上显著提升准确率",
        "给出粒度命题的受控证据：在列合并方向上，把谱系从列级下沉到段位表达式，"
        f"执行等价率由 {_mg['c3m']['rate']:.1f}% 升至 {_mg['c3mp']['rate']:.1f}%"
        f"（证据效应 {_m(D['merge']['evidence_effect_pt'])}pt）；"
        "列拆分方向该效应为零，故命题限于「一列内多段编码」"],
       ["贡献二", "语言模型能够理解 schema 变化并自动完成迁移",
        "把误差来源定位到「谱系能否被正确发现」：自动谱系在开发库上与构造映射逐条等价，"
        f"在 7 个未见库上留下 {_gap_ho:.1f}pt 的缺口，"
        "且缺口全部表现为执行报错而非静默错"],
       ["贡献三", "（无）",
        f"给出确定性重写这条强基线：给定同一份段位谱系时，规则改写高出所测模型 "
        f"{_d_mg:.1f}pt（合并）与 {_d_sp:.1f}pt（拆分）、{_d_hm:.1f}pt（留出合并），"
        "且模型在开发库上无独占正确条目。**该差值限于本构造**——"
        "gold SQL 由同一条替换规则生成，规则侧上界被构造抬高，"
        "LLM 侧只测一个基座与一套模板，因此不构成「确定性方法普遍优于 LLM」的结论"],
       ["作用域", "（无）",
        "证据来自两段可逆拼接的合成编码；真实数据扫描未覆盖空值与值内分隔符；"
        "未实现检索；投影星号题在本构造下无法被忠实评价"]],
      [1.5, 4.6, 9.8], ["c", "l", "l"], font=8.3)
NOTE("贡献三这一条第九轮之后又加了一层限定。原来的写法已经承认「构造抬高上界」，"
     "但摘要层面只给了数字。现在的写法把「限于本构造」放进同一句话里，"
     "读者不必读到第六节才知道边界。这不是措辞软化，是把限定语从注释放到主张里。")

# ===================================================================== 六
H("六、结论段")
PB(f"本文的证据支持一个窄命题：当 schema 演化把多列合并为带分隔符的一列时，"
   f"列级谱系不足以维持历史 SQL 的语义，"
   f"把谱系下沉为指明段位的可执行表达式可以把执行等价率从 {_mg['c3m']['rate']:.1f}% "
   f"提升到 {_mg['c3mp']['rate']:.1f}%；而在列拆分方向，同一提升不存在。"
   "在此基础上，本文给出两条对后续工作更有用的读数："
   f"其一，在本构造中，给定同一份谱系，确定性重写的执行等价率高出所测模型 "
   f"{_d_mg:.1f} 个百分点——由于 gold 由同一替换规则机械生成，"
   "这一差值衡量的是「执行器在本构造下的差异」，"
   "说明该任务的瓶颈不在 SQL 生成而在谱系发现，但不能外推为工程选型结论；"
   f"其二，自动谱系在开发库上与构造映射逐条等价，在未见库上仍有 {_gap_ho:.1f}pt "
   "的硬失败缺口，因此自动发现才是这条路线当前的主要成本。")
PB("这些结论的外推受四重限制，均已在正文量化：样本为 11 个库上的合成演化，"
   "独立的列重整事件只有 10 次；留出确认仅 7 个库、165 条，"
   "且存在一个负向库（regional_sales），只能支持「总体方向可迁移」；"
   "编码固定为两段可逆拼接，真实数据扫描中空值分量与值内分隔符从未出现，"
   "故不覆盖此类取值分布；含投影星号的查询在本构造下无法被忠实评价，"
   "而样本中恰好没有此类条目。"
   "本文因此不主张跨库泛化，不主张已解决真实历史 schema 演化，"
   "也不主张确定性方法在工程上优于语言模型。")

# ===================================================================== 七
H("七、投稿信（Applied Intelligence 版，可直接投递）")
P("这封信的写法与常规不同：主动把「规则基线更强」写在第二段。"
  "评审已明确指出坦诚披露是本文可信度的来源；若把这条留给审稿人自己发现，"
  "整封信的可信度都会被折价。")
PB("Dear Editor-in-Chief,", size=10.5)
PB("We submit for consideration in Applied Intelligence the manuscript entitled "
   "“After Merging and Splitting: Lineage Granularity and a Deterministic Rewriting "
   "Baseline for Migrating Historical SQL”.", size=10.5)
PB("Schema evolution invalidates historical SQL. Our study isolates a failure mode that "
   "column-level lineage cannot repair: when several columns are merged into one delimited "
   "column, a column-level mapping returns the concatenated string instead of the original "
   "value, so the migrated query executes without error and returns wrong data. We show "
   "that moving lineage from the column level to an executable expression that pins segment "
   "positions raises execution equivalence from 2.6% to 79.6% on merged columns, and that "
   "the effect is absent for splits. Across seven held-out databases the direction recurs "
   "in six of seven independent change events, and we report the single negative case in "
   "the main results rather than in the limitations.", size=10.5)
PB("Because a reviewer may reasonably ask whether a language model is needed at all, the "
   "manuscript includes a deterministic rewriter built on the same lineage. Given identical "
   "lineage it outperforms the tested model by 20.4 points on merged columns and 22.1 on "
   "splits, and no development item is solved by the model alone. We state plainly that the "
   "gold SQL in our construction is generated by the same substitution rule the rewriter "
   "applies, so the rewriter's ceiling is inflated by construction; we therefore present "
   "this comparison as a paired difference within our setting, not as a general engineering "
   "recommendation. The finding that matters for the community is diagnostic: the bottleneck "
   "lies in discovering lineage, not in generating SQL.", size=10.5)
PB("The work fits the journal's scope as an applied AI study with a controlled evaluation "
   "protocol and a released artefact. All scripts, per-item outputs, statistical routines "
   "and the environment specification are provided; a Data Availability Statement is "
   "included. The manuscript is original, is not under consideration elsewhere, and all "
   "authors have approved the submission. We declare no competing interests.", size=10.5)
PB("Sincerely,", size=10.5)
PB("（通讯作者姓名、单位、邮箱、ORCID）", size=10.5)
NOTE("第二段是这封信的实质内容，第三段是主动交代。**不要在信中写「LLM 提升了迁移能力」**——"
     "这与第三章的摘要口径冲突，而编辑会同时看到两处。"
     "第四段的 artefact 与 Data Availability 必须真实存在，"
     "投稿前先完成材料包归档与 DOI 申请（见第九节）。")

# ===================================================================== 八
H("八、必须与摘要同进同出的四条限定")
P("以下四条，在摘要、结论、投稿信与贡献点中只要缺一条，就会重新出现第一节表里的不一致。")
PL(f"**① 编码可逆性**：写「两段、可逆、非空、不含分隔符的取值分布」，不写「一般的列合并」。"
   f"真实数据侧的证据是 21.5 万行扫描、0 行越界——"
   f"这恰好说明构造没有对这类型数据做压力测试，而不是说明数据干净。")
PL(f"**② 构造内上界**：规则 100% 必须与「gold 由同一条替换规则机械生成」同句出现，"
   f"或紧邻一句内出现。**只写数字不写来源，是本轮评审的第一条意见。**"
   f"这一条对摘要、贡献三、结论、投稿信四处同时生效。")
PL(f"**③ 单元与方向**：留出库一律写「{_ev['positive']} 正 {_ev['negative']} 负」，"
   f"不写「方向完全重现」；EvoSchema 对照一律按映射单元写（合并档仅 3 个映射、"
   f"最大一个占 92.7%，单元级 p = 0.5），不用条目级 p = 0.0018 替代事件级证据。")
PL(f"**④ 星号与评测范围**：一律写「投影星号题在本构造下无法被忠实评价，样本中恰为 0 条」，"
   f"不写「已支持 SELECT *」「经探针验证了星号能力」。"
   f"合成探针只说明实现没写错，不构成评测证据。")
NOTE("第四条是本轮新增。第九轮评审对它的措辞是「不能暗示星号能力已经通过基准验证」——"
     "这类暗示往往不是刻意写的，而是「能力已实现并经探针验证」这种顺手的句子造成的。"
     "上一版 7.10 的已闭环项里就有一句这样的表述，已在章内 v10 改掉。")

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

CAP("表 9-2  Applied Intelligence 的三项强制或高概率要求")
TABLE(["要求", "状态", "动作"],
      [["Data Availability Statement（原创研究强制）",
        "材料包 v4 已具备（19 统计产物、30 脚本、逐题输出 4,490 条）",
        "归档至 Zenodo/GitHub 并取得 DOI，在正文末加声明段"],
       ["代码可用性（非强制但必被问）",
        "同上，含 ENVIRONMENT.md 与最短复现路径", "在 Data Availability 中一并给出链接"],
       ["摘要 150–250 词、关键词 4–6 个",
        "__ABSLENOK__", "已按实测值对齐，无需再调"]],
      [5.0, 6.0, 5.2], ["l", "l", "l"], font=8.2)
NOTE("另需注意：该刊不在中科院预警名单（2020–2026 各版均未列入），"
     "但 2025 版分区表中小类「计算机:人工智能」为 4 区（大类计算机科学 3 区）。"
     "若毕业或考核按小类认定，需先与本单位确认口径，这不是稿件层面的问题。")

CAP("表 9-3  Data Availability Statement 模板（按实际情况删改）")
PB("The datasets analysed are derived from the BIRD benchmark, which is publicly "
   "available. The evolution operators, item-level outputs, and all scripts required to "
   "reproduce every reported number are archived at [repository DOI]. The environment "
   "specification, including interpreter and dependency versions and the shortest "
   "reproduction path, is included in the archive. No restricted or proprietary data were "
   "used.", size=10.5)

# ===================================================================== 十
H("十、数字 → 出处自检表")
P("粘贴前逐行核对。任一行对不上，说明章内章外已经开始分叉。")
TABLE(["数字／表述", "取值", "出处"],
      [["主实验样本", f"{N_ALL} 条 = 改名 {_n('rename_all')} + 合并 {_n('merge')} + "
        f"拆分 {_n('split')}", "stats_ch7_v5.json main"],
       ["合并档阶梯", f"C2 {_mg['c2']['rate']:.1f} / C3m {_mg['c3m']['rate']:.1f} / "
        f"C3mp {_mg['c3mp']['rate']:.1f}", "表 7-21（合并行）"],
       ["拆分档阶梯", f"C2 {_sp['c2']['rate']:.1f} / C3m {_sp['c3m']['rate']:.1f} / "
        f"C3mp {_sp['c3mp']['rate']:.1f}", "表 7-21（拆分行）"],
       ["规则·自动谱系", f"合并 {_rm['c3r_infer']['rate']:.1f} / 拆分 "
        f"{_rs['c3r_infer']['rate']:.1f} / 留出合并 {_rhm['c3r_infer']['rate']:.1f} / "
        f"留出拆分 {_rsh['c3r_infer']['rate']:.1f}", "表 7-21"],
       ["规则 − LLM", f"+{_d_mg:.1f}（{P_mg['a_only']} : {P_mg['b_only']}）／"
        f"+{_d_sp:.1f}（{P_sp['a_only']} : {P_sp['b_only']}）／"
        f"+{_d_hm:.1f}（{P_hm['a_only']} : {P_hm['b_only']}）", "表 7-22"],
       ["两因子分解", f"合并 证据 {_m(D['merge']['evidence_effect_pt'])} / 执行器 "
        f"{_m(D['merge']['executor_effect_pt'])}；拆分 证据 "
        f"{_m(D['split']['evidence_effect_pt'])}", "表 7-23"],
       ["规则失败画像", f"dev 零失败；留出合并漏抽映射致执行报错；"
        f"COUNT(*) 修正 {RWA['count_star_items']} 条", "表 7-24、表 7-25"],
       ["留出库", f"7 库 165 条；事件 {_ev['positive']} 正 {_ev['negative']} 负；"
        f"合并 {_rhm['c3mp']['rate']:.1f} vs 列级 {_rhm['c3m']['rate']:.1f}",
        "表 7-33、表 7-34"],
       ["自动谱系缺口", f"oracle − 自动 = {_gap_ho:.1f}pt（留出合并）", "表 7-21"],
       ["编码边界", "21.5 万行扫描、0 行越界；探针 NULL_SIB 0/2、SEP_IN_A 0/2",
        "表 7-35、表 7-36"],
       ["星号探针", f"投影星号 {RWA['projection_star_items']} 条；P1–P4 自检通过",
        "表 7-26"],
       ["EvoSchema 单元级", "合并 3 映射、最大占 92.7%、p = 0.5", "表 7-42"]],
      [2.5, 9.4, 3.6], ["l", "l", "l"], font=8.2)
NOTE("表中「表 7-xx」为第 7 章 v10 重排后的编号，与 docx 内一致。"
     "v9 到 v10 只增了两段文字、未增删表格，故表号与 v9 相同；"
     "若后续增删表格触发重排，须同步本表。")

# ===================================================================== 十一
H("十一、本文件不做的事")
PL("不改第 7 章的数值与表格结构；章内证据见《投稿版_第7章_实验结果_v10.docx》。")
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
           "第三段中标出的两个句子不可删："
           "「within this construction…upper bound」承担第八节第 ② 条义务，"
           "「no item is solved by the model alone」是开发集上唯一对模型不利的硬事实。"
           "若需压到 200 词以内，删第二段的 276/230/104 三个分项数，保留总数 610。")
_CNNOTE = (f"实测 {_CN_N} 字（含标点）。此处已修正上一版「约 400 字」的标注错误——"
           f"上一版标注 400 字、实测 576 字，属文件内部前后矛盾，第九轮评审已指出。"
           "另需说明：Applied Intelligence 只约束英文摘要 150–250 词，"
           "中文摘要不受该刊字数限制，仅供中文版本或作者自用，故此处只报实测值、不设目标值。"
           "若需压缩，删第一段第二个设问与第三段的区间表述，"
           "但「构造内上界」与「6 正 1 负」必须同时保留。")
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
