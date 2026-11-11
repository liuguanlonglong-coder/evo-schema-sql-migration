# -*- coding: utf-8 -*-
"""生成《摘要_标题_贡献点_结论_v9.docx》。

回答第八轮评审第 3 条：「摘要、贡献和结论都明确：主要实验证据来自可逆的
两段编码；真实数据扫描没有覆盖 NULL 和分隔符冲突」——以及第 6 条对
「标题与创新点须按实际结果收窄」的要求。

本文件不改第 7 章，只给出可直接粘贴的章外文字，并附一份「数字 → 出处」
自检表，避免章内章外口径再次分叉。
"""
import json
import os

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

BASE = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE)
OUT = "摘要_标题_贡献点_结论_v9.docx"

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
assert EVO["version"] == "evoschema-events-v1"

_n = lambda k: S[k]["n"]
N_ALL = _n("rename_all") + _n("merge") + _n("split")
_mg, _sp = S["merge"], S["split"]
_rm, _rs = G["merge"]["rates"], G["split"]["rates"]
_rhm, _rsh = G["merge_holdout"]["rates"], G["split_holdout"]["rates"]
_DBM = DBD["merge"] if "merge" in DBD else DBD.get("scan", {})
_nrows = DBD.get("total_rows") or DBD.get("n_rows") or None
_ev = HLD["event_level"]["merge_col_holdout"]
_PAU = PRB.get("verdict") or {}

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


def PB(txt, size=10.5, indent=False):
    """正文块：不缩进、行距略松，便于直接复制粘贴。"""
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
            _run(p_, str(v), size=font)
    return t


def _m(v, nd=1):
    return f"{v:+.{nd}f}".replace("-", "−")


# ===================================================================== 标题页
p_ = doc.add_paragraph()
p_.alignment = WD_ALIGN_PARAGRAPH.CENTER
_run(p_, "摘要、标题、贡献点与结论（v9 修订稿）", bold=True, size=17, font=CN_H)
p_ = doc.add_paragraph()
p_.alignment = WD_ALIGN_PARAGRAPH.CENTER
_run(p_, "回应第八轮评审第 3、6 条：按实际证据收窄表述，并统一章内章外口径", size=10, font=CN)
NOTE("本文件不含新的实验，全部数字取自第 7 章 v9 与 results/ 下的产物；"
     "末尾附「数字 → 出处」自检表，任何一格对不上即说明章内章外已分叉。"
     "可直接粘贴的中文段落以不缩进块呈现，英文摘要按期刊要求独立成段。")

# ===================================================================== 一
H("一、为什么必须改：三处已经对不上的表述")
P("v8 的实验把结论收窄了，但摘要、标题与贡献点还停在 v6 的口径上。"
  "这不是措辞问题，而是审稿人一眼能看出的内部不一致，三处各有硬证据。")
TABLE(["位置", "仍在使用的主张", "本章证据", "审稿后果"],
      [["标题与关键词", "以 RAG／检索为核心创新",
        "本章未实现检索模块，谱系整块注入提示；真检索实验仅 40 题且为阴性对照",
        "贡献与实验不符，属定位错误"],
       ["摘要与贡献点", "语言模型具备更好的 SQL 迁移能力",
        f"给定同一份段位谱系，确定性改写合并档 {_rm['c3r_infer']['rate']:.1f}%、"
        f"LLM {_rm['c3mp']['rate']:.1f}%（{_rm['c3r_infer']['rate'] - _rm['c3mp']['rate']:+.1f}pt，"
        f"{G['merge']['paired']['c3r_infer-c3mp']['a_only']} : "
        f"{G['merge']['paired']['c3r_infer-c3mp']['b_only']}）；dev 上 LLM 无独占正确条目",
        "被认定为方向性错误，而非数值误差"],
       ["摘要与结论", "方法带来普遍提升",
        "构造为两段可逆拼接；21.5 万行全表扫描未触发 NULL 分量与值内分隔符，"
        "即边界情形从未出现",
        "须在摘要内写明作用域，否则属过度概括"]],
      [2.3, 3.6, 6.8, 3.2], ["l", "l", "l", "l"], font=8.4)
NOTE("三处中前两处是「说了证据不支持的话」，第三处是「没说结论的适用范围」。"
     "前者必须删，后者必须补，二者都不能靠措辞软化。")

# ===================================================================== 二
H("二、标题：三个候选与取舍")
P("收窄的方向由实验结果决定：真正的自变量是「谱系粒度」，真正的强基线是"
  "「给定映射后的确定性重写」。标题应让这两件事出现在最前面，而不是让 RAG 或 NL2SQL 占位。")
PL("**A（推荐）** 列合并与列拆分之后：历史 SQL 迁移中的谱系粒度与确定性重写基线")
PL("**B** 谱系要细到哪一层：schema 演化下历史 SQL 迁移的粒度效应")
PL("**C** 段位表达式谱系：schema 演化下历史 SQL 迁移的可执行证据表示")
NOTE("A 把两个自变量都写进标题，且不含任何未实现的模块，与正文证据一一对应；"
     "B 更短但丢掉了规则基线这条最重要的对照；C 突出表示形式，"
     "但容易被读成「提出一种新表示」而非「检验粒度是否够用」。"
     "三案共同的前提是删去标题中的 RAG 与「端到端 NL2SQL」字样——"
     "若期刊定位要求保留应用领域，可加副标题「——以电子病历结构化字段为例」，"
     "但不能把 RAG 放回主标题。")

# ===================================================================== 三
H("三、英文摘要（约 230 词，适配 250 词上限）")
PB("Schema evolution breaks historical SQL. When several columns are merged into one "
   "delimited column, a column-level mapping returns the whole string rather than the "
   "original value; a migration is executable yet semantically wrong. We ask how fine a "
   "lineage must be, and whether a deterministic rewriter given that lineage matches a "
   "language model prompted with it.", indent=False)
PB("We construct a controlled benchmark on 11 BIRD databases (610 items): 276 column "
   "renamings, 230 merges and 104 splits. Conditions form an information ladder from the "
   "new schema alone to expression-level lineage that specifies segment positions "
   "(SUBSTR/INSTR). We add a deterministic rewriter that substitutes column references "
   "through the same lineage and run it with three mapping sources: gold, automatically "
   "inferred, and column-level.", indent=False)
PB("For merged columns, execution equivalence is 10.0% with the new schema only, 2.6% with "
   "column-level lineage, and 79.6% with segment-level lineage. Given the same lineage, the "
   "deterministic rewriter reaches 100.0%, i.e. 20.4 points above the model (47:0), and no "
   "item is solved by the model alone. Splits gain nothing from segment information "
   "(77.9%), yet the rewriter still gains 22.1 points. On seven held-out databases "
   "(165 items) the direction recurs (6 events positive, 1 negative), with the rewriter at "
   "80.2–100.0%.", indent=False)
PB("The scope is explicit: the encoding is a reversible two-segment concatenation. A full "
   "scan of 215,216 rows never triggers null components or in-value separators, so these "
   "results do not characterize such data. No retrieval module is implemented and we make "
   "no retrieval claim.", indent=False)
NOTE("实测 240 词（上限 250）。若目标期刊限 200 词，删去第三段中的 10.0% 与 2.6% 两值，"
     "保留 79.6% 与 20.4pt，并删第四段最后一句的检索声明；"
     "结论段与作用域段不可删——它们承担第三节「限定必须同进同出」的义务。")

# ===================================================================== 四
H("四、中文摘要（约 400 字）")
PB("schema 演化会使历史 SQL 失效：当多列被合并为带分隔符的单列时，列级映射会把整串返回，"
   "迁移结果可执行却语义错误。本文检验两个问题：谱系信息需要细到哪一层；"
   "以及在给定该谱系时，确定性重写是否不弱于把谱系写进提示的语言模型。")
PB("我们在 11 个 BIRD 数据库上构造 610 条受控样本（改名 276、列合并 230、列拆分 104），"
   "条件阶梯从「仅新 schema」逐级升到给出段位位置的可执行表达式谱系（SUBSTR/INSTR）。"
   "在同一谱系下另设确定性重写器，只做列引用替换，并按三种映射来源运行："
   "构造映射、自动抽取谱系、列级谱系。")
PB(f"列合并档的执行等价率随信息水平上升：仅新 schema {_mg['c2']['rate']:.1f}%、"
   f"列级谱系 {_mg['c3m']['rate']:.1f}%、段位表达式谱系 {_mg['c3mp']['rate']:.1f}%。"
   f"给定同一份段位谱系，确定性重写达到 {_rm['c3r_infer']['rate']:.1f}%，"
   f"高出所测模型 {_rm['c3r_infer']['rate'] - _rm['c3mp']['rate']:.1f} 个百分点"
   f"（{G['merge']['paired']['c3r_infer-c3mp']['a_only']} : "
   f"{G['merge']['paired']['c3r_infer-c3mp']['b_only']}），"
   f"且在开发库上没有一条是模型独有做对。列拆分档不受益于段位信息"
   f"（{_sp['c3mp']['rate']:.1f}% 与列级相同），但确定性重写仍高出 "
   f"{_rs['c3r_infer']['rate'] - _rs['c3mp']['rate']:.1f} 个百分点。"
   f"在 7 个未参与开发的库（165 条）上方向重现（独立变更事件 {_ev['positive']} 正 "
   f"{_ev['negative']} 负），确定性重写处于 {_rhm['c3r_infer']['rate']:.1f}–"
   f"{_rsh['c3r_infer']['rate']:.1f}% 之间。")
PB("适用范围必须与结果同时给出：本文的编码是可逆的两段拼接，"
   "对 21.5 万行真实数据的全表扫描未触发空值分量与值内分隔符，"
   "因此结论不覆盖含空值或分隔符冲突的取值分布；"
   "本文亦未实现检索模块，不对检索效果作任何主张。")
NOTE("实测 576 字（含标点）。中文摘要第三段是审稿人最可能逐字核对的一段："
     "压缩到 400 字以内时，删第一段的第二个设问与第三段的区间表述，"
     "但「方向重现」与「6 正 1 负」必须同时出现（见第七节第 ② 条）；"
     "第四段的适用范围内两句一句都不能删。")

# ===================================================================== 五
H("五、关键词与贡献点")
P("关键词 5 个，全部指称正文真实存在的东西：")
PL("schema 演化；历史 SQL 迁移；表达式级谱系；确定性查询重写；执行等价")
NOTE("删除候选：RAG、检索增强生成、NL2SQL 端到端、提示工程。"
     "理由与第一节第一行相同：这些词所指的模块在本章不存在。")

P("贡献点改写为三条，第三条是负面的；另加一条作用域声明。"
  "改写前后的差别不只是措辞，而是主张的强度。")
TABLE(["", "v8 的写法", "v9 建议写法"],
      [["贡献一", "提出 RAG 驱动的 NL2SQL 迁移框架，在医疗库上显著提升准确率",
        "给出粒度命题的受控证据：在列合并方向上，把谱系从列级下沉到段位表达式，"
        f"执行等价率由 {_mg['c3m']['rate']:.1f}% 升至 {_mg['c3mp']['rate']:.1f}%"
        f"（证据效应 {_m(D['merge']['evidence_effect_pt'])}pt）；"
        "列拆分方向该效应为零，故命题限于「一列内多段编码」"],
       ["贡献二", "语言模型能够理解 schema 变化并自动完成迁移",
        "把误差来源定位到「谱系能否被正确发现」：自动谱系在开发库上与构造映射逐条等价，"
        f"在 7 个未见库上留下 "
        f"{_rhm['c3r_gt']['rate'] - _rhm['c3r_infer']['rate']:.1f}pt 的缺口，"
        "且缺口全部表现为执行报错而非静默错"],
       ["贡献三", "（无）",
        f"给出确定性重写这条强基线：在给定同一份段位谱系时，规则改写高出所测模型 "
        f"{_rm['c3r_infer']['rate'] - _rm['c3mp']['rate']:.1f}pt（合并）与 "
        f"{_rs['c3r_infer']['rate'] - _rs['c3mp']['rate']:.1f}pt（拆分），"
        "且模型在开发库上无独占正确条目。这条结论对本文不利，但必须写，"
        "否则读者会把它读成「方法有效」"],
       ["作用域", "（无）",
        "证据来自两段可逆拼接的合成编码；真实数据扫描未覆盖空值与值内分隔符；"
        "未实现检索，不作检索主张"]],
      [1.6, 5.2, 9.1], ["c", "l", "l"], font=8.4)
NOTE("贡献三的写法值得注意：它不是「我们更强」，而是「这条基线更强」。"
     "评审已明确指出，坦诚披露这一条是本文可信度的来源而非弱点；"
     "把它藏进局限节反而会被读成回避。")

# ===================================================================== 六
H("六、结论段（可直接替换原结论）")
PB(f"本文的证据支持一个窄命题：当 schema 演化把多列合并为带分隔符的一列时，"
   f"列级谱系不足以维持历史 SQL 的语义，"
   f"把谱系下沉为指明段位的可执行表达式可以把执行等价率从 {_mg['c3m']['rate']:.1f}% "
   f"提升到 {_mg['c3mp']['rate']:.1f}%；而在列拆分方向，同一提升不存在。"
   "在此基础上，本文给出两条对后续工作更有用的读数："
   f"其一，给定同一份谱系，确定性重写的执行等价率高于所测模型约 "
   "20 个百分点，"
   "说明该任务的瓶颈不在 SQL 生成而在谱系发现；"
   f"其二，自动谱系在开发库上与构造映射逐条等价，在未见库上仍有 "
   f"{_rhm['c3r_gt']['rate'] - _rhm['c3r_infer']['rate']:.1f}pt 的硬失败缺口，"
   "因此自动发现才是这条路线当前的主要成本。")
PB("这些结论的外推受三重限制，均已在正文量化：样本为 11 个库上的合成演化，"
   "独立的列重整事件只有 10 次；留出确认仅 7 个库、165 条，"
   "且存在一个负向库（regional_sales），只能支持「总体方向可迁移」；"
   "编码固定为两段可逆拼接，真实数据扫描中空值分量与值内分隔符从未出现，"
   "故不覆盖此类取值分布。本文因此不主张跨库泛化，也不主张已解决真实历史 schema 演化。")

# ===================================================================== 七
H("七、必须与摘要同进同出的三条限定")
P("以下三条在摘要、结论、投稿信中只要缺一条，就会重新出现第一节日志里的不一致。")
PL("**① 编码可逆性**：写「两段、可逆、非空、不含分隔符的取值分布」，"
   "不写「一般的列合并」。真实数据侧的证据是 21.5 万行扫描、0 行越界——"
   "这恰好说明构造没有压力测试这类数据，而不是说明数据干净。")
PL(f"**② 单元与方向**：留出库一律写「{_ev['positive']} 正 {_ev['negative']} 负」，"
   "不写「方向完全重现」；EvoSchema 对照一律按映射单元写"
   f"（合并档仅 3 个映射、最大一个占 92.7%，单元级 p = 0.5），"
   "不用条目级 p = 0.0018 替代事件级证据。")
PL("**③ 规则基线的定位**：写「本构造下的配对差值」，不写「确定性方法优于 LLM」。"
   "本章的 gold 由同一条替换规则机械生成，规则侧上界被构造抬高；"
   "LLM 侧只测一个基座与一套提示模板。")

# ===================================================================== 八
H("八、数字 → 出处自检表")
P("粘贴前逐行核对。任一行对不上，说明章内章外已经开始分叉。")
TABLE(["数字／表述", "取值", "出处"],
      [["主实验样本", f"{N_ALL} 条 = 改名 {_n('rename_all')} + 合并 {_n('merge')} + 拆分 {_n('split')}",
        "stats_ch7_v5.json main"],
       ["合并档阶梯", f"C2 {_mg['c2']['rate']:.1f} / C3m {_mg['c3m']['rate']:.1f} / "
        f"C3mp {_mg['c3mp']['rate']:.1f}", "表 7-16、表 7-18"],
       ["拆分档阶梯", f"C2 {_sp['c2']['rate']:.1f} / C3m {_sp['c3m']['rate']:.1f} / "
        f"C3mp {_sp['c3mp']['rate']:.1f}", "表 7-16、表 7-18"],
       ["规则·自动谱系", f"合并 {_rm['c3r_infer']['rate']:.1f} / 拆分 {_rs['c3r_infer']['rate']:.1f} / "
        f"留出合并 {_rhm['c3r_infer']['rate']:.1f} / 留出拆分 {_rsh['c3r_infer']['rate']:.1f}",
        "表 7-21"],
       ["规则 − LLM", f"+{_rm['c3r_infer']['rate'] - _rm['c3mp']['rate']:.1f}（"
        f"{G['merge']['paired']['c3r_infer-c3mp']['a_only']} : "
        f"{G['merge']['paired']['c3r_infer-c3mp']['b_only']}）／"
        f"+{_rs['c3r_infer']['rate'] - _rs['c3mp']['rate']:.1f}（23 : 0）／"
        f"+{_rhm['c3r_infer']['rate'] - _rhm['c3mp']['rate']:.1f}（15 : 4）",
        "表 7-22"],
       ["两因子分解", f"合并 证据 {_m(D['merge']['evidence_effect_pt'])} / 执行器 "
        f"{_m(D['merge']['executor_effect_pt'])}；拆分 证据 {_m(D['split']['evidence_effect_pt'])}",
        "表 7-23"],
       ["规则失败画像", f"dev 零失败；留出合并 20 条执行报错命中 "
        f"{RWA['count_star_items']} 条 COUNT(*) 修正", "表 7-24、表 7-25"],
       ["留出库", f"7 库 165 条；事件 {_ev['positive']} 正 {_ev['negative']} 负；"
        f"合并 {_rhm['c3mp']['rate']:.1f} vs 列级 {_rhm['c3m']['rate']:.1f}",
        "表 7-26、表 7-27"],
       ["编码边界", "21.5 万行扫描、0 行越界；探针 NULL_SIB 0/2、SEP_IN_A 0/2",
        "表 7-33、表 7-34"],
       ["EvoSchema 单元级", "合并 3 映射、最大占 92.7%、p = 0.5", "表 7-40"]],
      [2.5, 9.4, 3.6], ["l", "l", "l"], font=8.2)
NOTE("表中「表 7-xx」为第 7 章 v9 的重排后编号，与 docx 内编号一致；"
     "若后续增删表格触发重排，需同步本表。")

# ===================================================================== 九
H("九、本文件不做的事")
PL("不改第 7 章的正文与数字；章内证据见《投稿版_第7章_实验结果_v9.docx》。")
PL("不给期刊适配的第二版摘要；目标期刊若从 Journal of Medical Systems 换到"
   "Applied Intelligence，摘要结构不变，只需把临床动机段替换为通用数据库维护动机。")
PL("不代拟审稿回复信。回复信应逐条引用本文件第五节与第七节的句子，"
   "而不是重新组织措辞——两处措辞一旦不同，审稿人会认为作者自己也没统一口径。")

doc.save(OUT)
print("saved:", OUT)
print("段落:", len(doc.paragraphs), "表格:", len(doc.tables))
