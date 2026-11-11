# -*- coding: utf-8 -*-
"""生成《第7章v11修订说明_回应第十轮评审.docx》。

逐条对应第九轮评审的五条意见与三项优先事项。每条给出：原意见、改在哪、
改成了什么、以及哪些只能限定不能消除。
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
OUT = "第7章v11修订说明_回应第十轮评审.docx"

EN, CN, CN_H = "Times New Roman", "宋体", "黑体"

RVR = json.load(open("results/rule_vs_llm.json", encoding="utf-8"))
G, D = RVR["groups"], RVR["decomposition"]
RWA = json.load(open("results/rule_rewrite.json", encoding="utf-8"))["audit"]
PSX = json.load(open("results/probe_star_expand.json", encoding="utf-8"))
HLD = json.load(open("results/holdout_stats.json", encoding="utf-8"))
EVE = json.load(open("results/evoschema_events.json", encoding="utf-8"))
S = json.load(open("results/stats_ch7_v5.json", encoding="utf-8"))["main"]
ABM = json.load(open("results/ablation_matching.json", encoding="utf-8"))
assert RVR["version"] == "rule-vs-llm-v2"
assert ABM["version"] == "ablation-matching-v1"
_AB = ABM["groups"]
_EB, _EM = _AB["M1+形态指纹"]["段位档"], _AB["M1+形态指纹"]["非段位档"]
_B2M, _B3M = _AB["B2 值域重叠"]["非段位档"], _AB["B3 嵌入语义"]["非段位档"]
_ABP = ABM["paired"]["M1+形态指纹 vs B3 嵌入语义"]

_ev = HLD["event_level"]["merge_col_holdout"]
_rhm, _rsh = G["merge_holdout"]["rates"], G["split_holdout"]["rates"]
_rm, _rs = G["merge"]["rates"], G["split"]["rates"]
_pm = G["merge"]["paired"]["c3r_infer-c3mp"]
_ps = G["split"]["paired"]["c3r_infer-c3mp"]
_phm = G["merge_holdout"]["paired"]["c3r_infer-c3mp"]
_gap = _rhm["c3r_gt"]["rate"] - _rhm["c3r_infer"]["rate"]
_mg, _sp = S["merge"], S["split"]

doc = Document()
sec = doc.sections[0]
sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
sec.left_margin = sec.right_margin = Cm(2.4)
sec.top_margin = sec.bottom_margin = Cm(2.4)
st = doc.styles["Normal"]
st.font.name = EN
st.font.size = Pt(10.5)
st._element.rPr.rFonts.set(qn("w:eastAsia"), CN)
st.paragraph_format.line_spacing = 1.3
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
    p.paragraph_format.space_before = Pt(14 if lvl == 1 else 9)
    p.paragraph_format.space_after = Pt(5)
    p.paragraph_format.keep_with_next = True
    _run(p, txt, bold=True, size={1: 14, 2: 11.5}[lvl], font=CN_H)
    return p


def P(txt, indent=True, size=10.5):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    if indent:
        p.paragraph_format.first_line_indent = Pt(21)
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
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.keep_with_next = True
    _run(p, txt, bold=True, size=9)
    return p


def NOTE(txt):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(8)
    emit(p, txt, 8.6)
    return p


def TABLE(head, rows, widths, aligns, font=8.4):
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



# ===================================================================== 标题
p_ = doc.add_paragraph()
p_.alignment = WD_ALIGN_PARAGRAPH.CENTER
_run(p_, "第 7 章 v11 修订说明", bold=True, size=17, font=CN_H)
p_ = doc.add_paragraph()
p_.alignment = WD_ALIGN_PARAGRAPH.CENTER
_run(p_, "回应第十轮评审：补匹配基线对照、主张重摆、统一分析单位", size=11, font=CN)
NOTE("本轮新增一个脚本（ablation_matching.py）、一张章内表（7-16）与一次全文主张重摆。"
     "全部改动中，只有一项是新增数据，其余是重新组织已经存在的数据与句子。"
     "最重要的一点写在第二节：我在第十轮提出的三条意见里，第一条经核查是错的——"
     "那三条匹配基线早已实现、早已跑过、也早已写进正文。")

# ===================================================================== 一
H("一、三条意见的核查结果")
TABLE(["意见", "提出时的判断", "核查后的实际情况", "本轮处理"],
      [["补 M1 与值域/实例级匹配方法的对照",
        "认为缺这项实验，创新性无法评估",
        "三条基线（名字、值域重叠、嵌入）早已实现并跑过，表 7-15 已报告题级恢复率；"
        "缺的不是实验而是口径",
        "按 op 级重报并拆出「类型一致」，新增表 7-16；三条基线配对 15:0"],
       ["主张重摆：以粒度为核心",
        "认为摘要首句仍以「比较两种执行器」为主",
        "属实。且规则 100% 因构造循环性不能承担卖点",
        "摘要、贡献、结论、投稿信、标题五处重写"],
       ["统一分析单位、补跨模型矩阵与解码参数",
        "认为单位混用、矩阵残缺、参数未报",
        "解码参数（temperature=0、max_tokens）早已在表 7-9；"
        "单位确实无选择规则；矩阵确实只有 4/2/1",
        "7.1.5 增固定规则；7.1 增矩阵分级声明"]],
      [3.6, 3.6, 5.6, 3.8], ["l", "l", "l", "l"], font=8.2)
NOTE("第一行的教训值得单独记：审稿意见也会错，而错的方式往往是「按论文的印象提要求」，"
     "不是「按论文的内容提要求」。收到意见第一件事应是核对正文与产物，而不是直接动手补实验——"
     "若照意见重做一遍基线，会得到同一批数字，然后误以为「补上了」。")

# ===================================================================== 二
H("二、第一条：一个被口径掩盖的硬事实")
P("三条基线在 baselines_lineage.py 中实现，产物在 results/base_acc_*.json，"
  "表 7-15 报告的是**题级**恢复率。问题出在「题级」这个口径上："
  "它把「映射到哪个列」与「这一列由哪几段构成」合成了一个数，"
  "于是「列对上了但段位全错」这种情形，在表里表现为「恢复率低一点」。")
CAP("表 A  谱系抽取的 op 级质量（同一 gold、同一判定函数）")
_LAB = {"B1 列名相似": "B1 列名", "B2 值域重叠": "B2 值域重叠",
        "B3 嵌入语义": "B3 嵌入", "M1 值域指纹": "M1 值域指纹",
        "M1+形态指纹": "M1+形态指纹"}
_rows = []
for _l in ["B1 列名相似", "B2 值域重叠", "B3 嵌入语义", "M1 值域指纹", "M1+形态指纹"]:
    _r = [_LAB[_l]]
    for _g, _k in (("段位档", "recovered"), ("段位档", "type_ok"),
                   ("非段位档", "recovered"), ("非段位档", "type_ok")):
        _s = _AB[_l][_g]
        _pct = _s["recall_pct"] if _k == "recovered" else _s["type_pct"]
        _r.append(f"{_s[_k]}/{_s['n']}（{_pct:.1f}）")
    _rows.append(_r)
TABLE(["方法", "段位档 目标列", "段位档 类型一致", "非段位档 目标列", "非段位档 类型一致"],
      _rows, [3.4, 3.1, 3.1, 3.1, 3.1], ["l", "r", "r", "r", "r"], font=8.2)
NOTE(f"段位档 {_EB['n']} op（合并 10 + 拆分 5），非段位档 {_EM['n']} op。"
     f"配对：M1+形态与三条基线在段位档的 op 类型一致上均为 "
     f"{_ABP['ours_only']}:{_ABP['base_only']}，精确符号检验 p = {_ABP['p']:.1e}，"
     "即 n = 15 时的最小可达值。括号内为百分比。")

P(f"这张表说明的事比「补了一条基线」重要。B2/B3 在段位档的目标列恢复各 {_AB['B2 值域重叠']['段位档']['recovered']}/15，"
  f"看起来是「找到了一半」；但类型一致是 **0/15** ——一条段位都没恢复。"
  "它们找到的是「哪两列并进了哪一列」，而端到端环节需要的是「第几段」。"
  "一个不给出段位的谱系，与完全不提供谱系等价。")
P(f"这就解释了全章最强的那组数字：合并档 C3m {S['merge']['c3m']['rate']:.1f}% 与 "
  f"C3mp {S['merge']['c3mp']['rate']:.1f}% 相差 "
  f"{S['merge']['c3mp']['rate'] - S['merge']['c3m']['rate']:.1f}pt。"
  "v10 的表述是「列级映射精度不足」，读者据此会以为多给些数据就能补上；"
  "真实的机制是列级谱系在结构上无法表达段位，多给数据也补不上。"
  "**口径选软了，最强结论跟着变弱。**")
P(f"同一张表也第一次量出了 M1 的代价：非段位档 {_EM['type_pct']:.1f}%，"
  f"低于 B2 的 {_B2M['type_pct']:.1f}% 与 B3 的 {_B3M['type_pct']:.1f}%。"
  "若任务只要求列对，嵌入基线是更好的选择。"
  "这一点已写进 7.2 节、贡献二与摘要第四段，位置与 15/15 并列。")

# ===================================================================== 三
H("三、第二条：主张重摆的落点")
TABLE(["位置", "v10 的写法", "v11 的写法"],
      [["摘要首句", "「我们检验两个问题：谱系要多细；以及给定谱系时确定性重写是否不弱于模型」",
        "「要问的核心问题不是用哪个模型重写，而是谱系必须细到哪一层」"],
       ["核心数字", "规则改写 100%（承担卖点）",
        "同一执行器：列级 0.9% vs 段位 100.0%（承担卖点）；规则 100% 降为佐证"],
       ["M1 定位", "「把误差来源定位到谱系发现」",
        "「段位档 15/15 对 0/15（p=6.1e-5），但非段位档 91.8% 低于嵌入的 100%」"],
       ["标题", "A 案：列合并与列拆分之后：谱系粒度与确定性重写基线（两变量并列）",
        "A 案：谱系要细到哪一层：schema 演化下历史 SQL 迁移的粒度效应"],
       ["贡献排序", "粒度证据第一、误差定位第二、规则基线第三",
        "粒度证据第一（改为同执行器对照）、M1 带边界第二、执行器效应第三"]],
      [2.6, 6.6, 7.4], ["l", "l", "l"], font=8.2)
P("重摆的依据是一条逻辑，不是偏好：规则 100% 由任务定义决定——"
  "gold SQL 是用同一条替换规则机械生成的——所以它不能作为发现。"
  "而「同一执行器，列级 0.9%、段位 100%」不是同义反复："
  "两侧跑的是同一个改写器、同一套判定，唯一变化的是谱系表示。"
  "这是全文唯一能同时绕开「构造循环性」与「LLM 不如规则」两个质疑的结果。")
P("把它放到首句，论文的主张才与证据同构。此前的问题不是数字不真，"
  "而是最强的那句话放错了位置，导致读者从最弱的一处进入全文。")

# ===================================================================== 四
H("四、第三条：三项里两项需要补，一项本来就是清楚的")
TABLE(["项", "核查结果", "本轮处理"],
      [["解码参数", "表 7-9 已写 temperature = 0、max_tokens 4000/3000，"
        "并在正文说明服务端无随机种子",
        "无需改。新增一句：这是商用 API 复现的固有限制"],
       ["分析单位", "条目、op、事件、库、映射五种混用，每处有说明但无选择规则",
        "7.1.5 增固定规则：描述性按条目；推断按最粗可用单位；条目级显著仅为必要条件"],
       ["跨模型矩阵", "实际覆盖为改名 4 模型、合并 2 模型、拆分 1 模型",
        "7.1 增分级声明，「四模型一致」不得外推到拆分档"]],
      [2.4, 8.8, 5.4], ["l", "l", "l"], font=8.2)
P("分析单位那条之所以能站住，是因为它能指出自己第一次生效的位置："
  "EvoSchema 外部对照的条目级 p = 0.0018 被降为映射单元级 p = 0.5，"
  "正文据此把「显著优于」改成了「方向一致但样本不足」。"
  "事后为自洽而定的规则，读者看得出来；有先例的规则才有约束力。")
P("跨模型矩阵必须分级说，因为真实覆盖就是 4/2/1。"
  "把这条写进 7.1 之后，读者不会把改名档的四模型一致性读成全任务稳健，"
  "也不会在拆分档追问一个本章没做过的复现。")

# ===================================================================== 五
H("五、两个不利事实的位置调整")
PL(f"**① M1 在通用匹配上不占优**：非段位档 {_EM['type_pct']:.1f}%，低于 B3 的 "
   f"{_B3M['type_pct']:.1f}%。此前只在 7.3 的文献边界段用「本章不主张列匹配精度更优」带过，"
   f"现在给出数字，并与 15/15 并列写进贡献二。")
PL(f"**② 规则 100% 是构造内上界**：gold SQL 由同一替换规则生成，规则侧上界被构造抬高。"
   f"此前已声明，但声明之后该数字仍在承担卖点。本轮降为佐证，"
   f"并在摘要、贡献、结论、投稿信四处与列级 0.9% 成对出现。")
NOTE("两条都不新，新的是它们的表述强度被提到与事实相符。"
     "主动写清的好处很实际：审稿人自己算出 91.8% 那一行时，不会连带怀疑其余数字；"
     "而被他自己发现作者略去，整段正向结果都要重新打折。")

# ===================================================================== 六
H("六、变与不变")
TABLE(["变了", "没变"],
      [["新增章内表 7-16（op 级分段对照）与 4 段正文",
        "全部原有数值，包括 C1–C3mp 阶梯、规则基线四档、因素分解、留出库、EvoSchema 单元级"],
       ["7.2 节两段改写：把「精度不足」改为「能力缺失」，并把 M1 的不利数字写入",
        "结论方向：粒度在合并方向有效、拆分方向为零、规则基线限于构造内"],
       ["7.1.5 增分析单位规则；7.1 增矩阵分级声明",
        "表 7-9 的解码参数与复现说明"],
       ["章外五处重写：标题、英文摘要（249 词）、中文摘要（623 字）、贡献点、结论、投稿信",
        "样本量、库数、统计方法、判定口径"]],
      [6.2, 9.8], ["l", "l"], font=8.2)
NOTE("表号因新增一张表而后移一位。章外自检表因此改为按表题从 docx 动态反查表号，"
     "不再手写常量——上一版的教训是：正文与表题由脚本同步改写，写死的表号不会，"
     "而自检表自己分叉过一次（v9 把「合并档阶梯」标成表 7-16、7-18）。")

# ===================================================================== 七
H("七、章外同步：五处已重写，可直接粘贴")
PL("**标题**：换为 A 案「谱系要细到哪一层：schema 演化下历史 SQL 迁移的粒度效应」。"
   "v10 的 A 案降为备选，因为它并列了两个自变量，而数据显示表示效应大一个数量级。")
PL("**英文摘要**：249 词，五段结构（问题／方法／粒度对照／M1 定位／范围）。"
   "0.9% 与 100.0% 同句出现，这是本轮最关键的技术性要求。")
PL("**中文摘要**：623 字，段序与英文一致，第三、四段不可互换。")
PL("**贡献点**：三条重排，第二条改为带边界的正面贡献。")
PL("**投稿信**：第二段给主命题，第三段主动交代规则基线更强与 M1 的短板。")
NOTE("全部在《摘要_标题_贡献点_结论_v11.docx》中，含一处「数字 → 出处」自检表（表号已动态核对）。"
     "需要你自己填的只有 DOI、作者信息与 ORCID。")

# ===================================================================== 八
H("八、仍未闭环：六条，一条没少")
PL("① 需重写谓词结构而非替换标识符的迁移未测（口径改变、单位换算、等值谓词变区间）。"
   "**这是唯一可能反转「规则 > LLM」这一比较域的实验**。")
PL("② 三段以上合并未在真实库验证（实现支持任意段数，样本只有两段）。")
PL("③ 构造未覆盖 NULL 分量与值内分隔符（真实数据扫描 0 越界，探针每类仅 1–3 例）。")
PL("④ 无第二编码者与 kappa（单作者编码，附编码手册供复核）。")
PL("⑤ 检索实验未做，本章不主张检索效果。")
PL("⑥ 含投影星号的题目在本构造下无法被忠实评价，样本中恰为 0 条。")
NOTE("另一项虽不在六条内，但需要在投稿前处理：44 张表、2.2 万字的密度。"
     "Applied Intelligence 无字数上限，不等于没有可读性下限；"
     "两个最重要的结果（粒度效应、规则强基线）目前仍淹没在稳健性检验里，"
     "这属于结构问题，需单独一轮处理。")
P("就目前状态给一句判断：三条意见落实后，方法贡献那一格从「无法评估」变成「有数字、有边界」。"
  "但创新性仍偏增量——它建立在「段位恢复」这一个未被前人处理的子问题上，"
  "而不是一个新的迁移框架。投稿信与摘要已按这个强度写，不再往上加。")

# ===================================================================== 保存
doc.save(OUT)
print("saved:", OUT)
print("段落:", len(doc.paragraphs), "表格:", len(doc.tables))
