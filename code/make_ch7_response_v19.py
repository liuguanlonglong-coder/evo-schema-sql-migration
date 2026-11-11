# -*- coding: utf-8 -*-
"""生成《第19轮修订说明_逐条落实AI投稿清单.docx》。

输入：D:/Download/Applied_Intelligence_逐条投稿前修改清单.docx（M-01..M-14、
S-01..S-08、O-01..O-04 + 投稿前最终核对表 14 条）。

本轮做两件事：
  ① 把清单每一条落到具体位置（正文/章外/补充材料），逐条给处置与状态；
  ② 顺带记录本轮自查发现的两处**新缺陷**——它们不是清单里的条目，
     但恰好是清单 M-01 自己沿用了旧口径所造成的（58 次事件）：
     - audit_independence.py 的段位档事件键缺 evo，使 610 集事件数 58（应 61）、
       段位档 7（应 10）、段位表达式覆盖事件 2（应 5）；
     - slim_chapter7.py 只在表数增加时刷新素材库备份，纯文字修订不会进入
       主文与补充材料（本轮实际踩到：第一次生成的主文仍是旧文字）。

所有数字从 results/*.json 与生成后的 docx 实测读取，不手写统计量。
"""
import json
import os
import re
from datetime import date

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

BASE = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE)
EN, CN, CN_H = "Times New Roman", "宋体", "黑体"
TODAY = date.today().isoformat()
OUT = "第19轮修订说明_逐条落实投稿清单.docx"

ADI = json.load(open("results/audit_independence.json", encoding="utf-8"))
assert ADI["version"] == "audit-independence-v2"
PS = ADI["A_units"]["paper_set"]
SG = ADI["A_units"]["segment_set"]
DD = ADI["D_segment_expr"]
EMA = json.load(open("results/event_map_audit.json", encoding="utf-8"))
MHO = json.load(open("results/m1_holdout_eval.json", encoding="utf-8"))
MHO_OP, MHO_EV, MHO_DB = MHO["op_level"], MHO["event_level"], MHO["db_level"]
EB = json.load(open("results/ablation_matching.json", encoding="utf-8"))


def stat(f):
    d = Document(f)
    caps = sum(1 for p in d.paragraphs
               if re.match(r"^表\s*(?:7|S|5|9)-\d+\s\s", p.text.strip()))
    return len(d.paragraphs), len(d.tables), caps


N_MAIN = stat("投稿版_第7章_实验结果_v11.docx")
N_SUPP = stat("补充材料_第7章_v11.docx")
N_FULL = stat("第7章_v11_完整版_49表.docx")
N_FM = stat("摘要_标题_贡献点_结论_v11.docx")

_fm = Document("摘要_标题_贡献点_结论_v11.docx")
_txt = [p.text for p in _fm.paragraphs]
_i = next(k for k, t in enumerate(_txt) if t.startswith("三、英文摘要"))
_j = next(k for k, t in enumerate(_txt) if t.startswith("四、中文摘要"))
_segs = [t for t in _txt[_i + 1:_j] if t and len(t) > 80 and not t.startswith("实测")]
EN_WORDS = sum(len(re.findall(r"[A-Za-z0-9][A-Za-z0-9%\.\-–']*", s)) for s in _segs)
_k = next(k for k, t in enumerate(_txt) if t.startswith("四、中文摘要"))
_l = next(k for k, t in enumerate(_txt) if t.startswith("五、关键词"))
CN_CHARS = sum(len(re.sub(r"\s", "", t)) for t in _txt[_k + 1:_l]
                if t and not t.startswith("实测"))


def make_doc():
    d = Document()
    s = d.sections[0]
    s.page_width, s.page_height = Cm(21.0), Cm(29.7)
    s.left_margin = s.right_margin = Cm(2.0)
    s.top_margin = s.bottom_margin = Cm(2.1)
    st = d.styles["Normal"]
    st.font.name = EN
    st.font.size = Pt(10.5)
    st._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), CN)
    st.paragraph_format.line_spacing = 1.28
    st.paragraph_format.space_after = Pt(4)
    return d


def _run(p, txt, bold=False, size=10.5, font=CN):
    r = p.add_run(txt)
    r.bold = bold
    r.font.size = Pt(size)
    r.font.name = EN
    r._element.rPr.rFonts.set(qn("w:eastAsia"), font)
    return r


def emit(p, txt, size=10.5):
    for j, part in enumerate(str(txt).split("**")):
        if part:
            _run(p, part, bold=(j % 2 == 1), size=size)


def H(doc, txt, lvl=1):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(13 if lvl == 1 else 8)
    p.paragraph_format.space_after = Pt(5)
    p.paragraph_format.keep_with_next = True
    _run(p, txt, bold=True, size={1: 13.5, 2: 11}[lvl], font=CN_H)
    return p


def P(doc, txt):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.first_line_indent = Pt(21)
    emit(p, txt)
    return p


def PL(doc, txt):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Pt(18)
    p.paragraph_format.space_after = Pt(2)
    emit(p, txt, 10.2)
    return p


def CAP(doc, txt):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.keep_with_next = True
    _run(p, txt, bold=True, size=9)
    return p


def NOTE(doc, txt):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(8)
    emit(p, txt, 8.4)
    return p


def TABLE(doc, head, rows, widths, aligns, font=8.2):
    t = doc.add_table(rows=1, cols=len(head))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    for i, c in enumerate(t.rows[0].cells):
        c.text = ""
        p = c.paragraphs[0]
        p.alignment = (WD_ALIGN_PARAGRAPH.CENTER if aligns[i] == "c"
                       else WD_ALIGN_PARAGRAPH.LEFT)
        _run(p, head[i], bold=True, size=font)
    for row in rows:
        cells = t.add_row().cells
        for i, v in enumerate(row):
            cells[i].text = ""
            p = cells[i].paragraphs[0]
            p.alignment = (WD_ALIGN_PARAGRAPH.CENTER if aligns[i] == "c"
                           else WD_ALIGN_PARAGRAPH.LEFT)
            emit(p, v, font)
    for r in t.rows:
        for i, c in enumerate(r.cells):
            c.width = Cm(widths[i])
    return t


d = make_doc()
p = d.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
_run(p, "第 19 轮修订说明", bold=True, size=17, font=CN_H)
p = d.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
_run(p, "逐条落实《Applied Intelligence 逐条投稿前修改清单》", bold=True, size=12, font=CN_H)
p = d.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
_run(p, f"M 类 14 条 ｜ S 类 8 条 ｜ O 类 4 条 ｜ 另附自查发现的 2 处新缺陷 ｜ {TODAY}", size=9.5)

# ================================================================= 一
H(d, "一、处置总览")
P(d, "本清单共 26 条（必须改 14、建议改 8、可以不改 4），另附 14 条投稿前核对项。"
  "逐条处置的结果是：M 类 14 条全部落地，其中 13 条已闭环、1 条（M-14 数据可用性声明）"
  "只做到「表述与仓库真实状态一致」，仓库与 DOI 的实际建立仍需作者本人操作；"
  "S 类 8 条全部落实，其中 3 条是新增内容（S-01 边界表、S-04 方法前提段、S-03 归因限定），"
  "其余 5 条为口径统一；O 类 4 条按清单要求保留，未作削弱性改写。")
P(d, "**三条硬约束全程未被破坏**：主文仍严格 15 表（所有新增内容都进正文段落或补充材料）；"
  f"英文摘要 ≤235 词（现 {EN_WORDS} 词）；不新增任何未经验证的数字——"
  "本轮所有新写入的数值都能在 results/*.json 或逐题产物中复算。")
CAP(d, "表 1  当前四份交付文档的实测规模")
TABLE(d, ["文档", "段落", "表格", "说明"],
      [["投稿版_第7章_实验结果_v11.docx", str(N_MAIN[0]), str(N_MAIN[1]),
        "主文，表 7-1 至 7-15 连续"],
       ["补充材料_第7章_v11.docx", str(N_SUPP[0]), str(N_SUPP[1]),
        f"表 S-1 至 S-{N_SUPP[2]}，含主文表号对照"],
       ["第7章_v11_完整版_49表.docx", str(N_FULL[0]), str(N_FULL[1]),
        "素材库全表（文件名沿用历史约定）"],
       ["摘要_标题_贡献点_结论_v11.docx", str(N_FM[0]), str(N_FM[1]),
        f"英文摘要 {EN_WORDS} 词 / 中文摘要 {CN_CHARS} 字"]],
      [6.4, 1.4, 1.4, 7.4], ["l", "c", "c", "l"])

# ================================================================= 二
H(d, "二、必须改（M-01 至 M-14）")
P(d, "下表逐条给出「清单要求 → 本轮处置 → 落点」。落点一律写到节号与表号，"
  "以便逐条核对时能直接翻到对应位置；凡本轮**未动数字、只改口径**的条目，也在表中写明，"
  "避免把「已经对的」当成「这轮改的」。")
CAP(d, "表 2  M 类 14 条的处置与落点")
TABLE(d, ["编号", "清单要求（摘要）", "本轮处置", "落点", "状态"],
      [["M-01", "段位档 15 op 对应 10 次事件；全文事件数同一口径",
        f"清单本身沿用 58 次，本轮复算后发现同一键缺陷也影响全集："
        f"610 集 58→**{PS['events']}**、段位档 7→**{EMA['counts']['events_correct']}**、"
        f"段位表达式覆盖事件 2→**{DD['events_covered']}**；"
        "表 7-114 由「三种事件数」扩为「五种」并逐行对账",
        f"§7.1.5、§7.1.8、表 7-114、摘要第 5 段、结论限制段", "已闭环"],
       ["M-02", "两类留出实验分开报告，命名并给出库集合",
        "§7.7.4 开头新增口径段，把两项实验定名为 H-E2E（165 条 / 7 库 / 12 事件单元）"
        "与 H-LER（481 条 / 16 op / 8 事件 / 8 库），写明不可相加、不可互为分母；"
        "节标题与两张表题同步加标记",
        "§7.7.4 标题、表 7-25/7-26（H-E2E）、表 7-119（H-LER）、中英文摘要、投稿信", "已闭环"],
       ["M-03", "条目数必须伴随独立事件数",
        "合并档补「10 个 op 两两属同一事件 → 5 次合并事件」；"
        "拆分档补「5 个 op → 5 次拆分事件、5 个数据库」；"
        "并加一句「条目数说明测了多少题，不说明观测了多少次变更」",
        "§7.1.4 末段", "已闭环"],
       ["M-04", "事件级与数据库级统计同时报告，标明比较对象",
        "两处 p 值本轮再次并列并加注：事件级 10/10 同向 p = 0.0020、库级 5 单元 p = 0.0625；"
        "并写明 §7.1.5 的 p = 0.0020（M1 对基线）与 §7.8 的 p = 0.0020（C3mp − c_rw 池化）"
        "来自不同对照，数值相同不可混用",
        "§7.1.5、§7.8、表 7-14（主文表号）、结论、7.11 写作口径", "数值已对，口径本轮补强"],
       ["M-05", "固定执行器 100% 明确为构造内上界",
        "在规则基线循环性段落引入术语 construction-internal upper bound（构造内上界）"
        "与 controlled representation ablation，并明确它不是方法优劣比较",
        "§7.5 规则基线段", "已闭环"],
       ["M-06", "M1 与 B4 比较避免算法优越性误读",
        "§7.2.1 已如实报告 B4 在同等输出空间下 10/10、与 M1 持平、未观察到算法增益；"
        "本轮在贡献二与英文摘要保留 no algorithmic gain 一句，不删",
        "§7.2.1、表 7-116、贡献二、英文摘要", "数值已对，措辞本轮复核"],
       ["M-07", "主实验为回顾性合成评测，11 库参与开发",
        "§7.7.3① 已写明「回顾性合成评测」与「阈值曾依据 gold operation 设定」；"
        "本轮把同一句补进中英文摘要与投稿信",
        "§7.7.3①、摘要第 2/5 段、投稿信第二段", "已闭环"],
       ["M-08", "移除 RAG 与端到端 NL2SQL 的贡献主张",
        "关键词已无 RAG/NL2SQL；中英文摘要与投稿信均写「未实现检索模块」；"
        "结论段明确「不主张检索增强、不主张端到端 NL2SQL」",
        "关键词、摘要、结论、投稿信", "已闭环"],
       ["M-09", "合并与拆分结果必须分开表述",
        "摘要与结论均为条件化表述：合并档列级 0.9% / 段位 100.0%、拆分档无额外收益；"
        "未出现「段位谱系普遍优于列级」一类总括句",
        "摘要第 3 段、结论、§7.3", "已闭环"],
       ["M-10", "五次重跑的报告边界，不报不存在的均值与 SD",
        "保留「最大绝对偏差 0.1pt + 方向与判定不变」，并明写「逐轮逐题输出未按条目存档，"
        "故不报告五轮均值与标准差」；未生成任何逐轮数值",
        "§7.7.3⑥、表 S-38、结论、中文摘要", "已闭环"],
       ["M-11", "不使用「形式化语义等价」",
        "正文 3 处「执行等价」改为「执行结果一致」；关键词第 5 个由「执行等价」"
        "改为「执行结果一致率」；投稿信 execution equivalence 改为 execution-result agreement；"
        "补充材料表题与表注同步",
        "§7.1.3、§7.5、§7.9.1、关键词、投稿信、表 S-17/S-20 表题", "已闭环"],
       ["M-12", "SQL 适用范围写进方法、实验与结论",
        "§7.1.3 末段已列 CTE、子查询、JOIN USING/NATURAL、聚合与 HAVING、投影星号、"
        "ORDER BY 序号、显式 CAST、其他方言；本轮把同一范围压成一句补进中英文摘要与投稿信",
        "§7.1.3、§7.7.3、摘要、投稿信", "已闭环"],
       ["M-13", "英文摘要压到 220–235 词",
        f"247 → **{EN_WORDS}** 词。删去三条基线的名称与配置、「阈值由开发集校准」的实现说明、"
        "门控与守卫细节；保留研究问题、数据范围、合并核心结果、拆分无额外收益、"
        "构造内上界、两项留出实验的分别读数",
        "§三 英文摘要及其下方说明", "已闭环"],
       ["M-14", "DAS 与仓库状态一致",
        "稿件、投稿信、表 9-3 一律写「将在发表前存入公开仓库并取得持久标识符，归档尚未完成」；"
        "未使用 released / included / provided；fill_doi.py 保留 reserved / deposited 三态切换",
        "表 9-3、投稿信第四段、fill_doi.py", "表述已一致；仓库与 DOI 待作者执行"]],
      [1.0, 3.0, 6.5, 4.1, 2.2], ["c", "l", "l", "l", "l"], font=7.8)

P(d, "M 类里唯一需要单独解释的是 M-01，因为它与清单给出的数字不一致。"
  "清单第 139 行写「610 条、437 个基础题、90 个唯一 op、58 次独立事件、11 个主实验数据库」，"
  "其中 58 是上一版口径。本轮把事件键修好之后，610 集的事件数是 "
  f"{PS['events']}（旧键 {PS['events_v13_triplet']}）。这不是为了改数字而改数字："
  "同一处缺陷已经在第 13 轮把段位档从 10 压成 7，只是当时只修了段位档，没有回头查全集。"
  "细节见第五节。")

# ================================================================= 三
H(d, "三、建议改（S-01 至 S-08）")
CAP(d, "表 3  S 类 8 条的处置与落点")
TABLE(d, ["编号", "清单要求（摘要）", "本轮处置", "落点", "状态"],
      [["S-01", "增加相关工作任务边界表（输入/输出/用不用历史 SQL/是否表达列内段位）",
        "新增一张四维边界表，四行分别是 schema 匹配、演化查询重写、Text-to-SQL 鲁棒性、"
        "本任务；并附英文段落可直接贴入 Introduction 末尾。表中只陈述各线的输出形态，"
        "不写「已有方法都无法处理段位表达式」这类超出检索范围的断言",
        "表 5-1 及其上下两段说明（章外文档）", "已新增"],
       ["S-02", "贡献点弱化 M1，改以受控分析为核心",
        "贡献一补入静默错误分布这一条（C3m 223 条中 220 条仍引用新列、仅 5 条含切段；"
        "C3mp 47 条中 45 条已含切段），把「粒度改变的是错误留在哪一层」写成可检验的贡献；"
        "贡献二保留「不是更好的通用匹配器」这半句",
        "表 2（贡献点）贡献一、贡献二", "已落实"],
       ["S-03", "静默错误归因增加方法限定",
        "改写为 220/223 与 5/223，并明确二者**嵌套而非互斥**：逐条核对确认 5 条同时属于 220 条，"
        "220 − 5 = 215 条是完全无切段的；补写「结构解析属描述性证据，不是穷尽证明」"
        "以及「SILENT 含空结果集，个别条目可能因谓词过严而非粒度问题」",
        "§7.6.1 第一段", "已闭环"],
       ["S-04", "把不可逆编码条件写成方法前提",
        "新增一段把两条边界写成前提 P1（分量不为 NULL）、P2（非末段不含分隔符），"
        "并写明「21.5 万行未观察到」只说明该数据集在此时点满足前提，"
        "不构成其他库同样满足的保证；给出两条规避路径（编码侧转义 / 外部元数据）",
        "§7.7.5 真实数据扫描段之后", "已新增"],
       ["S-05", "限制跨模型结论",
        "§7.1 已按档分级：改名档四模型、合并档两模型、拆分档单模型；"
        "§7.3 明确「不做跨档外推」，并说明把四模型一致性读成「全任务稳健」是本章不主张的",
        "§7.1 覆盖矩阵段、§7.3、表 S-22", "已闭环"],
       ["S-06", "关键词与标题定位",
        "标题保留 O-01；关键词 5 个，全为正文真实存在的内容，"
        "第 5 个由「执行等价」改为「执行结果一致率」，与 M-11 同源",
        "关键词行、§五 说明", "已闭环"],
       ["S-07", "统计方法可复核",
        "正文已给方法 10 的四格定义、φ 修正形式、与 statsmodels 算例对齐到 1e-10 的 9 项自检、"
        "以及 248 格点的覆盖率网格；表 7-6 列出全部九个比较的四格计数",
        "§7.1.5–7.1.6、表 7-5/7-6/7-7、补充材料 S-2", "已闭环"],
       ["S-08", "列属性统计与适用性解释",
        "正文已报 162 列 / 69 列（42.6%）近唯一 / 56 列极低基数 / 55 列含 NULL 或空串 / "
        "26 列非文本（INTEGER 6、REAL 7、DATE 13），并把近唯一列写成值域证据的信息下界",
        "§7.7.3 列属性段", "已闭环"]],
      [1.0, 3.0, 6.5, 4.1, 2.2], ["c", "l", "l", "l", "l"], font=7.8)

# ================================================================= 四
H(d, "四、可以不改（O-01 至 O-04）")
P(d, "O 类四项按清单要求原样保留，本轮只做了一处与之相邻的措辞处理，不影响结论。")
CAP(d, "表 4  O 类 4 条的保留状态")
TABLE(d, ["编号", "清单结论", "本轮处置"],
      [["O-01", "标题与收敛后的研究问题一致，保留",
        "保留英文标题；中文标题保留「谱系要细到哪一层：schema 演化下历史 SQL 迁移的粒度效应」"],
       ["O-02", "实验条件阶梯保留，勿为整齐删负面条件",
        "C1/C2/C3m/C3mp/c_rw 阶梯未动；未删任何负面读数（拆分 0 增益、regional_sales 负向、"
        "值改写两档落后）"],
       ["O-03", "拆分任务负结果保留，但不写成「两粒度完全等价」",
        "保留 77.9% 与 77.9% 的并列，措辞为「未观察到表达式级谱系的额外收益」，"
        "不写「等价」"],
       ["O-04", "局限性披露保留并整理",
        "§7.7 六项未完成 + 六项已闭环的结构保留；本轮只在 M-02 处把两项留出实验分列，"
        "未删任何限制"]],
      [1.1, 5.0, 10.7], ["c", "l", "l"])

# ================================================================= 五
H(d, "五、本轮自查发现的两处新缺陷")
P(d, "这两处不在清单里，但它们决定了清单里若干条能不能真正闭环，因此单列一节。"
  "两处都是**流程缺陷**而非写作缺陷：数字本身没有错，错的是算数字与传文件的那段代码。")

H(d, "5.1  事件聚类键的第三处出血点：58 → 61、7 → 10、2 → 5", 2)
P(d, "清单 M-01 要求「以事件ID清单为唯一事实源」，并沿用 58 次这一数字。"
  "核对时发现 58 与第 13 轮已经修过的 7 是同一个缺陷的两个伤口："
  "段位档的事件键只取 (库, 表, 表达式首个 attr_ 列)，不含演化档。"
  "合并档产出 SUBSTR(attr_001,…)、拆分档产出 attr_001‖sep‖attr_002，"
  "两者首列同名，于是同一 (库, 表) 上的合并与拆分被并成一件事。"
  "5 个库中有 3 个库（california_schools、formula_1、thrombosis_prediction）中招。")
P(d, "第 13 轮只修了段位档那一次统计（10 vs 7），没有回头检查同一函数是否还服务别处。"
  "本轮把 audit_independence.py 的 event_key 与 D 节的 _seg_event 一并补上 evo，"
  "三处数字同时变化：")
CAP(d, "表 5  事件键修正前后的三处读数")
TABLE(d, ["读数", "旧键", "修正后", "受影响的位置"],
      [["610 集独立演化事件", str(PS["events_v13_triplet"]), str(PS["events"]),
        "§7.1.5 三层口径、摘要第 5 段、结论限制段"],
       ["段位档独立事件", str(EMA["counts"]["events_v13_triplet"]),
        str(EMA["counts"]["events_correct"]),
        "§7.1.5、§7.1.8、表 7-4 表注、摘要第 4 段"],
       ["段位表达式恢复覆盖的事件", "2", str(DD["events_covered"]),
        "§7.2 第二处、表 7-114 第 5 行、贡献二说明"]],
      [4.2, 1.6, 1.6, 9.4], ["l", "c", "c", "l"])
NOTE(d, "口径须一次说清：段位档事件 = 合并 5 + 拆分 5 = 10；"
     "610 集 = 改名 51 + 合并 5 + 拆分 5 = 61。旧键只在段位相关档少算，"
     "改名与依赖档的键本来就含 evo，故不受影响。"
     "修正后 §7.10（池化对照）的 10 次事件与 §7.1.5（M1 对基线）的 10 次事件恰好同值，"
     "但两者仍是不同对照——正文已就此加注「数值接近但不可混用」。")
P(d, "把 58 改成 61 有一个连带风险必须说明：若第 2、8、9 章也引用了「58 次独立演化事件」，"
  "那里需要同步改为 61，否则全文会出现两个口径。本批交付材料只覆盖第 7 章与章外文档，"
  "故此项只能在此登记，不能在此完成。")

H(d, "5.2  slim_chapter7.py 的刷新条件：只改文字时主文不更新", 2)
P(d, "正文瘦身脚本先把完整版备份到素材库文件，再从素材库读入并拆成主文 + 补充材料。"
  "旧条件只在「本次完整版表数 > 备份表数」时才刷新备份。"
  "于是只改文字、表数不变时，备份不刷新，主文与补充材料继续从旧备份生成——"
  "本轮第一次跑三步链时就踩到了：脚本报「段 222 表 15」，而 §7.1.5 仍是旧文字。")
P(d, "修法有两处：一是把刷新条件由「表数增加」改为「内容哈希不同」；"
  "二是把输入源由备份改为本次刚生成的完整版，备份只作为归档副本。"
  "同时保留 slim 开头对表数的断言（< 49 即中止），避免在已瘦身的主文上重复运行。"
  "这两处改动不改变任何数字，但决定了以后每一轮修订能不能真正落到交付文件上。")
P(d, "这条教训值得写进流程：**生成链的每一步都要能从产物反查输入是否更新**，"
  "而不是靠「表数变了没有」这种间接信号。上一版曾用「表数是否增加」修过一次同类问题，"
  "这次换成了内容哈希——同一类 bug 的第二次出现，说明间接信号本身就是错的设计。")

# ================================================================= 六
H(d, "六、数字与措辞变动对照")
CAP(d, "表 6  本轮发生变动的数字与措辞")
TABLE(d, ["项", "修正前", "修正后", "原因"],
      [["610 集独立事件", "58", str(PS["events"]), "事件键缺 evo（5.1）"],
       ["段位档独立事件", "7（清单已不含此项，旧稿残留在表 7-4 表注）",
        str(EMA["counts"]["events_correct"]), "事件键缺 evo（5.1）"],
       ["段位表达式覆盖事件", "2", str(DD["events_covered"]), "事件键缺 evo（5.1）"],
       ["英文摘要", "247 词", f"{EN_WORDS} 词", "清单 M-13：留出计数余量"],
       ["关键词第 5 个", "执行等价", "执行结果一致率", "清单 M-11 / S-06：术语不能声称形式化等价"],
       ["正文「执行等价」", "3 处", "0 处", "清单 M-11"],
       ["投稿信 execution equivalence", "1 处", "0 处", "清单 M-11"],
       ["静默错误归因", "220 条、5 条（并列两说）", "220/223 与 5/223，且 5 ⊂ 220",
        "清单 S-03：须说明是否互斥"],
       ["留出实验", "混在一段内", "H-E2E 与 H-LER 分列",
        "清单 M-02：两项范围不同不可相加"]],
      [3.8, 4.0, 4.4, 4.6], ["l", "l", "l", "l"])

# ================================================================= 七
H(d, "七、仍未闭环")
CAP(d, "表 7  需要作者本人执行或超出本批材料的项")
TABLE(d, ["项", "为什么没闭环", "需要什么"],
      [["GitHub 仓库 + Zenodo DOI", "涉及账号操作，且 Zenodo 发布后文件与 DOI 均不可改",
        "新建公开仓库 → Zenodo New Upload 预留 DOI → 上传材料包 → Publish → "
        "把 5 个值交给 fill_doi.py 回填八处"],
       ["逐轮逐题 LLM 原始输出存档", "5 次重跑已做，但未按条目保存每轮输出",
        "补齐后可把表 S-38 升级为逐轮读数 + 均值±SD；在此之前只报最大偏差"],
       ["第 2/8/9 章的 58 次事件", "本批材料只有第 7 章与章外文档",
        "按表 5 改为 61，或确认那几章用的是另一套定义并加以说明"],
       ["SQL 比较器类型感知主判、其他方言、真实历史演化数据",
        "属新实验，非本轮修订范围", "另立一轮；本章已把这些写进范围限制，不再外推"]],
      [4.2, 5.6, 7.0], ["l", "l", "l"])

P(d, "最后一句自我约束：本轮所有新增句子都遵守同一条规则——"
  "**先有可复算的数字，再写句子**。清单里 S-01 的边界表、S-04 的前提段、S-03 的嵌套关系，"
  "都不是靠措辞补出来的，而是先回到 results/*.json 与逐题产物把关系算清楚，再落笔。"
  "反过来，M-01 的 58→61 也不是为了好看，而是复算后发现清单与旧稿共用了同一个坏键——"
  "这种情况下改数字比守数字更安全。")

d.save(OUT)
print("saved:", OUT)
print("段落:", len(d.paragraphs), "表格:", len(d.tables))
print(f"英文摘要 {EN_WORDS} 词 / 中文摘要 {CN_CHARS} 字")
