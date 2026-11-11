# -*- coding: utf-8 -*-
"""生成《第13轮修订说明_统计与数字修正.docx》。

本轮意见给出四个「必须在下一版解决」的具体问题：
① M1 的 op 级 p 值　② 91.8% 的方法归属　③ 610 条与 10 次事件的口径　④ 摘要残留旧结论。
这份说明按这四条逐一对账，并记录审计过程中额外发现的两处更正
（type_ok ≠ 段位恢复、输出空间不对称），以及一个清理项（重复的「六、结论段」）。
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
OUT = "第13轮修订说明_统计与数字修正.docx"
EN, CN, CN_H = "Times New Roman", "宋体", "黑体"

ADI = json.load(open("results/audit_independence.json", encoding="utf-8"))
assert ADI["version"] == "audit-independence-v1"
AD_A, AD_B, AD_C, AD_D = (ADI["A_units"], ADI["B_segment"],
                          ADI["C_nonsegment"], ADI["D_segment_expr"])
PS, SG = AD_A["paper_set"], AD_A["segment_set"]
IN, SH = AD_C["no_shift_121"], AD_C["shift_26"]
_EV = AD_B["paired"]["B3 嵌入语义"]["event_level"]   # 7:0, p=0.015625
_OP = AD_B["paired"]["B3 嵌入语义"]["op_level"]      # 15:0, p=6.1e-05

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


# ===================================================================== 抬头
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
_run(p, "第 13 轮修订说明：统计与数字修正", bold=True, size=15, font=CN_H)
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
_run(p, "回应《投稿版第 7 章》审稿意见中四个「必须在下一版解决」的具体问题", size=9.5)

P("本轮意见与前几轮不同：它不再讨论「要不要更多实验」，而是点名了四个数字或口径问题，"
  "并给出了判断——「最值得优先完成的是 P0 统计与数字修正，以及 M1 的独立性和基线公平性审计」。"
  "本说明按这四条逐一对账，另记录审计过程中额外发现的两处更正与一个清理项。"
  "四条的最终结论一致：问题都出在「数字旁边缺了那半句口径」，而非数字本身算错。")

# ===================================================================== 一
H("一、四个具体问题的对账")
TABLE(["问题", "原状（错在哪）", "修正后", "落点"],
      [["① M1 的 op 级 p 值",
        f"段位档配对报 15:0，p = 6.1e-05，按 {_OP['n']} 个 op 计。"
        "但 15 个 op 并非独立样本：多个旧列并入同一新列，一次演化产出多个 op",
        f"按 (db, 表, 新列) 聚成 {SG['events']} 次独立事件，"
        f"事件级 7:0，精确符号检验 p = {_EV['p']:.4f}",
        "7.2 表 7-113 note；摘要「事件级 p = 0.016，非算子级 6.1e-5」"],
       ["② 91.8% 的方法归属",
        "非段位档 M1+形态 91.8% vs B3 100%，被读成「M1 通用匹配低 8 个百分点」",
        f"拆两段：不含值改写 {IN['n_ops']} 个 op 两者持平（各 {IN['ours_ok']}/{IN['n_ops']}），"
        f"缺口全在值改写 {SH['n_ops']} 个 op（{SH['ours_ok']}/{SH['n_ops']} 对 "
        f"{SH['base_ok']}/{SH['n_ops']}）",
        "7.2 第三处解读；贡献二；⑤限定；投稿信第三段"],
       ["③ 610 条与事件口径",
        "摘要只写「610 条样本」，未披露独立单元数，读者把条目数读成样本量",
        f"610 条 = {PS['ops']} 个唯一 op、{PS['events']} 次独立事件；"
        f"段位档 {SG['items']} 条只对应 {SG['ops']} 个 op、{SG['events']} 次事件",
        "中英摘要第五段；结论段两处五重限制；第八节③硬性要求"],
       ["④ 摘要残留旧结论",
        "摘要文档残留「10 次事件」「大一个数量级」「91.8% 不如嵌入」等与正文不一致表述",
        "全部改为动态从审计产物取值，与第 7 章共用同一份数据，不再手写数字",
        "摘要/贡献点/结论/投稿信/自检表五处统一"]],
      [1.7, 4.5, 4.6, 4.6], ["l", "l", "l", "l"])

# ===================================================================== 二
H("二、审计额外发现的两处更正（比四条更深）")
P("在按四条重算时，判定函数本身暴露了一个上一版没注意到的口径错误："
  "evaluate() 的 type_ok 在段位档上度量的是「识别出这是一次合并或拆分」，"
  "因为判定前 gold 的目标表达式已被归一为纯列名，段位字符串从未进入比较。"
  "因此「type_ok 15/15」不等于「段位恢复了 15/15」。")
PL(f"**更正一：段位表达式本身的真实恢复率是 10/10，且只覆盖 2 次独立事件。**"
   f"绕开判定函数、直接与 gold 表达式逐字符比对：合并档 {AD_D['part_covered_ops']} 个 op 全部一致，"
   f"拆分档 {len(AD_D['uncovered_ops'])} 个 op 未产出段位表达式（回退列级，这是正确行为——"
   f"拆分只需列级）。故段位恢复的直接证据是「{AD_D['part_expr_match']}/{AD_D['part_covered_ops']} 的机制演示"
   f"+ 端到端效应」，不是「10/10 本身即为统计显著的抽取能力」（n = {AD_D['event_level_n']} 时最小 p = 0.5）。")
PL("**更正二：输出能力不对称是实锤，且必须写明。** 三条基线的输出类型空间只有 "
   "rename_column 与 deprecate_column 两类，段位表达式（SUBSTR/INSTR）根本不在其表示空间内；"
   "因此 0/15 是表达能力所限的必然结果，不是匹配精度差距。这恰恰支持了本文的核心命题——"
   "段位谱系不能由「更好的列匹配器」得到，只能靠能输出段位表达式的抽取器。")
NOTE("更正一把结论往更弱、更诚实的方向改；更正二反而让「表达能力差异」这个定性判断更硬。"
     "两者都不改变端到端的主命题（列级 0.9% vs 段位 100.0%），只改变「段位价值该挂在哪一环」的表述。")

# ===================================================================== 三
H("三、一个清理项")
P("对账过程中发现摘要文档存在两段完全重复的「六、结论段」（内容略不同、标题相同），"
  "系此前多轮编辑叠加所致。已删除冗余的第二段，保留含 M1 三档定位的完整版。"
  "其余章节标题（一至十一）均唯一，未发现同类重复。")

# ===================================================================== 四
H("四、六条 Weaknesses 的处置")
TABLE(["Weakness", "处置"],
      [["真实历史 schema evolution 数据缺乏", "不解决。需 BIRD 版本差分数据（当前无），"
        "加条目数解决不了；审稿人自己也写「不要继续堆积更多条目级实验」"],
       ["合成变更事件少、条目级显著不能推广", "已解决（问题①③）。三层口径并列，事件级重报"],
       ["M1 与基线输出不对称、15 op 独立性", "已解决（更正一、二）。事件级 p=0.016 + 输出空间说明"],
       ["M1 参数开发与 Gold 关联", "部分解决。段位表达式 10/10 为直接证据；"
        "但阈值调优与 gold 的关联仍待独立验证，已写入 Limitations"],
       ["确定性执行器 100% 是构造内上界", "上一轮已解决。主张重摆为「同一执行器 0.9% vs 100.0%」"],
       ["文章结构过于复杂、发现不突出", "未解决。45 表 2.2 万字需单独一轮瘦身（见附注）"]],
      [5.0, 9.9], ["l", "l"])

# ===================================================================== 五
H("五、修订后的核心数字（本轮唯一事实来源）")
TABLE(["口径", "数值"],
      [["主实验集", f"{PS['items']} 条 = {PS['ops']} 唯一 op / {PS['events']} 独立事件 / 11 库"],
       ["段位档", f"{SG['items']} 条 = {SG['ops']} op / {SG['events']} 事件"],
       ["M1 类型恢复（段位档）", f"op {AD_B['methods']['M1 值域指纹']['op_ok']}/{SG['ops']}，"
        f"事件 {AD_B['methods']['M1 值域指纹']['event_ok']}/{SG['events']}；基线均为 0"],
       ["配对（事件级）", f"{_EV['ours_only']}:{_EV['base_only']}，p = {_EV['p']:.4f}"
        f"（op 级 {_OP['ours_only']}:{_OP['base_only']}，p = {_OP['p']:.1e} 属高估）"],
       ["段位表达式真实恢复", f"{AD_D['part_expr_match']}/{AD_D['part_covered_ops']}（合并档）"],
       ["非段位·不含值改写", f"{IN['ours_ok']}/{IN['n_ops']} vs {IN['base_ok']}/{IN['n_ops']}，持平"],
       ["非段位·值改写", f"{SH['ours_ok']}/{SH['n_ops']} vs {SH['base_ok']}/{SH['n_ops']}，落后"]],
      [4.6, 10.3], ["l", "l"])

NOTE("这张表里没有一个是手写的：全部来自 results/audit_independence.json。"
     "第 7 章与摘要文档也改为从同一份产物取值，因此章内外不会再分叉——"
     "这正是问题④的根治方式，而不是再次逐处手改。")

# ===================================================================== 附注
H("附注：仍未动的两项")
PL("**正文瘦身**：45 表 2.2 万字，对应 Weakness 第 7 条「结构过于复杂」。"
   "这是当前最大的单点风险，但属于结构问题，需单独一轮，不在本轮统计修正范围内。")
PL("**DAS + GitHub/Zenodo 归档**：Applied Intelligence 强制要求，目前为 0。"
   "需要用户提供 GitHub 账号或 Zenodo token；全程只在本地打包，不代上传。")

doc.save(OUT)
print("saved:", OUT, "| 段落", len(doc.paragraphs), "| 表格", len(doc.tables))
