# -*- coding: utf-8 -*-
"""生成《第17轮修订说明_A层数据入稿.docx》。

本轮把第 16 轮产出的三项 A 层复核结果写入正文与章外文档：
  ① 留出库 M1 段位表达式直接恢复 14/16 = 87.5%（对己不利，必须主动披露）
  ② 判定器全严规则下 8/8 主对照符号稳定、最大偏移 1.9pt
  ③ to_expr 末段分支缺陷修复：末段含分隔符 20.0% → 100.0%

设计要点（写进说明，供后续复现）：
  - 三张新表追加在正文最末尾，占位号 7-118/119/120，重排后为 50/51/52；
    原有 1–49 表号零漂移，slim_chapter7.py 的 KEEP 15 张表号无需改动；
  - 瘦身脚本改为动态表数，新表自动映射为 S-35/S-36/S-37，主文仍严格 15 表；
  - 章外摘要/结论/限定条款同步，英文摘要压缩后 247 词（期刊 150–250）。

所有数值从 results/*.json 与生成后的 docx 实测读取，不手写统计量。
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

MHO = json.load(open("results/m1_holdout_eval.json", encoding="utf-8"))
JSA = json.load(open("results/judge_strict_audit.json", encoding="utf-8"))
EBV = json.load(open("results/encoding_bounds_v2.json", encoding="utf-8"))
EBL = json.load(open("results/encoding_bounds_v2_legacy.json", encoding="utf-8"))
OP = MHO["op_level"]
JSRE = JSA["reexecution"]


def stat(f):
    d = Document(f)
    caps = sum(1 for p in d.paragraphs
               if re.match(r"^表\s*(?:7|S)-\d+\s\s", p.text.strip()))
    return len(d.paragraphs), len(d.tables), caps


N_MAIN = stat("投稿版_第7章_实验结果_v11.docx")
N_SUPP = stat("补充材料_第7章_v11.docx")
N_FULL = stat("第7章_v11_完整版_49表.docx")
N_FM = stat("摘要_标题_贡献点_结论_v11.docx")

# 英文摘要实测词数（按脚本同款正则，只数 5 个正文段）
_fm = Document("摘要_标题_贡献点_结论_v11.docx")
_txt = [p.text for p in _fm.paragraphs]
_i = next(k for k, t in enumerate(_txt) if t.startswith("三、英文摘要"))
_j = next(k for k, t in enumerate(_txt) if t.startswith("四、中文摘要"))
_segs = [t for t in _txt[_i + 1:_j]
         if t and len(t) > 80 and not t.startswith("实测")]
EN_WORDS = sum(len(re.findall(r"[A-Za-z0-9][A-Za-z0-9%\.\-–']*", s)) for s in _segs)


def make_doc():
    d = Document()
    s = d.sections[0]
    s.page_width, s.page_height = Cm(21.0), Cm(29.7)
    s.left_margin = s.right_margin = Cm(2.3)
    s.top_margin = s.bottom_margin = Cm(2.2)
    st = d.styles["Normal"]
    st.font.name = EN
    st.font.size = Pt(10.5)
    st._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), CN)
    st.paragraph_format.line_spacing = 1.3
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
    for j, part in enumerate(txt.split("**")):
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
    p.paragraph_format.left_indent = Pt(19)
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
    emit(p, txt, 8.6)
    return p


def TABLE(doc, head, rows, widths, aligns, font=8.5):
    t = doc.add_table(rows=1, cols=len(head))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    for i, c in enumerate(t.rows[0].cells):
        c.width = Cm(widths[i])
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        _run(c.paragraphs[0], str(head[i]), bold=True, size=font)
    for row in rows:
        cs = t.add_row().cells
        for i, v in enumerate(row):
            cs[i].width = Cm(widths[i])
            p_ = cs[i].paragraphs[0]
            p_.alignment = {"l": WD_ALIGN_PARAGRAPH.LEFT,
                            "c": WD_ALIGN_PARAGRAPH.CENTER}[aligns[i]]
            p_.paragraph_format.space_after = Pt(1)
            emit(p_, str(v), font)
    return t


D = make_doc()
p = D.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
_run(p, "第 17 轮修订说明：A 层数据入稿", bold=True, size=16, font=CN_H)
p = D.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
_run(p, f"三项独立复核写入正文与章外文档 · 主文仍守 15 表（{TODAY}）", size=9.5)

H(D, "一、这轮做什么、不做什么")
P(D, "第 16 轮已完成三项零 API 成本的本地复核，但结果只躺在《第16轮修订说明》里，"
     "正文、摘要与结论都没有引用。本轮把它们写入稿件，属于**入稿**而非**补实验**："
     "不新增任何模型调用、不改动任何历史数值、不新增主张。")
P(D, "同时处理一个结构约束：第 15 轮刚把主文压到 15 表，直接插表会破坏瘦身后的编号体系。"
     "本轮的解法是把三张新表**追加在正文最末尾**，使原有 1–49 表号完全不漂移，"
     "KEEP 的 15 张主文表号继续有效；瘦身脚本改为动态表数后，"
     f"新表自动落到补充材料 S-35 至 S-{N_SUPP[2]}，主文仍严格 15 表。")

CAP(D, "表 1  文档规模（本轮生成后实测）")
TABLE(D, ["文件", "段落", "表格", "表题", "说明"],
      [["投稿版_第7章_实验结果_v11.docx", str(N_MAIN[0]), str(N_MAIN[1]), str(N_MAIN[2]),
        "主文，7-1 至 7-15 连续"],
       ["补充材料_第7章_v11.docx", str(N_SUPP[0]), str(N_SUPP[1]), str(N_SUPP[2]),
        f"S-1 至 S-{N_SUPP[2]} 连续（含 1 张对照表）"],
       ["第7章_v11_完整版_49表.docx", str(N_FULL[0]), str(N_FULL[1]), str(N_FULL[2]),
        "素材库，7-1 至 7-52 连续；文件名沿用历史约定"],
       ["摘要_标题_贡献点_结论_v11.docx", str(N_FM[0]), str(N_FM[1]), str(N_FM[2]),
        f"英文摘要 {EN_WORDS} 词（期刊 150–250）"]],
      [6.6, 1.5, 1.5, 1.5, 5.2], ["l", "c", "c", "c", "l"], font=8.3)

H(D, "二、三项结果的入稿位置与措辞")
CAP(D, "表 2  三项复核：数据、正文位置、补充材料表、章外同步")
TABLE(D, ["复核项", "关键读数", "正文位置", "补充表", "章外同步"],
      [["留出库 M1 段位表达式",
        f"{OP['matched']}/{OP['n_units']} op（{OP['rate_pct']:.1f}%，"
        f"Wilson [{OP['wilson95_pct'][0]:.1f}, {OP['wilson95_pct'][1]:.1f}]）；"
        f"事件 {MHO['event_level']['rate_pct']:.1f}%、库 {MHO['db_level']['rate_pct']:.1f}%",
        "7.7.4 末新增段；7.7.3 边界①", "S-36", "中英摘要、结论、限定⑤"],
       ["判定器全严规则",
        f"{JSA['sign_stable_count']} 主对照符号不变，最大偏移 "
        f"{JSA['max_abs_delta_pt']:.1f}pt；{JSRE['n']:,} 条重执行 "
        f"{JSRE['ex_current_pct']:.1f}% → {JSRE['ex_strict_pct']:.1f}%",
        "7.1.3 改写；7.7.3 边界⑤", "S-35", "中文摘要、结论"],
       ["编码压力测试 v2 / to_expr 修复",
        f"末段含分隔符 {EBL['by_case']['sep_in_last']:.1f}% → "
        f"{EBV['by_case']['sep_in_last']:.1f}%；nseg ≥ 4 差异格 "
        f"{len(EBL['nseg_ge4_changes'])} → {len(EBV['nseg_ge4_changes'])}",
        "7.7.3 边界④改写；7.7.5 新增段", "S-37", "中文摘要、限定①"]],
      [3.2, 4.6, 3.2, 1.4, 3.0], ["l", "l", "l", "c", "l"], font=8.0)

P(D, "三项中有**两项对己不利或主动收紧**，这是本轮最需要保留的性质：")
PL(D, f"· 留出库 M1 从 dev 的 10/10 掉到 {OP['rate_pct']:.1f}%（事件级 "
      f"{MHO['event_level']['rate_pct']:.1f}%）。两个失败都是 no_infer，"
      "不是给出错误表达式。这把「dev ↔ 未见库差距」从被动承认变成带数值的主动披露；"
      f"Wilson 上限 {OP['wilson95_pct'][1]:.1f}% 已写入正文，"
      "明确不能读作通用恢复能力。")
PL(D, "· 判定器复核结论写成「符号稳定性」，不是「三类风险已消除」。"
      "现行主判仍为 r0，绝对 EX 一律按 r0 双报，且仍限定为"
      "「在所评估 SQLite 实例上的执行结果一致率」。")
PL(D, "· 编码修复只触及 nseg ≥ 3 的末段分支；2,384 个真实 gold 段位表达式全是两段 "
      "flat 形态、零嵌套，故 v1 三段探针、留出库 M1、B4 公平基线在修正前后逐字不变。"
      "NULL 分量与非末段含分隔符两类仍 0%，属分隔符拼接的信息论限制，"
      "限定语不因实现修复而放宽。")

H(D, "三、为什么主文仍是 15 表")
P(D, "第 15 轮瘦身时 `slim_chapter7.py` 的 KEEP 按**原最终表号**写死了 15 张"
     "（1、3、4、16、18、19、20、22、25、28、33、34、39、42、48）。"
     "如果在正文中间插表，其后所有表号会顺延，KEEP 就会指向错误的表。"
     "本轮把三张新表放在 7.12 节（全章末尾），占位号 7-118/119/120，"
     "重排后恰为 50/51/52——前 49 张表号零漂移，KEEP 无需改动。")
CAP(D, "表 3  瘦身链的三处改造")
TABLE(D, ["位置", "原实现", "本轮改为"],
      [["素材库备份刷新", "仅在备份不存在时 copy2，残留旧表数",
        "按表题数比较 SRC 与 BAK，完整版表数变大则自动刷新；重复运行幂等"],
       ["表数断言", "assert len(ALL) == 49 / all(1<=n<=49)",
        "动态 N_ALL、断言表号连续 1..N_ALL、KEEP 长度 15"],
       ["迁移与校验", "MOVED 固定 34、audit expect 34、引用上限 34",
        f"动态 N_MOVED = {N_SUPP[2]}，S-1..S-{N_SUPP[2]}，引用上限同步"]],
      [3.4, 5.4, 6.6], ["l", "l", "l"], font=8.2)
P(D, "章外 `_T()` 仍是双文档反查，本轮把两份文档的表数表述改为从实际 docx 读取"
     "（_count_caps），不再写死 S-34 / 49 表——否则章外会再次先于章内分叉。")

H(D, "四、英文摘要的词数处理")
P(D, f"补入 A 层披露后英文摘要一度涨到 262 词，超出 Applied Intelligence 的 250 词上限。"
     f"处理方式不是删读数，而是把「全严判定 {JSA['sign_stable_count']} 稳定」留在正文与结论，"
     "摘要只保留信息量最高、且对己不利的留出库 M1 回落，并对末段做无损压缩"
     f"（The 610 items rest... / On 7 held-out databases...），最终 {EN_WORDS} 词，"
     "落在 150–250 区间内且留有余量。中文摘要与结论段则完整保留三项读数。")

H(D, "五、校验结果")
PL(D, f"· 三份第 7 章文档表号连续：主文 7-1..7-{N_MAIN[2]}、"
      f"补充材料 S-1..S-{N_SUPP[2]}、完整版 7-1..7-{N_FULL[2]}。")
PL(D, "· 主文对新增三表的引用全部重定向为 S-35 / S-36 / S-37；"
      "占位号 7-118/119/120 零残留。")
PL(D, "· 章外自检表 14 行出处逐行命中，无空值、无「未找到」。")
PL(D, f"· 英文摘要 {EN_WORDS} 词合规；四份文档 ** 残留 0、悬空表号 0。")
PL(D, "· 复现链顺序不变：make_chapter7_v11.py → slim_chapter7.py → "
      "make_front_matter_v11.py；只跑第 1 步会得到 52 表完整版，只跑第 3 步会报错。")

H(D, "六、仍未闭环（与第 16 轮相同，不因本轮改变）")
PL(D, "· 真实历史 schema 演化数据；第二位标注者 / Cohen's kappa"
      "（gold 为程序构造，κ 无对应物，只能用等效独立复核 + 编码手册回应）；"
      "LLM 核心条件重复运行 3 次（需 API 预算）；SQL 比较器类型感知主判"
      "（本轮只做全严规则下的符号稳定性复核，主判仍 r0）；其他 SQL 方言验证。")
PL(D, "· GitHub 仓库与 Zenodo DOI 仍需作者本人创建；归档前 DAS、投稿信、"
      "表 9-2、表 9-3 保持「将归档」口径。")
NOTE(D, "本轮不新增实验、不新增主张、不改动任何历史数值。"
        "三项结果的价值在于把审稿人点名的「dev↔未见库差距」「判定器过宽」"
        "「三段以上编码覆盖」从定性承认升级为可复核的量化界定，"
        "其中两项结论对方法不利，保留它们比删掉更安全。")

D.save("第17轮修订说明_A层数据入稿.docx")
print("saved: 第17轮修订说明_A层数据入稿.docx")
print("段落", len(D.paragraphs), "表格", len(D.tables), "| 英文摘要", EN_WORDS, "词")
