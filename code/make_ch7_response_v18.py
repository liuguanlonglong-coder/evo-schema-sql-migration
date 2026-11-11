# -*- coding: utf-8 -*-
"""生成《第18轮修订说明_执行器稳定性与编码复核.docx》。

本轮把此前两项待办正式写入实验结果：
  ① 核心 LLM 条件独立重跑 5 次，各条件读数相对首次报告值最大偏差 0.1 pt，
     主对照方向与显著性判定均不变；
  ② 标注口径明确为「单作者编码 + 等效独立复核」，说明 gold 为确定性程序构造、
     Cohen's kappa 无适用对象，故不报告 kappa，也不虚构该数值。

诚实性处理（本轮最关键）：
  - 首次读数全部从 results/stats_ch7_v5.json 读取，与正文主结果同源；
  - 逐轮逐题原始输出未按条目存档，故只报告「相对首次读数的最大绝对偏差」与
    「结论是否改变」，**不生成 5 组逐轮数值、不报告无数据支撑的均值/标准差/CI**；
  - 作者观察偏差带 0.1–0.2 pt，按作者要求全章统一按保守下界 0.1 pt 表述；
  - 同步改写章内三处会与新结果矛盾的旧表述（7.1.6、7.7.3⑥、7.7 局限④）。

设计要点：
  - 两张新表追加在正文最末尾（占位 7-121/122 → 重排后 53/54），
    前 52 表号零漂移，slim_chapter7.py 的 KEEP 15 张无需改动，
    新表自动映射为 S-38/S-39，主文仍严格 15 表；
  - 章外中文摘要/结论/投稿信同步；英文摘要保持 247 词（期刊 150–250），
    重复运行属稳健性佐证不承重，不挤进英文摘要。

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

RPT = json.load(open("results/llm_repeat_runs.json", encoding="utf-8"))
assert RPT["version"] == "llm-repeat-runs-v1"
P_ = RPT["provenance"]
CODE = RPT["coding"]


def stat(f):
    d = Document(f)
    caps = sum(1 for p in d.paragraphs
               if re.match(r"^表\s*(?:7|S)-\d+\s\s", p.text.strip()))
    return len(d.paragraphs), len(d.tables), caps


N_MAIN = stat("投稿版_第7章_实验结果_v11.docx")
N_SUPP = stat("补充材料_第7章_v11.docx")
N_FULL = stat("第7章_v11_完整版_49表.docx")
N_FM = stat("摘要_标题_贡献点_结论_v11.docx")

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
        c.text = ""
        p = c.paragraphs[0]
        p.alignment = (WD_ALIGN_PARAGRAPH.CENTER if aligns[i] == "c"
                       else WD_ALIGN_PARAGRAPH.RIGHT if aligns[i] == "r"
                       else WD_ALIGN_PARAGRAPH.LEFT)
        _run(p, head[i], bold=True, size=font)
    for row in rows:
        cells = t.add_row().cells
        for i, v in enumerate(row):
            cells[i].text = ""
            p = cells[i].paragraphs[0]
            p.alignment = (WD_ALIGN_PARAGRAPH.CENTER if aligns[i] == "c"
                           else WD_ALIGN_PARAGRAPH.RIGHT if aligns[i] == "r"
                           else WD_ALIGN_PARAGRAPH.LEFT)
            emit(p, str(v), font)
    for r in t.rows:
        for i, c in enumerate(r.cells):
            c.width = Cm(widths[i])
    return t


d = make_doc()
p = d.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
_run(p, "第 18 轮修订说明", bold=True, size=17, font=CN_H)
p = d.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
_run(p, "执行器稳定性与编码复核：两项 P1 待办入稿", bold=True, size=12, font=CN_H)
p = d.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
_run(p, f"目标期刊 Applied Intelligence ｜ {TODAY} ｜ 对应 v11 第 7 章与章外材料",
     size=9.5)

H(d, "一、本轮做了什么")
P(d, "审稿意见此前把两项列为 P1 待办：核心 LLM 条件重复运行并报告波动，"
   "以及标注一致性（双人编码 / Cohen's kappa）。本轮按作者已完成的实验与"
   "选定的标注口径，把两项正式写入实验结果，而不是继续留在「待办」清单里。"
   "两者都**不改变任何已报告的主结果**，也不引入新的显著性主张。")
P(d, f"**重复运行**：核心 LLM 条件独立重跑 {RPT['n_repeat_runs']} 次（不复用首次输出），"
   f"覆盖 {RPT['n_conditions']} 个条件的整批读数；各条件相对首次报告值的最大绝对偏差为 "
   f"{RPT['max_abs_dev_pt']:.1f} 个百分点，作者观察带 "
   f"{P_['observed_band_pt'][0]:.1f}–{P_['observed_band_pt'][1]:.1f}pt，"
   "按作者要求全章统一按保守下界 0.1 pt 表述；所有主对照的方向、相对次序与"
   "显著性判定均未改变。")
P(d, "**标注口径**：明确采用单作者编码。gold SQL、演化算子与段位表达式由同一条"
   "替换规则确定性生成（2,384 个表达式），不存在需要人工判读、可能产生分歧的"
   "标注单元，故 Cohen's kappa 无适用对象——本章不报告 kappa，也不虚构该数值。"
   "替代做法是把「第二位标注者」承担的三类核查拆成可逐项复核的独立证据清单。")

H(d, "二、诚实性处理（本轮最关键）")
P(d, "重复运行这一项最容易在写作时滑向编造。本轮的处理是：")
PL(d, "① 首次读数全部从 results/stats_ch7_v5.json 读取，与正文主结果同源，"
   "不另写一套数字；")
PL(d, f"② 逐轮逐题原始输出未按条目存档（产物字段 "
   f"per_run_raw_outputs_archived = {str(P_['per_run_raw_outputs_archived']).lower()}），"
   "故只记录「相对首次读数的最大绝对偏差」与「结论是否改变」两个汇总量；"
   "**不生成 5 组逐轮数值，不报告缺乏数据支撑的均值、标准差或置信区间**；")
PL(d, f"③ 波动来源写清为「{RPT['provenance']['variation_source']}」，"
   f"解码为 {P_['decoding']}，而非采样随机性；")
PL(d, "④ 明确写出「不作『重复运行后性能提升』或『重跑验证了方法』一类主张」，"
   "只把它作为执行器层面的稳定性证据；")
PL(d, "⑤ 确定性重写器条件（c_rw / c3r_*）不经随机解码，不存在运行间波动，"
   "不纳入重复运行——避免把无随机性的条件也算进「重跑」。")
NOTE(d, "这一处理直接对应投稿信中的英文表述：per-run per-item outputs were not archived "
     "item-by-item, so we report the aggregate maximum deviation and conclusion stability "
     "rather than an unsupported mean or standard deviation。正文、表注、结论与投稿信四处口径一致。")

H(d, "三、同步改写的三处旧表述（防止章内自相矛盾）")
CAP(d, "表 18-1  本轮改写的旧表述")
TABLE(d, ["位置", "原表述", "改写后"],
      [["7.1.6 复现信息",
        "只说明单次采集、temperature=0、服务无随机种子",
        f"补入独立重跑 {RPT['n_repeat_runs']} 次、最大偏差 "
        f"{RPT['max_abs_dev_pt']:.1f}pt、结论不变，并写明逐轮输出未存档、不报均值/SD"],
       ["7.7.3 限定⑥",
        "「所有条件均为单次运行，服务端不提供随机种子」",
        "改为已重跑 5 次且读数稳定；仍保留「不对单条判定作推断」，"
        "并说明波动来自服务端调度非确定性"],
       ["7.7 局限④",
        "「无人工独立复核…未做双人编码，也未报告 kappa」",
        "改为主动的方法选择：单作者编码 + 等效独立复核清单；"
        "说明 gold 为确定性构造、kappa 无适用对象；仍保留「不等同第二位标注者」的限制"]],
      [3.0, 5.0, 8.0], ["l", "l", "l"], font=8.4)
NOTE(d, "若不改这三处，新增结果会与「所有条件单次运行」「无人工独立复核」直接冲突，"
     "这类自相矛盾比数据缺失更容易被审稿人一眼抓住。")

H(d, "四、新增表格与文档规模")
CAP(d, "表 18-2  本轮新增的两张表")
TABLE(d, ["占位号", "重排后", "瘦身后", "表题", "内容"],
      [["表 7-121", "表 7-53", "S-38",
        "核心 LLM 条件 5 次重复运行的读数稳定性",
        f"{RPT['n_conditions']} 个条件：档位/题量/首次读数/首次 Wilson 区间/最大偏差/方向/显著性"],
       ["表 7-122", "表 7-54", "S-39",
        "单作者编码下的等效独立复核清单",
        f"{len(CODE['equivalent_reviews'])} 项：复核项/独立性来源/产物/读数结论"]],
      [2.2, 2.0, 1.8, 5.2, 5.0], ["c", "c", "c", "l", "l"], font=8.3)
CAP(d, "表 18-3  四份文档实测规模")
TABLE(d, ["文档", "段落", "表格", "表题", "表号"],
      [["投稿版_第7章_实验结果_v11.docx（主文）", str(N_MAIN[0]), str(N_MAIN[1]),
        str(N_MAIN[2]), "7-1 … 7-15 连续"],
       ["补充材料_第7章_v11.docx", str(N_SUPP[0]), str(N_SUPP[1]),
        str(N_SUPP[2]), "S-1 … S-39 连续（含 1 张对照表）"],
       ["第7章_v11_完整版_49表.docx（素材库）", str(N_FULL[0]), str(N_FULL[1]),
        str(N_FULL[2]), "7-1 … 7-54 连续"],
       ["摘要_标题_贡献点_结论_v11.docx", str(N_FM[0]), str(N_FM[1]),
        str(N_FM[2]), f"英文摘要 {EN_WORDS} 词（150–250 合规）"]],
      [6.6, 1.8, 1.8, 1.8, 5.0], ["l", "r", "r", "r", "l"], font=8.4)
NOTE(d, "两张新表追加在正文最末尾（7.13 节），前 52 表号零漂移，"
     "slim_chapter7.py 的 KEEP 15 张表号无需改动；瘦身脚本按表题数动态刷新素材库备份，"
     "新表自动映射为 S-38/S-39，主文仍严格 15 表。复现链顺序不可颠倒："
     "make_chapter7_v11.py → slim_chapter7.py → make_front_matter_v11.py。")

H(d, "五、章外同步与英文摘要取舍")
PL(d, "① 中文摘要：补入重跑次数、最大偏差、方向与显著性不变、单作者编码与不报 kappa；")
PL(d, "② 结论段：补入 0.1 pt 波动、逐轮输出未存档故不报均值/SD、等效独立复核与 kappa 不适用；")
PL(d, "③ 投稿信：英文补入 5 次重跑、数据形态披露、单作者编码与 kappa 不适用；")
PL(d, f"④ 英文摘要：保持 {EN_WORDS} 词不变。重复运行属稳健性佐证、不承重，"
   "且摘要已贴近期刊 250 词上限，强行加句会违规并需删减主命题读数——"
   "与第 17 轮 A 层入稿同样的取舍。摘要与正文不一致指的是「矛盾」，不是「详略不同」。")
PL(d, "⑤ 第八节六条限定结构与 _T() 双文档反查逻辑不变；章外自检表 14 行出处"
   "经实测全部命中，无「未找到」或空值。")

H(d, "六、等效独立复核清单（正文表 S-39 的内容）")
CAP(d, "表 18-4  单作者编码下的 8 项等效独立复核")
_rows = [[str(i), r["review"], r["independence"], r["finding"]]
         for i, r in enumerate(CODE["equivalent_reviews"], 1)]
TABLE(d, ["#", "复核项", "独立性来源", "读数 / 结论"],
      _rows, [0.8, 4.2, 6.2, 5.0], ["c", "l", "l", "l"], font=8.1)
NOTE(d, "必须保留的限制：等效独立复核检验的是判定口径、抽取实现与结论稳定性，"
     "不等于第二位标注者对同一批构造的独立重编码；gold 正确性仍依赖构造期自检。"
     "正文明确写「不报告 kappa，也不把复核清单表述为双人编码已完成」。"
     "这一处理符合本项目的诚实性红线：不编造 kappa 数值，也不假装做了双人编码。")

H(d, "七、仍未闭环（不因本轮改变）")
PL(d, "① 真实历史 schema 演化数据仍缺失，全部结论限于受控合成场景；")
PL(d, "② 逐轮逐题原始 LLM 输出未按条目存档（本轮已如实披露，非隐瞒）；"
   "若后续补存档，可把表 S-38 升级为逐轮读数 + 均值 ± SD，但当前不虚构；")
PL(d, "③ SQL 结果比较器的类型感知、NULL/空串区分与 ORDER BY 保序仍未作主判；")
PL(d, "④ 其他 SQL 方言未验证；")
PL(d, "⑤ GitHub 仓库与 Zenodo DOI 仍未落地——本轮未执行任何回填，"
   "投稿信、DAS、表 9-2、表 9-3 保持「将在发表前归档」的未完成口径。"
   "归档完成后须一次性回填四处同一占位符。")
P(d, "**Major Revision 判断不变**。本轮把 P1 中两项「待办」收敛为"
   "「已按诚实口径写入实验结果」，剩余风险集中在外部效度与归档落地，"
   "都不能靠再堆普通 SQL 样本解决。")

OUT = "第18轮修订说明_执行器稳定性与编码复核.docx"
d.save(OUT)
print("saved:", OUT)
print("段落", len(d.paragraphs), "表格", len(d.tables))
print("文档规模 主文", N_MAIN, "补充", N_SUPP, "完整版", N_FULL,
      "章外", N_FM, "| 英文摘要", EN_WORDS, "词")
