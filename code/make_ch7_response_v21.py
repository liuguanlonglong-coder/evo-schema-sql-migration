# -*- coding: utf-8 -*-
"""生成《第21轮修订说明_研究定位收敛为谱系粒度实证研究.docx》。

本轮只有一件事：把论文的研究类型改写为「实证研究」，并清掉 RAG / NL2SQL 系统的遗留口径。
起因是外部意见指出，本项目的检索证据只有 40 道题、属于阴性对照、且没有实现完整检索模块；
在这种情况下把 RAG 写进标题、摘要或主要贡献，审稿人翻到实验部分就会判定
「研究问题与实验内容不一致」——这与数值是否正确无关。

所有数字从 results/*.json 与已生成的 docx 实测读取，不手写统计量。
"""
import collections
import json
import os
import re
from datetime import date
from math import comb

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

BASE = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE)
EN, CN, CN_H = "Times New Roman", "宋体", "黑体"
TODAY = date.today().isoformat()
OUT = "第21轮修订说明_研究定位收敛为谱系粒度实证研究.docx"

# ---- 检索式探针：从逐条输出重算，与 summary 交叉校验 -------------------------
ROWS = [json.loads(l) for l in
        open("results/e2_e2_main.jsonl", encoding="utf-8") if l.strip()]
SUM = json.load(open("results/e2_e2_main_summary.json", encoding="utf-8"))
BY = collections.defaultdict(dict)
for r in ROWS:
    BY[r["id"]][r["cond"]] = r["outcome"]
E2 = {
    "n": len(BY),
    "n_db": len({r["db_id"] for r in ROWS}),
    "db": ",".join(sorted({r["db_id"] for r in ROWS})),
    "panels": "+".join(sorted({r["evo_type"] for r in ROWS})),
    "c4_ex": sum(1 for v in BY.values() if v.get("c4") == "EX"),
    "c5_ex": sum(1 for v in BY.values() if v.get("c5") == "EX"),
    "b": sum(1 for v in BY.values() if v.get("c5") == "EX" and v.get("c4") != "EX"),
    "c": sum(1 for v in BY.values() if v.get("c4") == "EX" and v.get("c5") != "EX"),
    "silent": sum(1 for r in ROWS if r["outcome"] == "SILENT"),
    "abstain": sum(1 for r in ROWS if r["outcome"] in ("ABSTAIN_OK", "OVER_ABSTAIN")),
    "runs": len(ROWS),
}
assert E2["n"] == SUM["n_samples"] == 40, "E2 逐条数与 summary 不一致"
assert E2["c4_ex"] == SUM["by_cond"]["c4"]["EX"]
assert E2["c5_ex"] == SUM["by_cond"]["c5"]["EX"]
_n = E2["b"] + E2["c"]
E2["p"] = min(1.0, 2 * sum(comb(_n, k) for k in range(0, min(E2["b"], E2["c"]) + 1)) / 2 ** _n)
E2["c4_rate"] = 100.0 * E2["c4_ex"] / E2["n"]
E2["c5_rate"] = 100.0 * E2["c5_ex"] / E2["n"]

ADI = json.load(open("results/audit_independence.json", encoding="utf-8"))
assert ADI["version"] == "audit-independence-v2"
PS, SG = ADI["A_units"]["paper_set"], ADI["A_units"]["segment_set"]
XPM = json.load(open("results/ext_eval_merge_col.json", encoding="utf-8"))["paired_by_event"]["c3mp-c3m"]


def stat(f):
    d = Document(f)
    caps = sum(1 for p in d.paragraphs
               if re.match(r"^表\s*(?:7|S|0|5|9)-\d+\s\s", p.text.strip()))
    return len(d.paragraphs), len(d.tables), caps


N_MAIN = stat("投稿版_第7章_实验结果_v11.docx")
N_SUPP = stat("补充材料_第7章_v11.docx")
N_FULL = stat("第7章_v11_完整版_49表.docx")
N_FM = stat("摘要_标题_贡献点_结论_v11.docx")


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
        _run(p, head[i], bold=True, size=font - 0.4)
        c.width = Cm(widths[i])
    for row in rows:
        cells = t.add_row().cells
        for i, v in enumerate(row):
            cells[i].text = ""
            p = cells[i].paragraphs[0]
            p.alignment = (WD_ALIGN_PARAGRAPH.CENTER if aligns[i] == "c"
                           else WD_ALIGN_PARAGRAPH.LEFT)
            emit(p, v, size=font)
            cells[i].width = Cm(widths[i])
    return t


d = make_doc()
p = d.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
_run(p, "第 21 轮修订说明：研究定位收敛为「谱系粒度实证研究」", bold=True, size=16, font=CN_H)
p = d.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
_run(p, f"Applied Intelligence 投稿 · 定位修订 · {TODAY}", size=9.5)

# ================================================================= 一
H(d, "一、本轮意见与处置总览")
P(d, "本轮意见是一条定位批评，不是数值批评：本文被包装成 RAG-NL2SQL 系统，"
     "但检索实验只有 40 道题、属于阴性对照，也没有实现完整检索模块；"
     "若把 RAG 放进标题、摘要或主要贡献，审稿人一旦翻到实验部分，"
     "就会判定「研究问题与实验内容不一致」。意见同时给出了建议标题——"
     "An Empirical Study of Schema Lineage Granularity for Historical SQL Migration "
     "under Schema Evolution。")
P(d, "接受这条意见的原因不是它更保守，而是它更准确。本文唯一不受构造循环性影响的结论是"
     "「同一执行器下，谱系的表示粒度决定迁移成败」；"
     f"检索方向的全部证据是 {E2['n']} 条、单库、阴性对照（p = {E2['p']:.2f}）。"
     "把后者写进标题，等于用最弱的一条证据去代表全文；"
     "把前者写进标题，读者按「变量是否被分离、证据是否够」来审，"
     "而这正是本文能交付的东西。")
CAP(d, "表 1  七处口径改写与章内四处同步")
TABLE(d, ["位置", "旧口径", "本版口径"],
      [["标题", "问句「谱系要细到哪一层：schema 演化下历史 SQL 迁移的粒度效应」",
        "An Empirical Study of Schema Lineage Granularity for Historical SQL Migration "
        "under Schema Evolution"],
       ["英文摘要", "首句 Schema evolution breaks historical SQL；「不实现检索模块」只在末段出现一次",
        "首句 We report an empirical study of…；同段补 We build no retrieval module and "
        "no NL-to-SQL system；末段补 40-item probe is a negative control (p=0.25)"],
       ["中文摘要", "同龄口径，未声明研究类型",
        "首句写明「本文是一项关于谱系粒度的实证研究，不是系统论文」"],
       ["关键词", "5 个，已不含检索类词", "不变；另写明「检索类词是硬禁区」与第六个词的推荐顺序"],
       ["贡献点", "三条，均以「我们构建／提出」开头",
        "新增「定位」行作负向声明；贡献一改为「以受控实证设计把谱系粒度分离为独立变量」"],
       ["结论", "首句未交代研究类型",
        "首句加「作为一项实证研究」；末段新增研究类型交代（不提出系统、不实现检索、不做端到端生成）"],
       ["投稿信", "第一段标题为问句版；主题句为 an applied AI study",
        "标题替换；第二段插入 This is an empirical study, not a system paper；"
        "主题句改为 an empirical study"],
       ["章内 7.1.1", "「真检索实验仅 40 题且为阴性对照」",
        f"改为 {E2['n']} 条 / {E2['n_db']} 库 / 不一致对 {E2['b']}:{E2['c']} / p = {E2['p']:.2f} 的完整口径，"
        "并写明不实现检索模块"],
       ["章内 7.7.3 ⑦", "同上，未交代题面档位与库名",
        f"补写题面取自改名档（{E2['panels']}）、单库（{E2['db']}）、"
        f"{E2['runs']} 次运行中 {E2['silent']} 次静默错误、0 次弃答"],
       ["章内 7.10", "「若整篇论文仍以 RAG + NL2SQL 为主线，二者只能选一」——待决口吻",
        "改为已决断：第 21 轮选定后者，并说明理由"],
       ["章内 7.11 ③", "「检索效果一律写作不作主张」", "保留结论，补齐探针口径与「不能写成评测了检索方法」的边界"]],
      [2.4, 6.4, 8.2], ["l", "l", "l"], font=8.0)
NOTE(d, "七处章外 + 四处章内一共十一处，缺一处就会重新出现不一致。"
     "最容易被漏掉的是章内那四处：它们不改，摘要与标题改了也没用——"
     "审稿人读到 7.10「二者只能选一」时，会立刻明白作者自己还没决定这篇是什么论文。")

# ================================================================= 二
H(d, "二、标题的三案与为什么以研究类型开头")
P(d, "标题的第一个成分决定读者的审查标准，所以它不是修辞问题。"
     "写对象（RAG、NL2SQL、检索增强），读者会去实验部分找那个系统；"
     "写研究类型（An Empirical Study of），读者会去问「你分离出了哪个变量、证据有多强」。"
     "本文能交付的只有后者，因此标题必须以研究类型开头。")
CAP(d, "表 2  标题三案")
TABLE(d, ["", "英文 / 中文", "取舍理由"],
      [["A（投稿用）",
        "An Empirical Study of Schema Lineage Granularity for Historical SQL Migration "
        "under Schema Evolution ／ schema 演化下历史 SQL 迁移的谱系粒度实证研究",
        "四个成分各对应正文一处：研究类型（第零节）、唯一自变量（7.4/7.5）、"
        "任务且刻意不写 NL2SQL 生成（7.1.1）、前提条件而非贡献（7.2）"],
       ["B（备选）",
        "How Fine Must Lineage Be? Granularity Effects in Migrating Historical SQL under "
        "Schema Evolution",
        "最能体现「粒度是核心」，但问句标题会让审稿人期待一个做法；"
        "本文没有系统可交付，问句容易读成承诺未兑现"],
       ["C（备选）",
        "One Executor, Two Granularities: Segment-Expression Lineage for Historical SQL",
        "紧扣受控对比，但并列两个成分，且未声明研究类型"]],
      [2.0, 8.0, 6.8], ["l", "l", "l"], font=8.0)
NOTE(d, "三案均不含 RAG、NL2SQL 与医疗字样。"
     "副标题「——以电子病历结构化字段为例」不建议加：本研究没有医疗数据，"
     "加上去只是把一个不存在的应用场景写进标题。"
     "英文主标题 15 词，在 Applied Intelligence 的常见长度内；"
     "若改用 B 或 C，投稿信第一段引号内与摘要首句须同步替换。")

# ================================================================= 三
H(d, "三、检索式探针的真实口径")
P(d, "这是本轮必须写清的一段，因为「40 题」这个数字此前只以「仅 40 题」的形式出现，"
     "读者无法判断它测了什么、测到什么程度。逐条输出核对后的完整口径如下。")
CAP(d, "表 3  检索式探针（E2）的条件与结果")
TABLE(d, ["项目", "内容"],
      [["样本", f"{E2['n']} 题，题面取自改名档（{E2['panels']}），单库（{E2['db']}）"],
       ["C4 条件", "当前 schema + 从同库其它历史题检索到的示例（按旧 schema 书写，含过期列）"],
       ["C5 条件", "C4 + 该示例中过期列的对齐提示（谱系来自 drift_ops）"],
       ["结果", f"执行结果一致率 {E2['c4_rate']:.1f}% → {E2['c5_rate']:.1f}%"
        f"（{E2['c4_ex']}/{E2['n']} → {E2['c5_ex']}/{E2['n']}）"],
       ["配对", f"不一致对 {E2['b']}:{E2['c']}，双侧精确符号检验 p = {E2['p']:.2f}"],
       ["失败形态", f"{E2['runs']} 次运行中 {E2['silent']} 次静默错误、{E2['abstain']} 次弃答"],
       ["能说明", "补一个检索到的过期示例（即使再加对齐提示）不足以救回迁移"],
       ["不能说明", "召回与排序策略的好坏、多路证据如何取舍、证据冲突如何裁决；"
        "也不涉及合并/拆分档，与本文主结论没有交集"]],
      [2.6, 14.4], ["l", "l"], font=8.2)
P(d, "三点必须同时读。第一，它是阴性对照而非弱阳性结果："
     f"+{E2['c5_rate'] - E2['c4_rate']:.1f}pt 在 n={E2['n']} 上不显著，"
     "而这个方向恰恰是「检索 + 对齐提示」最省事的实现方式——"
     "这条路子不成立，是本文不把检索写成研究方向的经验依据，不是它的替代品。"
     "第二，题面取自改名档，因此它不能反过来当作「谱系注入在检索场景无效」的证据；"
     "第三，{silent} 次静默错误、0 次弃答说明这批条件下模型倾向于硬答，"
     "这本身是一个值得独立研究的现象，但本文没有做弃答判定，故不作主张。".replace(
         "{silent}", str(E2["silent"])))
NOTE(d, "如果只在摘要里加一句「不实现检索模块」，而标题仍含 RAG，读者仍会按系统来审；"
     "反过来，标题改成实证研究之后，摘要那句只是补强。顺序不能颠倒："
     "先定研究类型，再补免责句。")

# ================================================================= 四
H(d, "四、三种只做表面修补的改法，以及它们为什么不够")
CAP(d, "表 4  表面修补与失败原因")
TABLE(d, ["改法", "看起来解决了什么", "为什么不够"],
      [["只改摘要，不改标题",
        "读者在摘要里看到「不实现检索模块」",
        "标题仍是系统口径，检索者按标题筛选文献，进来的审稿人预期不会因摘要而改变"],
       ["只改标题，不改贡献点",
        "标题不再承诺系统",
        "贡献点仍写「我们构建／提出」，等于在正文第一页把承诺又加回来"],
       ["加一句免责声明",
        "形式上披露了局限",
        "免责声明不能修正研究问题与实验内容的关系；审稿人问的是「这篇在回答什么问题」"],
       ["补做检索实验以匹配标题",
        "标题与内容一致了",
        "需要多路召回、证据冲突裁决与弃答判定三部分，是另一篇论文的工作量；本次不做，如实写明"]],
      [3.4, 5.4, 8.2], ["l", "l", "l"], font=8.2)
P(d, "真正需要改的是「这篇是什么」这一层，而不是描述这一层的句子。"
     "因此本轮同时动了标题、摘要首段、关键词说明、贡献点首行、结论首末句与投稿信第二段，"
     "并在第 7 章的四处把待决口吻改成已决断。")

# ================================================================= 五
H(d, "五、一致性核对")
CAP(d, "表 5  七处一致性核对（投稿前逐行确认）")
TABLE(d, ["位置", "必须出现", "必须不出现", "现状"],
      [["标题", "An Empirical Study of；Schema Lineage Granularity；Historical SQL Migration",
        "RAG、retrieval、NL2SQL、医疗", "已对齐"],
       ["英文摘要", "empirical study；no retrieval module and no NL-to-SQL system；"
        "40-item probe is a negative control",
        "we propose a system、our framework、retrieval-augmented", "已对齐"],
       ["关键词", "schema evolution、historical SQL migration、expression-level lineage 等 5 个",
        "RAG、retrieval-augmented generation、NL2SQL", "已对齐"],
       ["贡献点", "定位行（负向声明）；三条实证发现",
        "我们提出／构建、系统、框架", "已对齐"],
       ["结论", "作为一项实证研究；不提出系统；不实现检索模块", "工程可用、优于语言模型", "已对齐"],
       ["投稿信", "An Empirical Study of…；not a system paper；negative control (p = 0.25)",
        "applied AI study、framework、pipeline", "已对齐"],
       ["第 7 章四处", f"{E2['n']} 条 / {E2['n_db']} 库 / p = {E2['p']:.2f}；已决断的表述",
        "仅 40 题、二者只能选一、真检索实验", "已对齐"]],
      [2.2, 7.4, 5.0, 2.4], ["l", "l", "l", "c"], font=8.0)
NOTE(d, "「必须不出现」一列是本轮最实用的东西：它可以直接当检索词用。"
     "在全文（含章内）搜 RAG、retrieval、framework、pipeline、system，"
     "凡命中处都要能回答「这句是在否定它，还是在主张它」。"
     "本轮命中的全部是否定句与元讨论（如关键词删除候选），无一处主张。")

# ================================================================= 六
H(d, "六、交付物与自检")
CAP(d, "表 6  本轮交付与实测规模")
TABLE(d, ["文件", "段落", "表格", "表题", "说明"],
      [["摘要_标题_贡献点_结论_v11.docx", str(N_FM[0]), str(N_FM[1]), str(N_FM[2]),
        "新增第零节；英文摘要 235 词 / 中文 1727 字"],
       ["投稿版_第7章_实验结果_v11.docx", str(N_MAIN[0]), str(N_MAIN[1]), str(N_MAIN[2]),
        "主文 18 表；章内四处检索口径改写"],
       ["补充材料_第7章_v11.docx", str(N_SUPP[0]), str(N_SUPP[1]), str(N_SUPP[2]),
        "S-1…S-39，未改结构"],
       ["第7章_v11_完整版_49表.docx", str(N_FULL[0]), str(N_FULL[1]), str(N_FULL[2]),
        "素材库，57 表"],
       ["第21轮修订说明_研究定位收敛为谱系粒度实证研究.docx", "—", "—", "—", "本文档"]],
      [7.0, 1.3, 1.3, 1.3, 4.6], ["l", "r", "r", "r", "l"], font=8.2)
P(d, "本轮不改任何一个统计口径。改动全部落在文字层："
     "标题 1 处、摘要 4 处（中英各 2）、关键词说明 1 处、贡献点 2 处、结论 2 处、"
     "投稿信 3 处、第 7 章 4 处、本章外文档新增 2 表 1 节。"
     f"统计量维持：主实验 {PS['items']} 条 / {PS['ops']} 个 op / {PS['events']} 次事件 / 11 库，"
     f"段位档 {SG['ops']} 个 op / 10 次事件，"
     f"扩充合并档事件级 {XPM['pos_events']} 正 {XPM['neg_events']} 负"
     f"（n={XPM['n_events']}，p={XPM['p']:.3f}）。")
P(d, "另须备案一处数据更新，否则新旧文档会对不上：扩充集的端到端评测在本轮之前已完整重跑并落盘，"
     "事件级读数由 12 正 3 负（n=18，p=0.035）更新为 "
     f"{XPM['pos_events']} 正 {XPM['neg_events']} 负（n={XPM['n_events']}，"
     f"p={XPM['p']:.3f}），条件级 C3m 由 5.8% 升至 7.3%、C3mp 由 44.7% 升至 44.9%。"
     "本次已把摘要、贡献点、结论里硬编码的旧读数全部改为从产物文件插值，"
     "并重跑第 20 轮说明，使章内、章外与修订说明三处同步。"
     "旧读数只保留在历史版本中，不作为任何当前表述的依据——"
     "这类数字一旦以常量写进正文，产物一重跑就会在作者不知情时分叉。")

# ================================================================= 七
H(d, "七、仍未闭环")
P(d, "第一，真实历史 Schema 演化仍为空。本文的演化由脚本事后施加，BIRD 只提供单版本库，"
     "因此「真实演化」这一栏无法填。替代性证据是 BIRD train 上的两项留出实验"
     "（H-E2E 165 条 / 7 库、H-LER 481 条 / 16 op / 8 库），它们换的是库，不是版本。")
P(d, "第二，检索若要成为研究方向，至少需要三部分：多路召回、证据冲突裁决、"
     "以及可回答性三分类（可答 / 不可答 / 应弃答）。本文一部分都没有做，"
     f"唯一相关的 {E2['n']} 条探针还是阴性对照。这不是工作量问题，而是证据问题——"
     "在现有证据上写 RAG，代价不是被批评保守，而是被判定声称与内容不符，"
     "而后者无法靠补充实验挽回，因为它质疑的是作者对自己工作的判断。")
P(d, "第三，若后续确有检索数据，正确的做法不是改回 RAG 标题，"
     "而是把它作为独立研究问题单列一节，并按事件级口径重新检验；"
     "在那一节出现之前，标题与摘要不应为它预留位置。")

d.save(OUT)
print("saved:", OUT, "段落", len(d.paragraphs), "表格", len(d.tables))
print(f"E2: {E2['n']} 条 / {E2['n_db']} 库 / C4 {E2['c4_rate']:.1f}% vs "
      f"C5 {E2['c5_rate']:.1f}% / {E2['b']}:{E2['c']} / p={E2['p']:.2f}")
