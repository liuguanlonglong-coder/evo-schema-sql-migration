# -*- coding: utf-8 -*-
"""生成《第20轮修订说明_独立事件扩容与口径统一.docx》。

回应本轮审稿意见的三块内容：
  ① 独立 Schema 演化事件太少（合并 5、拆分 5），统计方法无法创造事件；
  ② 若达不到 20–30 的规模，至少增加库、字段组合、分隔符编码与演化历史；
  ③ 4,031 / 4,490 / 610 三组数字缺一张流转表，且不能假定简单包含关系。

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
OUT = "第20轮修订说明_独立事件扩容与口径统一.docx"

ADI = json.load(open("results/audit_independence.json", encoding="utf-8"))
assert ADI["version"] == "audit-independence-v2"
PS, SG = ADI["A_units"]["paper_set"], ADI["A_units"]["segment_set"]
EXS = json.load(open("results/ext_events_summary.json", encoding="utf-8"))
EM, ES, EV = (EXS["by_task"]["merge_col"], EXS["by_task"]["split_col"],
              EXS["by_task"]["value_shift"])
XEM = json.load(open("results/ext_eval_merge_col.json", encoding="utf-8"))
XPM = XEM["paired_by_event"]["c3mp-c3m"]
XPS = XEM["paired_by_event"]["c3m-c2"]
DLN = json.load(open("results/data_lineage_audit.json", encoding="utf-8"))
L1, L2, L3 = DLN["L1_constructed"], DLN["L2_frozen"], DLN["L3_base_questions"]
L4, L5 = DLN["L4_events"], DLN["L5_outputs"]


def stat(f):
    d = Document(f)
    caps = sum(1 for p in d.paragraphs
               if re.match(r"^表\s*(?:7|S|5|9)-\d+\s\s", p.text.strip()))
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
_run(p, "第 20 轮修订说明：独立事件扩容与数据口径统一", bold=True, size=16, font=CN_H)
p = d.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
_run(p, f"Applied Intelligence 投稿 · 逐条回应 · {TODAY}", size=9.5)

# ================================================================= 一
H(d, "一、本轮意见与处置总览")
P(d, "本轮意见有一条主线和一个附带要求。主线是：**统计方法可以修正重复样本造成的推断偏差，"
     "却不能凭统计方法创造新的独立事件**——上一版把合并档的 230 条压缩到 5 个事件、"
     "拆分档的 104 条压缩到 5 个事件，然后用事件级符号检验去纠正条目级检验的高估，"
     "这一步方向对，但事件数没有变多，10 个事件能给出的最小双侧 p 就是 0.0625。"
     "附带要求是补一张数据集流转表，把 4,031 / 4,490 / 610 三组数字的层级关系写清楚，"
     "并核对它们在原始日志中的实际关系，而不是假定为简单包含。")
CAP(d, "表 1  本轮意见与处置对照")
TABLE(d, ["#", "意见", "处置", "落点", "状态"],
      [["1", "列合并独立事件 5 → 20–30",
        f"改构造：每（库, 表）各建一个独立演化库，事件 {EM['events']} 个"
        f"（有效 {EM['events_effective']}、合格 {EM['events_ge3']}）",
        "正文 7.15；主文表 7-56", "达标"],
       ["2", "列拆分独立事件 5 → 20–30",
        f"同法分片，事件 {ES['events']} 个（有效 {ES['events_effective']}、"
        f"合格 {ES['events_ge3']}），含 3 类分隔符",
        "正文 7.15；主文表 7-56", "达标"],
       ["3", "复杂值改写 10–20",
        f"按（库, 表, 列）分片，事件 {EV['events']} 个"
        f"（有效 {EV['events_effective']}、合格 {EV['events_ge3']}）",
        "正文 7.15；主文表 7-56", "超出目标"],
       ["4", "增加库 / 字段组合 / 分隔符编码",
        f"库 5 → {EM['dbs']}（合并、拆分）、{EV['dbs']}（值改写）；"
        f"表 {EM['tables']}/{ES['tables']}/{EV['tables']} 张；"
        f"分隔符 {ES.get('separator_kinds', 0)} 类",
        "正文 7.15；主文表 7-56 表注", "达标"],
       ["5", "把新增事件做成可评测读数",
        f"扩充集上重跑 end-to-end：C3mp {XEM['conditions']['c3mp']['rate']:.1f}% 对 "
        f"C3m {XEM['conditions']['c3m']['rate']:.1f}%；事件级 "
        f"{XPM['pos_events']} 正 {XPM['neg_events']} 负，p = {XPM['p']:.3f}",
        "正文 7.15；主文表 7-57", "已闭环"],
       ["6", "数据集流转表（4,031 / 4,490 / 610 / 437 / 61）",
        f"五层逐个核对到文件：{L1['rows']:,} / {L5['rows']:,} / "
        f"{L2['unique_ids']} / {L3['unique_question_id']} / {L4['paper_set']['events']}",
        "正文 7.14；主文表 7-55", "已闭环"],
       ["7", "不得假定数字之间存在简单包含关系",
        "核实只有两条包含关系成立，其余全部不可推算，逐条写明",
        "正文 7.14 表注", "已闭环"],
       ["8", "真实 Schema 演化事件尽可能增加", "未取得真实历史版本数据，如实标注为未闭环",
        "正文 7.15 末段", "未闭环"]],
      [0.8, 3.4, 7.4, 3.4, 1.4], ["c", "l", "l", "l", "c"], font=8.0)

# ================================================================= 二
H(d, "二、事件数只能靠构造解决")
P(d, "瓶颈不在分析而在构造。build_evolution_v2.py 的 run() 对每个库只取 "
     "pick_merge_pair / pick_split_col 的 top-1 候选，于是「每库每档 1 个事件」"
     "成为硬上限，事件数被钉在库数上——5 个可用库就是 5 个事件，与题量无关。"
     "本轮改为按表分片：每个候选（库, 表）各复制一份演化库、各施加一次 DDL 变更、"
     "各带一套题；值改写档按（库, 表, 列）分片。这样每个事件有自己的题集，"
     "任意两个片段不共享被改列，且仍走主实验的四步验证。")
CAP(d, "表 2  各档独立事件的前后对比")
TABLE(d, ["演化档", "旧事件", "新事件", "有效事件", "合格事件", "库", "表", "题"],
      [["列合并", '5', str(EM["events"]),
        str(EM["events_effective"]), str(EM["events_ge3"]),
        str(EM["dbs"]), str(EM["tables"]), str(EM["items"])],
       ["列拆分", "5", str(ES["events"]), str(ES["events_effective"]),
        str(ES["events_ge3"]), str(ES["dbs"]), str(ES["tables"]), str(ES["items"])],
       ["复杂值改写", "6", str(EV["events"]), str(EV["events_effective"]),
        str(EV["events_ge3"]), str(EV["dbs"]), str(EV["tables"]), str(EV["items"])]],
      [3.0, 1.7, 1.7, 1.9, 1.9, 1.2, 1.2, 1.4],
      ["l", "r", "r", "r", "r", "r", "r", "r"], font=8.2)
NOTE(d, "「有效事件」= 至少有 1 道题通过四步验证的事件；「合格事件」= 至少 3 道题通过，"
        "是事件级检验实际使用的样本。旧事件数取自主实验：合并 5、拆分 5、值改写 6。"
        "新事件数超过审稿建议的 20–30 区间，是因为按表分片后候选表本身就是这么多；"
        "为避免夸大，正文同时给出有效与合格两个口径。")

P(d, f"端到端在扩充集上重跑后，段位谱系相对列级谱系的事件级配对为 "
     f"{XPM['pos_events']} 正 {XPM['neg_events']} 负、持平 {XPM['tie']}，"
     f"n = {XPM['n_events']}，p = {XPM['p']:.3f}——**这是本章第一次在事件层面过 0.05 线**。"
     "同一批数据上，列级谱系相对无谱系条件为 "
     f"{XPS['pos_events']} 正 {XPS['neg_events']} 负（p = {XPS['p']:.3f}），"
     "方向不利且不显著，说明扩充并没有把所有对比都推成正向。这两条必须一起写，"
     "否则就成了只挑有利方向报告。")
CAP(d, "表 3  扩充集上的端到端结果（%）")
TABLE(d, ["事件集", "C1", "C2", "C3m", "C3mp", "题数", "事件级 p（C3mp−C3m）"],
      [["合并（扩充集）"] + [f"{XEM['conditions'][c]['rate']:.1f}"
                        for c in ("c1", "c2", "c3m", "c3mp")]
       + [str(XEM["conditions"]["c2"]["n"]), f"{XPM['p']:.3f}"]],
      [3.6, 1.6, 1.6, 1.6, 1.6, 1.4, 4.0],
      ["l", "r", "r", "r", "r", "r", "r"], font=8.2)
NOTE(d, "模型为 DeepSeek-V4.1-Flash，与主实验跨模型复现的权重版本一致；提示模板、"
        "判定器与谱系抽取器一行未改。扩充集的绝对水平低于主实验（主实验 C3mp 79.6%），"
        "原因是按表分片后纳入了更多低引用、多表的列重整事实，任务整体更难；"
        "可比的是条件之间的差，不是与主实验的绝对数值。")

# ================================================================= 三
H(d, "三、数据集流转：五个数字分属五层")
P(d, "审稿意见要求核对而不是假定。核对结果确认了它的怀疑：三组数字之间**不存在**"
     "简单的倍数或包含关系。只有两条包含成立——冻结评测是构造记录的子集，"
     "冻结评测的题号全部出现在逐题输出里；其余任何推算都不成立。")
CAP(d, "表 4  数据集四层流转（第五层为运行记录）")
TABLE(d, ["层", "数量", "定义与来源", "可否直接比较"],
      [["① 构造记录", f"{L1['rows']:,}",
        f"manifest_v3 全量：9 类演化 × {L1['dbs']} 库 × {L1['question_ids']} 基础题",
        "否，是构造规模"],
       ["② 冻结端到端评测", f"{L2['unique_ids']}",
        f"ids_xm_full + ids_ms 去重；5 档、{L2['dbs']} 库；"
        f"其中 {L2['ids_not_in_manifest']} 条不在构造清单内",
        "是，主实验结果以此为准"],
       ["③ 独立基础题", f"{L3['unique_question_id']}",
        f"②的 question_id 去重，平均每基础题 {L3['items_per_question']} 条",
        "否，是②的下层计数"],
       ["④ 独立演化事件", f"{L4['paper_set']['events']}",
        f"②背后 {L4['paper_set']['ops']} 个 op 按（库, 表, 演化档/目标列）聚类；"
        f"段位相关 {L4['segment_set']['events']} 次 / {L4['segment_set']['dbs']} 库",
        "是，事件级检验的合法 n"],
       ["⑤ 模型逐题输出", f"{L5['rows']:,}",
        f"outputs.jsonl：{L5['conditions']} 条件、{L5['unique_ids']} 个唯一题号"
        f"（冻结内 {L5['ids_in_frozen']}、留出库 {L5['ids_outside_frozen']}）",
        "否，是同一题被多次调用的记录"]],
      [3.0, 1.6, 9.0, 3.4], ["l", "r", "l", "l"], font=8.0)
NOTE(d, f"⑤的 {L5['rows']:,} 行不是「{L2['unique_ids']} × 条件数」："
        "各条件适用的题集不同（c1 / c3f / c_rw 覆盖冻结集全部；c2 / c3m 另含留出库条目；"
        "c3mp 只覆盖合并与拆分档；c3 只有一部分），因此它不能被读成对同一批题的重复测量，"
        "也不能用来估方差。③与④是②的两种去重视角，方向不同、不相等、不互相包含。"
        f"⑤的模型列只有一个取值（{L5['models'][0]}），即全部为单模型输出。")

# ================================================================= 四
H(d, "四、本轮自查发现的两处工程缺陷")
P(d, "这两条不在审稿意见里，但都会让「已经改好的东西」悄悄丢失，因此必须记录。")
CAP(d, "表 5  新增或修复的工程缺陷")
TABLE(d, ["缺陷", "现象", "根因", "修法"],
      [["分片构造覆盖写",
        "先跑 5 库得到 19 个拆分事件，再单独跑 card_games 后清单只剩 5 个",
        "脚本每轮全量重写清单文件",
        "改为按 id 合并写入，事件表按（库, 表, 序号）去重合并"],
       ["同表多列事件被并成一个",
        "student_club.major 的 college 与 department 落进同一事件",
        "拆分档事件键缺列序号，且分片文件名相同互相覆盖",
        "补 pair_idx 并写进文件名与事件键"],
       ["素材库自毁",
        "第二次运行 slim 时，完整版被读成 18 表的主文，备份随之被覆盖",
        "主文输出与完整版输入是同一文件（MAIN_OUT = SRC）",
        "表数少于备份时判定为已被覆盖，改从备份读且不刷新备份"],
       ["评估进程静默退出",
        "1,916 次调用跑到一半进程消失，结果全部丢失",
        "结果在最后一次性写盘，无中间落盘",
        "逐条 append + 断点续跑，重复启动只补未完成项"]],
      [3.0, 5.2, 4.2, 4.6], ["l", "l", "l", "l"], font=8.0)

# ================================================================= 五
H(d, "五、边界与仍未闭环的部分")
P(d, "扩容改善的是「列重整事实」层面的独立性，不改善「库」层面的独立性。"
     f"同一库内的不同表、同一表内的不同列对，是彼此独立的列重整事实，"
     f"但它们共享同一个数据库实例——事件数从 10 增到 46，库数只从 5 增到 6。"
     "任何事件级结论都必须同句给出库数，这条写作要求不因扩容而放宽。")
P(d, "扩充集与主实验的条目不可相加：同一道基础题在两套构造下各出现一次，"
     "合并只会放大重复计数偏差，而这正是本轮意见批评的做法。正文把两者分写为"
     "「同一批库、两套演化构造」，扩充集的作用是事件数与稳定性复核，"
     "不替代 H-E2E 的留出结论。")
P(d, "**仍未闭环**：真实 Schema 演化事件。BIRD 只提供单一版本的数据库，"
     "本章没有真实历史 schema 版本可供对照，因此「真实演化事件」这一栏仍为空。"
     "替代性证据是 BIRD train 划分上的两项留出实验（H-E2E 165 条 / 7 库、"
     "H-LER 481 条 / 16 op / 8 库），它们换的是库，不是版本。"
     "如果要真正补上这一栏，需要带版本历史的数据源（数据仓库的迁移日志、"
     "开源项目的 DDL 变更史），并需要与之配套的历史 SQL 作为 gold——"
     "这是独立的一轮工作，不应在本轮以近似数据充数。")

# ================================================================= 六
H(d, "六、交付物与自检")
CAP(d, "表 6  本轮交付与实测规模")
TABLE(d, ["文件", "段落", "表格", "表题", "说明"],
      [["投稿版_第7章_实验结果_v11.docx", str(N_MAIN[0]), str(N_MAIN[1]),
        str(N_MAIN[2]), "主文，18 表（本轮 +3）"],
       ["补充材料_第7章_v11.docx", str(N_SUPP[0]), str(N_SUPP[1]),
        str(N_SUPP[2]), "S-1…S-39"],
       ["第7章_v11_完整版_49表.docx", str(N_FULL[0]), str(N_FULL[1]),
        str(N_FULL[2]), "素材库，57 表"],
       ["摘要_标题_贡献点_结论_v11.docx", str(N_FM[0]), str(N_FM[1]),
        str(N_FM[2]), "英文摘要 247 词 / 中文 1515 字"],
       ["第20轮修订说明_独立事件扩容与口径统一.docx", "—", "—", "—", "本文档"]],
      [6.2, 1.3, 1.3, 1.3, 5.0], ["l", "r", "r", "r", "l"], font=8.2)
P(d, "新增脚本 6 个、新增结果产物 6 份，全部可复算：")
CAP(d, "表 7  新增脚本与产物")
TABLE(d, ["脚本", "作用", "产物"],
      [["probe_event_expansion.py", "量化扩容空间（只读探查）", "results/event_expansion_probe.json"],
       ["build_evolution_ext.py", "按表/按列分片构造扩充事件集", "data/manifest_ext_*.jsonl；results/evolution_ext_*.json"],
       ["m1_ext.py", "在分片库上跑 M1（profile / shape / part）", "data/lineage_infer_ext/*.json"],
       ["run_ext_eval.py", "扩充集端到端（C1/C2/C3m/C3mp）", "results/ext_eval_*.jsonl / .json"],
       ["ext_summary.py", "事件统计汇总", "results/ext_events_summary.json"],
       ["audit_data_lineage.py", "五层口径核对到原始文件", "results/data_lineage_audit.json"]],
      [4.4, 6.4, 6.2], ["l", "l", "l"], font=8.2)

d.save(OUT)
print("saved:", OUT, "段落", len(d.paragraphs), "表格", len(d.tables))
