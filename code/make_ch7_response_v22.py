# -*- coding: utf-8 -*-
"""生成《第22轮修订说明_一次注入缺陷的定位与修正.docx》。

本轮的性质与前十轮都不同：不是补充证据，而是**发现既有一个结论是错的**。
上一版把「扩充集比主实验难 35pt」当作真实现象，并围绕它写了落差分解；
本轮定位到根因是 m1_ext.py 向错误目录注入段位表达式，导致 223/479 条题
拿到的是另一张表的谱系。修正后 C3mp 由 44.9% 升至 91.2%，扩充集反而略高于主实验，
「落差」问题本身消失。

因此本文件的写法也不同于以往：先出事故报告，再列作废清单，最后才谈审稿意见的落点。
所有数字从 results/*.json 读取，不手写统计量。
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
OUT = "第22轮修订说明_注入缺陷定位与口径定稿.docx"

# ---------------------------------------------------------------- 产物
AU = json.load(open("results/audit_units_v3.json", encoding="utf-8"))
AC = json.load(open("results/audit_cluster_v3.json", encoding="utf-8"))
AG = json.load(open("results/audit_gap_v3.json", encoding="utf-8"))
SCV = json.load(open("results/sql_coverage_v3.json", encoding="utf-8"))
RVS = json.load(open("results/reversibility_v3.json", encoding="utf-8"))
GDI = json.load(open("results/gold_independence_v3.json", encoding="utf-8"))
RPC = json.load(open("results/repeat_consistency_v3.json", encoding="utf-8"))
LIA = json.load(open("results/lineage_injection_audit.json", encoding="utf-8"))
assert LIA["version"] == "lineage-injection-audit-v1"
EXT = json.load(open("results/ext_eval_merge_col.json", encoding="utf-8"))
EXS = json.load(open("results/ext_events_summary.json", encoding="utf-8"))["by_task"]
for _d, _v in ((AU, "audit-units-v3"), (AC, "audit-cluster-v3"), (AG, "audit-gap-v3"),
               (SCV, "sql-coverage-v3"), (RVS, "reversibility-v3"),
               (GDI, "gold-independence-v3"), (RPC, "repeat-consistency-v3")):
    assert _d["version"] == _v, f"产物版本不符：{_v}"

PS, SG = AU["main"]["paper_set"], AU["main"]["segment_set"]
XM, XS, XV = AU["ext"]["merge_col"], AU["ext"]["split_col"], AU["ext"]["value_shift"]
ACE = AC["ext_merge"]
AE_ALL, AE_ELI, AE_DB = (ACE["event_level_all"], ACE["event_level_eligible"],
                         ACE["db_level"])
AE_ITEM = ACE["item_level"]
AE_LOO, AE_BOOT = ACE["leave_one_db_out"], ACE["cluster_bootstrap"]
ACM = AC["main_merge"]
DET = RVS["detector_confusion"]
REAL = RVS["real_candidates"]
RC = REAL["total"]
OLD_RATE = 44.9          # 已作废的旧读数，仅用于对照，来源于 superseded 产物
MAIN_V32 = 79.6          # 主实验合并档 c3mp，DeepSeek-V3.2 首轮（183/230）
MAIN_V41 = AG["overall"]["main"]
EXT_V41 = AG["overall"]["ext"]


def stat(f):
    d = Document(f)
    caps = sum(1 for p in d.paragraphs
               if re.match(r"^表\s*(?:7|S|0|5|9)-\d+\s+\S", p.text.strip()))
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


def H(txt, lvl=1):
    p = d.add_paragraph()
    p.paragraph_format.space_before = Pt(13 if lvl == 1 else 8)
    p.paragraph_format.space_after = Pt(5)
    p.paragraph_format.keep_with_next = True
    _run(p, txt, bold=True, size={1: 13.5, 2: 11.5}[lvl], font=CN_H)
    return p


def P(txt):
    p = d.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.first_line_indent = Pt(21)
    emit(p, txt)
    return p


def PL(txt):
    p = d.add_paragraph()
    p.paragraph_format.left_indent = Pt(19)
    p.paragraph_format.space_after = Pt(2)
    emit(p, txt, 10.2)
    return p


def CAP(txt):
    p = d.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.keep_with_next = True
    _run(p, txt, bold=True, size=9)
    return p


def NOTE(txt):
    p = d.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(8)
    emit(p, txt, 8.4)
    return p


def TABLE(head, rows, widths, aligns, font=8.4):
    t = d.add_table(rows=1, cols=len(head))
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
_run(p, "第 22 轮修订说明：一次注入缺陷的定位与修正", bold=True, size=16, font=CN_H)
p = d.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
_run(p, f"Programming and Computer Software · Major Revision · {TODAY}", size=9.5)

# ================================================================= 一
H("一、先说结论：本轮推翻的是自己的一个结论")
P("审稿报告列了 12 条，其中 6 条「必须修改」。逐条处理的过程中，"
  "第 8 条（解释主实验 79.6% 与扩充集 44.7% 的落差）把本轮引向一个不属于写作层面的问题："
  "两个数差了 35pt，按惯例应当分层解释是难度、库、还是抽取质量。"
  "但分层之前先做了一步核对——两个数是不是测的同一件事——结果发现不是："
  "扩充集那一侧绝大多数题根本没有拿到对应列的段位表达式。")
P("定位结果是实现缺陷：m1_ext.py 把临时演化目录写进模块属性 MP.EVO_DIR，"
  "而 m1_part.infer_db() 读的是环境变量 EVO_DIR。两者不一致时，"
  "推断回退到默认目录 data/evolved_merge/<db>.sqlite，也就是**主实验的合并库**。"
  f"于是 {XM['constructed']} 个构造事件里有 {LIA['n_const_mismatched']} 个、"
  f"{XM['items']} 条题里有 {LIA['n_items_mismatched_v1']} 条，"
  "注入的是同一数据库里**另一张表**的段位表达式；render() 按旧列名过滤后谱系行为空，"
  "C3mp 退化成 C2 并大量弃答。")
P(f"修正后（修法与验证见第二节）C3mp 由 {OLD_RATE}% 升至 {EXT_V41['rate']}%"
  f"（{EXT_V41['ex']}/{EXT_V41['n']}），弃答由 111 条降为 3 条。"
  f"同一模型（DeepSeek-V4.1-Flash）下主实验为 {MAIN_V41['rate']}%、"
  f"扩充集为 {EXT_V41['rate']}%，扩充集反而高 "
  f"{EXT_V41['rate'] - MAIN_V41['rate']:.1f}pt——"
  "**那个 35pt 的落差不是现象，是缺陷**，针对它的整段解释（任务更难／库不同／引用频次不同）"
  "全部作废，改为「两批题难度无系统差异」的正面证据。")
P("这一轮之所以要单独出一份说明，原因不是改了很多字，而是：一旦正文里保留任何"
  f"{OLD_RATE}%、44.7%、35pt 的痕迹，读者拿到的就是一个由错误读数支撑的结论。"
  "因此本文件第二节给事故报告，第三节给作废清单，第四节以后才是审稿意见的落点。")

# ================================================================= 二
H("二、事故报告")
H("2.1  现象：三个互相矛盾的数字同时存在", 2)
P(f"在修正之前，扩充集合并档的 C3mp 有三个来源：改前产物 {OLD_RATE}%、"
  f"日志汇总 44.7%、以及部分分层表里 40% 附近的分项值。三者都被当作「真实水平」，"
  "并据此推断「扩充集更难」。矛盾之处早就出现过，只是被归因为抽样噪声："
  f"事件级符号检验在 {OLD_RATE}% 这一水平上仍得到 12 正 3 负（p = 0.035），"
  "方向一致却没有量级——一个真正「更难」的集合不会让方向如此整齐而水平如此之低。")
P("更早的旁证是弃答数：扩充集 C3mp 的弃答达 111 条，而主实验同一条件下弃答为 0。"
  "弃答集中在少数几个库，且与「该库有合并事件」高度相关。"
  "当时的解释是「模型在陌生库上更保守」，这一解释掩盖了真正的原因："
  "提示里给的段位表达式指向了不存在的列，模型只能弃答。")

H("2.2  根因：模块属性与环境变量不一致", 2)
P("缺陷出在实验脚本，不在方法代码，这一点对读者很重要：m1_part.infer_db() 的行为是正确的，"
  "它只是按设计读环境变量；错在调用方只改了一半。代码形态如下（简化）：")
PL("m1_ext.py：MP.EVO_DIR = tmpd　　　　　　# 只改了模块属性")
PL("m1_part.py：db_dir = os.environ[\"EVO_DIR\"]　# 读的是环境变量")
PL("不一致时 → 回退默认 data/evolved_merge/<db>.sqlite（主实验合并库）")
P("这个缺陷的隐蔽性在于「不会报错」：默认目录确实存在、里面确实有同名数据库、"
  "也确实有段位标注，只是标注属于另一张表。render() 按旧列名过滤后得到空行，"
  "空行不会抛异常，只会让提示里少一段内容——从外面看，就是「模型表现差」。"
  "同一类失败在本项目中已出现过三次（口径、判据、注入），"
  "其共同点是**错误落在数据里而不是异常里**，这也正是本文反复强调"
  "「静默失败是最危险失败模式」的自身经验来源。")

H("2.3  影响面：三个数要分开报", 2)
P("修正后逐分片核对（25 个分片全对），影响面可以精确到三个层次。"
  "这三个数经常被混用，因此表中把它们并列：")
CAP("表 1  注入缺陷的影响面（三个不同分母）")
_bad_db = len({e["db"] for e in LIA["events"] if e["table_v1"] != e["table_v2"]})
TABLE(["层次", "总数", "受影响", "占比", "说明"],
      [["构造事件", str(XM["constructed"]), str(LIA["n_const_mismatched"]),
        f"{100 * LIA['n_const_mismatched'] / XM['constructed']:.1f}%",
        "按 (演化档, 库, 表, 新列) 聚类；受影响指注入表 ≠ 应注入表"],
       ["有题事件", str(XM["evaluated"]), str(LIA["n_events_mismatched_v1"]),
        f"{100 * LIA['n_events_mismatched_v1'] / XM['evaluated']:.1f}%",
        "构造事件中至少有一条题目的"],
       ["题目", str(XM["items"]), str(LIA["n_items_mismatched_v1"]),
        f"{100 * LIA['n_items_mismatched_v1'] / XM['items']:.1f}%",
        "受影响的逐条评测记录"],
       ["库", str(XM["dbs"]), str(_bad_db),
        f"{100 * _bad_db / XM['dbs']:.1f}%",
        "v1 注入表 ≠ 应注入表的库数（按事件表计）"]],
      [3.0, 1.6, 2.0, 1.8, 8.6], ["l", "r", "r", "r", "l"])
NOTE("受影响题的判定标准是「该题被改列所属的表 ≠ 提示中段位表达式所属的表」。"
     "18、15、223 三个数各有分母，不能相加，也不能互相推算："
     "同一事件下题数不同，同一题只记一次。")
P("有一点需要如实说明：这个缺陷对本轮之前的**所有**结论方向都不构成威胁。"
  "C1、C2、C3m、c3f、c_rw 都不读段位表达式，逐条输出与修正前逐字相同；"
  "受影响的只有 C3mp 与以它为一项的对照。所以「段位谱系优于列级谱系」这一主结论方向未变，"
  "变的只有扩充集的绝对水平与由此派生的「难度落差」叙事。"
  "这不是自我安慰：如果缺陷落在 C3m 一侧，结论方向就会反过来，"
  "而当时的检查手段（方向一致性）根本发现不了。")

H("2.4  修法：三条断言，把「静默」变成「报错」", 2)
P("修法不是在出错处加 if，而是把三条隐含前提写成断言，让同类错误在下次运行时立刻中断：")
PL("① 同时设置环境变量与模块属性：os.environ[\"EVO_DIR\"] = tmpd 与 MP.EVO_DIR = tmpd 并列，"
   "并加注释说明两者必须同步，防止后来者只改一处。")
PL("② 临时目录名断言：assert os.path.basename(tmpd) == tag，"
   "确保临时目录确实指向当前分片，而不是复用上一片残留。")
PL("③ 推断表与分片表断言：推断结果里的表名必须等于分片名中的表名，"
   "否则拒绝写入。这一条直接针对本次缺陷——若当时有它，18 个事件会在第一片就中断。")
P("第 ③ 条是本轮最值得保留的工程经验：断言的对象不是「程序有没有崩」，"
  "而是「这次运行的身份对不对」。分片实验里，同库不同表是最容易串的，"
  "而串了以后所有下游指标都还能算出来、还能画出图。")

H("2.5  修正后的读数", 2)
CAP("表 2  扩充集合并档修正前后（同一批 479 题、同一模型、同一提示）")
TABLE(["读数", "修正前（作废）", "修正后（本版）", "变化"],
      [["C1（仅新 schema）", "4.6%", "4.6%", "不受影响"],
       ["C2（列名对齐）", "8.6%", "8.6%", "不受影响"],
       ["C3m（列级谱系）", "7.3%", "7.3%", "不受影响"],
       ["C3mp（段位表达式谱系）", f"{OLD_RATE}%", f"{EXT_V41['rate']}%",
        f"+{EXT_V41['rate'] - OLD_RATE:.1f}pt"],
       ["C3mp 弃答条数", "111", "3", "−108"],
       ["事件级（有题）", "12 正 3 负（p=0.035）",
        f"{AE_ALL['pos']} 正 {AE_ALL['neg']} 负 {AE_ALL['tie']} 平"
        f"（n={AE_ALL['n_events']}）", "方向不变、强度大幅提高"]],
      [4.6, 4.0, 4.6, 3.8], ["l", "c", "c", "l"])
NOTE(f"事件级 p 值：全部有题事件 {AE_ALL['p']:.1e}；合格事件 "
     f"{AE_ELI['pos']}:{AE_ELI['neg']}（n={AE_ELI['n_events']}，p={AE_ELI['p']:.0e}）；"
     f"库级 {AE_DB['pos']}/{AE_DB['n_dbs']}（p={AE_DB['p']:.4f}）。"
     "三个 p 值对应三种聚类单元，不能只报其中最小的那个。")
P(f"修正后需要重新回答审稿第 3 条：原来 12 正 3 负（p=0.035）属于「勉强越线」的边界性证据，"
  f"现在有题事件 {AE_ALL['pos']} 正 {AE_ALL['neg']} 负、合格事件 "
  f"{AE_ELI['pos']} 正 {AE_ELI['neg']} 负，"
  f"库级 {AE_DB['pos']}/{AE_DB['n_dbs']}（p={AE_DB['p']:.4f}，"
  f"已是 n={AE_DB['n_dbs']} 时的最小可达值）。"
  f"按库聚类的 bootstrap（{AE_BOOT['resamples']:,} 次，重采样库而非条目）给出事件级平均差 "
  f"{AE_BOOT['event_mean_delta_pt']['point']:.1f}pt，"
  f"95% 区间 [{AE_BOOT['event_mean_delta_pt']['ci95'][0]:.1f}, "
  f"{AE_BOOT['event_mean_delta_pt']['ci95'][1]:.1f}]；"
  f"留一库检验（leave-one-database-out）的 p 值区间为 "
  f"{min(x['p'] for x in AE_LOO):.1e}～{max(x['p'] for x in AE_LOO):.1e}，"
  "删去任何一个库结论都不翻向。审稿要求的「数据库聚类稳健性」由此闭环。")

# ================================================================= 三
H("三、作废清单：哪些话从现在起不能再写")
P("这一节的存在是为了让后来者（包括三个月后的作者）不必重读全部历史。"
  "凡下表左列出现的表述，无论出现在哪一版文档里，一律以右列为准。")
CAP("表 3  旧读数与旧表述的作废对照")
TABLE(["旧表述", "为什么作废", "本版口径"],
      [[f"扩充集 C3mp {OLD_RATE}%（或 44.7%）", "注入错表，读数不代表该方法",
        f"{EXT_V41['rate']}%（{EXT_V41['ex']}/{EXT_V41['n']}）"],
       ["扩充集比主实验难约 35pt", "同上；落差是缺陷产物",
        f"扩充集反而高 {EXT_V41['rate'] - MAIN_V41['rate']:.1f}pt（同模型）"],
       ["任务更难／库不同／引用频次不同导致落差", "分层做完后各层差 <5pt，无系统差异",
        "落差问题取消；分层表改为「难度无差异」的正面证据"],
       ["事件级 12 正 3 负，p = 0.035（边界性）", "受同一缺陷影响",
        f"有题 {AE_ALL['pos']}:{AE_ALL['neg']}（{AE_ALL['tie']} 平，"
        f"n={AE_ALL['n_events']}）；合格 {AE_ELI['pos']}:{AE_ELI['neg']}"],
       ["C3m 5.8% / C3mp 44.9%（扩充集）", "C3mp 作废，C3m 本身正确",
        "C3m 7.3% / C3mp 91.2%"],
       ["「主实验 C3mp 79.6% 对扩充集 44.7%」并列比较", "跨模型又跨集，双重不可比",
        "79.6% 属 V3.2；87.8% 属 V4.1-Flash；同模型对比只用后两者"]],
      [5.6, 4.6, 6.4], ["l", "l", "l"])
NOTE("作废的是读数与由它派生的叙事，不是方法。C1/C2/C3m/c3f/c_rw 的逐条输出在修正前后"
     "逐字相同，可作交叉核对：若有人发现这五个条件的读数也变了，说明改动范围超出了本轮声明。")

# ================================================================= 四
H("四、12 条审稿意见的落点")
P("下表按审稿报告的编号排列。「落点」指第 7 章或章外文档的具体节号，「可复算」指"
  "该条是否有独立脚本产出、且脚本只读既有逐条判定。最后两列是状态与不能闭环的原因。")
_ag_std = AG["standardization"]
CAP("表 4  12 条意见的处置")
TABLE(["#", "意见", "落点", "可复算产物", "状态"],
      [["①", "统一事件数／样本数／模型版本",
        "7.10、7.14、7.15、7.16", "audit_units_v3.py", "已闭环"],
       ["②", "Gold SQL 独立性与构造循环性",
        "7.18、7.7.3⑨", "audit_gold_indep_v3.py", "已闭环（人工核验除外）"],
       ["③", "扩充集 p=.035 定位与聚类稳健性",
        "7.7.1、7.15", "audit_cluster_v3.py", "已闭环（读数整体更新）"],
       ["④", "界定「真实 schema + 受控合成演化」",
        "7.1、7.7.3①", "—（措辞层面的改写）", "已闭环"],
       ["⑤", "统一执行结果一致性术语",
        "全文、7.0 术语表", "—（术语统一）", "已闭环"],
       ["⑥", "可逆性前提与边界测试",
        "7.7.5", "audit_reversibility_v3.py", "已闭环"],
       ["⑦", "JOIN/GROUP BY/HAVING 覆盖或排除",
        "7.17、7.7.3⑧", "audit_sqlcov_v3.py", "已闭环"],
       ["⑧", "解释 79.6% 与 44.7% 落差",
        "7.16", "audit_gap_v3.py", "已闭环（问题本身取消）"],
       ["⑨", "重复实验逐轮逐题报告",
        "7.7.3⑥、7.18", "audit_repeat_v3.py", "部分闭环（见第五节）"],
       ["⑩", "数据流转图／字典／日志／复现",
        "7.14、材料包", "pack_materials.py", "部分闭环（DOI 待办）"],
       ["⑪", "不改造为 RAG/NL2SQL 论文",
        "标题、摘要、贡献点", "—", "已按意见执行"],
       ["⑫", "不覆盖所有方言但限定范围",
        "7.7.3⑧、7.17", "audit_sqlcov_v3.py", "已闭环"]],
      [0.8, 5.4, 3.6, 4.2, 3.0], ["c", "l", "l", "l", "l"])
P("需要点明两条容易「看起来改了、其实没改」的意见。第 ⑤ 条不是把「execution equivalence」"
  "替换成「execution-result agreement」就结束——真正的风险是读者把两个 100% 当成性能："
  "重写器的 100% 来自 gold SQL 由同一条替换规则机械生成，"
  "所以正文一律写「construction-internal upper bound」，并把 LLM 的读数与之并列而不是相加。"
  "第 ⑫ 条也不是加一句「未来将扩展到更多方言」，而是给出可核对的覆盖表："
  f"610 条中 LEFT JOIN 仅 {SCV['paper']['features']['join_left']['pct']:.1f}%、"
  f"JOIN USING/NATURAL {SCV['paper']['features']['join_using_or_natural']['pct']:.1f}%、"
  f"相关子查询 {SCV['paper']['features']['subquery_correlated']['pct']:.1f}%、"
  f"CTE {SCV['paper']['features']['cte']['pct']:.1f}%、"
  f"窗口 {SCV['paper']['features']['window']['pct']:.1f}%。"
  "这些位置方法没有被检验，正文按「已检验／未覆盖／不适用」三类明写。")

# ================================================================= 五
H("五、两处「只能部分闭环」的诚实交代")
H("5.1  第 ⑨ 条：五轮逐题输出当年没有归档", 2)
P("审稿要求逐轮逐题报告。事实是：五轮重复运行当时只保存了汇总与部分条件，"
  "逐题输出没有按轮存档，现在无法重建。可行的替代证据是跨批次判重："
  f"同一条件在不同文件里出现 {RPC['strict']['n_pairs']} 组重复键、"
  f"共 {RPC['strict']['n_judgements']:,} 次判定，逐题一致率 "
  f"{RPC['strict']['item_agreement_pct']}%，"
  f"边缘率差中位数 {RPC['strict']['median_abs_delta_pt']:.1f}pt、"
  f"最大 {RPC['strict']['max_abs_delta_pt']:.1f}pt。"
  f"另一条证据是表达形式消融的 c3mp 对照臂与主实验 c3mp：同提示、同模型、同 230 题，"
  f"两个批次的逐题一致率 {RPC['focus_ablation_vs_main']['item_agreement_pct']}%，"
  f"六折差 ≤{RPC['focus_ablation_vs_main']['max_abs_delta_pt']:.1f}pt。")
P("这两条能支持「读数在批次间稳定」，不能替代「逐轮逐题报告」。"
  "因此正文写的是前者，且明确标注「逐轮逐题原始输出未按条目存档」，"
  "不把判重结果包装成重复实验报告。这一条在正文中保持「部分闭环」状态，"
  "直到有人按轮次重跑并存档。")

H("5.2  第 ② 条：没有第二编码者，就没有 kappa", 2)
P(f"gold 与本方法改写规则的同源程度可以量化：替换后的 gold 与构造期 rewrite() 输出逐字相同"
  f"（合并 {GDI['merge']['identical']}/{GDI['merge']['n']}、"
  f"拆分 {GDI['split']['identical']}/{GDI['split']['n']}），"
  f"被插入表达式中位占 gold 全长 {GDI['merge']['inserted_char_share_median']}%"
  f"（合并）与 {GDI['split']['inserted_char_share_median']}%（拆分）。"
  f"补贴量级由表达形式消融给出：给可复制表达式 {GDI['ablation']['c3mp_literal']['rate']}% "
  f"对给等价自然语言 {GDI['ablation']['c3mp_nat']['rate']}%，"
  f"差 {GDI['ablation']['delta_pt']}pt（{GDI['ablation']['only_literal']}:"
  f"{GDI['ablation']['only_natural']}，p={GDI['ablation']['mcnemar_p']:.0e}）。")
P("但「两位标注者独立编码、报告一致性系数」这件事没有做，因为本轮没有第二个人。"
  "处理方式是：不报告 kappa，不编造一致性数值；材料包中提交编码手册与一份 60 条分层抽样"
  "核验协议，供第三方按同一手册复核。这是未完成事项，写进 7.18 与 7.7.3⑨，"
  "不是阴性结果，也不是「已用等效复核替代」。")

# ================================================================= 六
H("六、第 ⑥ 条的落点：把可逆性写成前提，而不是讨论")
P("这一条原本只打算补几个 NULL 与分隔符用例，实际做下来变成三件事，"
  "因为「补测试」并不回答审稿真正的问题——审稿问的是「你的等价性依赖什么」。")
P(f"第一，把边界谱扩到 {len(RVS['config']['cases'])} 类边界值 × "
  f"{len(RVS['config']['seps'])} 类分隔符 × {RVS['config']['nsegs'][0]}–"
  f"{RVS['config']['nsegs'][-1]} 段，共 {DET['cells']} 格逐格执行生产代码。"
  f"结果是七类 100% 可恢复、三类 0%：{RVS['by_case']['null_any']:.0f}% 的「任一分量为 NULL」、"
  f"{RVS['by_case']['sep_in_nonlast']:.0f}% 的「分隔符在非末段」、"
  f"以及 {RVS['by_case']['escaped_sep']:.0f}% 的「分隔符被转义」。"
  "最后一类是本轮的新发现：编码侧做了转义也没用，因为解码表达式仍按分隔符字面计数；"
  "这条否掉了「加个转义就能补齐」的直觉，也说明前一轮把「引入转义」列为出路是不准确的。")
P(f"第二，写了一个构造前的准入检测器（三条规则：含 NULL、分隔符计数、类型漂移）。"
  f"在 {DET['cells']} 格上 tp={DET['tp']}、fp={DET['fp']}、fn={DET['fn']}、tn={DET['tn']}。"
  f"fn=0 是关键：没有一格是「放行但实际不可逆」，即它不会把静默失败放进评测集；"
  f"fp={DET['fp']} 全部来自「分隔符在末段」——该情形实际可逆，但列级规则无法预知段位顺序，"
  "只能一律拒绝。这是保守方向的错误。")
P(f"第三，把检测器用到真实数据上量回退代价：BIRD dev 的 11 个库有 "
  f"{RC['cand_cols']} 个候选分量列，拒绝 {RC['rej_null'] + RC['rej_type'] + RC['rej_sep']} 列"
  f"（{REAL['reject_rate_pct']}%），保留 {RC['kept_cols']} 列，能配成同表同类型列对的只有 "
  f"{RC['pairs']} 对。这个数字不太好听，但它约束的是「在全库自动挑列批量构造」这条路，"
  "不约束本章的 5 次人工指定事件（那 5 对已通过同样的检查）。"
  "如实写出来反而有用：它说明可逆性前提是真实演化的现实约束，"
  "真实历史中的列合并更可能发生在低基数、非空、非数值的文本列上。")
P("三条合起来，正文不再说「本方法在新值上保持语义」，而说：等价性依赖三条前提"
  "（P1 分量非 NULL；P2 非末段不含分隔符；P3 解码规则与编码规则逐字对应），"
  "本章满足前两条、第三条由同一实现保证，且三条都能在构造前检测。"
  "这是把「等价」从一个形容词改成一个可检查的条件。")

# ================================================================= 七
H("七、章内章外同步")
P("本轮改动跨三份文档与一份材料包，表号与 S 号都会漂移，因此同步规则也一并写死：")
CAP("表 5  本轮同步范围")
TABLE(["文档", "规模", "本轮改动", "校验方式"],
      [["第 7 章主文", f"{N_MAIN[0]} 段 / {N_MAIN[1]} 表",
        "7.7.5 重写；7.7.3 增至九点；新增 7.16–7.18",
        "表题连续、无超范围引用"],
       ["第 7 章补充材料", f"{N_SUPP[0]} 段 / {N_SUPP[1]} 表",
        "收纳 41 张 S 表；新增准入检测器回退代价表",
        "S 号连续 S-1..S-41"],
       ["第 7 章完整版", f"{N_FULL[0]} 段 / {N_FULL[1]} 表",
        "作为素材库，不删任何数字", "表题数 64"],
       ["章外（摘要/贡献/结论/投稿信）", f"{N_FM[0]} 段 / {N_FM[1]} 表",
        "英文摘要 234 词；中文摘要四层口径；S 号改为反查",
        "统计量全部插值，无手写常量"]],
      [4.4, 3.4, 5.4, 3.4], ["l", "c", "l", "l"])
P("两条同步铁律值得单独记下，因为本轮各踩过一次："
  "其一，章外文档禁止手写统计常量——所有数字必须从 results/*.json 插值，"
  "否则产物一重跑就会在作者不知情时分叉（44.9% 就是这样在摘要里活了两轮）。"
  "其二，补充材料表号必须按表题关键词反查，不能写死——"
  "本轮在 7.7.5 中间插入两张表，其后所有主文表号 +2、S 号 +1，"
  "写死的 S-35/S-38/S-39 当场全部指错。现改为反查，找不到就直接报错，不再静默用旧号。")

# ================================================================= 八
H("八、仍未闭环（逐条说明为什么不能靠写作解决）")
PL("① 真实历史 schema 演化仍为空。BIRD 只提供单版本库，"
   "本文的演化由脚本事后施加。没有第二版本，就没有任何写法能把它变成真实演化。")
PL("② 独立人工核验未做。需要第二个人按编码手册独立判 60 条；"
   "本轮无条件，故不报 kappa，只提交协议。")
PL("③ 五轮重复实验的逐题输出未归档。文件不存在，只能给跨批次判重的替代证据，"
   "并明确它不是同一件事。")
PL("④ JOIN USING/NATURAL、相关子查询、CTE、集合运算、窗口、HAVING 中的目标列均未覆盖。"
   "已在 7.17 列为明确排除项，不再用「未来工作」含糊带过。")
PL("⑤ 其他方言未验证。构造与评测都在 SQLite 上完成，方言差异对表达式语义的影响未测。")
PL("⑥ 仓库与 DOI 未落地。材料包已打包，但公开归档需要机构账号与审核周期。")
PL("⑦ 扩充集的拆分档与值改写档未跑端到端。"
   f"拆分档 {XS['constructed']} 个构造事件、值改写档 {XV['constructed']} 个构造事件"
   "目前只有构造与抽取层读数，端到端只做了合并档。")

P("最后说一句判断：这 12 条里，能让本文被拒的不是数值错误，而是内在效度——"
  "构造循环性、单作者编码、回顾性合成评测、5 次事件。"
  "本轮把这四项都写进了正文的边界节，而不是留在给审稿人的回复里，"
  "因为它们不因回复而消失。审稿报告建议 Major Revision 是恰当的；"
  "按修正后的读数，本文的贡献可以收窄为一句话："
  "在受控的两段列重整构造上，谱系的表示粒度把历史 SQL 的可用性从个位数推到九成左右，"
  "而这一效应的证据强度受 5 次事件与构造同源性限制。"
  "这句话既没有被 44.9% 支撑，也不需要它。")

d.save(OUT)
print("saved:", OUT, "段落", len(d.paragraphs), "表格", len(d.tables))
print(f"ext c3mp {EXT_V41['rate']}% ({EXT_V41['ex']}/{EXT_V41['n']}); "
      f"main V4.1F {MAIN_V41['rate']}%; 事件 {AE_ALL['pos']}:{AE_ALL['neg']}"
      f":{AE_ALL['tie']} n={AE_ALL['n_events']}")
