# -*- coding: utf-8 -*-
"""生成《第12轮修订说明_意见对账.docx》。

本轮意见针对的是 v10，而稿件已推进到 v11 并完成两轮修订。
因此这份说明的第一件事不是"改了什么"，而是"哪几条已经改过、哪一条真的还没解决"——
否则会重复劳动，也会把未解决的藏进已完成里。
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
OUT = "第12轮修订说明_意见对账.docx"
EN, CN, CN_H = "Times New Roman", "宋体", "黑体"

EBD = json.load(open("results/encoding_bounds.json", encoding="utf-8"))
ABM = json.load(open("results/ablation_matching.json", encoding="utf-8"))
HLD = json.load(open("results/holdout_stats.json", encoding="utf-8"))
RVR = json.load(open("results/rule_vs_llm.json", encoding="utf-8"))
S = json.load(open("results/stats_ch7_v5.json", encoding="utf-8"))["main"]
assert EBD["version"] == "encoding-bounds-v1"
assert ABM["version"] == "ablation-matching-v1"
_AB = ABM["groups"]
_SEG, _NSEG = _AB["M1+形态指纹"]["段位档"], _AB["M1+形态指纹"]["非段位档"]
_B3 = _AB["B3 嵌入语义"]["非段位档"]
_ev = HLD["event_level"]["merge_col_holdout"]
_EBM = EBD["real_data"]["merge_side"]
_rm = RVR["groups"]["merge"]["rates"]
_mg = S["merge"]

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
_run(p, "第 12 轮修订说明：与审稿意见的逐条对账", bold=True, size=15, font=CN_H)
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
_run(p, "—— 对象为《投稿版_第7章_实验结果》与《摘要_标题_贡献点_结论》", size=9.5)
NOTE("说明：本意见明确以 v10 两份文件为对象，而稿件已推进到 v11 并完成两轮修订。"
     "因此第一节先做对账——五条主要问题里哪些已经改过、改在哪、还差什么；"
     "第二节才列本轮的实际改动。把已改的重复改一遍是浪费，"
     "而把未解决的混进已完成里，是这类修订说明最常见也最伤信誉的错。")

# ===================================================================== 一
H("一、五条主要问题的对账")
P("先给结论：五条中四条已在 v10→v11 两轮修订中闭环，一条属定位问题、已按建议执行；"
  "本轮真正新增的只有一件事——把「关键边界情况缺失」从「承认未覆盖」改为「逐类界定」，"
  "并在摘要层面披露独立事件数。")

CAP("表 12-1  五条主要问题的对账结果")
TABLE(["审稿意见", "当前状态", "证据位置", "本轮动作"],
      [["① 样本数看起来较大，但独立 schema 演化事件偏少",
        "正文已处理；**摘要未披露**",
        "7.1.5 分析单位规则；op 级 10/5；事件级 6 正 1 负；EvoSchema 单元级 p=0.5",
        "摘要中英各加一句「610 条背后只有 10 次独立事件」"],
       ["② 合成数据效果好，但合成规则与方法同源",
        "已闭环（v10–v11）",
        "7.4.2 循环性说明；构造内上界；比较域三档（第十一轮）",
        "无（已改过）"],
       ["③ 自动谱系在开发集与未见库差距明显",
        "已闭环（v8–v11）",
        "留出 7 库 165 条；硬失败缺口 "
        f"{RVR['groups']['merge_holdout']['rates']['c3r_gt']['rate'] - RVR['groups']['merge_holdout']['rates']['c3r_infer']['rate']:.1f}pt；7.4.2 专段",
        "无（已改过）"],
       ["④ 确定性重写基线存在构造内循环验证",
        "已闭环（v10–v11）",
        "主张重摆：列级 0.9% 对段位 100.0%（同一执行器）；规则 100% 降为佐证",
        "无（已改过）"],
       ["⑤ 创新点较窄；建议彻底放弃 RAG 作为主要创新",
        "已执行",
        "标题、摘要、关键词、贡献点、结论、投稿信均无 RAG；"
        "正文写明「本章实验不足以支撑 RAG 定位，二者只能选一」",
        "无（已改过）"]],
      [4.6, 3.2, 6.0, 4.2], ["l", "l", "l", "l"], font=8.2)
NOTE("表内第一列是意见原文的压缩，未改其意。第四列写「无」的三行不是忽略，"
     "而是对应内容已在 v10/v11 落纸，重复修改只会让章内出现两处说法。")

P("第①条值得单独说明，因为它是本轮唯一真正补上的缺口。"
  "正文侧其实早有处理：7.1.5 立了固定的分析单位规则（描述性按条目、推断按最粗可用单位、"
  "条目级显著仅为必要条件），7.2 把段位档拆到 op 级，7.8 并列五种单位，"
  "EvoSchema 对照也按映射单元重报过。问题不在正文，在摘要——"
  "中文摘要只写「610 条样本」和「7 库方向重现」，没有一个字提独立事件数。"
  "审稿人只看摘要就能得出「样本大而独立性弱」的印象，而这个印象其实是对的，"
  "只是正文里有更细的说明。本轮把这条补进摘要，等于把正文的诚实搬到摘要，"
  "代价是英文摘要要腾出空间——为此压缩了四处措辞，词数仍守在 249。")

# ===================================================================== 二
H("二、评分表逐项看：靶心只有三项")
P("九项评分中，三项明显低于其余：泛化能力 2.2、数据集质量 2.3、创新性 2.8。"
  "实验设计 3.6 与统计严谨性 3.8 反而是强项——这与「不要继续堆积条目级实验」的判断一致。"
  "因此本轮只在最低的那一项上动手，另外两项属结构问题，靠加实验解决不了。")

CAP("表 12-2  三项最低分的成因与本轮处置")
TABLE(["维度", "分数", "成因", "本轮处置", "能否靠加实验解决"],
      [["泛化能力", "2.2", "无真实历史 schema 演化数据；留出库仍是脚本合成",
        "不处置，明确写「不主张跨真实演化泛化」", "**不能**——需外部数据集，非工作量问题"],
       ["数据集质量", "2.3", "关键边界情况被认为缺失",
        "边界已逐类界定；三段合并补齐；真实频次同表给出", "已做，低成本且已闭环"],
       ["创新性", "2.8", "增量创新，非新框架",
        "定位收窄为「段位恢复」这一个子问题，并主动写出 M1 在通用匹配上不如嵌入基线",
        "**不能**——这是命题范围，不是证据量"]],
      [2.4, 1.4, 4.6, 5.6, 5.0], ["l", "c", "l", "l", "l"], font=8.2)
NOTE("把「不能」写进表里是刻意的。这两项如果靠堆实验硬撑，"
     "结果会是样本更大、结论更弱——因为审稿人已经明说问题在独立性和真实性，不在条目数。"
     "承认它们不能靠本轮解决，比给出一个「又加了 200 条」的回应更可信。")

# ===================================================================== 三
H("三、本轮实际改动：边界从「未覆盖」到「逐类界定」")
P("审稿意见把数据集质量打到 2.3，理由是「关键边界情况缺失」。"
  "核对后发现，正文 7.7.5 其实已有隔离库探针（表 7-36，九类边界）与真实数据扫描"
  f"（表 7-37，{_EBM['events']} 个合并事件、{_EBM['rows']:,} 行）。"
  "所以缺的不是实验，而是两件事：一是这些边界没有在摘要与未完成项里被正确定位——"
  "原文写「构造未覆盖 NULL 与值内分隔符」，读起来像不知道会发生什么；"
  "二是三段合并确实没在库上跑过。本轮把这两处都补上。")

H("3.1  独立实现复核（新增 probe_encoding_bounds.py）", 2)
P("原探针在 Python 里直接切分字符串，结论依赖探针自身的实现。"
  "新脚本换一条完全不同的通道：在真实 SQLite 表上按构造规则 UPDATE 出编码列，"
  "再把生产代码 m1_part.to_expr 渲染的表达式回填进 SELECT 执行，逐行比较能否恢复。"
  "六个可对应情形的判定方向与表 7-36 逐类一致，说明那张表的结论不依赖探针实现。")

CAP("表 12-3  复核结果：六类边界的可逆性（整行判据）")
_rows = []
_lab = {"B0": "常规值", "B1a": "左分量为 NULL", "B1b": "右分量为 NULL",
        "B2a": "左分量为空串", "B2b": "右分量为空串",
        "B3a": "分隔符在左分量内", "B3b": "分隔符在右分量内"}
for k, lab in _lab.items():
    v = EBD["two_seg"][k]
    typ = {"ok": "可恢复", "fail_null": "整体 NULL（显式丢失）",
           "fail_value": "静默错值"}[
        "ok" if v["ok"] == v["n"] else ("fail_null" if v["fail_null"] else "fail_value")]
    _rows.append([lab, f"{v['ok']}/{v['n']}", typ])
TABLE(["边界情形", "可恢复", "失败性质"], _rows, [6.4, 2.6, 6.0],
      ["l", "c", "l"], font=8.6)
NOTE("判据是「整行全部分量同时恢复」，不是逐列恢复。这一条是本轮纠正的口径："
     "若按单列判，单个分量为 NULL 会被误判为可逆（NULL 反解仍是 NULL），"
     "实际是另一分量的信息随整体 NULL 一并丢失。"
     "这个误判方向恰好会把最危险的情形报成最好的结果，所以必须钉死。")

H("3.2  三段合并补齐（此前唯一没在库上验证的段数）", 2)
P("旧版 to_expr 在 nseg ≥ 3 时把 idx > 0 一律当作末段，第 1 段与第 2 段会渲染出"
  "完全相同的表达式——这才是此前「只验证两段」的真实原因，不是数据选择。"
  "该分支已改为逐段剥前缀，本轮在库上执行验证。")

CAP("表 12-4  三段合并的三类取值")
_rows = []
for k, lab in (("S3a", "三段全为常规值"), ("S3b", "中段含分隔符"), ("S3c", "首段为 NULL")):
    v = EBD["three_seg"][k]
    typ = {"ok": "可恢复", "fail_null": "整体 NULL", "fail_value": "静默错值"}[
        "ok" if v["ok"] == v["n"] else ("fail_null" if v["fail_null"] else "fail_value")]
    _rows.append([lab, f"{v['ok']}/{v['n']}", typ])
TABLE(["三段取值", "可恢复", "失败性质"], _rows, [6.4, 2.6, 6.0],
      ["l", "c", "l"], font=8.6)
NOTE("三段与两段共享同一条可逆条件，差别只在「非末段」的位置从一个增加到两个。"
     "需要限定：真实 BIRD 数据中不存在三段合并事件，这三行只证明实现分支正确，"
     "不构成三段场景下的效果证据——这句话已写进正文与未完成项①。")

H("3.3  真实数据的边界频次与「零越界」的正确读法", 2)
P(f"合并侧逐列扫描覆盖 {_EBM['events']} 个事件、{_EBM['rows']:,} 行："
  f"NULL 分量出现 {_EBM['null']} 次，值内含分隔符出现 {_EBM['has_sep_char']} 次，"
  f"空串出现 {_EBM['empty']} 次，有损行 {_EBM['lossy_rows']} 行。"
  "空串这一项此前被当作风险，实测全部可恢复（表 12-3），因此它不构成问题；"
  "真正致命的 NULL 与内部分隔符在真实数据中恰好为 0。"
  "这个「零」必须正确解读：它说明构造从未产生这类行，而不是方法处理得好。"
  "正文与摘要现在都按这个读法写，不再写「未覆盖」。")

H("3.4  未完成项的计数口径也改了", 2)
P("原文写「共六项，其中五项是尚未做的实验，第六项是本构造自身的评价边界」。"
  "第三节改完后这个说法已经不准：编码边界不再是「未做」，而是「已界定但本构造无法触发」。"
  "现改为三项仍未做（谓词重写、人工复核、检索）、三项属构造与评价范围边界（三段、编码、星号），"
  "并把「后者不是工作量问题，而是本构造无法产生该类情形，故只能界定、不能实测」写进了正文。")

H("3.5  摘要与限定条款的同步", 2)
PL("中文摘要末段新增：610 条样本背后只有 10 次独立的列重整事件，"
   "故合并与拆分另以演化算子与变更事件为单位复核，条目级显著只作必要条件。")
PL("中文摘要末段改写：编码边界已逐类测定（两类不可恢复、四类可恢复），"
   "真实数据中前两类出现 0 次，故零越界应读作构造未触发。")
PL("英文摘要：同等披露 10 次独立事件与逐类边界，词数由 289 压回 249，仍在 150–250 内。")
PL("第八节限定条款：③ 新增硬性要求——凡出现条目数必须同句给出独立单元数；"
   "① 补入逐类测定结果；⑥ 比较域条款保持。")
PL("结论段两处「编码固定为两段可逆拼接 / 未出现」改为「已逐类测定 + 两类致命情形出现 0 次」。")

# ===================================================================== 四
H("四、唯一仍然不能解决的一条")
P("泛化能力 2.2 是九项里最低的，也是本轮唯一没有实质动作的。原因不是回避，"
  "而是它需要的东西不在本研究的可控范围内：真实的历史 schema 演化数据"
  "（同一库的两个真实版本，且新旧 SQL 并非机械对应）。BIRD 提供 95 个库但没有版本差分，"
  "本章的演化全部由脚本注入。留出 7 库只能解决「换库」的问题，解决不了「换演化来源」的问题。")
P("有一点必须说清：这条不解决，不影响已经报告的数字，但影响这些数字能走多远。"
  "因此正文把主张收窄到「在本文构造下」——"
  "同一执行器下列级 0.9% 对段位 100.0% 这个对照，即使换到真实演化数据上，"
  "其方向仍可能保持（因为列级映射在合并方向上信息不足是结构性的），"
  "但绝对水平一定会下降。这个判断已写进 7.4.2 与结论，"
  "而不是留一句「未来工作」了事。")

# ===================================================================== 五
H("五、给编辑与审稿人的回应稿（可直接改写使用）")
P("我们感谢审稿人指出数据独立性与真实场景有效性这两点。"
  "在修订中我们核对了五条主要意见：关于合成规则与方法同源、开发库与未见库差距、"
  "确定性重写基线的循环性三条，前两轮修订已分别落实为"
  "「比较域三档界定」「留出 7 库并量化 19.8pt 硬失败缺口」"
  "「主张重摆为同一执行器下 0.9% 对 100.0%」，本轮不再重复修改。"
  "关于 RAG 定位，我们已按建议彻底移除：标题、摘要、关键词、贡献点与结论中不再出现，"
  "并在正文明确写出「本章实验不足以支撑检索主线，二者只能选一」。")
P("本轮新增两处。其一，独立事件数已进入摘要：610 条样本背后只有 10 次列重整事件，"
  "因此关键对照另以演化算子与变更事件为单位复核。其二，边界情形从「未覆盖」改为「逐类界定」："
  "我们用一套独立实现复核了六类边界的可逆性（结论方向与隔离库探针逐类一致），"
  "补齐了此前未在库上验证的三段合并，并给出真实数据中的出现频次"
  f"（{_EBM['events']} 个合并事件、{_EBM['rows']:,} 行，两类致命情形均为 0）。"
  "相应地，正文的「零越界」明确改写为「构造未触发」，而非「方法处理好」。")
P("关于泛化能力，我们不打算用增加条目数来回应。审稿人指出问题在独立性与真实性而非样本量，"
  "我们同意：真实历史演化数据的缺失是本构造的硬边界，"
  "因此把主张限定为「在本文构造下」，并逐条写明哪些结论不能外推。"
  "若期刊认为这一限定不足以支撑发表，我们愿意将其降格为明确的边界声明，"
  "而不通过扩大合成样本来掩盖。")

# ===================================================================== 六
H("六、本轮之后的剩余风险")
PL("**泛化能力仍为 2.2 量级**：这是构造决定的，除非引入真实版本差分数据，否则不会改善。")
PL("**创新性仍偏增量**：建立在「段位恢复」这一个子问题上，不是新框架。"
   "摘要与投稿信已按此强度写，不再往上加。")
PL("**正文 45 表、约 2.2 万字**：AI 无字数上限，不等于没有可读性下限。"
   "两个最重要的结果目前仍淹没在稳健性检验里，属结构问题，需单独一轮。")
PL("**无 DAS 与归档 DOI**：Applied Intelligence 对原创研究强制要求，提交前必须补。")
NOTE("这四项里，前两项不可消除只可界定，后两项可消除但需单独处理。"
     "本轮把可界定的部分写到了位，可消除的部分留待结构轮次。")

doc.save(OUT)
print("saved:", OUT)
print("段落:", len(doc.paragraphs), "表格:", len(doc.tables))
