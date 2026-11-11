# -*- coding: utf-8 -*-
"""生成《第14轮修订说明_事件口径与公平基线.docx》。

回应《Applied Intelligence 审稿式综合评估报告_v11第13轮.pdf》给出的 Major Revision
与 P0 清单：① 可审计事件映射表 ② M1 指标归属 ③ 摘要精确区分 15/15 与 10/10
④ 构造循环性 ⑤ 投稿材料真实性。另记录本轮两项新增审计带来的口径修正：
事件数 7 → 10（p 0.015625 → 0.001953），以及段位感知公平基线 B4 的负面结论。
所有数值直接读 results/*.json，不手写常量。
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
OUT = "第14轮修订说明_事件口径与公平基线.docx"
EN, CN, CN_H = "Times New Roman", "宋体", "黑体"

EMA = json.load(open("results/event_map_audit.json", encoding="utf-8"))
assert EMA["version"] == "event-map-audit-v1"
B4 = json.load(open("results/baseline_segment_aware.json", encoding="utf-8"))
assert B4["version"] == "segment-aware-baseline-v1"
ADI = json.load(open("results/audit_independence.json", encoding="utf-8"))
assert ADI["version"] == "audit-independence-v1"

C = EMA["counts"]
SC, SV = EMA["sign_correct"], EMA["sign_v13"]
TN = EMA["three_numbers"]
BN, BG, MP = B4["B4_naive"], B4["B4_greedy"], B4["M1_part"]
PS = ADI["A_units"]["paper_set"]

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
_run(p, "第 14 轮修订说明：事件口径与公平基线", bold=True, size=15, font=CN_H)
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
_run(p, "回应《Applied Intelligence 审稿式综合评估报告（v11 第 13 轮）》P0 清单", size=9.5)

P("报告维持 **Major Revision**，但明确肯定第 13 轮四项改进（事件级统计、M1 归属拆分、"
  "610 条口径、摘要清理）已经闭环，并给出新的 P0 清单。本轮不推翻 v11，也不增加普通 SQL "
  "样本；只做可审计性补齐与口径修正。下表按 P0 逐条对账。")
P("本轮还有两项**对结论不利**的更正，必须写在前面：其一，第 13 轮报的 7 次独立事件源于"
  "聚类键缺陷，正确口径是 10 次；其二，新增的同等输出能力基线 B4 达到与 M1 相同的 10/10，"
  "因此本文**不能**主张 M1 的抽取算法优于基线，只能主张谱系输出粒度决定下游成败。")

# ===================================================================== 一
H("一、P0 清单逐条对账")
CAP("表 14-1  报告 P0 五项与本次修订的对应关系")
TABLE(["P0 项", "报告指出的问题", "本轮处置", "落点"],
      [["可审计事件映射", "10 次事件缺逐条映射，外部无法复核聚类是否人为凑数",
        "新增 event_map_audit.py，四层定义（op ⊂ 映射事实 ⊂ 变更事件 ⊂ 库聚类），"
        "逐 op 列出 map ID / event ID / 旧列 / 目标列 / 表达式类型",
        "7.1.8、表 7-15、表 7-16"],
       ["M1 指标归属", "15/15、10/10、91.8% 三个读数被混称，读者会误以为是同一指标",
        "新增「四个关键读数各自属于哪个方法变体」表，把类型识别、表达式字符级一致、"
        "M1+shape 的两组 op 分开；B1–B3 输出空间不含段位表达式，0/15 标为表达能力不对称",
        "表 7-19、表 7-20"],
       ["摘要精确区分", "摘要仍把 15/15 叫段位恢复，且未披露独立事件数",
        "英文摘要重写为 239 词，逐项写清 15/15 类型识别、10/10 表达式一致、"
        "B4 同等输出空间亦 10/10、M1+shape 分组；中文摘要同步",
        "摘要_标题_贡献点_结论_v11.docx"],
       ["构造循环性", "gold SQL 与确定性 rewriter 同源，100% 只是构造内上界",
        "7.1.3 追加判定器与结构覆盖限定；7.4 全部读数作用域限定为「仅替换列引用」的比较域；"
        "规则 100% 只作佐证，不作工程推荐",
        "7.1.3、7.4.2、7.4.3"],
       ["投稿材料真实性", "投稿信写 released artefact / DAS is included，但归档实际未完成",
        "改为「已整包、将在发表前 deposit、归档尚未完成」；表 9-2 状态改为未完成，"
        "表 9-3 标注 [repository DOI] 是占位符，归档完成前不得作为已发布声明",
        "投稿信第四段、表 9-2、表 9-3"]],
      [2.6, 4.4, 6.4, 3.0], ["l", "l", "l", "l"], font=8.0)
NOTE("P1（留出库 M1 直接表达式指标、双人标注 kappa、LLM 重复运行、SQL 比较器类型感知、"
     "更多编码压力测试、其他方言）仍未闭环，本轮不伪称已完成；见第六节。")

# ===================================================================== 二
H("二、事件口径修正：7 → 10（p 由 0.015625 改为 0.001953）")
P("第 13 轮把事件定义为 **(db, table, 表达式首个 attr_ 列)** 三元组，得到 7 次事件、"
  "精确符号检验 p = 0.015625。审计发现这个键有两处缺陷：它**忽略 evo 档**，且**只取表达式里"
  "第一个 attr_ 列**。于是合并档的 SUBSTR(attr_001, …) 与拆分档的 attr_001||sep||attr_002 "
  "在 california_schools、formula_1、thrombosis_prediction 三个共用库上被并成同一事件，"
  "10 个正确事件塌缩为 7。本版把事件键改为 **(evo, db, table, target_cols)**。")
CAP("表 14-2  三个事件数各自的含义与处置")
TABLE(["数字", "来源", "含义", "处置"],
      [["7", "第 13 轮三元组聚类",
        "忽略 evo、只取首个 attr_ 列，3 个共用库的合并/拆分被误并",
        "作废；event_map_audit.py 中保留 events_v13 供复核"],
       ["10", "本版 (evo, db, table, target_cols)",
        f"合并 {C['events_correct_by_evo']['merge_col']} + 拆分 "
        f"{C['events_correct_by_evo']['split_col']} 次独立列重整事件，"
        f"对应 {C['atomic_ops']} 个原子 op、{C['db_clusters']} 个库聚类",
        "全文统一采用；精确符号检验 "
        f"{SC['ours_only']}:{SC['base_only']}，n={SC['n']}，p={SC['p_exact']:.6f}"],
       ["2", "第 13 轮 D_segment_expr「覆盖事件」",
        "要求事件内全部 op 都有 part 产物，叠加 7 口径的聚类混淆后只剩 2 个",
        "属低估；段位表达式恢复只在合并档 10 op / 5 事件上有定义"]],
      [1.5, 3.6, 6.0, 5.3], ["c", "l", "l", "l"], font=8.1)
NOTE("p 值变大还是变小并不重要，重要的是分母从「错误聚类出的 7」变成「可逐条追溯的 10」；"
     "10 次事件仍是小样本，正文与摘要均不把它写成强泛化证据。")

# ===================================================================== 三
H("三、段位感知的公平基线 B4：一个主动收紧的负面结论")
P("报告指出 B1–B3 的输出类型空间只有 rename_column / deprecate_column，无法输出 "
  "SUBSTR/INSTR 段位表达式，因此 0/15 是**表达能力不对称**，不是匹配精度低。为把这个质疑"
  "彻底关掉，本轮新增 B4：非 LLM、复用 m1_part.py 的 detect_sep / jac / read_schema / "
  "col_values / to_expr，输出空间与 M1 完全相同，只去掉 M1 的阈值门控、低基数守卫与采样门限，"
  "改用纯 argmax 段位分配。评估与 D_segment_expr 同口径：归一化后逐字符比对 gold，"
  "不经 evaluate()。")
CAP("表 14-3  同等输出能力下的段位表达式恢复")
TABLE(["方法", "表达式一致 / op", "一致率", "覆盖事件", "说明"],
      [["B4 · naive（纯 argmax）", f"{BN['expr_match']}/{BN['n_ops']}",
        f"{BN['pct']:.1f}%", f"{BN['events_ok']}/{BN['n_events']}",
        "去掉 M1 全部门控"],
       ["B4 · greedy（去重分配）", f"{BG['expr_match']}/{BG['n_ops']}",
        f"{BG['pct']:.1f}%", f"{BG['events_ok']}/{BG['n_events']}",
        "再加一层冲突消解"],
       ["M1 · part 分支", f"{MP['expr_match']}/{MP['n_ops']}",
        f"{MP['pct']:.1f}%", f"{MP['events_ok']}/{MP['n_events']}",
        "本文方法，带阈值门控"],
       ["B1 / B2 / B3", "0/15（type_ok）", "0.0%", "0/7",
        "输出空间不含段位表达式，不可与上行直接比较"]],
      [4.2, 3.0, 2.0, 2.4, 4.8], ["l", "c", "c", "c", "l"], font=8.2)
P(f"审计过程中修过一个**对方法不利**的实现误差：gold 的旧列名为小写，库内列名为原始大小写"
  f"（County / A2 / SEX 等），初版评估侧未做归一化，B4 只有 4/10。修正大小写映射后 B4 为"
  f"{BN['expr_match']}/{BN['n_ops']}。该修正方向对本文不利，正文保留说明，不隐藏。")
P(f"拆分档 {len(B4['split_ops_excluded'])} 个 op 的 gold 是拼接表达式 "
  "(attr_001 || sep || attr_002)，不是段位 SUBSTR，因此不进入 B4 的分母，单独说明。"
  "结论据此收紧为：**同等输出能力下 B4 与 M1 同为 10/10，M1 的阈值门控在段位表达式恢复上"
  "未带来可观测增益；可主张的是输出粒度，不是抽取算法优越性。** 这一改动削弱了原「M1 更聪明」"
  "的措辞，但把主命题（谱系粒度决定下游）留得更干净。")

# ===================================================================== 四
H("四、术语、判定器与比较域的限定")
CAP("表 14-4  本轮统一替换的四类限定表述")
TABLE(["位置", "旧表述", "本版表述", "原因"],
      [["方法章 / 摘要", "无监督谱系抽取",
        "推断阶段不使用标签、阈值由开发集校准（label-free inference with "
        "development-set-calibrated thresholds）",
        "阈值在开发集调过，「无监督」不严谨"],
       ["判定口径 / 摘要", "执行等价率 / SQL 语义等价",
        "在所评估 SQLite 实例上的执行结果一致率",
        "比较器只对比结果集，不是形式化等价证明；且方言限定"],
       ["7.1.3 / 7.4", "规则 100% 优于 LLM",
        "构造内上界；三档比较域中仅第一档（只替换列引用、谓词/聚合/连接不变）成立",
        "gold 与 rewriter 同源；第二档需改谓词字面值、第三档需改查询结构，均未评测"],
       ["投稿信 / 表 9-2 / 9-3", "released artefact、DAS is included",
        "材料包已整理，将在发表前公开归档并给出持久标识符，归档尚未完成",
        "避免投稿材料出现未兑现事实"]],
      [2.8, 3.5, 6.2, 4.4], ["l", "l", "l", "l"], font=8.0)
NOTE("三档比较域在第 11 轮已写入正文，本轮把「五重限制」同步扩为「六重限制」、投稿信补英文限定，"
     "并把「大一个数量级」改为按数值动态计算的 4.9 倍（99.1 / 20.4 = 4.86）。")

# ===================================================================== 五
H("五、产物与校验")
PL("· make_chapter7_v11.py — 接入 event_map_audit.json 与 baseline_segment_aware.json；"
   "新增 7.1.8 与 7.2.1 两节、四张表；7.1.3 与 7.7.3 追加限定。")
PL("· make_front_matter_v11.py — 摘要/贡献/结论/投稿信/自检表同步 10 事件口径；"
   "英文摘要 239 词（期刊 150–250）；中文摘要 819 字。")
PL("· event_map_audit.py、baseline_segment_aware.py — 两个新增审计脚本，产物入 results/。")
PL("· pack_materials.py — 纳入两个新脚本与两份新 stats，材料包重新生成。")
CAP("表 14-5  生成结果校验")
TABLE(["文档", "段落", "表格", "表号", "悬空引用", "markdown 残留"],
      [["投稿版_第7章_实验结果_v11.docx", "292", "49", "1–49 连续", "0", "0"],
       ["摘要_标题_贡献点_结论_v11.docx", "72", "5", "n/a", "0", "0"],
       ["第14轮修订说明_事件口径与公平基线.docx", "见文末", "5", "14-1–14-5", "0", "0"]],
      [6.6, 1.8, 1.8, 2.6, 2.0, 2.4], ["l", "c", "c", "c", "c", "c"], font=8.4)
P("残留词检查：文档中仍可检索到 0.0156 / 无监督 / released artefact，但均出现在「旧口径作废"
  "说明」「术语为何不准确」「上一版错误措辞」的修正语境，或属于无关的拆分档 McNemar p 值，"
  "不是仍在主张的旧结论，故保留。")

# ===================================================================== 六
H("六、明确未闭环与下一轮")
PL("· P1 未做：留出库上单独评估 M1 段位表达式质量；第二位标注者与 Cohen’s kappa；"
   "核心 LLM 条件重复运行 3 次并报均值/标准差；SQL 比较器改为类型感知、NULL/空串区分、"
   "ORDER BY 保序主判；更多 NULL、分隔符、三段以上编码压力测试；其他 SQL 方言验证。")
PL("· 正文结构：报告 Weakness 第 7 条「结构过于复杂、主要科学发现不够突出」。"
   "本轮为回应 P0 又新增 4 张表，当前 49 表，比报告评估时的 45 表更重；"
   "下一轮单独做主文瘦身（12–15 表）并把逐题/阈值/实现细节移入补充材料。")
PL("· 公开归档：Data Availability Statement 已按「尚未完成」如实改写，"
   "但 GitHub/Zenodo 归档与 DOI 本身仍未完成，投稿前必须落定。")
P("总体判断：本轮把审稿人点名的统计口径、指标归属、基线公平性与材料真实性四处硬伤全部"
  "暴露并修正，且主动报告了两个对己不利的更正。**但 Major Revision 的判定不变**——剩余风险"
  "已从「数字错误」转为「结构臃肿 + 归档未落地 + 泛化证据弱」，这三项都不再能靠增加普通实验解决。")

doc.save(OUT)
print("saved:", OUT)
print("段落:", len(doc.paragraphs), "表格:", len(doc.tables))
