# -*- coding: utf-8 -*-
"""生成《第7章v9修订说明_回应第八轮评审.docx》。

逐条对应第八轮评审的五个要求与其余观察，每条给出：改了什么、落在哪、
以及哪些是「改不了只能限定」的。
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
OUT = "第7章v9修订说明_回应第八轮评审.docx"

EN, CN, CN_H = "Times New Roman", "宋体", "黑体"

RVR = json.load(open("results/rule_vs_llm.json", encoding="utf-8"))
G, D, V12 = RVR["groups"], RVR["decomposition"], RVR["v1_vs_v2"]
RWA = json.load(open("results/rule_rewrite.json", encoding="utf-8"))["audit"]
PSX = json.load(open("results/probe_star_expand.json", encoding="utf-8"))
HLD = json.load(open("results/holdout_stats.json", encoding="utf-8"))
EVE = json.load(open("results/evoschema_events.json", encoding="utf-8"))
DBD = json.load(open("results/data_bounds.json", encoding="utf-8"))
S = json.load(open("results/stats_ch7_v5.json", encoding="utf-8"))["main"]
assert RVR["version"] == "rule-vs-llm-v2"
assert PSX["version"] == "star-probe-v1"

_ev = HLD["event_level"]["merge_col_holdout"]
_rhm, _rsh = G["merge_holdout"]["rates"], G["split_holdout"]["rates"]
_rm, _rs = G["merge"]["rates"], G["split"]["rates"]
_pm = G["merge"]["paired"]["c3r_infer-c3mp"]
_ps = G["split"]["paired"]["c3r_infer-c3mp"]
_phm = G["merge_holdout"]["paired"]["c3r_infer-c3mp"]
_psh = G["split_holdout"]["paired"]["c3r_infer-c3mp"]
_fhd = RVR["failure_profile"]["merge_col_holdout"]["c3r_infer"]

doc = Document()
sec = doc.sections[0]
sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
sec.left_margin = sec.right_margin = Cm(2.4)
sec.top_margin = sec.bottom_margin = Cm(2.4)
st = doc.styles["Normal"]
st.font.name = EN
st.font.size = Pt(10.5)
st._element.rPr.rFonts.set(qn("w:eastAsia"), CN)
st.paragraph_format.line_spacing = 1.34
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
    p.paragraph_format.space_before = Pt(15 if lvl == 1 else 10)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.keep_with_next = True
    _run(p, txt, bold=True, size={1: 15, 2: 12, 3: 11}[lvl], font=CN_H)
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
    p.paragraph_format.space_before = Pt(9)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.keep_with_next = True
    _run(p, txt, bold=True, size=9)
    return p


def NOTE(txt):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(8)
    emit(p, txt, 8.5)
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
            _run(p_, str(v), size=font)
    return t


def _m(v, nd=1):
    return f"{v:+.{nd}f}".replace("-", "−")


# ===================================================================== 题头
p_ = doc.add_paragraph()
p_.alignment = WD_ALIGN_PARAGRAPH.CENTER
_run(p_, "第 7 章 v9 修订说明", bold=True, size=17, font=CN_H)
p_ = doc.add_paragraph()
p_.alignment = WD_ALIGN_PARAGRAPH.CENTER
_run(p_, "回应第八轮评审：确定性基线的公平性、因果表述的边界、章内外口径统一",
     size=10.5, font=CN)
NOTE("对应文件：《投稿版_第7章_实验结果_v9.docx》（257 段、44 表、表号连续）与"
     "《摘要_标题_贡献点_结论_v9.docx》。本说明逐条回答五项要求，"
     "并单列本轮自查发现、评审未提及但同样会致命的一处问题（第二节）。")

# ===================================================================== 一
H("一、结论摘要：这一轮不是补实验，是修一个 bug 并收窄三处表述")
P("评审要求的五件事里，第一件最重。按它去核查时发现，问题比评审设想的更严重："
  "上一版判为「星号投影不可改写」的 21 条，**全部是 COUNT(*)**，没有一条是投影星号。"
  "也就是说，v8 里「规则基线的唯一系统失败是星号投影」「LLM 在合并档补上 6/6」"
  "「并联上界 230/230」这三句话，建立在一个实现缺陷之上，必须整段撤销。")
TABLE(["评审要求", "本版处置", "落在"],
      [["① 支持 SELECT * 展开的规则基线",
        "不仅补上展开能力，还修了 COUNT(*) 误判；规则侧 dev 由 97.4/95.2 升至 100.0；"
        "撤销 v8 的互补性结论", "7.4.2、7.4.4、表 7-25、表 7-26"],
       ["② 限定「证据效应／执行器效应」为当前构造下的配对差值",
        "写明三条限定，并把幅度随档位的波动全部披露（+99.1 / +64.4 / 0 / 0）",
        "7.4.3、表 7-23"],
       ["③ 摘要·贡献·结论写明两段可逆编码与未覆盖边界",
        "章外三处给出可粘贴文字；章内 7.10 未完成项由 5 条增至 6 条",
        "另册文件；7.10 ⑥"],
       ["④ 保留 6 正 1 负与映射级分析，不用条目级 p 替代",
        "regional_sales 从「不能略过」升格为主结果段，给出净差；条目级 p 不再出现在结论层",
        "7.7.4、7.11 ④"],
       ["⑤ 表 7-37 前后文字与表内行数统一",
        "「两项近邻工作」改为「三项」，并在 7.1 与 7.9 两处同步点名",
        "7.1、7.9"]],
      [4.4, 7.6, 3.6], ["l", "l", "l"])

# ===================================================================== 二
H("二、第 1 条：SELECT * 展开基线 —— 实查结果比评审预期更严重")
H("2.1  问题不是「不够强」，而是「判错」", 2)
P("v1 的判据是「AST 中是否出现星号节点」。SQLite 的 COUNT(*) 在 sqlglot 中同样表现为"
  "一个 Star 节点，因此被判为不可改写。审计结果如下：")
TABLE(["档位", "被判为星号投影（条）", "其中含 COUNT(*)", "真正的投影星号", "投影星号展开次数"],
      [[k, RWA["per_evo"][k]["projection_star"] + RWA["per_evo"][k]["count_star"],
        RWA["per_evo"][k]["count_star"], RWA["per_evo"][k]["projection_star"],
        RWA["per_evo"][k]["expanded"]]
       for k in ("merge_col", "split_col", "merge_col_holdout", "split_col_holdout")]
      + [["合计", RWA["projection_star_items"] + RWA["count_star_items"],
          RWA["count_star_items"], RWA["projection_star_items"], RWA["star_expanded"]]],
      [3.6, 4.0, 3.2, 3.2, 3.4], ["l", "r", "r", "r", "r"])
NOTE("判失败的 21 条与含 COUNT(*) 的 21 条完全重合；投影星号 0 条。"
     "COUNT(*) 在列替换下必须保持不变，把它判为失败等于把可以做对的题主动判错。")
P("本版把判据改为「只看 SELECT 投影位上的 * 或 t.*」，并补上真正的展开通道："
  "以原始库（演化前）为据、按 cid 顺序取旧表列清单，把 * 或 t.* 展开为逐列引用，"
  "再交给统一的替换通道。两种模式并存（expand 展开、keep 保留），"
  "在本样本上数值完全相同——因为投影星号出现 0 次。")

H("2.2  修正后的数值：规则侧在 dev 上达到 100%", 2)
TABLE(["档位", "规则·oracle", "规则·自动谱系", "规则·列级", "v8 报告值（oracle／自动／列级）"],
      [["合并（230）", f"{_rm['c3r_gt']['rate']:.1f}", f"{_rm['c3r_infer']['rate']:.1f}",
        f"{_rm['c3r_name']['rate']:.1f}", "97.4 ／ 97.4 ／ 0.9"],
       ["拆分（104）", f"{_rs['c3r_gt']['rate']:.1f}", f"{_rs['c3r_infer']['rate']:.1f}",
        f"{_rs['c3r_name']['rate']:.1f}", "95.2 ／ 95.2 ／ 95.2"],
       ["合并·留出（101）", f"{_rhm['c3r_gt']['rate']:.1f}", f"{_rhm['c3r_infer']['rate']:.1f}",
        f"{_rhm['c3r_name']['rate']:.1f}", "95.0 ／ 79.2 ／ 15.8"],
       ["拆分·留出（64）", f"{_rsh['c3r_gt']['rate']:.1f}", f"{_rsh['c3r_infer']['rate']:.1f}",
        f"{_rsh['c3r_name']['rate']:.1f}", "92.2 ／ 92.2 ／ 92.2"]],
      [3.4, 2.4, 2.8, 2.4, 5.0], ["l", "r", "r", "r", "c"])
NOTE("三处应当注意。第一，oracle 映射在四个档位上都达到 100.0%，"
     "这不是方法性能，而是构造决定的：gold 的 new_sql 由同一条替换规则机械生成。"
     "第二，自动谱系在 dev 上也达到 100.0%，说明 M1 在这 6 个库上已把映射抽全——"
     "但这是样本内结果。第三，唯一有缺口的仍是留出库合并档（80.2%），"
     f"缺口 {_rhm['c3r_gt']['rate'] - _rhm['c3r_infer']['rate']:.1f}pt 全部为执行报错"
     f"（{_fhd.get('DB_ERROR', 0)} 条），性质是漏抽映射而非替换逻辑错误。")

H("2.3  连带撤销：v8 的互补性结论作废", 2)
P("修正后规则侧在 dev 上没有任何失败，于是 v8 那段「规则唯一失败是星号投影、"
  "LLM 在合并档补上 6/6、拆分档补上 2/5」的互补结构不存在了。"
  f"现在的读数是：合并档仅 LLM 对 {_pm['b_only']} 条、规则对而 LLM 错 {_pm['a_only']} 条；"
  f"拆分档 {_ps['b_only']} 条与 {_ps['a_only']} 条；并联上界等于单用规则。"
  f"唯一出现「仅 LLM 对」的是留出库合并档（{_phm['b_only']} 条），"
  f"来源是规则侧 {_fhd.get('DB_ERROR', 0)} 条漏抽映射导致的执行报错——"
  "模型在这些条目上并非更会改写，而是不会因为缺一条映射就整体失败。"
  "这个区别写进了正文，因为它把「失败模式不同」与「能力更强」分开了。")
P(f"执行器效应随之由 +17.8pt 修正为 {_m(_rm['c3r_infer']['rate'] - _rm['c3mp']['rate'])}pt，"
  f"拆分档为 {_m(_rs['c3r_infer']['rate'] - _rs['c3mp']['rate'])}pt。"
  "所有引用旧值的地方（7.4、7.10 表 7-32、7.11 表 7-33、投稿信口径）已一并替换。")

H("2.4  展开能力必须被验证，否则等于没写", 2)
P("样本里投影星号 0 条，意味着展开路径从未在真实数据上执行过。"
  "为免「写了但没跑过」的质疑，另建一个两步演化的合成库做探针："
  "旧表 Patient(pid, name, sex, admit) 经 name+sex 合并、admit 改名，"
  "以「旧库上的执行结果」为语义基准，比较两种模式。")
TABLE(["探针用例", "expand 重建旧结果", "keep 重建旧结果", "expand 与构造 gold 一致"],
      [["裸 *（单表）", "是", "否", "否"],
       ["T.* （带表限定）", "是", "否", "否"],
       ["裸 *（两表 JOIN）", "是", "否", "否"],
       ["COUNT(*)", "是", "是", "是"],
       ["COUNT(*) + 被演化列谓词", "是", "是", "是"]],
      [6.0, 3.0, 3.0, 3.6], ["l", "c", "c", "c"])
NOTE("前五行回答「展开是否实现正确」，第六列回答「本构造能否评价含投影星号的题」——"
     "答案为否：gold 只替换列引用、不触碰 *，故 gold 在新库上返回新列集合，"
     "与旧语义不等价；展开模式给出语义正确的旧列集合，反而与 gold 不一致。"
     "这两种读法都不能当正确答案，因此**含投影星号的题目在本构造下无法被忠实评价**。"
     "样本中此类题目恰为 0 条，故不影响任何已报告数字，但它写入 7.10 第 ⑥ 条作为构造限制。")

# ===================================================================== 三
H("三、第 2 条：把两个「效应」限定为构造内的配对差值")
P("评审要求不要把证据效应与执行器效应扩展成普遍因果结论。本版做了三件事。")
PL("**其一，改掉了一个量级判断。** 上一版写「两者相差一个量级」，"
   f"但修正后合并档是 {_m(D['merge']['evidence_effect_pt'])}pt 对 "
   f"{_m(D['merge']['executor_effect_pt'])}pt，约为 4.9 倍，不构成量级差异。"
   "正文改为「前者约为后者的 4.9 倍」，并删去「量级显著」一类措辞。")
PL(f"**其二，把幅度波动全部披露。** 证据效应在四档分别为 "
   f"{_m(D['merge']['evidence_effect_pt'])} / {_m(D['merge_holdout']['evidence_effect_pt'])} / "
   f"{_m(D['split']['evidence_effect_pt'])} / {_m(D['split_holdout']['evidence_effect_pt'])}，"
   f"执行器效应为 {_m(D['merge']['executor_effect_pt'])} / "
   f"{_m(D['merge_holdout']['executor_effect_pt'])} / "
   f"{_m(D['split']['executor_effect_pt'])} / {_m(D['split_holdout']['executor_effect_pt'])}。"
   "两个因子都不是常数：执行器效应在留出库合并档降到 +10.9，"
   "原因是自动谱系先丢了 19.8pt，规则侧失分让两者差距缩小。"
   "正文据此写成「只主张在合并方向、本构造下，表示差异比执行器选择大数倍」。")
PL("**其三，写明作用域。** 7.4.3 末尾新增一段，列出三条不可外推的理由："
   "gold 由同一规则机械生成故规则上界被抬高；LLM 侧只测一个基座与一套提示模板；"
   "样本以演化事件为独立单元高度聚集。"
   "结论句改为「本节全部结论的作用域是：本构造、本实现、本样本」。")

# ===================================================================== 四
H("四、第 3 条：摘要、贡献、结论的作用域统一（章外另册）")
P("这一条不能在第 7 章里解决——摘要与贡献点不在本章。"
  "因此另出一份可粘贴文字：《摘要_标题_贡献点_结论_v9.docx》，含英文摘要（240 词）、"
  "中文摘要（576 字）、三个标题候选、5 个关键词、三条贡献点、结论段，"
  "以及一张「数字 → 出处」自检表。三处实质修改是：")
TABLE(["位置", "改法"],
      [["摘要", "删除「RAG 驱动」「语言模型能更好地迁移」两类表述；"
                "新增独立一段写作用域：两段可逆拼接、21.5 万行扫描未触发空值与值内分隔符、不作检索主张"],
       ["标题与关键词", "删去 RAG 与「端到端 NL2SQL」字样，把「谱系粒度」与「确定性重写基线」提到主标题；"
                        "推荐案：列合并与列拆分之后：历史 SQL 迁移中的谱系粒度与确定性重写基线"],
       ["贡献点", "三条重写：粒度命题（含拆分方向为零的反向边界）、误差来源定位、"
                  "确定性重写为强基线（负面结论，明写）。另加一条作用域声明"]],
      [3.0, 12.6], ["l", "l"])
NOTE("为什么把贡献三写成「这条基线更强」而不是藏进局限节："
     "评审已明确表示，坦诚披露这一条是可信度来源；藏起来会被读成回避，"
     "而写在贡献里则把「我们测出了自己的上界在哪」变成一项工作。")

# ===================================================================== 五
H("五、第 4 条：6 正 1 负与映射级分析保持在事件单位")
P("这一条本版没有改数值，只改位置与权重。三个动作：")
PL(f"**① regional_sales 升格为主结果。** 原文写法是「这个反例不能略过」，"
   f"属于承认但降格。现在改为「必须作为主结果报告，而不是归入局限节」，"
   f"并给出该库的净差（C3mp − C3m 为负）与解释（合并列取值较短、"
   f"C3m 在该库已达 {HLD['per_db']['merge_col_holdout|regional_sales']['c3m']['rate']:.1f}%）。"
   "同时点明它界定的是主张形式：「总体与多数库的方向」，不是逐库规律。")
PL(f"**② 事件单位不变。** 留出库合并档仍报 {_ev['positive']} 正 {_ev['negative']} 负、"
   f"单元级 p = {_ev['p']:.3f}，未用条目级 p 替代。"
   "EvoSchema 对照仍按映射单元报（合并档 3 个映射、最大占 92.7%、单元级 p = 0.5），"
   "条目级 p = 0.0018 保留但明确标注为高估。")
PL("**③ 结论层不再出现条目级 p。** 7.10 与 7.11 的所有显著性表述已核过一遍，"
   "事件级与库级并列给出，条目级 p 只出现在表格与 7.8 的单元分解里。")

# ===================================================================== 六
H("六、第 5 条：表 7-37 前后的数量对不上")
P("确认属实。正文写「必须与两项近邻工作正面划界」，而表内为三行"
  "（EvoSchema、How Far Can They Map?、PRISM/PRISM++）。"
  "本版把两处都改为三项，并在 7.9 的引导句里直接点名三项，"
  "使读者不必先看表才知道「两项」指谁。7.1 节提到该表的句子同步改为三项。")
NOTE("这类错误的风险不在数字，而在信号：审稿人看到「两项」与三行表格并列，"
     "会怀疑其余计数也未经核对。已顺带核过全章其他数量词（「五处差别」「三条判断」"
     "「六条已闭环」），与对应段落数目一致。")

# ===================================================================== 七
H("七、评审未点名但本轮自查出的两处")
H("7.1  M1 相对值域匹配方法的新颖性（评审第 6 节末提出）", 2)
P("评审要求把 M1 相对已有值域匹配方法的贡献讲清楚。本版在 7.3 的基线比较后新增一段，"
  "给出三条分界，并明确写出一句否定："
  "「本章不主张 M1 在列匹配精度上优于实例级匹配方法——这个比较本章没有做，"
  "也不该用 B1–B3 的结果代替」。三条分界是：")
PL("输出形态不同：那一族方法输出列对，合并场景输出「新列 + 段位」的表达式，"
   "因为两个旧列映射到同一个新列，列对形式信息不足——问题不等价，不是精度差异。")
PL("场景不同：那一族多做跨库匹配，本章是同库前后版本，可用「算子只作用于少数列」的先验收缩空间，"
   "这也是合并档零误报的原因。")
PL("证据用途不同：把值域与形态证据用于「恢复段位」而非「匹配列名」，是本章的分界点。")
P("必须说明一处待补：这一族方法需要有具体文献落点，本版只写到「方法类别」层面，"
  "没有编造具体引文。**建议在投稿前补 2–3 篇实例级 schema 匹配的原始文献**"
  "（值域重叠／取值分布一类），插在 7.3 该段末尾。我不代拟引文，因为编造的 DOI 一旦被查到，"
  "后果远超这一段的收益。")

H("7.2  边界探针的定位（评审第 2 节末提出）", 2)
P("评审指出边界探针每类只有 1–3 个例子，应作机制反例而非现实失败比例的估计。"
  "本版在 7.10 ③ 中补了一句限定："
  "「探针每类只有 1–3 个例子，只作机制反例，不能用来估计真实数据中的失败比例」，"
  "并把主实验的适用范围写成条件句——"
  "「只适用于两分量非空且不含分隔符这一前提成立的取值分布」。")
P("留出库合并档的负向、探针的致命情形、真实数据的 0 行越界，这三者构成一组互补证据："
  "真实数据没触发边界（0 行越界），不等于数据干净，只说明构造没有压力测试这类分布。"
  "这句话现在同时出现在 7.4.4、7.7.5 与 7.10，避免读者只看到「0 行越界」而误读为「数据无边界问题」。")

# ===================================================================== 八
H("八、变与不变：一张表看清哪些数字动了")
TABLE(["项目", "v8", "v9", "原因"],
      [["规则·oracle（合并／拆分／留出合并／留出拆分）", "97.4／95.2／95.0／92.2",
        "100.0／100.0／100.0／100.0", "COUNT(*) 误判修复"],
       ["规则·自动谱系（同上）", "97.4／95.2／79.2／92.2",
        "100.0／100.0／80.2／100.0", "同上（留出合并 +1.0 后仍有 20 条漏抽映射）"],
       ["规则·列级（同上）", "0.9／95.2／15.8／92.2", "0.9／100.0／15.8／100.0",
        "同上；合并档列级仍为 0.9"],
       ["执行器效应（合并）", "+17.8pt", _m(_rm['c3r_infer']['rate'] - _rm['c3mp']['rate']) + "pt",
        "规则侧修正后的直接结果"],
       ["证据效应（合并）", "+96.5pt", _m(D['merge']['evidence_effect_pt']) + "pt", "同上"],
       ["互补性结论", "规则失败 6+5 条全为星号投影，LLM 补 6/6 与 2/5", "撤销",
        "失败本不存在，属误判"],
       ["并联上界", "230/230 与 101/104", "四个档位分别 100%／100%／84.2%／100%",
        "重算"],
       ["LLM 侧全部数值（C1–C3mp）", "不变", "不变", "本轮未重新调用模型"],
       ["EvoSchema 单元级结果", "不变", "不变", "未重算"],
       ["留出库事件 6 正 1 负", "不变", "不变（位置改为主结果）", "未重算"]],
      [5.4, 4.2, 4.4, 3.6], ["l", "c", "c", "l"])
NOTE("LLM 侧一个数都没动，说明本轮改动集中在规则基线与表述层，"
     "不存在「为了配合新结论而重算旧结果」的情况。这一点应在回复信中明写。")

# ===================================================================== 九
H("九、仍然不闭环的六项，以及我对这份稿子的判断")
P("六项与 7.10 一致，按对结论的影响从大到小排列：")
PL("① **需要重写谓词结构而非替换标识符的迁移未测。** 这是规则基线与 LLM 之间唯一可能"
   "反转的比较域，也是当前稿件最大的空白。补它需要构造口径改变（单位、编码、区间）的演化，"
   "工作量大，但它决定「LLM 是否在某一类迁移上不可替代」这个问题能否被回答。")
PL("② 三段以上合并与分隔符嵌套：to_expr 支持任意段数，但未在真实库验证。")
PL("③ 无人工独立复核：单作者编码，未做双人编码与 kappa，只附编码手册供复核。")
PL("④ 检索实验未做：真检索实验仅 40 题且为阴性对照。")
PL("⑤ 含投影星号的题目在本构造下无法忠实评价（样本中 0 条）。")
PL("⑥ 探针每类 1–3 例，只能作机制反例。")
P("我的判断是：v9 解决的问题比 v8 更关键——它把「规则基线被低估」这个可能被用来"
  "否定整章结论的漏洞补上了，代价是承认 LLM 在本章场景里没有增量。"
  "这个交换是划算的：现在的稿件可以被称为一篇诚实的实证研究，"
  "但它也因此不再是「以 LLM 或 RAG 为主要贡献」的论文。"
  "若标题与贡献点按第四节收窄，我认为值得继续投稿准备；"
  "若坚持保留 RAG 作为核心创新，本轮所有修正都不足以挽回，因为问题不在证据而在定位。")

doc.save(OUT)
print("saved:", OUT)
print("段落:", len(doc.paragraphs), "表格:", len(doc.tables))
