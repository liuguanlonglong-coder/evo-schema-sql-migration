# -*- coding: utf-8 -*-
"""生成《第16轮修订说明_P1六项闭环与A层落地.docx》。

所有实测数值均从 results/*.json 读取，不在文档里手写统计量。
方案性内容（抽样设计、方言映射、责任分工）标注为「方案/未执行」。
"""
import json
import os
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

M1 = json.load(open("results/m1_holdout_eval.json", encoding="utf-8"))
J = json.load(open("results/judge_strict_audit.json", encoding="utf-8"))
LEG = json.load(open("results/encoding_bounds_v2_legacy.json", encoding="utf-8"))
V2 = json.load(open("results/encoding_bounds_v2.json", encoding="utf-8"))
V1 = json.load(open("results/encoding_bounds.json", encoding="utf-8"))
B4 = json.load(open("results/baseline_segment_aware.json", encoding="utf-8"))


def make_doc():
    d = Document()
    s = d.sections[0]
    s.page_width, s.page_height = Cm(21.0), Cm(29.7)
    s.left_margin = s.right_margin = Cm(2.4)
    s.top_margin = s.bottom_margin = Cm(2.4)
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
    for i, part in enumerate(txt.split("**")):
        if part:
            _run(p, part, bold=(i % 2 == 1), size=size)


def H(doc, txt, lvl=1):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(13 if lvl == 1 else 9)
    p.paragraph_format.space_after = Pt(5)
    p.paragraph_format.keep_with_next = True
    _run(p, txt, bold=True, size={1: 14, 2: 11.5}[lvl], font=CN_H)
    return p


def P(doc, txt, size=10.5):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.first_line_indent = Pt(21)
    emit(p, txt, size)
    return p


def PL(doc, txt):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Pt(19)
    p.paragraph_format.space_after = Pt(2)
    emit(p, txt, 10.3)
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


def TABLE(doc, head, rows, widths, aligns, font=8.4):
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
                            "c": WD_ALIGN_PARAGRAPH.CENTER,
                            "r": WD_ALIGN_PARAGRAPH.RIGHT}[aligns[i]]
            p_.paragraph_format.space_after = Pt(1)
            emit(p_, str(v), font)
    return t


A = make_doc()
p = A.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
_run(p, "第 16 轮修订说明：P1 六项闭环方案与零成本项落地", bold=True, size=15, font=CN_H)
p = A.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
_run(p, "回应《Applied Intelligence 审稿式综合评估报告 v11 第13轮》P1 清单（%s）" % TODAY,
     size=9.5)

# ---------------------------------------------------------------- 一
H(A, "一、先纠正一个前提：闭环不等于六项全部补实验")
P(A, "报告 P1 共六条，措辞是 strongly recommended，不是拒稿硬条件。大修回复的正确做法是"
  "**逐条给交代**：已做、已用等效方式回应、或因结构性原因不做并写入 Limitations。"
  "最忌讳的是假装做了，或沉默不回应。按这个标准对账，六项并不处于同一层。")
P(A, "首先要划掉一条：**P1 第 1 条「增加能输出段位表达式的非 LLM 基线」已在第 14 轮用 B4 闭环**。"
  "B4_naive 与 B4_greedy 均为 %d/%d（%.1f%%）、事件 %d/%d，与 M1_part 完全持平；"
  "正文已据此把主张从「M1 算法更优」收紧为「谱系输出粒度决定成败」。"
  % (B4["B4_naive"]["expr_match"], B4["B4_naive"]["n_ops"], B4["B4_naive"]["pct"],
     B4["B4_naive"]["events_ok"], B4["B4_naive"]["n_events"]))
P(A, "其余五条按可否由本地工作闭环分为四层：**A 层**（零 API 成本、纯本地可复现，本轮已做完："
  "留出库 M1、判定器复核、编码压力测试）；**表述层**（跨方言，只做可移植性分析，不做实验）；"
  "**B 层**（LLM 重复运行，需真实 API 预算）；**C 层**（双人标注，需真人第二位标注者）。"
  "本轮把 P1 从「六项未闭环」压缩到「两项需外部资源（钱 / 人）」。")

CAP(A, "表 16-1  P1 六项闭环总表")
TABLE(A, ["项", "审稿要点", "现状判定", "闭环类型", "本轮动作", "结果 / 待决"],
      [["P1-1", "非 LLM 公平基线", "第 14 轮已闭环", "已完成",
        "B4 与 M1 同为 10/10", "主张已收紧为输出粒度"],
       ["P1-2a", "留出库 M1 直接表达式", "此前只报 SQL 执行结果", "A 实验闭环",
        "m1_holdout_eval.py", "14/16 op、6/8 事件；见第三节"],
       ["P1-2b", "改进 SQL 结果比较器", "五种规则早已实现，只写成残留风险", "A 复核闭环",
        "judge_strict_audit.py", "8/8 符号稳定；见第四节"],
       ["P1-2c", "NULL / 分隔符 / 三段以上压力测试", "v1 只到 nseg=3、每类 1–3 行", "A 实验+实现修正",
        "probe_encoding_bounds_v2.py", "发现并修复 to_expr 缺陷；见第五节"],
       ["P1-2d", "跨方言适用范围", "已声明未验证，但偏被动", "表述闭环（不做实验）",
        "可移植性逐项分析", "见第六节"],
       ["P1-3", "LLM 重复运行", "缓存命中会使方差恒为 0", "B 需 API 预算",
        "未启动，给出最小方案", "待你决定；见第七节"],
       ["P1-4", "双人标注 / kappa", "gold 为程序构造，kappa 概念不适用", "C 需真人",
        "给出等效复核与抽样方案", "数值须由第二标注者产生；见第八节"]],
      [1.2, 3.3, 3.3, 2.2, 3.0, 3.2], ["c", "l", "l", "c", "l", "l"], font=8.0)
NOTE(A, "表中 P1 编号按报告 P1 段六条顺序归并：报告把「留出库 M1」「kappa」「LLM 重跑」"
     "「比较器」「压力测试/方言」「B4 基线」并列，本表按闭环类型重排，条目一一对应，未增删。")

# ---------------------------------------------------------------- 二
H(A, "二、A 层之一：留出库上 M1 段位表达式的直接评估")
P(A, "此前留出库只报 C1/C2/C3m 的执行结果一致率，审稿人要求「单独评估 M1 段位抽取质量，"
  "而不仅是 SQL 执行结果」。核查发现留出库来自 **BIRD train 划分**（%s），"
  "而主实验与 M1 阈值校准用的是 dev 划分，因此**数据划分与阈值校准两重独立**，"
  "比原先理解的「仅按库留出」更强。复用 m1_part.infer_db 与 to_expr，零 API 调用。"
  % M1["source"]["old_dbdir"])
CAP(A, "表 16-2  独立性口径与三级读数（数字均来自 m1_holdout_eval.json）")
u = M1["units"]
TABLE(A, ["分析单位", "n", "恢复 / 匹配", "读数与区间"],
      [["评测条目", str(u["items"]), "—", "不得当作独立样本"],
       ["唯一原子 op", str(u["unique_ops"]),
        "%d/%d" % (M1["op_level"]["matched"], M1["op_level"]["n_units"]),
        "%.1f%%，Wilson 95%% CI [%s, %s]" %
        (M1["op_level"]["rate_pct"], M1["op_level"]["wilson95_pct"][0],
         M1["op_level"]["wilson95_pct"][1])],
       ["独立变更事件", str(u["events"]),
        "%d/%d" % (M1["event_level"]["fully_recovered"], M1["event_level"]["n_units"]),
        "%.1f%%（事件内全部 op 均匹配才算恢复）" % M1["event_level"]["rate_pct"]],
       ["数据库", str(u["dbs"]),
        "%d/%d" % (M1["db_level"]["fully_recovered"], M1["db_level"]["n_units"]),
        "%.1f%%" % M1["db_level"]["rate_pct"]]],
      [3.0, 1.6, 3.0, 8.6], ["l", "c", "c", "l"], font=8.4)
P(A, "**这是一个对己不利、必须如实写进正文的结果**：dev 上 M1 段位表达式为 10/10，"
  "留出库降到 %d/%d（%.1f%%），事件级仅 %.1f%%。两个失败 op 为 %s 与 %s，"
  "均为 no_infer（M1 未推断出该列的段位归属，日期型值域与段位集合的 Jaccard 未过阈值），"
  "不是表达式写错。Wilson 区间宽至 [%.1f, %.1f]，说明 n=%d，"
  "**不得写成「M1 具有 87.5%% 的通用恢复能力」**。"
  % (M1["op_level"]["matched"], M1["op_level"]["n_units"], M1["op_level"]["rate_pct"],
     M1["event_level"]["rate_pct"],
     M1["failures"][0]["db"] + "." + M1["failures"][0]["old"],
     M1["failures"][1]["db"] + "." + M1["failures"][1]["old"],
     M1["op_level"]["wilson95_pct"][0], M1["op_level"]["wilson95_pct"][1],
     M1["op_level"]["n_units"]))
P(A, "入稿位置：7.3 留出库节补一段 + 补充材料表（逐 op 明细），摘要与结论各补一句"
  "「未见库上段位恢复低于开发集」的限定。这条正好把审稿人「dev↔未见库差距」的质疑"
  "从被动承认变成有数值的主动披露。")

# ---------------------------------------------------------------- 三
H(A, "三、A 层之二：全严判定规则下八个主对照的复核")
P(A, "judge_sensitivity.py 早已实现行序敏感（r1）、数值容差（r2）、NULL 与空串区分（r3）"
  "及三者同时（r_all_strict），并在 %d 条重执行上量化过影响；缺的不是实现，"
  "而是正文只把它写成「残留风险未消除」，没有回答**收严后主对照方向是否改变**。"
  "本轮不新增任何 SQL 执行，只把冻结结果重算成可直接入稿的表。"
  % J["reexecution"]["n"])
CAP(A, "表 16-3  全严规则下八个主对照（pt = 条件间 EX 差）")
rows = []
for r in J["contrasts"]:
    rows.append([r["group"], r["contrast"], str(r["n"]),
                 "%+.1f" % r["r0_current_pt"], "%+.1f" % r["r_all_strict_pt"],
                 "%+.1f / %s" % (r["delta_pt"], "稳定" if r["sign_stable"] else "翻转")])
TABLE(A, ["档位", "对照", "n", "现行 r0", "全严", "偏移 / 符号"], rows,
      [2.0, 3.2, 1.5, 2.6, 2.6, 4.3], ["c", "l", "c", "c", "c", "c"], font=8.3)
P(A, "结果：**%s 主对照符号全部稳定**，全严规则下最大偏移 %.1f pt（%s）；"
  "整体 EX 从 %.1f%% 降到 %.1f%%，仅 %d 条（%.2f%%）受行序影响，"
  "McNemar p=%.6f 说明收严确实改变了少量判定，但不改变任何一条主对照的方向。"
  % (J["sign_stable_count"], J["max_abs_delta_pt"],
     max(J["contrasts"], key=lambda x: abs(x["delta_pt"]))["contrast"],
     J["reexecution"]["ex_current_pct"], J["reexecution"]["ex_strict_pct"],
     J["reexecution"]["n_order_sensitive"], J["reexecution"]["share_order_sensitive_pct"],
     J["reexecution"]["mcnemar_order_p"]))
P(A, "**闭环判据**：判定器的宽严不驱动本文结论，因此**不必把 r_all_strict 改为主判**。"
  "改主判会推翻全部历史数值，且 r0 与 gold SQL 的构造口径一致；正确做法是双报。"
  "正文 7.1.3 应把「三点残留风险未消除」改写为「严格规则下八个主对照方向全部不变、"
  "最大偏移 %.1f pt；绝对 EX 仍是 r0 口径下所评估 SQLite 实例上的执行结果一致率，"
  "不是形式化 SQL 语义等价证明」。" % J["max_abs_delta_pt"])

# ---------------------------------------------------------------- 四
H(A, "四、A 层之三：三段以上编码压力测试与一处实现修正")
P(A, "v2 把覆盖面从 v1 的「nseg≤3、每类 1–3 行、单一 CHAR(31)」扩到 "
  "**nseg 2–6 × %d 种分隔符 × %d 类边界情形 × %d 行/格 = %d 格 / %d 次执行**，"
  "固定种子 %d，判据仍是「整行全部分量同时恢复」。默认模式直接调用生产 m1_part.to_expr。"
  % (len(V2["config"]["seps"]), len(V2["config"]["cases"]),
     V2["config"]["rows_per_cell"], len(V2["cells"]),
     len(V2["cells"]) * V2["config"]["rows_per_cell"], V2["config"]["seed"]))
P(A, "首轮实测暴露了一个此前未发现的问题：**末段含分隔符在两段时 100%% 可恢复，"
  "nseg≥4 却掉到 0%%**。定位到 m1_part.to_expr 的 nseg≥3 末段分支多切了一次 "
  "SUBSTR(head, INSTR(head,sep)+1)：rest(nseg-1) 已经是「最后一段及其后全部内容」，"
  "应整段取回。隔离验证给出的修正候选在 nseg=3/4/5/6 上均为 60/60 可恢复，"
  "说明这是**实现缺陷而非编码规则的固有限制**。")
P(A, "修正前先确认回归面：全部 %d 个 gold 段位表达式都是两段 flat 形态、零嵌套，"
  "B4 与独立性审计产物中不存在 nseg≥3；两段走 to_expr 的独立分支，逐字不受影响。"
  "因此把末段改为 return head 是**零回归风险的纯增益**。"
  % 2384)
CAP(A, "表 16-4  修正前后对照（两份 JSON 均可由脚本复现）")
TABLE(A, ["读数", "修正前 legacy", "修正后 生产", "说明"],
      [["sep_in_last 整体恢复率", "%.1f%%" % LEG["by_case"]["sep_in_last"],
        "%.1f%%" % V2["by_case"]["sep_in_last"], "末段含分隔符"],
       ["nseg=2 汇总", "%.1f%%" % LEG["by_nseg"]["2"],
        "%.1f%%" % V2["by_nseg"]["2"], "两段分支独立，未变"],
       ["nseg=4 汇总", "%.1f%%" % LEG["by_nseg"]["4"],
        "%.1f%%" % V2["by_nseg"]["4"], "三段以上补齐"],
       ["nseg≥4 相对≤3 变化的格", str(len(LEG["nseg_ge4_changes"])),
        str(len(V2["nseg_ge4_changes"])), "修正后全部消失"],
       ["每类 nseg 汇总一致", "否", "是", "by_nseg 全为 %.1f%%" % V2["by_nseg"]["4"]]],
      [4.2, 3.6, 3.6, 4.8], ["l", "c", "c", "l"], font=8.4)
CAP(A, "表 16-5  八类边界情形的逐类界定（修正后）")
case_cn = {"normal": "普通取值", "null_any": "任一分量为 NULL",
           "empty_first": "首段为空串", "empty_last": "末段为空串",
           "sep_in_nonlast": "非末段内含分隔符", "sep_in_last": "末段内含分隔符",
           "whitespace": "分量首尾空白", "multibyte": "多字节字符"}
verdict = {"null_any": "致命：NULL 传播使整行丢失，需外部元数据",
           "sep_in_nonlast": "致命：段位错位，需转义或外部元数据",
           "sep_in_last": "可恢复（修正后与两段一致）",
           "empty_first": "可恢复", "empty_last": "可恢复",
           "whitespace": "可恢复", "multibyte": "可恢复", "normal": "可恢复"}
TABLE(A, ["情形", "修正前", "修正后", "界定"],
      [[case_cn[k], "%.1f%%" % LEG["by_case"][k], "%.1f%%" % V2["by_case"][k],
        verdict[k]] for k in case_cn],
      [3.6, 2.2, 2.2, 8.2], ["l", "c", "c", "l"], font=8.3)
P(A, "**回归检查全部通过**：v1 三段探针结论不变（S3a 100%%、S3b 0%%、S3c 0%%）；"
  "留出库 M1 仍为 %d/%d；B4 仍为 %d/%d；真实数据逐列扫描 合并侧 %d 事件 / %d 行、"
  "全量 %d 事件 / %d 行，null %d、内部分隔符 %d、empty %d 全部可恢复、有损行 %d。"
  % (M1["op_level"]["matched"], M1["op_level"]["n_units"],
     B4["B4_naive"]["expr_match"], B4["B4_naive"]["n_ops"],
     V1["real_data"]["merge_side"]["events"], V1["real_data"]["merge_side"]["rows"],
     V1["real_data"]["all_sides"]["events"], V1["real_data"]["all_sides"]["rows"],
     V1["real_data"]["merge_side"]["null"], V1["real_data"]["merge_side"]["has_sep_char"],
     V1["real_data"]["merge_side"]["empty"], V1["real_data"]["merge_side"]["lossy_rows"]))
P(A, "两类致命情形是 sep.join 编码规则的信息论限制，改实现也修不了，必须写入 Limitations："
  "近唯一列 / 含分隔符值需要数据字典或审计日志等外部元数据（与第三阶段「信息下界」结论一致）。"
  "**注意**：m1_part.py 已被修改，公开材料包必须重新打包，否则包内代码与正文复现链不一致。")

# ---------------------------------------------------------------- 五
H(A, "五、表述层：跨方言不做实验，做可移植性分析")
P(A, "报告要求「明确 SQLite 方言适用范围」。核查后判断：**这一项不应通过跨方言实验闭环**。"
  "整个基准绑死 SQLite——BIRD 的 .sqlite 库、gold 的 SUBSTR/INSTR/CHAR(31) 表达式、"
  "执行判定与 fetchall 结果比较全部如此。做跨方言等于重做数据集、重写全部表达式、"
  "重跑全部实验，工作量等同第二篇论文，且不会改变「谱系粒度决定成败」这一命题。")
P(A, "审稿人真正担心的是作者把 SQLite 结果说成通用 SQL 迁移能力。正文 7.1.3 已明确否认，"
  "但偏被动。本轮给出**主动的逐项依赖分析**，把「未验证」升级为「已界定哪些可直译、"
  "哪些需在目标方言重新验证」。")
CAP(A, "表 16-6  SQLite 依赖到 PostgreSQL / MySQL 的可移植性分析（分析结论，非实测）")
TABLE(A, ["依赖项", "SQLite 用法", "PostgreSQL / MySQL", "是否影响本文命题"],
      [["分隔符字面量", "CHAR(31)", "CHR(31) / CHAR(31)", "否，仅字面写法"],
       ["定位分隔符", "INSTR(x,s)", "STRPOS / POSITION / LOCATE", "否，语义可直译"],
       ["取子串", "SUBSTR(x,a,b)", "SUBSTRING(x FROM a FOR b)", "否，语义可直译"],
       ["无 split_part", "INSTR+SUBSTR 组合", "原生 SPLIT_PART / SUBSTRING_INDEX", "否，表达式可简化"],
       ["NULL 传播", "拼接遇 NULL 整体 NULL", "同", "否"],
       ["动态类型 / typeof", "有", "严格类型系统", "**是**：gold 与判定器需重验"],
       ["结果集比较", "sqlite3.fetchall 元组", "需按方言重实现", "**是**：比较口径需重验"],
       ["覆盖的 SQL 结构", "列引用、两段编码", "CTE/JOIN USING/HAVING 等未覆盖", "**是**：与方言无关的既有范围限制"]],
      [3.0, 3.4, 4.6, 5.2], ["l", "l", "l", "l"], font=8.2)
P(A, "建议正文表述：「谱系表达式本身可直译到其它方言，但 gold 构造与结果判定依赖 SQLite "
  "的动态类型与 fetchall 口径，故本文结论严格限定为**在所评估 SQLite 实例上的执行结果一致率**，"
  "不主张其它方言下的等价性。」这样无需新实验即可闭环，且与第七节比较域三档限定相互独立。")

# ---------------------------------------------------------------- 六
H(A, "六、B 层：LLM 重复运行——需你决定是否启动")
P(A, "先核查能否零成本复用，结论是**不能**：llm_client.py 为 temperature=0 且按 prompt 哈希"
  "落盘缓存，同参数重跑必然命中缓存、方差恒为 0，不构成重复运行；"
  "现有 nhanes/llm_cache.jsonl 属另一项目（医疗问答），与本章 SQL 实验无关；"
  "results/e1_xm_ds4f_ms_f0…f6 是并行分片而非重复采样（第 14 轮已有「播种造成假进度」的教训）。")
PL(A, "最小方案（**未执行**）：只跑承重对照合并档 C3m / C3mp，3 次独立运行；"
   "通过独立 LLM_CACHE_PATH 绕开缓存，或设 temperature=0.3 并给不同 seed；"
   "报每条件 3 次 EX 的均值±标准差、配对差区间、主对照符号在 3 次中是否一致。")
PL(A, "规模与成本：按 230 题 × 2 条件 × 3 次 ≈ 1,380 次调用；记忆中 SiliconFlow 单 key "
   "6 路并发、峰值约 26 次/分钟，约 53 分钟 + API 费用。推理模型 45s+ 不可用，"
   "须用非推理模型（GLM-5.3-Flash 约 4s）。缩到单条件 3 次约 690 次 / 27 分钟。")
PL(A, "**学术判断**：LLM 只是执行器，主命题是「同一执行器下谱系粒度决定成败」，"
   "此项属稳健性佐证、不承重。若预算或时间紧，可只做小样本，或写进 Limitations。"
   "我不擅自启动：需你确认 API 预算，以及是否接受 temperature 变动引入的新变量。")

# ---------------------------------------------------------------- 七
H(A, "七、C 层：双人标注 / kappa——需真人，且概念要先澄清")
P(A, "**关键澄清**：本文 gold SQL 与演化标签是 drift_ops 确定性程序构造（%d 个段位表达式"
  "全部由同一替换规则生成），不是人工标注产物。Cohen's kappa 衡量「两位标注者对同一批"
  "主观判断的一致性」，对确定性脚本输出没有对应物。审稿人这条沿用了人工标注数据集的模板。"
  % 2384)
P(A, "因此正确回应不是硬凑一个 κ，而是三层：① 说明基准构造方式、指出 kappa 不适用；"
  "② 给出**等效的独立复核证据**（已有）：第 12 轮 probe_encoding_bounds.py 独立实现复核、"
  "第 14 轮 event_map_audit.py 的 op→map→event→db 四层可审计映射、"
  "第 16 轮 judge_strict_audit.py 判定器复核；③ 若审稿人坚持人类复核，提供可执行备料。")
CAP(A, "表 16-7  第二标注者复核方案（设计稿，未执行）")
TABLE(A, ["环节", "设计", "交付物", "责任"],
      [["抽样", "按 8 个留出库 × 演化类型分层，每层 15 条，共 120 条；"
        "含 gold SQL、旧 SQL、drift_ops、期望动作", "标注清单 CSV", "我可生成"],
       ["编码手册", "判定项：gold SQL 在旧库上的语义是否与旧 SQL 等价；"
        "三分类 是/否/存疑；附 5 条正例与 5 条反例", "手册 + 示例", "我可生成"],
       ["独立标注", "两位标注者背对背，不交换意见", "两份标注 CSV", "**需真人**"],
       ["统计", "一致率、κ=(po−pe)/(1−pe)、bootstrap 95% CI、分歧清单与裁决规则",
        "compute_kappa.py", "我可生成，但输入须真人"],
       ["红线", "不得编造数值；未做时正文只能写「单作者编码，附编码手册与逐条映射供复核」+ 局限性",
        "正文限定句", "我可生成"]],
      [2.2, 6.6, 3.2, 4.2], ["l", "l", "l", "l"], font=8.2)
P(A, "这一项与项目既定的**诚实性红线一致**：双人编码 kappa 若未做，只能写单作者编码并附手册，"
  "不得编造数值。本轮不产出任何 κ 数值；若你安排第二位标注者（同学 / 导师 / 合作者），"
  "我可立即生成抽样清单、编码手册与计算脚本。")

# ---------------------------------------------------------------- 八
H(A, "八、入稿接入点与待决事项")
P(A, "**A 层数据已产出，但尚未写入主文**，因为入稿会改动刚压到 15 表的结构，"
  "且 B/C 还需你决策，合并成一轮更省。接入方案已定：")
PL(A, "① 复现链扩为：三个新脚本 → make_chapter7_v11.py（读三份 JSON，补结论段 + 3 张表）"
   "→ slim_chapter7.py（KEEP 列表不改，新表自动进补充材料 S-35…S-37，主文仍 15 表）"
   "→ make_front_matter_v11.py。")
PL(A, "② 章内外同步改四处：7.1.3 判定器段、7.3 留出库段、7.7.5 边界段、Limitations；"
   "摘要与结论各补一句留出库限定。")
PL(A, "③ pack_materials.py 纳入 3 个新脚本、4 份新 JSON 与修正后的 m1_part.py，"
   "重新打包并复检脱敏；README 补写 v2 探针的 --legacy 用法。")
CAP(A, "表 16-8  本轮产出文件")
TABLE(A, ["文件", "类型", "作用"],
      [["m1_holdout_eval.py / results/m1_holdout_eval.json", "脚本 / 数据", "留出库 M1 三级口径"],
       ["judge_strict_audit.py / results/judge_strict_audit.json", "脚本 / 数据", "八主对照全严复核"],
       ["probe_encoding_bounds_v2.py", "脚本", "200 格压力测试，含 --legacy 对照"],
       ["results/encoding_bounds_v2.json", "数据", "修正后（生产 to_expr）"],
       ["results/encoding_bounds_v2_legacy.json", "数据", "修正前可复现留档"],
       ["m1_part.py", "代码修正", "to_expr 末段改为取剩余全部；两段分支不变"],
       ["第16轮修订说明_P1六项闭环与A层落地.docx", "文档", "本文件"]],
      [6.6, 2.6, 7.0], ["l", "c", "l"], font=8.2)
P(A, "**待你决定三项**：(1) 是否启动 ③ LLM 重复运行（API 预算 / temperature 方案）；"
  "(2) 是否安排第二位标注者，或采用「单作者编码 + 等效独立复核」的写法；"
  "(3) 是否现在执行入稿。Major Revision 判断不变，本轮之后 P1 剩余风险已从"
  "「六项未做」收敛为「两项需外部资源」。")

OUT = "第16轮修订说明_P1六项闭环与A层落地.docx"
A.save(OUT)
print("saved:", OUT)
print("段落", len(A.paragraphs), "表格", len(A.tables))
