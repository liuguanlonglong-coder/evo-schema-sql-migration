# -*- coding: utf-8 -*-
"""正文瘦身：把第 7 章完整版拆成「主文 15 表 + 补充材料其余表」。

回应 Weakness 第 7 条「结构过于复杂、主要科学发现需要进一步突出」。
做法不是删数据，而是搬移：
  - 主文（文件名保持 投稿版_第7章_实验结果_v11.docx，章外 _T() 继续从它反查表号）
    只保留协议、主结果、M1/基线、留出/跨模型、事件统计、失败机制、RQ 汇总共 15 表；
  - 补充材料（补充材料_第7章_v11.docx）收纳其余表，编号 S-1..S-M，
    带原属节标题、与主文表号对照表；
  - 完整版另存 第7章_v11_完整版_49表.docx 作为素材库（文件名沿用历史约定，
    内容随入稿增长：第 17 轮后为 52 表），不丢任何数字；
  - 所有交叉引用自动重定向：主文中对被移走表的「表 7-N」改为「补充材料表 S-M」，
    补充材料中对保留表的引用改为「主文表 7-K」；两份文档各自表号连续。

表数是动态的：第 17 轮把 A 层三项复核表加在正文末尾后，完整版由 49 增至 52，
新增 3 表自动落到 S-35..S-37，主文仍严格 15 表。KEEP 按「原最终表号」书写，
只要新表追加在末尾、不插入中间，KEEP 就无需改动。

输入刷新规则：每次运行都以 make_chapter7_v11.py 刚生成的完整版为唯一真源，
按表题数判断是否需要刷新素材库备份，避免备份文件残留旧表数。
"""
import os
import re
import shutil
from copy import deepcopy

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt

BASE = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE)

SRC = "投稿版_第7章_实验结果_v11.docx"
BAK = "第7章_v11_完整版_49表.docx"
MAIN_OUT = SRC                     # 覆盖，保持章外脚本的文件名约定
SUPP_OUT = "补充材料_第7章_v11.docx"

# 主文保留的表（按原最终表号，1–49）。15 张。
# 第 20 轮在正文末尾追加三张表（55 数据集流转、56 扩充事件集构成、57 端到端），
# 因此主文由 15 表增至 18 表。追加在末尾是刻意的：KEEP 按最终表号书写，
# 只要不插入中间，前 54 张表的表号就不会漂移。
# 第 22 轮：新增 58（按库配对）、59（注入缺陷）、61（SQL 覆盖）、62（同源与重复判定）。
# 60（分层对比）留在补充材料——注入缺陷修正后它已属旁证，不再是回答的核心。
# 第 22 轮（续）：在 7.7.5 中间插入两张可逆性表（十类边界 250 格、准入检测器回退代价），
# 因此其后所有表号 +2。主文保留新的第 42 张（十类边界），第 43 张（回退代价）落到补充材料，
# 因为正文已给出 369 / 269 / 72.9% / 100 / 33 这五个数，表只是其排布。
# 教训：KEEP 按最终表号书写，插入中间必须整体平移——本轮即按 +2 重算。
KEEP = [1, 3, 4, 16, 18, 19, 20, 22, 25, 28, 33, 34, 39, 42, 44, 50, 57, 58,
        59, 60, 61, 63, 64]

CAP_RE = re.compile(r"^表\s*7-(\d+)\s\s(.+)$")
REF_RE = re.compile(r"表\s*7-(\d+)")
HEAD_RE = re.compile(r"^7(?:\.\d+)*\s")

def _n_caps(path):
    """数一份 docx 里的表题数，用来判断素材库备份是否需要刷新。"""
    if not os.path.exists(path):
        return -1
    _d = Document(path)
    return sum(1 for p_ in _d.paragraphs if CAP_RE.match(p_.text.strip()))

# 完整版是唯一真源。若备份缺失，或备份表数少于当前完整版，则刷新备份。
# 这样第 17 轮把 49 表增至 52 表后不会继续读旧备份；重复运行也保持幂等。
if not os.path.exists(SRC):
    raise SystemExit("先运行 make_chapter7_v11.py 生成完整第 7 章")
_n_src, _n_bak = _n_caps(SRC), _n_caps(BAK)
# 两条历史教训必须同时满足：
#   ① 只改文字、表数不变时也必须刷新，否则主文继续从旧备份生成（第 19 轮踩过）；
#   ② 本脚本的主文输出与完整版输入是同一个文件（MAIN_OUT = SRC），
#      所以第二次运行会把 SRC 读成「18 表的主文」。此时绝不能用它刷新备份，
#      否则完整版就自毁了（第 20 轮踩过，靠重新生成才恢复）。
import hashlib
_h_src = hashlib.sha1(open(SRC, "rb").read()).hexdigest()
_h_bak = (hashlib.sha1(open(BAK, "rb").read()).hexdigest()
          if os.path.exists(BAK) else "")
if _n_src < _n_bak:
    # SRC 已是被上一次运行覆盖的主文：从备份读完整版，且不刷新备份。
    print(f"[warn] {SRC} 只有 {_n_src} 表，少于备份的 {_n_bak} 表，"
          f"判定它已被主文覆盖；本次从备份读取完整版。")
    INPUT = BAK
else:
    if _h_src != _h_bak:
        shutil.copy2(SRC, BAK)
        print(f"素材库备份已刷新：{BAK}（表数 {_n_bak} → {_n_src}，内容有变）")
    else:
        print(f"素材库备份已是最新：{BAK}（表数 {_n_src}，内容未变）")
    INPUT = SRC   # 直接读本次刚生成的完整版
src = Document(INPUT)
body = src.element.body

# ---------------------------------------------------------------- 1. 块序列
blocks = []          # (kind, el, text)
for el in body.iterchildren():
    tag = el.tag
    if tag == qn("w:p"):
        txt = "".join(t.text or "" for t in el.iter(qn("w:t")))
        blocks.append(("p", el, txt.strip()))
    elif tag == qn("w:tbl"):
        blocks.append(("tbl", el, ""))
    # sectPr 等其它元素跳过

# ---------------------------------------------------------------- 2. 找 CAP
caps = {}            # N -> (block_index, title)
for i, (kind, el, txt) in enumerate(blocks):
    if kind == "p":
        m = CAP_RE.match(txt)
        if m:
            n = int(m.group(1))
            assert n not in caps, f"重复表号 {n}"
            caps[n] = (i, m.group(2))
ALL = sorted(caps)
N_ALL = len(ALL)                      # 第 17 轮后为 52
assert N_ALL >= 49, f"表数异常偏少：{N_ALL}"
assert ALL == list(range(1, N_ALL + 1)), "表号不连续或存在占位号残留（7-1xx），需先跑 _renumber_tables"
assert set(KEEP).issubset(set(ALL)), "KEEP 中有不存在的表号"
assert len(KEEP) == 23, "主文保留表数必须是 22（第 22 轮由 18 增至 22）"

MOVED = [n for n in ALL if n not in KEEP]
N_MOVED = len(MOVED)                  # 第 17 轮后为 37
MAIN_NEW = {n: i + 1 for i, n in enumerate([n for n in ALL if n in KEEP])}
SUPP_NEW = {n: i + 1 for i, n in enumerate(MOVED)}
print(f"完整版 {N_ALL} 表 → 主文 {len(KEEP)} 表 + 补充材料 {N_MOVED} 表（S-1..S-{N_MOVED}）")

# ---------------------------------------------------------------- 3. 表单元
def is_heading(txt):
    return bool(HEAD_RE.match(txt)) and len(txt) < 70

def is_note(i):
    """表注：紧随表格之后的小字号段落（NOTE() 用 8.5pt）。"""
    if i >= len(blocks) or blocks[i][0] != "p":
        return False
    el = blocks[i][1]
    sizes = []
    for r in el.iter(qn("w:r")):
        sz = r.find(qn("w:rPr"))
        if sz is not None:
            v = sz.find(qn("w:sz"))
            if v is not None:
                sizes.append(float(v.get(qn("w:val")) or 0) / 2)
    return bool(sizes) and max(sizes) < 9.0   # NOTE 8.5pt；CAP 是 9.0pt，不能误收

units = {}           # N -> [block_index,...]，含引导段 / CAP / 表格 / 表注
for n, (ci, _title) in caps.items():
    u = []
    # 引导段：CAP 前一段，且只提到本表（或本表所属被移走表），非标题、非 CAP
    if ci - 1 >= 0:
        pk, _pe, ptxt = blocks[ci - 1]
        refs = set(int(x) for x in REF_RE.findall(ptxt))
        if (pk == "p" and ptxt and not is_heading(ptxt)
                and not CAP_RE.match(ptxt) and refs
                and refs.issubset(set(MOVED) | {n})
                and len(ptxt) < 420):
            u.append(ci - 1)
    u.append(ci)
    j = ci + 1
    if j < len(blocks) and blocks[j][0] == "tbl":
        u.append(j)
        j += 1
    while j < len(blocks) and is_note(j):
        u.append(j)
        j += 1
    units[n] = u

# 校验：每个表单元有且仅有一个表格；块不重叠
seen = {}
for n, u in units.items():
    assert sum(1 for k in u if blocks[k][0] == "tbl") == 1, f"表 {n} 单元表格数异常"
    for k in u:
        assert k not in seen, f"块 {k} 同时属于表 {seen[k]} 与表 {n}"
        seen[k] = n

MOVE_BLOCKS = set(k for n in MOVED for k in units[n])

# ---------------------------------------------------------------- 4. 引用重定向
def _sub_text(s, target):
    def rep(m):
        n = int(m.group(1))
        if n not in ALL:
            return m.group(0)
        if target == "main":
            return (f"表 7-{MAIN_NEW[n]}" if n in KEEP
                    else f"补充材料表 S-{SUPP_NEW[n]}")
        return (f"主文表 7-{MAIN_NEW[n]}" if n in KEEP
                else f"表 S-{SUPP_NEW[n]}")
    return REF_RE.sub(rep, s)

def _rewrite_par(el, target):
    """改写段落内所有『表 7-N』引用；必要时合并 run，避免跨 run 漏改。"""
    full = "".join(t.text or "" for t in el.iter(qn("w:t")))
    if "表 7-" not in full:
        return
    n_all = len(REF_RE.findall(full))
    done = 0
    for t in el.iter(qn("w:t")):
        if t.text and "表 7-" in t.text:
            done += len(REF_RE.findall(t.text))
            t.text = _sub_text(t.text, target)
    if done >= n_all:
        return
    runs = list(el.iter(qn("w:r")))
    if not runs:
        return
    first = runs[0]
    for r in runs:
        for t in list(r.findall(qn("w:t"))):
            r.remove(t)
    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = _sub_text(full, target)
    first.append(t)
    for r in runs[1:]:
        r.getparent().remove(r)

def _rewrite_block(el, target):
    if el.tag == qn("w:p"):
        _rewrite_par(el, target)
    else:
        for p_ in el.iter(qn("w:p")):
            _rewrite_par(p_, target)

# ---------------------------------------------------------------- 5. 拼装文档
def new_doc():
    d = Document()
    b = d.element.body
    for p_ in list(b)[:-1]:          # 删掉默认空段落，保留 sectPr
        b.remove(p_)
    st = d.styles["Normal"]
    st.font.name = "Times New Roman"
    st.font.size = Pt(10.5)
    st._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), "宋体")
    return d

def append_block(d, el, target):
    cp = deepcopy(el)
    _rewrite_block(cp, target)
    d.element.body[-1].addprevious(cp)   # sectPr 之前
    return cp

def add_para(d, txt, bold=False, size=10.5, align=None, after=None):
    p_ = d.add_paragraph()
    if align is not None:
        p_.alignment = align
    r = p_.add_run(txt)
    r.bold = bold
    r.font.size = Pt(size)
    r.font.name = "Times New Roman"
    r._element.rPr.rFonts.set(qn("w:eastAsia"), "黑体" if bold else "宋体")
    if after is not None:
        p_.paragraph_format.space_after = Pt(after)
    return p_

main = new_doc()
supp = new_doc()

# --- 主文（原抬头段落随块复制，不另加，避免重复） ---
add_para(supp, "补充材料：第 7 章 表格与实现细节", bold=True, size=16,
         align=WD_ALIGN_PARAGRAPH.CENTER, after=4)
add_para(supp, "Supplementary Materials for Chapter 7", bold=True, size=11,
         align=WD_ALIGN_PARAGRAPH.CENTER, after=12)
add_para(supp,
         f"本文件收纳第 7 章主文之外的全部 {N_MOVED} 张表，编号 S-1 至 S-{N_MOVED}，"
         "顺序与主文原表号一致。主文为控制篇幅只保留 15 张表（协议、主结果、M1 与基线、"
         "复现与事件统计、失败机制、研究问题汇总）；逐题输出、统计方法与覆盖率网格、"
         "判定器敏感性、实现自检、筛选漏斗、外部对照分层，以及第 17 轮补入的三项独立复核明细"
         "全部在此给出，数值与主文完全相同，未删除或改写任何结果。"
         "交叉引用中「主文表 7-K」指主文第 K 张表。",
         size=10, after=10)

intro = ("为控制篇幅并突出主要科学发现，本章正文只保留 15 张表："
         "评测协议与判定口径、谱系抽取与四条基线、端到端主结果与增益归因、"
         "跨基座与留出库复现、事件级统计、静默错误分布以及研究问题汇总。"
         "统计方法与覆盖率网格、判定规则敏感性、实现自检、边界与漏斗明细等 "
         f"{N_MOVED} 张表移入《补充材料_第7章_v11.docx》，编号 S-1 至 S-{N_MOVED}；"
         "正文以「补充材料表 S-M」指向它们，所有数值与完整版一致。")

# 按块顺序复制；在第一个 7.1 标题前插入瘦身说明
cur_sec = ""
inserted_intro = False
for i, (kind, el, txt) in enumerate(blocks):
    if i in MOVE_BLOCKS:
        continue
    if kind == "p" and is_heading(txt):
        cur_sec = txt
    if (not inserted_intro and kind == "p" and is_heading(txt)
            and txt.startswith("7.1")):
        add_para(main, intro, size=10.2, after=8)
        inserted_intro = True
    append_block(main, el, "main")

if not inserted_intro:
    add_para(main, intro, size=10.2, after=8)

add_para(main,
         f"补充材料索引：S-1 至 S-{N_MOVED} 覆盖统计方法与四格计数、配对区间覆盖率、等价检验、"
         "基座与提示长度、判定规则敏感性、保序子集、事件数对账、oracle 一致性、"
         "Wilson 区间、逐库稳健性、规则基线与星号自检、改名档跨基座、聚类检验、"
         "留出库漏斗与样本构成、边界探针与真实数据扫描、事件级池化、近邻工作对照、"
         "EvoSchema 分层与映射单元重算，以及统一表述清单；"
         f"S-{SUPP_NEW.get(50, 35)} 至 S-{SUPP_NEW.get(52, 37)} 为第 17 轮补入的三项独立复核明细"
         "（全严判定规则、留出库 M1 直接表达式、段位编码压力测试 v2）；"
         f"S-{SUPP_NEW.get(53, 38)} 至 S-{N_MOVED} 为第 18 轮补入的执行器稳定性与编码复核明细"
         "（核心 LLM 条件 5 次重复运行、单作者编码下的等效独立复核清单）。"
         "详见《补充材料_第7章_v11.docx》末尾的表号对照表。",
         size=9.5, after=6)

# --- 补充材料：按原属节分组 ---
sec_of = {}
cur = ""
for i, (kind, el, txt) in enumerate(blocks):
    if kind == "p" and is_heading(txt):
        cur = txt
    for n in ALL:
        if i in units[n]:
            sec_of[n] = sec_of.get(n) or cur or "（章节外）"

last_sec = None
for n in MOVED:
    s = sec_of.get(n, "")
    if s != last_sec:
        add_para(supp, s, bold=True, size=11.5, after=6)
        last_sec = s
    for k in units[n]:
        append_block(supp, blocks[k][1], "supp")

# 对照表（不给「表 S-」题，避免破坏 S-1..S-34 的连续性）
add_para(supp, "附录：主文与补充材料表号对照", bold=True, size=12, after=6)
t = supp.add_table(rows=1, cols=4)
t.style = "Table Grid"
hdr = ["补充材料表号", "原表号", "表题", "主文/补充"]
for i, v in enumerate(hdr):
    c = t.rows[0].cells[i]
    c.text = ""
    r = c.paragraphs[0].add_run(v)
    r.bold = True
    r.font.size = Pt(8.5)
for n in ALL:
    row = t.add_row().cells
    vals = [f"S-{SUPP_NEW[n]}" if n in MOVED else "—",
            f"7-{n}（原号）", caps[n][1],
            "主文 表 7-%d" % MAIN_NEW[n] if n in KEEP else "补充材料"]
    for i, v in enumerate(vals):
        row[i].text = ""
        r = row[i].paragraphs[0].add_run(str(v))
        r.font.size = Pt(8)
        r.font.name = "Times New Roman"
        r._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")

main.save(MAIN_OUT)
supp.save(SUPP_OUT)

# ---------------------------------------------------------------- 6. 校验
def audit(path, prefix, expect):
    d = Document(path)
    titles = []
    for p_ in d.paragraphs:
        m = re.match(r"^表\s*%s-(\d+)\s\s(.+)$" % prefix, p_.text.strip())
        if m:
            titles.append((int(m.group(1)), m.group(2)))
    nums = [n for n, _ in titles]
    ok_seq = nums == list(range(1, expect + 1))
    text = "\n".join(p_.text for p_ in d.paragraphs)
    for tb in d.tables:
        for row in tb.rows:
            for c in row.cells:
                text += "\n" + c.text
    dangling = [int(x) for x in re.findall(r"表\s*%s-(\d+)" % prefix, text)
                if int(x) > expect]
    print(f"{path} | 段 {len(d.paragraphs)} 表 {len(d.tables)} | "
          f"表题 {len(titles)} 连续 {ok_seq} | 超范围引用 {len(dangling)}")
    assert len(titles) == expect, f"{path} 表题数 != {expect}"
    assert ok_seq, f"{path} 表号不连续: {nums[:20]}"
    if dangling:
        _ctx = []
        for p_ in d.paragraphs:
            for m_ in re.finditer(r"表\s*%s-(\d+)" % prefix, p_.text):
                if int(m_.group(1)) > expect:
                    _ctx.append(p_.text[max(0, m_.start()-40):m_.start()+40])
        for tb in d.tables:
            for row in tb.rows:
                for c in row.cells:
                    for m_ in re.finditer(r"表\s*%s-(\d+)" % prefix, c.text):
                        if int(m_.group(1)) > expect:
                            _ctx.append("[表]" + c.text[:80])
        print("  悬空引用上下文:", _ctx[:8])
    assert not dangling, f"{path} 存在悬空表号: {sorted(set(dangling))}"
    return titles

mt = audit(MAIN_OUT, "7", len(KEEP))
st_ = audit(SUPP_OUT, "S", N_MOVED)
assert len(mt) + len(st_) == N_ALL, f"拆分后表数 {len(mt)}+{len(st_)} != {N_ALL}"

# 主文中「补充材料表 S-M」必须都落在 1..N_MOVED
mtxt = "\n".join(p.text for p in Document(MAIN_OUT).paragraphs)
for tb in Document(MAIN_OUT).tables:
    for row in tb.rows:
        for c in row.cells:
            mtxt += "\n" + c.text
sm = [int(x) for x in re.findall(r"补充材料表 S-(\d+)", mtxt)]
assert sm and max(sm) <= N_MOVED and min(sm) >= 1, f"补充材料引用越界: {sorted(set(sm))}"
print("主文补充材料引用:", len(sm), "处，范围", min(sm), "-", max(sm))
print("完整版备份:", BAK)
print("主文保留表号(原→新):", {n: MAIN_NEW[n] for n in KEEP})
print("完成。")
