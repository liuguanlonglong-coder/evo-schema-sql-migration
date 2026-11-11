# -*- coding: utf-8 -*-
"""正文瘦身：把第 7 章 49 张表拆成「主文 15 表 + 补充材料 34 表」。

回应 Weakness 第 7 条「结构过于复杂、主要科学发现需要进一步突出」。
做法不是删数据，而是搬移：
  - 主文（文件名保持 投稿版_第7章_实验结果_v11.docx，章外 _T() 继续从它反查表号）
    只保留协议、主结果、M1/基线、留出/跨模型、事件统计、失败机制、RQ 汇总共 15 表；
  - 补充材料（补充材料_第7章_v11.docx）收纳其余 34 表，编号 S-1..S-34，
    带原属节标题、与主文表号对照表；
  - 原 49 表完整版另存 第7章_v11_完整版_49表.docx 作为素材库，不丢任何数字；
  - 所有交叉引用自动重定向：主文中对被移走表的「表 7-N」改为「补充材料表 S-M」，
    补充材料中对保留表的引用改为「主文表 7-K」；两份文档各自表号连续。

不重跑任何实验、不改动任何数值，只做文档结构重组。
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
KEEP = [1, 3, 4, 16, 18, 19, 20, 22, 25, 28, 33, 34, 39, 42, 48]

CAP_RE = re.compile(r"^表\s*7-(\d+)\s\s(.+)$")
REF_RE = re.compile(r"表\s*7-(\d+)")
HEAD_RE = re.compile(r"^7(?:\.\d+)*\s")

if not os.path.exists(BAK):
    if not os.path.exists(SRC):
        raise SystemExit("先运行 make_chapter7_v11.py 生成完整第 7 章")
    shutil.copy2(SRC, BAK)
INPUT = BAK          # 可重复运行：始终以 49 表完整版为输入
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
assert len(ALL) == 49, f"预期 49 表，实测 {len(ALL)}"
assert all(1 <= n <= 49 for n in ALL), "存在占位号残留（7-1xx），需先跑 _renumber_tables"
assert set(KEEP).issubset(set(ALL)), "KEEP 中有不存在的表号"

MOVED = [n for n in ALL if n not in KEEP]
MAIN_NEW = {n: i + 1 for i, n in enumerate([n for n in ALL if n in KEEP])}
SUPP_NEW = {n: i + 1 for i, n in enumerate(MOVED)}

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
         "本文件收纳第 7 章主文之外的全部 34 张表，编号 S-1 至 S-34，顺序与主文原表号一致。"
         "主文为控制篇幅只保留 15 张表（协议、主结果、M1 与基线、复现与事件统计、"
         "失败机制、研究问题汇总）；逐题输出、统计方法与覆盖率网格、判定器敏感性、"
         "实现自检、筛选漏斗、外部对照分层等细节全部在此给出，数值与主文完全相同，"
         "未删除或改写任何结果。交叉引用中「主文表 7-K」指主文第 K 张表。",
         size=10, after=10)

intro = ("为控制篇幅并突出主要科学发现，本章正文只保留 15 张表："
         "评测协议与判定口径、谱系抽取与四条基线、端到端主结果与增益归因、"
         "跨基座与留出库复现、事件级统计、静默错误分布以及研究问题汇总。"
         "统计方法与覆盖率网格、判定规则敏感性、实现自检、边界与漏斗明细等 "
         "34 张表移入《补充材料_第7章_v11.docx》，编号 S-1 至 S-34；"
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
         "补充材料索引：S-1 至 S-34 覆盖统计方法与四格计数、配对区间覆盖率、等价检验、"
         "基座与提示长度、判定规则敏感性、保序子集、事件数对账、oracle 一致性、"
         "Wilson 区间、逐库稳健性、规则基线与星号自检、改名档跨基座、聚类检验、"
         "留出库漏斗与样本构成、边界探针与真实数据扫描、事件级池化、近邻工作对照、"
         "EvoSchema 分层与映射单元重算，以及统一表述清单。详见《补充材料_第7章_v11.docx》"
         "末尾的表号对照表。",
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
    assert not dangling, f"{path} 存在悬空表号: {sorted(set(dangling))}"
    return titles

mt = audit(MAIN_OUT, "7", 15)
st_ = audit(SUPP_OUT, "S", 34)
assert len(mt) + len(st_) == 49, "拆分后表数不等于 49"

# 主文中「补充材料表 S-M」必须都落在 1..34
mtxt = "\n".join(p.text for p in Document(MAIN_OUT).paragraphs)
sm = [int(x) for x in re.findall(r"补充材料表 S-(\d+)", mtxt)]
assert sm and max(sm) <= 34 and min(sm) >= 1, f"补充材料引用越界: {sorted(set(sm))}"
print("主文补充材料引用:", len(sm), "处，范围", min(sm), "-", max(sm))
print("完整版备份:", BAK)
print("主文保留表号(原→新):", {n: MAIN_NEW[n] for n in KEEP})
print("完成。")
