# -*- coding: utf-8 -*-
"""生成《DOI落地操作指南.docx》——给作者本人照着做的分步清单。

本轮修正了旧清单的一个实质性顺序错误：Zenodo 发布后文件与 DOI 都不可修改，
旧清单却把「回填占位符」放在 Publish 之后。正确顺序为
「草稿阶段预留 DOI → 回填 → 上传 → Publish」，并新增 fill_doi.py 一键回填。
"""
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
ACCOUNT = "liuguanlonglong@gmail.com"
TODAY = date.today().isoformat()


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


def CODE(doc, txt):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Pt(19)
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(6)
    r = p.add_run(txt)
    r.font.name = "Consolas"
    r.font.size = Pt(8.6)
    r._element.rPr.rFonts.set(qn("w:eastAsia"), CN)
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
        c.width = Cm(widths[i])
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        _run(c.paragraphs[0], str(head[i]), bold=True, size=font)
    for row in rows:
        cs = t.add_row().cells
        for i, v in enumerate(row):
            cs[i].width = Cm(widths[i])
            p_ = cs[i].paragraphs[0]
            p_.alignment = {"l": WD_ALIGN_PARAGRAPH.LEFT,
                            "c": WD_ALIGN_PARAGRAPH.CENTER}[aligns[i]]
            p_.paragraph_format.space_after = Pt(1)
            emit(p_, str(v), font)
    return t


D = make_doc()
p = D.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
_run(p, "DOI 落地操作指南", bold=True, size=16, font=CN_H)
p = D.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
_run(p, f"Applied Intelligence 投稿 · Data Availability / Zenodo 归档（{TODAY}）", size=9.5)

H(D, "一、先划清责任边界")
P(D, f"你给的是 GitHub 账号邮箱 {ACCOUNT}，**邮箱不是凭据**。我不能代为登录 GitHub / Zenodo、"
     "不能代为推送、不能代为申请 DOI，也没有把任何 token 或密码写进文件。"
     "涉及身份的四步（建库、预留 DOI、上传、Publish）必须你本人操作；"
     "**回填六处口径、重新打包、重新生成文档这些事可以交给我，一条命令完成**。")
P(D, "Applied Intelligence 对原创研究强制要求 Data Availability Statement（DAS）。"
     "它需要一个可公开访问、带持久标识符的材料包。目前材料包已备好但尚未公开，"
     "稿件中所有「已发布 / released / included」措辞都已在第 14 轮删掉，这是审稿人点名过的真实性红线。")

CAP(D, "表 1  谁做什么")
TABLE(D, ["步骤", "谁做", "需要什么"],
      [["本地材料包与元数据准备", "助手（已完成）", "无"],
       ["建 GitHub public 仓库", "你本人", "GitHub 登录"],
       ["Zenodo 草稿 + 预留 DOI", "你本人", "Zenodo 登录（可用 GitHub 授权）"],
       ["回填六个文件 + 重新打包", "助手（一条命令）", "你给出的 5 个值"],
       ["上传最终 zip + Publish", "你本人", "Zenodo 登录"],
       ["Publish 后切为完成时口径", "助手（一条命令）", "发布日期、record 页面"]],
      [5.4, 3.4, 6.6], ["l", "c", "l"])

H(D, "二、一个必须避开的坑（本轮修正）")
P(D, "旧清单把「回填占位符」放在 Zenodo Publish **之后**，这是错的。"
     "Zenodo 官方规则：记录发布后**文件与 DOI 都不可修改**，只有 metadata 可编辑；"
     "若要改文件只能 New version，而新版本会生成**新 DOI**，稿件里引用的旧 DOI 就指向过期版本了。")
P(D, "正确顺序是：**草稿阶段预留 DOI → 用这个 DOI 回填文件 → 上传最终包 → Publish**。"
     "Zenodo 专门提供 Reserve a DOI：草稿里对「Do you already have a DOI?」选 No，"
     "点 Get a DOI now! 即可先拿到 DOI，官方明确说明「预留的 DOI 可以在上传前写进文件里」。")
NOTE(D, "另一个坑：Zenodo 的 **GitHub 自动集成流程不支持发布前预留 DOI**，"
        "只有手动 New Upload 可以。所以主路径是手动上传；GitHub 仓库只作为代码的可浏览镜像，"
        "填进 CITATION.cff 的 repository-code 字段，不负责生成 DOI。")

H(D, "三、你要做的五步")

H(D, "Step 0  前置确认", 2)
PL(D, "· 确认 GitHub 账号可登录；确认 Zenodo 可用（GitHub 账号授权即可，无需另注册）。")
PL(D, "· 决定许可证：建议**整体 CC BY 4.0**。Zenodo 表单只让选一个主许可；"
      "包内 LICENSE 目前写了「数据/文档 CC BY 4.0 + 代码 MIT 可选」，若选单一许可需删掉 MIT 段。")
PL(D, "· 确认作者姓名（姓/名的英文或拼音）、题名、年份。CITATION.cff 里现在是占位符。")

H(D, "Step 2  建 GitHub 仓库（可选但建议）", 2)
P(D, "登录 github.com 新建 public repository，例如 schema-lineage-sql-migration；"
     "**不要勾选 Add a README**（本地包已有，会冲突）。然后在 公开材料包/ 目录内执行：")
CODE(D, "git init\ngit add .\ngit commit -m \"Reproduction materials v11\"\n"
        "git remote add origin https://github.com/用户名/schema-lineage-sql-migration\n"
        "git branch -M main\ngit push -u origin main")
P(D, "记下两个值：仓库名 repo_name、仓库地址 repo_url。")

H(D, "Step 3  Zenodo 预留 DOI（关键，此时不要 Publish）", 2)
PL(D, "1. 登录 zenodo.org，点 **New Upload**，先不传文件、不发布。")
PL(D, "2. 填最少元数据：Title、Creators（作者 + ORCID 如有）、Description（可粘 README 摘要）、"
      "License 选 **CC BY 4.0**、Resource type 选 Dataset 或 Software。")
PL(D, "3. 在 Digital Object Identifier 区块，对「Do you already have a DOI?」选 **No**，"
      "点 **Get a DOI now!** → 得到形如 10.5281/zenodo.XXXXXXX 的预留 DOI。")
PL(D, "4. 抄下 DOI，**草稿保持不发布**；不要删草稿（删了预留 DOI 会丢，重点 Get DOI 会拿回同一个）。")

CAP(D, "表 2  回填需要你给出的值")
TABLE(D, ["值", "来源", "示例"],
      [["doi", "Step 3 预留所得", "10.5281/zenodo.1234567"],
       ["repo_name", "Step 2 仓库名", "schema-lineage-sql-migration"],
       ["repo_url", "Step 2 仓库地址", "https://github.com/用户名/schema-lineage-sql-migration"],
       ["author_family / author_given", "作者姓名", "Liu / Guanlong"],
       ["date_released / record_url", "Publish 之后才有", "2026-11-11 / https://zenodo.org/records/1234567"]],
      [4.6, 4.4, 6.4], ["l", "l", "l"])

H(D, "Step 4  把值交给我，一条命令回填", 2)
P(D, "你只需把上面几个值发给我，我运行下面这条命令（也可自己跑）。"
     "它会写 archive_info.json，并自动重建 CITATION.cff、DAS.md、UPLOAD_CHECKLIST.md、"
     "投稿信第四段、表 9-2、表 9-3、摘要文档与 zip，**六处口径一次同步**：")
CODE(D, "python fill_doi.py --doi 10.5281/zenodo.1234567 "
        "--repo-name schema-lineage-sql-migration "
        "--repo-url https://github.com/用户名/schema-lineage-sql-migration "
        "--author-family Liu --author-given Guanlong")
P(D, "此时 status=reserved：CITATION.cff 已写入真实 DOI（供上传），"
     "但**稿件里的 DAS 正文与投稿信仍保持「将在发表前归档」口径**——这是刻意的，"
     "因为 record 还没发布，不是漏改。脚本自带 8 项一致性校验，全 PASS 才放行。")

H(D, "Step 5  上传最终包并 Publish", 2)
PL(D, "1. 回到 Zenodo 草稿，上传**重新打包后的** 公开材料包_v11.zip"
      "（或上传 公开材料包/ 内全部文件，含 CITATION.cff 与 MANIFEST.json）。")
PL(D, "2. 核对元数据：Title、Creators、License = CC BY 4.0、DOI 与文件内一致。"
      "若同时放了 CITATION.cff 与 .zenodo.json，Zenodo 只认后者；本包只提供 CITATION.cff，无冲突。")
PL(D, "3. 点 **Publish**，DOI 正式生效。")
PL(D, "4. 把发布日期和 record 页面发给我，我运行：")
CODE(D, "python fill_doi.py --status deposited --date-released 2026-11-11 "
        "--record-url https://zenodo.org/records/1234567")
P(D, "这一步把投稿信改为「已归档于 DOI」、表 9-2 状态改「已完成」、"
     "表 9-3 换成真实 DOI、DAS.md 切为已发布声明。11 项校验全 PASS 后即可投稿。")

H(D, "四、不能做的事与回滚")
PL(D, "· **不要先 Publish 再回填文件**；发布后改文件只能发新版本并换 DOI，稿件里的 DOI 就废了。")
PL(D, "· 归档未完成时，DAS 只能写「材料可经合理请求向通讯作者获取」，"
      "绝不写 provided / included / released / 已公开。")
PL(D, "· 任何阶段想回到未归档状态：python fill_doi.py --reset（删除 archive_info.json 并重建）。")
PL(D, "· 只想按现有 archive_info.json 重建：python fill_doi.py --rebuild-only。")
PL(D, "· 包内不含 BIRD 数据库文件（第三方数据，CC BY-SA 4.0），DAS 已写明须从官方渠道获取。")

CAP(D, "表 3  当前状态与剩余动作")
TABLE(D, ["项", "状态"],
      [["本地材料包 / 元数据 / zip", "已完成（98 文件，1.01 MB，0 处硬编码凭据）"],
       ["GitHub 仓库", "未创建（需本人）"],
       ["Zenodo DOI", "未申请（需本人）"],
       ["稿件 DAS / 投稿信 / 表 9-2 / 表 9-3", "均为「尚未归档」口径，内部一致"],
       ["回填链路", "已就绪并通过 reserved / deposited / reset 三态测试（8/8、11/11、6/6 PASS）"],
       ["P1 三项 A 层数据入稿", "待你决定是否现在入稿（另一轮）"]],
      [6.4, 9.0], ["l", "l"])

NOTE(D, "本指南对应修正：旧 UPLOAD_CHECKLIST.md 的 Step 4/5 顺序错误已改正，"
        "并新增 fill_doi.py 与 archive_info.json 单点驱动。"
        "包内 README.md 与 UPLOAD_CHECKLIST.md 已同步更新。")
D.save("DOI落地操作指南.docx")
print("saved: DOI落地操作指南.docx")
print("段落", len(D.paragraphs), "表格", len(D.tables))
