# -*- coding: utf-8 -*-
"""为 Programming and Computer Software (PCS) 投稿准备公开归档元数据。

产出（写入 公开材料包/ 并在工作目录生成上传用 zip）：
  CITATION.cff          引用元数据
  LICENSE               许可说明（整体 CC BY 4.0，发布前需作者确认）
  DAS.md                Data Availability Statement 中英模板
  UPLOAD_CHECKLIST.md   GitHub / Zenodo 分步上传清单
  公开材料包_v11.zip    可直接上传 Zenodo 的压缩包

重要边界：用户只提供了 GitHub 账号邮箱 liuguanlonglong@gmail.com，
**邮箱不是凭据**。本脚本不登录、不推送、不申请 DOI，也不写任何 token 到文件。
所有需要身份的步骤在 UPLOAD_CHECKLIST.md 中列为「需作者本人执行」。
"""
import json
import os
import re
import zipfile
from datetime import date

BASE = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE)
OUT = "公开材料包"
ZIP = "公开材料包_v11.zip"
ACCOUNT = "liuguanlonglong@gmail.com"        # GitHub / Zenodo 注册邮箱（仅清单里提示用）
AUTHOR_EMAIL = "liuguanlong@lyvust.edu.cn"   # 稿件与 CITATION.cff 的通讯邮箱
AUTHOR_AFFIL = ("Department of Information Engineering, Linyi Vocational University "
                "of Science and Technology, Linyi, Shandong 276025, China")
INFO = "archive_info.json"

# archive_info.json 由 fill_doi.py 在作者拿到 Zenodo 预留 DOI 后写入。
# 文件不存在 = 尚未归档，所有模板保持占位符与「NOT yet deposited」口径（默认状态）。
# 文件存在 = 已预留/已发布，模板自动填入真实值，避免 DAS/投稿信/表9-2/表9-3 四处手改不同步。
AI = {}
if os.path.exists(INFO):
    AI = json.load(open(INFO, encoding="utf-8"))
    assert AI.get("status") in ("reserved", "deposited"), \
        "archive_info.json 的 status 须为 reserved 或 deposited"
DEPOSITED = bool(AI)

os.makedirs(OUT, exist_ok=True)

# ---------------------------------------------------------------- 1. 脱敏核查
SENSITIVE = re.compile(
    r"sk-[a-z0-9]{12,}|api[_-]?key\s*=\s*['\"][^'\"]{8,}|"
    r"token\s*=\s*['\"][^'\"]{8,}|password\s*=\s*['\"][^'\"]{4,}|"
    r"bearer\s+[a-z0-9]{12,}", re.I)
hits = []
for root, _dirs, fs in os.walk(OUT):
    for f in fs:
        if not f.endswith((".py", ".md", ".json", ".jsonl", ".csv", ".cff")):
            continue
        p = os.path.join(root, f)
        try:
            with open(p, encoding="utf-8", errors="ignore") as fh:
                for ln, line in enumerate(fh, 1):
                    if SENSITIVE.search(line):
                        hits.append((p, ln, line.strip()[:100]))
        except OSError:
            pass
print("脱敏核查:", "通过（0 处硬编码凭据）" if not hits else f"发现 {len(hits)} 处，需处理")
for h in hits[:10]:
    print("  ", h)

# ---------------------------------------------------------------- 2. CITATION.cff
# status=reserved 时 DOI 已可写入文件（Zenodo 允许预留 DOI 后回填再上传），
# 但 record 尚未 Publish，故 date-released / url 仍用占位；deposited 时全部落定。
_author_family = AI.get("author_family", "Liu")
_author_given = AI.get("author_given", "Guanlong")
_author_affil = AI.get("author_affiliation", AUTHOR_AFFIL)
_date_released = AI.get("date_released", "[YYYY-MM-DD]") if AI.get("status") == "deposited" else "[YYYY-MM-DD]"
_doi = AI.get("doi", "[REPOSITORY_DOI]")
_repo_url = AI.get("repo_url") or ""
# record 页面地址可由 DOI 直接推导（Zenodo 的 DOI 末段就是 record id），
# 不必等 Publish：发布后该地址即生效，先写上可以少一个占位符。
_record_url = AI.get("record_url") or ""
if not _record_url and _doi != "[REPOSITORY_DOI]":
    _m = re.match(r"10\.5281/zenodo\.(\d+)$", str(_doi).strip())
    if _m:
        _record_url = f"https://zenodo.org/records/{_m.group(1)}"
_cff_notes = (
    "Metadata are complete. BIRD database files are not redistributed here; "
    "obtain them from the official BIRD release under its original licence."
    if AI.get("status") == "deposited" else
    "DOI reserved; only date-released remains a placeholder and is filled on publication. "
    "BIRD database files are not redistributed here; obtain them from the official BIRD "
    "release under its original licence."
    if DEPOSITED else
    "Square-bracketed fields are placeholders and MUST be replaced before the record "
              "is made public. BIRD database files are not redistributed here; obtain them from "
              "the official BIRD release under its original licence.")

_repo_line = f'repository-code: "{_repo_url}"\n' if _repo_url else ""
_url_line = f'url: "{_record_url}"\n' if _record_url else ""

citation = f"""cff-version: 1.2.0
message: "If you use these materials, please cite both this archive and the associated article."
title: "Reproduction materials for: Schema Lineage Granularity for Historical SQL Migration under Schema Evolution"
abstract: >-
  Per-item judge outputs, evolution manifests, prompt specifications, statistical
  routines, boundary probes and the scripts that generate them, supporting a controlled
  study of schema-lineage granularity for historical SQL migration on BIRD dev/train
  databases (610 controlled items over 11 databases; column-level lineage versus
  segment-level expression lineage under a fixed deterministic rewriter). The manuscript
  itself (4 main-text tables, 9 figures) and its 45 supplementary tables (Table S-1 to
  Table S-45, published as Online Resource 1 with the article) are NOT included in this
  archive; the archive holds the data, scripts and statistical products needed to
  recompute every reported number.
type: dataset
authors:
  - family-names: "{_author_family}"
    given-names: "{_author_given}"
    affiliation: "{_author_affil}"
    email: "{AUTHOR_EMAIL}"
version: "v1.0"
date-released: "{_date_released}"
doi: "{_doi}"
license: "CC-BY-4.0"
{_repo_line}{_url_line}keywords:
  - schema evolution
  - schema lineage
  - text-to-SQL
  - SQL migration
  - SQLite
  - reproducibility
notes: >-
  {_cff_notes}
"""
open(os.path.join(OUT, "CITATION.cff"), "w", encoding="utf-8").write(citation)

# ---------------------------------------------------------------- 3. LICENSE
lic = """License for the reproduction materials (v1.0)

Unless stated otherwise, all items in this archive -- including items.csv,
outputs.jsonl, prompts.md, evolution_rules.md, stats/*.json, DAS.md and
UPLOAD_CHECKLIST.md -- are made available under the
Creative Commons Attribution 4.0 International licence (CC BY 4.0).
You may share and adapt them provided you give appropriate credit.

The scripts under code/ are additionally available under the MIT licence
(at the author's option); see below.

    MIT License (code/ only)
    Permission is hereby granted, free of charge, to any person obtaining a copy
    of the scripts to deal in them without restriction, including the rights to
    use, copy, modify and distribute, subject to inclusion of this notice.
    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND.

CC BY 4.0 full text: https://creativecommons.org/licenses/by/4.0/legalcode

IMPORTANT -- third-party data are NOT included:
  * BIRD dev/train SQLite databases are not redistributed in this archive.
    Obtain them from the official BIRD release (CC BY-SA 4.0) and cite it.
  * EvoSchema public column-level samples are likewise not redistributed.

CONFIRM BEFORE PUBLICATION: the author must confirm the licence choice (the
Zenodo upload form requires a single licence; if a single one is preferred,
select CC BY 4.0 for the whole record and remove the MIT paragraph above).
"""
open(os.path.join(OUT, "LICENSE"), "w", encoding="utf-8").write(lic)

# ---------------------------------------------------------------- 4. DAS.md
if DEPOSITED and AI.get("status") == "deposited":
    _das_status = ("STATUS: **archive deposited**. The Zenodo record is published and the "
                   "DOI below is live. This English statement may be pasted into the manuscript.")
    _das_cn = ("归档已完成。下方英文声明可直接写入投稿稿；投稿信、标题页与正文数据可用性声明 "
               "由 make_front_matter_v11.py 读取同一份 archive_info.json 自动生成，口径一致。")
elif DEPOSITED:  # reserved：DOI 已预留，record 尚未 Publish
    _das_status = ("STATUS: **DOI reserved, record NOT yet published**. The DOI has been "
                   "written into CITATION.cff as required by Zenodo's reserve-DOI flow, "
                   "but the record is still a draft. The English statement below keeps its "
                   "placeholder and MUST NOT be pasted into the manuscript yet. "
                   "Run fill_doi.py --status deposited after publishing.")
    _das_cn = (f"DOI 已预留（{AI.get('doi')}），但 Zenodo record 尚未 Publish。\n"
               "CITATION.cff 已写入真实 DOI（供上传使用），稿件中的 DAS 正文、投稿信、\n"
               "标题页与正文声明仍保持「将归档」口径，这是刻意的，不是漏改。\n"
               "Publish 之后运行 python fill_doi.py --status deposited 再一次性切换。\n"
               "**不要先 Publish 再回填文件**——发布后文件与 DOI 都不可修改。")
else:
    _das_status = ("STATUS: **archive NOT yet deposited**. Every square-bracketed field is a\n"
                   "> placeholder. Do NOT paste the English statement below into the manuscript\n"
                   "> before the GitHub repository and Zenodo DOI actually exist. The current\n"
                   "> cover letter therefore says the deposit will be completed prior to\n"
                   "> publication, and the data availability statements in the title page\n"
                   "> and the manuscript are marked \"not completed\".")
    _das_cn = ("归档尚未完成。下方英文声明不能在投稿稿中使用。\n"
               "注意顺序：Zenodo 允许在草稿阶段「预留 DOI」，拿到 DOI 后先回填本文件、\n"
               "CITATION.cff 与稿件，重新打包，再上传最终 zip 并 Publish。\n"
               "**不要先 Publish 再回填**——发布后文件与 DOI 都不可修改。")
_das_repo = AI.get("repo_name", "[REPOSITORY_NAME]")
# CITATION.cff 的 doi 在 reserved 阶段就填真实值（这正是「预留 DOI」的用途：
# 让待上传文件自带正确 DOI）；但 DAS.md 里的英文声明是给稿件的，
# 只有 record 真正 Publish（deposited）后才能换成真实 DOI。
_das_doi = (AI.get("doi", "[REPOSITORY_DOI]")
            if AI.get("status") == "deposited" else "[REPOSITORY_DOI]")

das = f"""# Data Availability Statement (DAS) -- template

> {_das_status}

## 中文说明（给作者）

Programming and Computer Software（PCS）要求提供 Data Availability Statement。
本声明要说明三件事：材料在哪、如何取得、有什么限制。
公开包 `公开材料包/` 内含逐题清单与输出、提示与演化规则、全部统计产物、
生成脚本本体、环境说明与 sha1 清单，**不含 BIRD 数据库文件**（第三方数据，
须从官方渠道获取并遵守其 CC BY-SA 4.0 许可）。

{_das_cn}

归档落地需同步修改的六处（建议用 `python fill_doi.py` 一次完成，不要手改）：
`CITATION.cff`、`DAS.md`、`UPLOAD_CHECKLIST.md`、投稿信，以及投稿标题页与正文的
数据可用性声明。

## English statement (paste into the manuscript AFTER deposit)

Data availability statement: The datasets generated and analysed during the
current study, together with all scripts, per-item outputs, statistical
routines, prompt specifications and the environment specification, are
available in the {_das_repo} repository at {_das_doi}.
The supplementary tables referenced by the manuscript (Table S-1 to Table S-45)
are provided as Online Resource 1 accompanying the article. The source SQLite
databases are drawn from the publicly
available BIRD benchmark (dev and train splits) and are not redistributed
here; they must be obtained from the official BIRD release under its
original licence. No new human-subject data were collected.

## 占位符清单

| 占位符 | 含义 | 何时填 |
|---|---|---|
| `[REPOSITORY_NAME]` | GitHub 仓库名 | 建库后 |
| `[REPOSITORY_DOI]` | Zenodo DOI | **预留 DOI 后即可填，不必等 Publish** |
| `[GITHUB_REPOSITORY_URL]` | 仓库地址 | 建库后 |
| `[ZENODO_RECORD_URL]` | Zenodo record 页面 | Publish 后 |
| `[YYYY-MM-DD]` | 发布日期 | Publish 当天 |
| 作者姓名 / 单位 / 邮箱 | 已在 CITATION.cff 填实（Liu Guanlong，信息工程系，`liuguanlong@lyvust.edu.cn`）；`fill_doi.py --author-family/--author-given` 可覆盖 | 已填 |

## 与稿件内部一致性的约束

1. 投稿信第四段目前写的是「材料包已整理，将在发表前公开归档」，归档完成后
   改为「已归档于 [REPOSITORY_DOI]」。
2. 标题页与正文的数据可用性声明由「未完成」改为「已完成」，
   `[repository DOI]` 占位符替换为真实 DOI。
3. 若归档最终未完成，DAS 必须改写为「材料可经合理请求向通讯作者获取」，
   且不得声称已有公开链接 —— 这是审稿人已经点名过的真实性红线。
"""
open(os.path.join(OUT, "DAS.md"), "w", encoding="utf-8").write(das)

# ---------------------------------------------------------------- 5. 上传清单
if DEPOSITED and AI.get("status") == "deposited":
    _state_line = (f"归档已发布，DOI = {AI.get('doi')}；Step 0–6 均已完成，"
                   "四处口径由 archive_info.json 统一驱动。")
elif DEPOSITED:
    _state_line = (f"DOI 已预留 = {AI.get('doi')}（status=reserved）；"
                   "Step 4 回填已完成，尚待 Step 5 上传并 Publish，之后把 status 改为 deposited。")
else:
    _state_line = ("Step 1 本地准备已完成；Step 0 / 2–6 **均未执行**，"
                   "GitHub 仓库与 Zenodo DOI 尚不存在。稿件中一切「已公开 / "
                   "released artefact / DAS is included」措辞均已删除，不要恢复。")

checklist = f"""# GitHub / Zenodo 归档操作清单（需作者本人执行）

账号邮箱：{ACCOUNT}
**邮箱不是凭据。** 本清单不含、也不要求把 token / 密码写进任何文件。
AI 助手可以本地生成材料包与元数据，但**不能代为登录 GitHub / Zenodo、
不能代为推送、不能代为申请 DOI**。以下所有涉及身份的步骤请本人操作。

---

## Step 0  前置确认（本人）

- [ ] 确认 GitHub 账号 {ACCOUNT} 可正常登录（网页或已配置的本地 git 凭据）。
- [ ] 确认 Zenodo 账号可用（可用 GitHub 账号授权登录，无需单独注册）。
- [ ] 决定许可证：建议整体 CC BY 4.0；若坚持代码 MIT，需在 Zenodo 表单里
      选一个主许可，并在 LICENSE 里保留双许可说明。
- [ ] 确认 CITATION.cff 中的作者姓名、题名、年份。

## Step 1  本地准备（可交由助手，已完成）

- [x] `python pack_materials.py` 生成 `公开材料包/`（脚本、数据、统计、README、ENVIRONMENT.md）。
- [x] `python prepare_archive.py` 生成 CITATION.cff / LICENSE / DAS.md /
      UPLOAD_CHECKLIST.md / `公开材料包_v11.zip`。
- [x] 脱敏核查：0 处硬编码 API key / token / password。
- [x] `MANIFEST.json` 记录每个文件与脚本的 sha1 与字节数。
- [ ] 本人复核 `公开材料包/outputs.jsonl` 与 `prompts.md` 中不含任何内部路径、
      私有数据或未脱敏标识。

## Step 2  建 GitHub 仓库（本人，可选但建议）

目的：**给代码一个可浏览的活链接**，填进 CITATION.cff 的 repository-code。
它不产生 DOI，DOI 由 Step 3 的 Zenodo 生成。

方式 A（网页，最稳妥）：
1. 登录 https://github.com ，新建 public repository，例如
   `schema-lineage-sql-migration`。
2. 不要勾选 "Add a README"（本地包已有 README.md，避免冲突）。
3. 在 `公开材料包/` 目录内执行（把 URL 换成实际仓库）：

       git init
       git add .
       git commit -m "Reproduction materials v1.0"
       git remote add origin [GITHUB_REPOSITORY_URL]
       git branch -M main
       git push -u origin main

方式 B（GitHub CLI，需本机已 `gh auth login`）：

       gh repo create schema-lineage-sql-migration --public --source=. --push

记录下两个值：`repo_name`（仓库名）与 `repo_url`（https://github.com/用户名/仓库名）。

## Step 3  Zenodo「预留 DOI」（本人，关键一步，尚未 Publish）

**用手动 New Upload，不要用 GitHub 自动集成。** 原因：Zenodo 的 GitHub
自动发布流程**不支持发布前预留 DOI**；只有手动 Upload 能在草稿阶段点
"Get a DOI now!" 先把 DOI 拿到手。

1. 登录 https://zenodo.org （可用 GitHub 账号授权）。
2. 点 **New Upload**，先不要上传文件、不要 Publish。
3. 填最少元数据：Title、Creators（作者姓名 + ORCID，如有）、
   Description（可粘 README.md 摘要）、License 选 **CC BY 4.0**、
   Resource type 选 Dataset 或 Software。
4. 在 **Digital Object Identifier** 区块，对
   "Do you already have a DOI for this upload?" 选 **No**，
   点 **Get a DOI now!** → 得到预留 DOI，形如 `10.5281/zenodo.XXXXXXX`。
5. **把这个 DOI 抄下来**，草稿保持不发布。不要删草稿（删了预留 DOI 会丢，
   但重新点 Get DOI 会拿回同一个）。

> Zenodo 官方文档明确：预留的 DOI「可以在上传前写进文件里」，正是为
> 本场景设计的。参见 help.zenodo.org/docs/deposit/describe-records/reserve-doi/ 。

## Step 4  回填并重新打包（可交助手，一条命令）

把 Step 2/3 得到的值交给助手，运行（写成一行，或用 ^ 续行）：

       python fill_doi.py --doi 10.5281/zenodo.XXXXXXX --repo-name schema-lineage-sql-migration --repo-url https://github.com/用户名/schema-lineage-sql-migration --author-family 姓 --author-given 名

脚本会写 `archive_info.json`，并自动重建以下全部材料，六处口径一次同步：

- [x] `CITATION.cff`：doi / repository-code / url / date-released / 作者姓名
- [x] `DAS.md`：STATUS 改已归档，替换仓库名与 DOI
- [x] `UPLOAD_CHECKLIST.md`：仓库地址与状态
- [x] 投稿信第四段：改为「已归档于 DOI」
- [x] 标题页与正文的数据可用性声明改为「已归档于 DOI」
- [x] 重新生成 `摘要_标题_贡献点_结论_v11.docx` 与 `公开材料包_v11.zip`

## Step 5  上传最终包并 Publish（本人）

1. 回到 Step 3 的 Zenodo 草稿，上传**重新打包后的** `公开材料包_v11.zip`
   （或上传 `公开材料包/` 内全部文件，含 CITATION.cff / MANIFEST.json）。
2. 核对元数据：Title、Creators、License = CC BY 4.0、DOI 与文件内一致。
   > 若同时放了 `CITATION.cff` 与 `.zenodo.json`，Zenodo **只认后者**；
   > 本包只提供 CITATION.cff，无冲突。
3. 点 **Publish**。此时 DOI 正式生效，形如 `10.5281/zenodo.XXXXXXX`。
4. 把 status 从 `reserved` 改为 `deposited`（连同发布日期、record 页面 URL
   一起交给助手再跑一次 `fill_doi.py`，见脚本 --help）。

重要：**发布后文件与 DOI 都不可修改**（Zenodo 官方规则：仅 metadata 可编辑）。
若之后要改文件，只能 New version，那会生成**新 DOI**，稿件里的旧 DOI
就会指向过期版本。所以务必「先回填、后 Publish」，顺序不能反。

## Step 6  最终核查（投稿前）

- [ ] DOI 链接可公开访问，且无需登录即可下载。
- [ ] Zenodo record 中可见 zip / 文件列表与 `MANIFEST.json`。
- [ ] record 里的 DOI == CITATION.cff 的 doi == DAS 的 DOI == 正文声明的 DOI。
- [ ] 稿件正文的「补充材料表 S-M」与公开包中的补充材料文档一致。
- [ ] DAS、投稿信、标题页、正文声明四处表述完全一致（都写「已归档」）。
- [ ] 若归档最终未完成，DAS 改为「可经合理请求获取」，并删除所有已发布措辞。

---

当前状态（{date.today().isoformat()}）：{_state_line}
"""
open(os.path.join(OUT, "UPLOAD_CHECKLIST.md"), "w", encoding="utf-8").write(checklist)

# ---------------------------------------------------------------- 5b. MANIFEST 增量刷新
# CITATION.cff / LICENSE / DAS.md / UPLOAD_CHECKLIST.md 在本脚本里晚于
# pack_materials.py 写 MANIFEST，不刷新的话这四条的 sha1 与字节数会与实际文件不符，
# 而 MANIFEST 正是用来核对「包内文件未被改动」的，过期即失去意义。
import hashlib


def _sha1_file(path):
    h = hashlib.sha1()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


_mpath = os.path.join(OUT, "MANIFEST.json")
if os.path.exists(_mpath):
    _man = json.load(open(_mpath, encoding="utf-8"))
    _files = _man.setdefault("files", {})
    for _n in ("CITATION.cff", "LICENSE", "DAS.md", "UPLOAD_CHECKLIST.md"):
        _p = os.path.join(OUT, _n)
        if os.path.exists(_p):
            _files[_n] = {"sha1": _sha1_file(_p), "bytes": os.path.getsize(_p)}
    json.dump(_man, open(_mpath, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print("MANIFEST.json: 已刷新 4 项归档元数据的 sha1")

# ---------------------------------------------------------------- 6. zip
with zipfile.ZipFile(ZIP, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as z:
    for root, dirs, fs in os.walk(OUT):
        dirs[:] = [d for d in dirs if d not in ("__pycache__", ".git")]
        for f in fs:
            p = os.path.join(root, f)
            # 条目从 OUT 起算：否则每条都带「公开材料包/」这个非 ASCII 前缀，
            # 解压后多一层中文目录，GitHub / Zenodo 的自动解包也可能乱码
            z.write(p, os.path.relpath(p, OUT))
print("归档包:", ZIP, round(os.path.getsize(ZIP) / 1024 / 1024, 2), "MB")
print("新增文件:", "CITATION.cff / LICENSE / DAS.md / UPLOAD_CHECKLIST.md")
print("提示: 邮箱不是凭据，Step 0/2-6 需作者本人执行；AI 未登录、未推送、未申请 DOI。")
