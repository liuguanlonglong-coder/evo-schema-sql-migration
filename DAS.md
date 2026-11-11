# Data Availability Statement (DAS) -- template

> STATUS: **archive deposited**. The Zenodo record is published and the DOI below is live. This English statement may be pasted into the manuscript.

## 中文说明（给作者）

Programming and Computer Software（PCS）要求提供 Data Availability Statement。
本声明要说明三件事：材料在哪、如何取得、有什么限制。
公开包 `公开材料包/` 内含逐题清单与输出、提示与演化规则、全部统计产物、
生成脚本本体、环境说明与 sha1 清单，**不含 BIRD 数据库文件**（第三方数据，
须从官方渠道获取并遵守其 CC BY-SA 4.0 许可）。

归档已完成。下方英文声明可直接写入投稿稿；投稿信、标题页与正文数据可用性声明 由 make_front_matter_v11.py 读取同一份 archive_info.json 自动生成，口径一致。

归档落地需同步修改的六处（建议用 `python fill_doi.py` 一次完成，不要手改）：
`CITATION.cff`、`DAS.md`、`UPLOAD_CHECKLIST.md`、投稿信，以及投稿标题页与正文的
数据可用性声明。

## English statement (paste into the manuscript AFTER deposit)

Data availability statement: The datasets generated and analysed during the
current study, together with all scripts, per-item outputs, statistical
routines, prompt specifications and the environment specification, are
available in the evo-schema-sql-migration repository at 10.5281/zenodo.24123813.
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
