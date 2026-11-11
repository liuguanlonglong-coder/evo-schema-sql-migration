# Data Availability Statement (DAS) -- template

> STATUS: **archive NOT yet deposited**. Every square-bracketed field is a
> placeholder. Do NOT paste the English statement below into the manuscript
> before the GitHub repository and Zenodo DOI actually exist. The current
> cover letter therefore says the deposit will be completed prior to
> publication, and Tables 9-2 / 9-3 in the front-matter file are marked
> "not completed".

## 中文说明（给作者）

Applied Intelligence 对原创研究要求提供 Data Availability Statement。
本声明要说明三件事：材料在哪、如何取得、有什么限制。
公开包 `公开材料包/` 内含逐题清单与输出、提示与演化规则、全部统计产物、
生成脚本本体、环境说明与 sha1 清单，**不含 BIRD 数据库文件**（第三方数据，
须从官方渠道获取并遵守其 CC BY-SA 4.0 许可）。

归档尚未完成。下方英文声明不能在投稿稿中使用。
注意顺序：Zenodo 允许在草稿阶段「预留 DOI」，拿到 DOI 后先回填本文件、
CITATION.cff 与稿件，重新打包，再上传最终 zip 并 Publish。
**不要先 Publish 再回填**——发布后文件与 DOI 都不可修改。

归档落地需同步修改的六处（建议用 `python fill_doi.py` 一次完成，不要手改）：
`CITATION.cff`、`DAS.md`、`UPLOAD_CHECKLIST.md`、投稿信第四段、
章外文档表 9-2 与表 9-3。

## English statement (paste into the manuscript AFTER deposit)

Data availability statement: The datasets generated and analysed during the
current study, together with all scripts, per-item outputs, statistical
routines, prompt specifications and the environment specification, are
available in the [REPOSITORY_NAME] repository at [REPOSITORY_DOI].
The complete 49-table version of the results chapter and the 34-table
supplementary material referenced by the main text are included in the
same record. The source SQLite databases are drawn from the publicly
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
| `[AUTHOR_FAMILY_NAME]` / `[AUTHOR_GIVEN_NAME]` | 作者姓名 | 提交前 |

## 与稿件内部一致性的约束

1. 投稿信第四段目前写的是「材料包已整理，将在发表前公开归档」，归档完成后
   改为「已归档于 [REPOSITORY_DOI]」。
2. 章外文档表 9-2 状态由「未完成」改为「已完成」；表 9-3 的
   `[repository DOI]` 占位符替换为真实 DOI。
3. 若归档最终未完成，DAS 必须改写为「材料可经合理请求向通讯作者获取」，
   且不得声称已有公开链接 —— 这是审稿人已经点名过的真实性红线。
