# GitHub / Zenodo 归档操作清单（需作者本人执行）

账号邮箱：liuguanlonglong@gmail.com
**邮箱不是凭据。** 本清单不含、也不要求把 token / 密码写进任何文件。
AI 助手可以本地生成材料包与元数据，但**不能代为登录 GitHub / Zenodo、
不能代为推送、不能代为申请 DOI**。以下所有涉及身份的步骤请本人操作。

---

## Step 0  前置确认（本人）

- [ ] 确认 GitHub 账号 liuguanlonglong@gmail.com 可正常登录（网页或已配置的本地 git 凭据）。
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

## Step 2  GitHub 仓库（本人，**已完成**）

仓库：https://github.com/liuguanlonglong-coder/evo-schema-sql-migration
已写入 `CITATION.cff` 的 `repository-code` 字段；它不产生 DOI，DOI 由 Step 3 的 Zenodo 预留。

推送到仓库（在 `公开材料包/` 目录内执行；仓库若已初始化过则跳过 init 三行）：

       git init
       git add .
       git commit -m "Reproduction materials v1.0"
       git remote add origin https://github.com/liuguanlonglong-coder/evo-schema-sql-migration
       git branch -M main
       git push -u origin main

推送前请确认：仓库可见性设为 **public**，且 `git status` 里没有 `archive_info.json`
（控制文件，不入库）、没有 `__pycache__` / `*.pyc`。

## Step 3  Zenodo「预留 DOI」（本人，关键一步，**已完成**：DOI 已预留，记录仍是草稿，尚未 Publish）

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

当前状态（2026-11-11）：归档已发布，DOI = 10.5281/zenodo.24123813；Step 0–6 均已完成，四处口径由 archive_info.json 统一驱动。
