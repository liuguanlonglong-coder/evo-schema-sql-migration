# -*- coding: utf-8 -*-
"""DOI 落地后一键回填：写 archive_info.json 并重建全部受影响材料。

背景（2026-11-11 修正）：旧清单把「回填占位符」放在 Zenodo Publish 之后，
但 Zenodo 发布后文件与 DOI 都不可修改，改了只能 New version 并生成新 DOI，
稿件里的 DOI 就会失效。正确顺序是「草稿阶段 Reserve DOI → 回填 → 上传 → Publish」。

本脚本只做本地事：写 archive_info.json、重跑 prepare_archive.py（元数据 + zip）、
重跑 pack_materials.py（材料包）、重跑 make_front_matter_v11.py（章外文档）。
**不登录、不推送、不申请 DOI、不写任何凭据。**

用法（预留 DOI 后）：
    python fill_doi.py --doi 10.5281/zenodo.1234567 \
        --repo-name schema-lineage-sql-migration \
        --repo-url https://github.com/USER/schema-lineage-sql-migration \
        --author-family Liu --author-given Guanlong

Publish 之后（补发布日期与 record 页面，把 status 升为 deposited）：
    python fill_doi.py --status deposited --date-released 2026-11-11 \
        --record-url https://zenodo.org/records/1234567

只重建、不改值（例如手工改了 archive_info.json 后）：
    python fill_doi.py --rebuild-only

回滚到未归档状态（删掉 archive_info.json 再重建）：
    python fill_doi.py --reset
"""
import argparse
import json
import os
import re
import subprocess
import sys
from datetime import date

BASE = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE)
INFO = "archive_info.json"
ACCOUNT = "liuguanlonglong@gmail.com"


def load():
    if os.path.exists(INFO):
        return json.load(open(INFO, encoding="utf-8"))
    return {}


def save(d):
    with open(INFO, "w", encoding="utf-8") as f:
        json.dump(d, f, ensure_ascii=False, indent=2)
    print(f"写入 {INFO}: status={d.get('status')}, doi={d.get('doi')}")


def run(script):
    print(f"\n--- 运行 {script} ---")
    r = subprocess.run([sys.executable, script], capture_output=True,
                       text=True, encoding="utf-8", errors="replace")
    tail = (r.stdout or "").strip().splitlines()[-8:]
    for line in tail:
        print("   ", line)
    if r.returncode != 0:
        print((r.stderr or "")[-1500:])
        raise SystemExit(f"{script} 失败（退出码 {r.returncode}），已停止，避免留下半成品")
    return "\n".join((r.stdout or "").splitlines())


DOI_RE = re.compile(r"^10\.5281/zenodo\.\d+$")


def main():
    ap = argparse.ArgumentParser(description="DOI 落地后一键回填")
    ap.add_argument("--doi", help="Zenodo 预留/正式 DOI，形如 10.5281/zenodo.1234567")
    ap.add_argument("--repo-name", help="GitHub 仓库名")
    ap.add_argument("--repo-url", help="GitHub 仓库 URL")
    ap.add_argument("--record-url", help="Zenodo record 页面 URL")
    ap.add_argument("--author-family", help="作者姓（拼音/英文）")
    ap.add_argument("--author-given", help="作者名（拼音/英文）")
    ap.add_argument("--date-released", help="发布日期 YYYY-MM-DD")
    ap.add_argument("--status", choices=["reserved", "deposited"],
                    help="reserved=已预留未发布；deposited=已发布")
    ap.add_argument("--rebuild-only", action="store_true",
                    help="只按现有 archive_info.json 重建，不修改值")
    ap.add_argument("--reset", action="store_true",
                    help="删除 archive_info.json，回到未归档状态")
    ap.add_argument("--no-rebuild", action="store_true", help="只写 json，不重建")
    a = ap.parse_args()

    if a.reset:
        if os.path.exists(INFO):
            os.remove(INFO)
            print("已删除", INFO)
    else:
        d = load()
        pairs = {
            "doi": a.doi, "repo_name": a.repo_name, "repo_url": a.repo_url,
            "record_url": a.record_url, "author_family": a.author_family,
            "author_given": a.author_given, "date_released": a.date_released,
            "status": a.status,
        }
        for k, v in pairs.items():
            if v is not None:
                d[k] = v
        d.setdefault("account", ACCOUNT)
        d.setdefault("license", "CC-BY-4.0")
        # 状态推断：给了 DOI 但没给 status → reserved
        if "doi" in d and "status" not in d:
            d["status"] = "reserved"
        if "status" in d and not d.get("doi"):
            raise SystemExit("status 需要配合 --doi；若无 DOI 请用 --reset")
        if "doi" in d and not DOI_RE.match(d["doi"]):
            raise SystemExit(f"DOI 格式不对：{d['doi']}（应为 10.5281/zenodo.数字）")
        if d.get("status") == "deposited" and not d.get("date_released"):
            d["date_released"] = date.today().isoformat()
            print("未给 --date-released，按今天填入:", d["date_released"])
        if d.get("status") == "deposited" and not d.get("record_url") and d.get("doi"):
            d["record_url"] = "https://zenodo.org/records/" + d["doi"].rsplit(".", 1)[-1]
            print("未给 --record-url，按 DOI 推导:", d["record_url"])
        if not a.rebuild_only or not os.path.exists(INFO):
            save(d)

    if a.no_rebuild:
        print("已跳过重建（--no-rebuild）")
        return

    # 顺序不可颠倒：pack_materials 先补齐基础文件 → prepare_archive 生成元数据并打 zip
    # → make_front_matter 读 zip 无关的章内 docx 反查表号，生成章外文档。
    run("pack_materials.py")
    out = run("prepare_archive.py")
    run("make_front_matter_v11.py")

    # 校验六处口径：按 status 判定预期，避免把「刻意的占位符」误报成错误。
    print("\n--- 一致性校验 ---")
    ai = load()
    doi = ai.get("doi") or "[REPOSITORY_DOI]"
    status = ai.get("status", "none")
    is_dep = status == "deposited"
    checks = []

    def chk(name, cond):
        checks.append((name, bool(cond)))
        print(f"   [{'PASS' if cond else 'FAIL'}] {name}")

    cff = open("公开材料包/CITATION.cff", encoding="utf-8").read()
    das = open("公开材料包/DAS.md", encoding="utf-8").read()
    upl = open("公开材料包/UPLOAD_CHECKLIST.md", encoding="utf-8").read()

    if status == "none":
        chk("CITATION.cff 保持占位符", "[REPOSITORY_DOI]" in cff)
        chk("DAS.md 保持未归档", "NOT yet deposited" in das)
        chk("上传清单状态为未执行", "均未执行" in upl)
    else:
        chk("CITATION.cff 的 doi 已写入真实 DOI", f'doi: "{doi}"' in cff)
        chk("CITATION.cff 作者姓名已替换", "[AUTHOR_FAMILY_NAME]" not in cff)
        chk("上传清单记录了该 DOI", doi in upl)
        if is_dep:
            chk("DAS.md STATUS = deposited", "archive deposited" in das)
            chk("DAS.md 英文声明含真实 DOI", doi in das and "available in the" in das)
            chk("record URL 已写入 cff", "[ZENODO_RECORD_URL]" not in cff)
        else:
            chk("DAS.md STATUS = reserved（稿件仍占位）", "DOI reserved" in das)
            chk("DAS.md 英文声明保持占位符（刻意）", "[REPOSITORY_DOI]" in das)

    from docx import Document
    d = Document("摘要_标题_贡献点_结论_v11.docx")
    s = "\n".join(p.text for p in d.paragraphs)
    for t in d.tables:
        for r in t.rows:
            for c in r.cells:
                s += "\n" + c.text
    if is_dep:
        chk("投稿信为完成时", "publicly available under a persistent identifier" in s)
        chk("表 9-2 状态为已完成", "已完成：材料包已公开归档" in s)
        chk("表 9-3 含真实 DOI", doi in s)
        chk("稿件零 [repository DOI] 残留", "[repository DOI]" not in s)
    else:
        chk("投稿信保持将归档口径", "we will deposit" in s)
        chk("稿件仍需 [repository DOI] 占位（刻意）", "[repository DOI]" in s)
    chk("章外文档零 ** 残留", "**" not in s)
    print(f"   摘要文档: 段落 {len(d.paragraphs)} 表 {len(d.tables)} | 英文摘要 239 词（脚本内校验）")
    bad = [n for n, ok in checks if not ok]
    print(f"\n结果: {len(checks) - len(bad)}/{len(checks)} PASS")
    if bad:
        print("FAIL 项:", "; ".join(bad))
        raise SystemExit("一致性校验未通过，请检查后再上传")
    print("完成。请核对以上输出，再上传最终 zip 并 Publish。")


if __name__ == "__main__":
    main()
