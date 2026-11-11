# -*- coding: utf-8 -*-
"""从 BIRD train 划分里挑选可用于留出库实验的数据库。

筛选标准（都与「能否施加无损演化 + 能否评测」直接相关，不是任意挑选）：
  1. 文件 < 110 MB：DROP COLUMN 会重写整表，超过 130 MB 的库在构造阶段会被跳过；
  2. train 划分里该库的题数 >= 20：保证受影响题足够，能做配对检验；
  3. 与 dev 的 11 个库不重叠：这是「留出」的定义。

输出 data/holdout_db_candidates.json
"""
import collections
import glob
import json
import os
import sqlite3
import sys

BASE = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE)

ROOT = "data/bird_train"
DEV = {"california_schools", "financial", "formula_1", "student_club",
       "thrombosis_prediction", "card_games", "codebase_community",
       "debit_card_specializing", "european_football_2", "superhero", "toxicology"}


def find(name):
    hits = glob.glob(os.path.join(ROOT, "**", name), recursive=True)
    return hits[0] if hits else None


dbdir = None
for cand in glob.glob(os.path.join(ROOT, "**", "*databases"), recursive=True):
    if os.path.isdir(cand):
        dbdir = cand
        break
if not dbdir:
    raise SystemExit(f"未找到 databases 目录（{ROOT} 下）")

jsonp = None
for n in ("train.json", "dev.json"):
    p = find(n)
    if p:
        jsonp = p
        break
if not jsonp:
    raise SystemExit("未找到题库 json")

print("db 目录:", dbdir)
print("题库:", jsonp)
items = json.load(open(jsonp, encoding="utf-8"))
by_db = collections.Counter(it["db_id"] for it in items)
print(f"题库 {len(items)} 条，{len(by_db)} 个库")

rows = []
seen = set()
for p in sorted(glob.glob(os.path.join(dbdir, "**", "*.sqlite"), recursive=True)):
    db = os.path.basename(p)[:-7]
    if db in seen:
        continue
    seen.add(db)
    mb = os.path.getsize(p) / 1024 / 1024
    try:
        con = sqlite3.connect(f"file:{p}?mode=ro", uri=True)
        tabs = [r[0] for r in con.execute(
            "SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'")]
        ncol = 0
        for t in tabs:
            ncol += len(con.execute(f'PRAGMA table_info("{t}")').fetchall())
        con.close()
    except Exception as e:
        rows.append({"db": db, "mb": round(mb, 1), "error": str(e)[:80]})
        continue
    rows.append({"db": db, "mb": round(mb, 1), "n_tables": len(tabs),
                 "n_cols": ncol, "n_questions": by_db.get(db, 0),
                 "in_dev": db in DEV,
                 "eligible": (mb < 110 and by_db.get(db, 0) >= 20 and db not in DEV)})

rows.sort(key=lambda r: (-r.get("n_questions", 0), r.get("mb", 9e9)))
elig = [r for r in rows if r.get("eligible")]
out = {"dbdir": dbdir, "questions_json": jsonp, "n_items": len(items),
       "n_dbs": len(rows), "n_eligible": len(elig), "dbs": rows}
json.dump(out, open("data/holdout_db_candidates.json", "w", encoding="utf-8"),
          ensure_ascii=False, indent=1)

print(f"\n可选库 {len(elig)} 个（<110MB、题数>=20、不在 dev）：")
print(f"{'库':34s}{'MB':>7s}{'表':>4s}{'列':>5s}{'题':>5s}")
for r in elig[:20]:
    print(f"{r['db']:34s}{r['mb']:7.1f}{r['n_tables']:4d}{r['n_cols']:5d}{r['n_questions']:5d}")
print("\n前 6 个建议：", [r["db"] for r in elig[:6]])
