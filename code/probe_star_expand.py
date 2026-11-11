# -*- coding: utf-8 -*-
"""星号展开的合成探针（第八轮评审第 1 条的配套证据）。

为什么需要这个探针
------------------
v2 实现了「按旧表列清单展开投影星号」，但 559 条真实样本中投影星号
出现 **0 次**，展开路径在真实数据上从未被触发。仅凭「已实现」不足以
说服评审，故用一个两步演化的合成库验证三件事：

  P1  展开模式能重建旧列集合（逐分量还原），keep 模式不能；
  P2  `t.*` 限定星号、以及裸 `*` 的多表展开都能正确处理；
  P3  `COUNT(*)` 在两种模式下都不被触碰、结果一致。

同时测出构造层的固有缺陷（这是本轮必须写进正文的负面事实）
----------------------------------------------------------
  P4  本构造的 gold 由「机械替换列引用」生成，`*` 不是列引用，
      故 gold 会原样保留 `*`，在演化后库上返回**新列集合**。
      即：含投影星号的题目在本构造下无法被忠实评价——展开模式给出
      语义正确的旧列集合，反而与 gold 不一致。

输入：无（自建临时库）
输出：results/probe_star_expand.json
"""
import json
import os
import sqlite3
import tempfile

import rule_rewrite as RR

BASE = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE)

SEP = "CHAR(31)"
TMP = tempfile.mkdtemp(prefix="starprobe_")

# 旧库：Patient(pid, name, sex, admit) + Dept(did, dname, pid)
# 演化：name、sex 合并为 attr_001 = name ‖ CHAR(31) ‖ sex
#       admit 改名 admission_date；Dept 不受影响
OLD_DDL = """
CREATE TABLE Patient(pid INTEGER, name TEXT, sex TEXT, admit TEXT);
CREATE TABLE Dept(did INTEGER, dname TEXT, pid INTEGER);
INSERT INTO Patient VALUES (1,'ann','F','2020-01-02'),(2,'bob','M','2021-03-04');
INSERT INTO Dept VALUES (10,'cardio',1),(11,'onco',2);
"""
NEW_DDL = """
CREATE TABLE Patient(pid INTEGER,
  attr_001 TEXT,
  admission_date TEXT);
CREATE TABLE Dept(did INTEGER, dname TEXT, pid INTEGER);
INSERT INTO Patient VALUES (1,'ann' || %(s)s || 'F','2020-01-02'),
                           (2,'bob' || %(s)s || 'M','2021-03-04');
INSERT INTO Dept VALUES (10,'cardio',1),(11,'onco',2);
""" % {"s": SEP}

MAPPING = {("patient", "name"): "SUBSTR(attr_001,1,INSTR(attr_001,CHAR(31))-1)",
           ("patient", "sex"): "SUBSTR(attr_001,INSTR(attr_001,CHAR(31))+1)",
           ("patient", "admit"): "admission_date"}


def build(ddl, name):
    p = os.path.join(TMP, name)
    if os.path.exists(p):
        os.remove(p)
    con = sqlite3.connect(p)
    con.executescript(ddl)
    con.commit()
    con.close()
    return p


def run(path, sql):
    rows, err = RR.q(path, sql)
    return ("ERR:" + err) if err else RR.norm(rows)


def main():
    old_db = build(OLD_DDL, "old.sqlite")
    new_db = build(NEW_DDL, "new.sqlite")
    schema = RR.old_schema_of(old_db)
    out = {"version": "star-probe-v1",
           "old_schema": {k: [c for c, _ in v] for k, v in schema.items()},
           "cases": {}}

    cases = [
        ("P1_bare_star", "SELECT * FROM Patient WHERE pid = 1"),
        ("P2_qualified_star",
         "SELECT T1.* FROM Patient AS T1 INNER JOIN Dept AS T2 ON T1.pid = T2.pid "
         "WHERE T1.pid = 1"),
        ("P2_bare_star_join",
         "SELECT * FROM Patient AS T1 INNER JOIN Dept AS T2 ON T1.pid = T2.pid "
         "WHERE T1.pid = 1"),
        ("P3_count_star", "SELECT COUNT(*) FROM Patient"),
        ("P3_count_star_where",
         "SELECT COUNT(*) FROM Patient WHERE sex = 'F'"),
    ]

    for name, sql in cases:
        # 旧库上的正确答案（语义基准）
        want = run(old_db, sql)
        rec = {"sql": sql, "old_result": want}
        for mode in ("expand", "keep"):
            new_sql, st, det = RR.rewrite(sql, MAPPING, schema, star=mode)
            got = run(new_db, new_sql) if new_sql else None
            rec[mode] = {"sql": new_sql, "status": st, "detail": det,
                         "result": got,
                         "matches_old": got == want if new_sql else False}
        # 构造口径的 gold：机械替换列引用、但不触碰任何星号
        # —— 这正是 keep 模式的行为，故直接取 keep 的输出作为 gold。
        rec["construct_gold"] = {"sql": rec["keep"]["sql"],
                                 "result": rec["keep"]["result"]}
        rec["expand_matches_gold"] = (rec["expand"]["result"]
                                      == rec["construct_gold"]["result"])
        out["cases"][name] = rec

    # 汇总判定
    p1 = out["cases"]["P1_bare_star"]
    out["verdict"] = {
        "P1_expand_reconstructs_old": p1["expand"]["matches_old"],
        "P1_keep_fails": not p1["keep"]["matches_old"],
        "P2_qualified_ok": out["cases"]["P2_qualified_star"]["expand"]["matches_old"],
        "P2_join_ok": out["cases"]["P2_bare_star_join"]["expand"]["matches_old"],
        "P3_count_star_unaffected":
            (out["cases"]["P3_count_star"]["expand"]["matches_old"]
             and out["cases"]["P3_count_star"]["keep"]["matches_old"]
             and out["cases"]["P3_count_star_where"]["expand"]["matches_old"]),
        "P4_star_item_not_faithfully_evaluable":
            not out["cases"]["P1_bare_star"]["expand_matches_gold"],
    }

    json.dump(out, open("results/probe_star_expand.json", "w", encoding="utf-8"),
              ensure_ascii=False, indent=1)

    print("旧 schema:", out["old_schema"])
    for name, rec in out["cases"].items():
        print("\n[%s] %s" % (name, rec["sql"]))
        print("  旧库(语义基准):", rec["old_result"])
        for mode in ("expand", "keep"):
            m = rec[mode]
            print("  %-6s matches_old=%-5s %s" % (mode, m["matches_old"], m["sql"]))
        print("  构造gold :", rec["construct_gold"]["result"],
              "| expand==gold:", rec["expand_matches_gold"])
    print("\n判定:")
    for k, v in out["verdict"].items():
        print("  %-40s %s" % (k, v))
    print("\n-> results/probe_star_expand.json")


if __name__ == "__main__":
    main()
