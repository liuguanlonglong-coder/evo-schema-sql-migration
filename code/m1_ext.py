# -*- coding: utf-8 -*-
"""在扩充事件集（每表/每列一个分片库）上跑 M1 谱系抽取。

为什么要单独一层
  m1_lineage.infer_db / m1_part.infer_db 都按「一个库一个演化副本」定位文件
  （data/evolved_<evo>/<db>.sqlite）。扩充事件集的目录是每事件一个分片
  （data/evolved_ext/<db>__<table>__<task>.sqlite），文件名对不上。
  这里为每个分片建一个只含该分片的临时目录，把分片硬链成 <db>.sqlite，
  再把两个模块的 EVO_DIR 指过去——抽取算法本身一行不改，
  保证扩充集上的谱系与主实验同算法、同阈值、同输出格式。

输出
  data/lineage_infer_ext/<task>_<tag>_profile.json
  data/lineage_infer_ext/<task>_<tag>_part.json     （仅 merge_col）

用法
  python m1_ext.py --task merge_col
  python m1_ext.py --task split_col --mode profile
"""
import argparse
import json
import os
import shutil
import sys

BASE = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE)
sys.path.insert(0, BASE)

import m1_lineage as M1          # noqa: E402
import m1_part as MP             # noqa: E402

EXT_DIR = os.path.join(BASE, "data", "evolved_ext")
OUTDIR = os.path.join(BASE, "data", "lineage_infer_ext")
TMP = os.path.join(BASE, "data", "_ext_tmp")


def shard_to_alias(shard_path, tag):
    """把分片库放成 <tmp>/<tag>/<db>.sqlite，并返回 (临时目录, db_id)。"""
    db = tag.split("__")[0]
    d = os.path.join(TMP, tag)
    os.makedirs(d, exist_ok=True)
    dst = os.path.join(d, f"{db}.sqlite")
    if os.path.exists(dst):
        os.remove(dst)
    try:
        os.link(shard_path, dst)          # 硬链，不占额外磁盘
    except OSError:
        shutil.copy2(shard_path, dst)     # 跨卷时回退为复制
    return d, db


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--task", required=True,
                    choices=["merge_col", "split_col", "value_shift"])
    ap.add_argument("--mode", default="", help="profile/shape；留空按任务默认")
    a = ap.parse_args()

    man = [json.loads(l) for l in
           open(os.path.join(BASE, "data", f"manifest_ext_{a.task}.jsonl"),
                encoding="utf-8")]
    tags = sorted({os.path.basename(r["ext_db_path"])[:-len(".sqlite")]
                   for r in man})
    os.makedirs(OUTDIR, exist_ok=True)

    modes = ([a.mode] if a.mode else
             (["profile", "shape"] if a.task == "value_shift" else ["profile"]))
    done = {"profile": 0, "shape": 0, "part": 0}

    for tag in tags:
        shard = os.path.join(EXT_DIR, f"{tag}.sqlite")
        if not os.path.exists(shard):
            print(f"[skip] 缺分片 {tag}", flush=True)
            continue
        tmpd, db = shard_to_alias(shard, tag)

        for mode in modes:
            # 只改路径映射，不改抽取逻辑
            M1.EVO_DIR[a.task] = tmpd
            obj = M1.infer_db(db, a.task, mode=mode)
            p = os.path.join(OUTDIR, f"{a.task}_{tag}_{mode}.json")
            obj["shard"] = tag
            json.dump(obj, open(p, "w", encoding="utf-8"),
                      ensure_ascii=False, indent=1)
            done[mode] += 1

        if a.task == "merge_col":
            # 缺陷修复（第 22 轮）：m1_part.infer_db 读的是 os.environ["EVO_DIR"]，
            # 不是模块属性 MP.EVO_DIR。只赋模块属性等于没赋，函数会退回默认目录
            # data/evolved_merge/<db>.sqlite——即「主实验那个合并库」。
            # 后果：28 个分片里有 18 个拿到的是另一张表的段位谱系（223/479 题），
            # 表达式被 render() 按旧列名过滤后为空，c3mp 退化成 c2 并大量弃答。
            # 两个都设，避免再次依赖某个模块的具体实现细节。
            os.environ["EVO_DIR"] = tmpd
            MP.EVO_DIR = tmpd
            assert os.path.basename(tmpd) == tag, "临时目录未指向当前分片"
            part = MP.infer_db(db) or {}
            # 硬约束：推断出的表必须就是本分片的表，否则宁可报错也不要写错文件
            shard_tab = tag.split("__")[1]
            bad = [t for t in part if t != shard_tab]
            assert not bad, f"{tag} 推断到别的表 {bad}，拒绝写入"
            # 转成与主实验 data/lineage_infer/<evo>_<db>_part.json 完全相同的
            # 结构（ops 列表 + to_expr 渲染），否则评估侧读不到段位表达式。
            ops = []
            for tab, d in part.items():
                for mc, meta in d.items():
                    for oc, pi in meta["parts"].items():
                        ops.append({
                            "type": "merge_columns", "table": tab, "old": oc,
                            "new": MP.to_expr(mc, pi["idx"], pi["nseg"], meta["sep"]),
                            "cols": [mc], "score": pi["score"],
                            "part_idx": pi["idx"], "nseg": pi["nseg"],
                        })
            p = os.path.join(OUTDIR, f"{a.task}_{tag}_part.json")
            json.dump({"db_id": db, "evo_type": a.task, "mode": "part",
                       "shard": tag, "sep_detected": True, "ops": ops},
                      open(p, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
            done["part"] += 1
        print(f"[ok] {tag} modes={modes}", flush=True)

    print(f"\n完成：分片 {len(tags)}  产物 {done}")
    shutil.rmtree(TMP, ignore_errors=True)


if __name__ == "__main__":
    main()
