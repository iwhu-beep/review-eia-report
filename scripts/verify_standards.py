"""
标准时效初筛脚本：从报告文本中识别标准号，对照本地库判现行/过期/未知。

用法：
    python scripts/verify_standards.py <报告文本文件> [--db scripts/standards_db.json] [--format md|json] [--out 结果文件]

说明：
  - 本脚本只做本地初筛，不能替代官网复核；
  - status=疑似过期/未知/缺年号 的条目，必须按 regulations.md 第十二节到官网复核；
  - 仅依赖标准库，可离线运行。

依赖：无（Python3.8+ 标准库）。
"""

import argparse
import json
import os
import re
import sys


def eprint(*args, **kwargs):
    print(*args, file=sys.stderr, **kwargs)


# GB/HJ 标准号：GB 16297-1996、GB/T 31962-2015、HJ 2.2-2018、HJ 169-2018
STD_RE = re.compile(
    r"(GB(?:/T)?\s*\d{4,5}(?:\.\d+)?(?:\s*-\s*\d{4})?|HJ\s*\d+(?:\.\d+)?(?:\s*-\s*\d{4})?)"
)
# 国家危险废物名录（2021年版）/ 2016版 / 无年号
CATALOG_RE = re.compile(r"国家危险废物名录[（(]?(\d{4})?年?版?[）)]?")
# 产业结构调整指导目录 2024年本 / 2019年本
CATALOG2_RE = re.compile(r"产业结构调整指导目录[（(]?(\d{4})?年?本?[）)]?")


def norm(code):
    """归一化：去空格、全角转半角、大写。"""
    code = code.replace("（", "(").replace("）", ")").replace("／", "/")
    code = re.sub(r"\s+", "", code).upper()
    return code


def load_db(db_path):
    with open(db_path, "r", encoding="utf-8") as f:
        return json.load(f)


def extract_candidates(text):
    found = []  # (原文写法, 归一化, 类型)
    for m in STD_RE.finditer(text):
        raw = m.group(1)
        found.append((raw, norm(raw), "std"))
    for m in CATALOG_RE.finditer(text):
        year = m.group(1) or ""
        found.append((m.group(0), "危废名录" + year, "catalog"))
    for m in CATALOG2_RE.finditer(text):
        year = m.group(1) or ""
        found.append((m.group(0), "产调目录" + year, "catalog2"))
    # 去重保序
    seen, uniq = set(), []
    for item in found:
        if item[1] not in seen:
            seen.add(item[1])
            uniq.append(item)
    return uniq


def check_one(raw, code, kind, db_index):
    has_year = bool(re.search(r"-\d{4}$", code)) or kind in ("catalog", "catalog2") and any(
        c.isdigit() for c in code
    )
    if kind == "catalog":
        key = "危废名录2021" if "2021" in code else ("危废名录2016" if "2016" in code else "危废名录?")
    elif kind == "catalog2":
        key = None
        if "2024" in code:
            return verdict(raw, code, "现行", "产业结构调整指导目录2024年本", "现行版本", "")
        elif "2019" in code:
            return verdict(raw, code, "疑似过期", "产业结构调整指导目录2019年本", "已出2024年本，请官网复核", "")
        else:
            return verdict(raw, code, "缺年号", "产业结构调整指导目录（未注明版本）", "须补年号并核对2024年本", "")
    else:
        key = code
    if kind == "catalog" and key == "危废名录?":
        return verdict(raw, code, "缺年号", "国家危险废物名录（未注明版本）", "须补2021年版并核对代码", "")
    rec = db_index.get(key)
    if rec is None:
        if kind == "std" and not re.search(r"-\d{4}$", code):
            return verdict(raw, code, "缺年号", "标准号未带年号", "须补年号后复核", "")
        return verdict(raw, code, "未知", "本地库无此条目", "须到官网核实是否为现行", "")
    status = rec["status"]
    if status == "现行" or status == "现行试行":
        return verdict(raw, code, "现行", f"{rec['name']}", rec.get("note", ""), "")
    return verdict(raw, code, "疑似过期", f"{rec['name']}", f"{rec.get('date','')}; {rec.get('note','')}".strip("; "), "")


def verdict(raw, code, status, name, note, url):
    return {"raw": raw, "code": code, "status": status, "name": name, "note": note}


def to_markdown(results):
    lines = ["# 标准时效初筛结果", "", "（本地库初筛；疑似过期/未知/缺年号须官网复核，见 regulations.md 第十二节）", ""]
    lines.append("| 序号 | 报告写法 | 状态 | 对应标准 | 说明 |")
    lines.append("|------|----------|------|----------|------|")
    for i, r in enumerate(results, 1):
        lines.append(f"| {i} | {r['raw']} | {r['status']} | {r['name']} | {r['note']} |")
    c = {"现行": 0, "疑似过期": 0, "未知": 0, "缺年号": 0}
    for r in results:
        c[r["status"]] = c.get(r["status"], 0) + 1
    lines.append("")
    lines.append(f"汇总：共 {len(results)} 项；现行 {c.get('现行',0)}；疑似过期 {c.get('疑似过期',0)}；未知 {c.get('未知',0)}；缺年号 {c.get('缺年号',0)}。")
    if c.get("疑似过期", 0) or c.get("未知", 0) or c.get("缺年号", 0):
        lines.append("复核入口：GB类→国家标准全文公开系统 openstd.samr.gov.cn；HJ/环境标准→生态环境部官网标准栏目。")
    return "\n".join(lines)


def parse_args(argv=None):
    ap = argparse.ArgumentParser(description="标准时效初筛（本地库，需官网复核）")
    ap.add_argument("input", help="报告文本文件（extract_docx.py 输出的 md/txt/json 均可）")
    ap.add_argument("--db", default="", help="标准库路径（默认 scripts/standards_db.json，同目录自动定位）")
    ap.add_argument("--format", choices=["md", "json"], default="md")
    ap.add_argument("--out", default="")
    return ap.parse_args(argv)


def main(argv=None):
    args = parse_args(argv)
    if not os.path.exists(args.input):
        eprint(f"错误：文件不存在 - {args.input}")
        sys.exit(2)
    db_path = args.db or os.path.join(os.path.dirname(os.path.abspath(__file__)), "standards_db.json")
    if not os.path.exists(db_path):
        eprint(f"错误：标准库不存在 - {db_path}")
        sys.exit(2)
    db = load_db(db_path)
    db_index = {norm(s["code"]): s for s in db.get("standards", [])}

    with open(args.input, "r", encoding="utf-8") as f:
        text = f.read()
    cands = extract_candidates(text)
    if not cands:
        eprint("提示：未识别到任何标准号（GB/HJ/名录/目录），请确认输入的是报告文本。")
        sys.exit(2)
    results = [check_one(raw, code, kind, db_index) for raw, code, kind in cands]

    output = json.dumps({"results": results}, ensure_ascii=False, indent=2) if args.format == "json" else to_markdown(results)
    if args.out:
        with open(args.out, "w", encoding="utf-8") as f:
            f.write(output)
        print(f"已写入：{args.out}（{len(results)}项）")
    else:
        print(output)


if __name__ == "__main__":
    main()
