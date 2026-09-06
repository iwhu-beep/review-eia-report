"""
从 Word (.docx) / PDF 中提取文本内容，用于环评报告表审核。

用法：
    python scripts/extract_docx.py "<文件路径>" [--format md|json|txt] [--out 结果文件]
    python scripts/extract_docx.py 报告.pdf --scan-check          # 仅做扫描件预检
    python scripts/extract_docx.py 报告.docx --split-chapters --split-out chaps/   # 按章节切分

功能：
  - 扫描件 PDF 自动识别（平均每页字符过少即告警并提示 OCR）；
  - 按章节切分输出，配合 SKILL.md 长文档分块审核策略；
  - 全链路异常捕获：损坏文件/加密PDF/权限问题均给出中文指引而非堆栈。

向后兼容：python scripts/extract_docx.py file.docx 仍可工作（默认 md 输出到 stdout）。

依赖：pip install -r scripts/requirements.txt
"""

import argparse
import datetime
import json
import os
import re
import sys
import traceback


SCAN_MIN_CHARS_PER_PAGE = 50  # 低于此值判疑似扫描件


def eprint(*args, **kwargs):
    print(*args, file=sys.stderr, **kwargs)


def load_docx(path):
    try:
        from docx import Document
    except ImportError:
        eprint("错误：缺少 python-docx，请执行 pip install -r scripts/requirements.txt")
        sys.exit(2)
    try:
        return Document(path)
    except Exception as exc:
        eprint(f"错误：Word 文档打不开（可能损坏或被加密）：{exc}")
        sys.exit(2)


def extract_docx_meta(doc, path):
    props = doc.core_properties
    try:
        meta = {
            "file": os.path.basename(path),
            "title": getattr(props, "title", "") or "",
            "author": getattr(props, "author", "") or "",
            "created": str(getattr(props, "created", "") or ""),
            "modified": str(getattr(props, "modified", "") or ""),
            "paragraph_count": len(doc.paragraphs),
            "table_count": len(doc.tables),
            "extracted_at": datetime.datetime.now().isoformat(timespec="seconds"),
        }
    except Exception as exc:
        eprint(f"警告：文档元数据读取异常（{exc}），已用空值代替")
        meta = {"file": os.path.basename(path)}
    return meta


HEADING_PREFIXES = (
    "一、", "二、", "三、", "四、", "五、", "六、", "七、", "八、", "九、", "十、",
    "十一、", "十二、", "概述", "结论",
)
HEADING_RE = re.compile(r"^(第[一二三四五六七八九十\d]+[章节条]|(\d+\.)+\d*\s*\S|附件|附图|附表)")


def is_heading_text(text):
    text = text.strip()
    if not text or len(text) > 60:
        return False
    if any(text.startswith(p) for p in HEADING_PREFIXES):
        return True
    if HEADING_RE.match(text):
        return True
    return False


def is_heading(paragraph):
    try:
        style_name = paragraph.style.name if paragraph.style else ""
    except Exception:
        style_name = ""
    if "Heading" in style_name or "标题" in style_name:
        return True
    try:
        text = paragraph.text.strip()
    except Exception:
        return False
    return is_heading_text(text)


def extract_docx(path, include_tables=True):
    doc = load_docx(path)
    meta = extract_docx_meta(doc, path)

    paragraphs = []
    for para in doc.paragraphs:
        try:
            text = para.text.strip()
        except Exception:
            continue
        if not text:
            continue
        paragraphs.append({"text": text, "is_heading": is_heading(para)})

    headers_footers = []
    try:
        for section in doc.sections:
            for part in (section.header, section.footer):
                for para in part.paragraphs:
                    t = para.text.strip()
                    if t and t not in headers_footers:
                        headers_footers.append(t)
    except Exception as exc:
        eprint(f"提示：页眉页脚提取跳过（{exc}）")

    tables = []
    if include_tables:
        try:
            for t_idx, table in enumerate(doc.tables):
                rows = []
                for row in table.rows:
                    cells = []
                    for cell in row.cells:
                        try:
                            c = cell.text.strip().replace("\n", " / ").replace("\r", "")
                        except Exception:
                            c = ""
                        cells.append(c)
                    rows.append(cells)
                tables.append({"index": t_idx + 1, "rows": rows, "n_rows": len(rows)})
        except Exception as exc:
            eprint(f"警告：部分表格提取失败，已跳过（{exc}）")

    comments = []
    try:
        for rel in doc.part.rels.values():
            if "comments" in (rel.target_ref or ""):
                comments.append(f"文档含批注部件：{rel.target_ref}（请在 Word 中查看批注）")
                break
    except Exception:
        pass

    return {"meta": meta, "headers_footers": headers_footers, "paragraphs": paragraphs, "tables": tables, "comments": comments}


def extract_pdf(path, include_tables=True):
    try:
        from pypdf import PdfReader
    except ImportError:
        eprint("错误：PDF 解析需要 pypdf，请执行 pip install pypdf 后重试；扫描件 PDF 请先 OCR。")
        sys.exit(2)
    try:
        reader = PdfReader(path)
    except Exception as exc:
        msg = str(exc).lower()
        if "encrypt" in msg:
            eprint(f"错误：PDF 已加密，请先解密后再提取（{exc}）")
        else:
            eprint(f"错误：PDF 打不开（可能损坏）：{exc}")
        sys.exit(2)
    try:
        npages = len(reader.pages)
    except Exception as exc:
        eprint(f"错误：PDF 页面读取失败：{exc}")
        sys.exit(2)
    paragraphs, page_chars = [], []
    for i, page in enumerate(reader.pages):
        try:
            text = page.extract_text() or ""
        except Exception as exc:
            eprint(f"提示：第 {i + 1} 页解析失败，已跳过（{exc}）")
            page_chars.append(0)
            continue
        page_chars.append(len(text.strip()))
        for line in text.splitlines():
            line = line.strip()
            if line:
                paragraphs.append({"text": line, "is_heading": is_heading_text(line), "page": i + 1})
    avg = sum(page_chars) / max(npages, 1)
    scan_suspected = npages > 0 and avg < SCAN_MIN_CHARS_PER_PAGE
    meta = {
        "file": os.path.basename(path),
        "pages": npages,
        "avg_chars_per_page": round(avg, 1),
        "scan_suspected": scan_suspected,
        "paragraph_count": len(paragraphs),
        "table_count": 0,
        "note": "PDF 表格按文本行提取，未做结构化还原" if include_tables else "",
        "extracted_at": datetime.datetime.now().isoformat(timespec="seconds"),
    }
    if scan_suspected:
        eprint(
            f"【扫描件告警】该 PDF 平均每页仅 {avg:.0f} 字符（阈值 {SCAN_MIN_CHARS_PER_PAGE}），"
            "疑似扫描件/图片型 PDF，文本提取不完整。请先 OCR（如 ocrmypdf / WPS 文字识别）后再提取审核。"
        )
    return {"meta": meta, "headers_footers": [], "paragraphs": paragraphs, "tables": [], "comments": []}


def split_into_chapters(data):
    """按标题段落切分章节；表格统一归入末章"附表"。返回 [(title, [paragraphs])]。"""
    chapters, cur_title, cur = [], "前言", []
    for p in data.get("paragraphs", []):
        if p.get("is_heading") and cur and not (len(cur) == 0):
            chapters.append((cur_title, cur))
            cur_title, cur = p["text"][:30], []
        elif p.get("is_heading") and not cur:
            cur_title = p["text"][:30]
            continue
        cur.append(p)
    if cur or not chapters:
        chapters.append((cur_title, cur))
    if data.get("tables") and len(chapters) > 1:
        pass  # 表格另起"附表"章，由调用方处理
    return chapters


def to_markdown(data):
    out = []
    meta = data.get("meta", {})
    out.append(f"# 文档提取结果：{meta.get('file', '')}\n")
    out.append("## 元数据")
    for k, v in meta.items():
        out.append(f"- {k}: {v}")
    if meta.get("scan_suspected"):
        out.append("- 警告：疑似扫描件，须 OCR 后重新提取")
    if data.get("headers_footers"):
        out.append("\n## 页眉页脚")
        for h in data["headers_footers"]:
            out.append(f"- {h}")
    out.append("\n=== 正文内容 ===\n")
    for p in data.get("paragraphs", []):
        t = p["text"]
        out.append(f"\n【{t}】" if p.get("is_heading") else t)
    if data.get("tables"):
        out.append("\n\n=== 表格内容 ===\n")
        for tb in data["tables"]:
            out.append(f"\n--- 表格 {tb['index']}（{tb['n_rows']}行） ---")
            for row in tb["rows"]:
                out.append(" | ".join(row))
    if data.get("comments"):
        out.append("\n\n=== 批注提示 ===\n")
        out.extend(f"- {c}" for c in data["comments"])
    return "\n".join(out)


def to_text(data):
    lines = []
    for p in data.get("paragraphs", []):
        lines.append(p["text"])
    for tb in data.get("tables", []):
        lines.append(f"[表格{tb['index']}]")
        for row in tb["rows"]:
            lines.append(" | ".join(row))
    return "\n".join(lines)


def write_chapters(data, out_dir, fmt):
    """章节切分输出：ch_01_标题.md + index。表格单独成"附表"章。"""
    try:
        os.makedirs(out_dir, exist_ok=True)
    except OSError as exc:
        eprint(f"错误：输出目录创建失败 {out_dir}（{exc}）")
        sys.exit(2)
    chapters = split_into_chapters(data)
    index, n = [], 0
    for title, paras in chapters:
        n += 1
        safe = re.sub(r'[\\/:*?"<>|]', "", title).strip()[:20] or f"第{n}章"
        fn = os.path.join(out_dir, f"ch_{n:02d}_{safe}.md")
        body = [f"# {title}", ""]
        for p in paras:
            t = p["text"]
            body.append(f"## {t}" if p.get("is_heading") and t != title else t)
        content = "\n".join(body) if fmt != "txt" else "\n".join(p["text"] for p in paras)
        try:
            with open(fn, "w", encoding="utf-8") as f:
                f.write(content)
        except OSError as exc:
            eprint(f"错误：章节文件写入失败 {fn}（{exc}）")
            sys.exit(2)
        index.append((fn, len(paras)))
    if data.get("tables"):
        n += 1
        fn = os.path.join(out_dir, f"ch_{n:02d}_附表.md")
        body = ["# 附表（表格单独完整保留，不截断）", ""]
        for tb in data["tables"]:
            body.append(f"\n--- 表格 {tb['index']}（{tb['n_rows']}行） ---")
            for row in tb["rows"]:
                body.append(" | ".join(row))
        try:
            with open(fn, "w", encoding="utf-8") as f:
                f.write("\n".join(body))
        except OSError as exc:
            eprint(f"错误：附表文件写入失败 {fn}（{exc}）")
            sys.exit(2)
        index.append((fn, sum(t["n_rows"] for t in data["tables"])))
    print(f"章节切分完成：{len(index)} 个文件 → {out_dir}")
    for fn, cnt in index:
        print(f"  - {os.path.basename(fn)}（{cnt}段/行）")
    print("用法：按 SKILL.md 长文档策略逐章审核，表格章单独完整载入。")


def parse_args(argv=None):
    ap = argparse.ArgumentParser(description="提取环评报告表文档文本（docx/pdf，含扫描预检与章节切分）")
    ap.add_argument("input", help="输入文件路径（.docx/.pdf）")
    ap.add_argument("--format", choices=["md", "json", "txt"], default="md", help="输出格式（默认md）")
    ap.add_argument("--out", default="", help="输出文件路径，缺省输出到stdout")
    ap.add_argument("--max-chars", type=int, default=0, help="截断提示：超过N字符时在末尾追加分块提示（0=不截断）")
    ap.add_argument("--include-tables", dest="include_tables", action="store_true", default=True)
    ap.add_argument("--no-tables", dest="include_tables", action="store_false", help="不提取表格")
    ap.add_argument("--scan-check", action="store_true", help="仅做扫描件预检（PDF），不输出全文")
    ap.add_argument("--split-chapters", action="store_true", help="按章节切分输出到目录（长文档分块审核）")
    ap.add_argument("--split-out", default="", help="章节输出目录（默认 <输入名>_chapters）")
    return ap.parse_args(argv)


def main(argv=None):
    try:
        args = parse_args(argv)
        path = args.input

        if not os.path.exists(path):
            eprint(f"错误：文件不存在 - {path}")
            return 2
        if not os.path.isfile(path):
            eprint(f"错误：不是有效文件 - {path}")
            return 2

        ext = os.path.splitext(path)[1].lower()
        if ext == ".docx":
            data = extract_docx(path, include_tables=args.include_tables)
        elif ext == ".pdf":
            data = extract_pdf(path, include_tables=args.include_tables)
        elif ext == ".doc":
            eprint("错误：.doc 为老格式，请先用 Word/WPS 另存为 .docx 后再提取。")
            return 2
        else:
            eprint(f"错误：不支持的文件类型 {ext}，仅支持 .docx/.pdf。")
            return 2

        if args.scan_check:
            m = data.get("meta", {})
            if m.get("scan_suspected"):
                print(f"预检结论：疑似扫描件（平均每页 {m.get('avg_chars_per_page')} 字符），请先 OCR。")
                return 4
            print(f"预检结论：文本型 PDF（平均每页 {m.get('avg_chars_per_page', '?')} 字符），可直接提取。")
            return 0

        if not data.get("paragraphs") and not data.get("tables"):
            eprint("错误：文档为空或无法提取到文本。")
            if ext == ".pdf":
                eprint("扫描件 PDF 请先 OCR（如 ocrmypdf / WPS 文字识别）后再提取。")
            return 2

        if args.split_chapters:
            if args.format == "json":
                eprint("提示：--split-chapters 仅支持 md/txt，已自动按 md 切分。")
            out_dir = args.split_out or os.path.splitext(path)[0] + "_chapters"
            write_chapters(data, out_dir, args.format)
            return 0

        try:
            if args.format == "json":
                output = json.dumps(data, ensure_ascii=False, indent=2)
            elif args.format == "txt":
                output = to_text(data)
            else:
                output = to_markdown(data)
        except Exception as exc:
            eprint(f"错误：输出格式化失败（{exc}）")
            return 2

        if args.max_chars and len(output) > args.max_chars:
            output = output[: args.max_chars]
            output += (
                f"\n\n…[截断：全文超 {args.max_chars} 字符，已截断。"
                "建议用 --split-chapters 按章节切分审核，表格章单独完整提取。]"
            )

        if args.out:
            try:
                with open(args.out, "w", encoding="utf-8") as f:
                    f.write(output)
            except OSError as exc:
                eprint(f"错误：结果文件写入失败 {args.out}（{exc}）")
                return 2
            print(f"已写入：{args.out}（{len(output)}字符）")
        else:
            print(output)
        if data.get("meta", {}).get("scan_suspected"):
            return 4  # 疑似扫描件：给出告警退出码，提醒 OCR
        return 0
    except Exception:
        eprint("错误：提取过程异常：")
        eprint(traceback.format_exc(limit=3))
        return 3


if __name__ == "__main__":
    sys.exit(main())
