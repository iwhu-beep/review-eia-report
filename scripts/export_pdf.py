"""
审核意见 Markdown 转 PDF（中文归档版）：把 SKILL 第 5 步输出的审核意见导出为 PDF 存档/送审。

用法：
    python scripts/export_pdf.py 审核意见.md --out 审核意见.pdf
    python scripts/export_pdf.py 审核意见.md --project "XX项目" --company "XX公司" --author "审核人"

说明：
  - 自动识别输入编码（UTF-8/GBK），根治中文乱码；
  - 中文字体自动发现：--font > EIA_PDF_FONT 环境变量 > 系统常见中文字体；
  - PDF 自带归档元信息（标题/作者/主题/关键词）、页眉（标题+日期）、页脚（页码+复核提示）；
  - 自动提取"项目信息"节生成归档信息表；末尾强制附风险声明（AI辅助，须人工复核）；
  - 支持标题/表格/列表/引用/代码块/分隔线的基础渲染。

依赖：fpdf2（pip install fpdf2）。
"""

import argparse
import datetime
import os
import re
import sys
import traceback

try:
    from fpdf.fonts import FontFace
except ImportError:  # 极旧版 fpdf2 兼容
    FontFace = None


DISCLAIMER = (
    "风险声明：本文件由 AI 辅助生成，仅供环评技术审查参考，不能替代注册环境影响评价工程师的 "
    "专业判断和法定审批程序；其中引用的法规标准状态以生态环境部、市场监管总局官网现行版本为准； "
    "全部审核意见须经人工复核确认后方可作为整改或报批依据。"
)

# PDF 常用中文字体缺字形的符号 → 文字兜底
GLYPH_FALLBACK = {
    "⛔": "[一票否决]",
    "★": "[必查]",
    "☆": "[抽查]",
    "☐": "[ ]",
    "☑": "[x]",
    "☒": "[x]",
    "✅": "[通过]",
    "❌": "[不通过]",
    "⚠": "[注意]",
    "➡": "->",
    "⬅": "<-",
}


def eprint(*args, **kwargs):
    print(*args, file=sys.stderr, **kwargs)


def safe_print(text):
    """控制台编码兜底（GBK 下特殊字符会炸 print，改走 UTF-8 字节流）。"""
    try:
        print(text)
    except UnicodeEncodeError:
        try:
            sys.stdout.buffer.write((text + "\n").encode("utf-8", errors="replace"))
            sys.stdout.buffer.flush()
        except Exception:
            print(text.encode(errors="replace"))


def safe_text(text):
    for k, v in GLYPH_FALLBACK.items():
        text = text.replace(k, v)
    # 去除控制字符（保留换行/制表）
    return re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]", "", text)


def read_text_smart(path):
    """智能编码读取：UTF-8 优先，失败回退 GB18030/GBK（中文 Windows 常见乱码根源）。"""
    last_err = None
    for enc in ("utf-8-sig", "utf-8", "gb18030", "gbk"):
        try:
            with open(path, "r", encoding=enc) as f:
                return f.read(), enc
        except (UnicodeDecodeError, UnicodeError) as exc:
            last_err = exc
        except OSError as exc:
            raise RuntimeError(f"无法读取文件 {path}：{exc}")
    raise RuntimeError(f"文件编码无法识别（已试 utf-8/gbk）：{last_err}")


FONT_CANDIDATES = [
    r"C:\Windows\Fonts\simsun.ttc",      # 宋体（公文标准首选）
    r"C:\Windows\Fonts\msyh.ttc",        # 微软雅黑
    r"C:\Windows\Fonts\simhei.ttf",      # 黑体
    r"C:\Windows\Fonts\simkai.ttf",      # 楷体
    r"/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
    r"/usr/share/fonts/truetype/noto/NotoSansSC-Regular.otf",
    r"/System/Library/Fonts/PingFang.ttc",
    r"/System/Library/Fonts/STHeiti Medium.ttc",
]


def find_font(explicit=""):
    cands = []
    if explicit:
        cands.append(explicit)
    env = os.environ.get("EIA_PDF_FONT", "")
    if env:
        cands.append(env)
    cands.extend(FONT_CANDIDATES)
    tried = []
    for c in cands:
        if not c or not os.path.exists(c):
            continue
        try:
            from fpdf import FPDF
            p = FPDF()
            p.add_font("cjk", "", c)
            return c
        except Exception as exc:
            tried.append(f"{c}（不可用：{exc}）")
    eprint("错误：未找到可用的中文字体。已尝试：")
    for t in tried or cands:
        eprint(f"  - {t}")
    eprint("解决：用 --font 指定字体文件，或设置环境变量 EIA_PDF_FONT。")
    sys.exit(2)


def clean_inline(text):
    """去装饰符号：加粗/行内代码/链接转纯文本。"""
    text = safe_text(text)
    text = re.sub(r"\*\*(.+?)\*\*", r"\1", text)
    text = re.sub(r"`(.+?)`", r"\1", text)
    text = re.sub(r"\[(.+?)\]\(.+?\)", r"\1", text)
    return text.strip()


def parse_blocks(lines):
    """Markdown 行 → 块列表：heading/para/list-item/table/quote/code/hr。"""
    blocks, i, n = [], 0, len(lines)
    while i < n:
        line = lines[i].rstrip()
        s = line.strip()
        if not s:
            i += 1
            continue
        if s.startswith("```"):
            buf, i = [], i + 1
            while i < n and not lines[i].strip().startswith("```"):
                buf.append(safe_text(lines[i].rstrip()))
                i += 1
            i += 1
            blocks.append(("code", buf))
            continue
        m = re.match(r"^(#{1,4})\s+(.*)", s)
        if m:
            blocks.append(("h" + str(len(m.group(1))), clean_inline(m.group(2))))
            i += 1
            continue
        if re.match(r"^(\*{3,}|-{3,}|_{3,})\s*$", s):
            blocks.append(("hr", None))
            i += 1
            continue
        if s.startswith("|") and s.endswith("|"):
            rows, j = [], i
            while j < n and lines[j].strip().startswith("|"):
                cells = [clean_inline(c) for c in lines[j].strip().strip("|").split("|")]
                if all(re.match(r"^:?-{2,}:?$", c.replace(" ", "")) for c in cells):
                    j += 1
                    continue  # 分隔行跳过
                rows.append(cells)
                j += 1
            if rows:
                blocks.append(("table", rows))
            i = j
            continue
        m = re.match(r"^(\d+)[.)]\s+(.*)", s)
        if m:
            blocks.append(("ol", (m.group(1), clean_inline(m.group(2)))))
            i += 1
            continue
        if re.match(r"^[-*+]\s+", s):
            blocks.append(("ul", clean_inline(re.sub(r"^[-*+]\s+", "", s))))
            i += 1
            continue
        if s.startswith(">"):
            blocks.append(("quote", clean_inline(s.lstrip(">").strip())))
            i += 1
            continue
        buf, j = [s], i + 1
        while j < n and lines[j].strip() and not re.match(
            r"^(#{1,4}\s+|\||\d+[.)]\s+|[-*+]\s+|>|```|\*{3,}|-{3,})", lines[j].strip()
        ):
            buf.append(lines[j].strip())
            j += 1
        blocks.append(("para", clean_inline(" ".join(buf))))
        i = j
    return blocks


def extract_project_info(blocks):
    """从"项目信息"节提取（字段，值）行；返回 (rows, skip_idx_set)。"""
    rows, skip = [], set()
    for idx, (kind, data) in enumerate(blocks):
        if kind in ("h1", "h2", "h3", "h4") and "项目信息" in data:
            skip.add(idx)
            j = idx + 1
            while j < len(blocks) and blocks[j][0] not in ("h1", "h2", "h3", "h4"):
                k, d = blocks[j]
                text = d if k in ("para", "ul", "quote") else (d[1] if k == "ol" else "")
                text = re.sub(r"^[-*•\d.)\s]+", "", text or "")
                if "：" in text or ":" in text:
                    sep = "：" if "：" in text else ":"
                    f, v = text.split(sep, 1)
                    rows.append((f.strip("-*• ").strip(), v.strip()))
                    skip.add(j)
                j += 1
            break
    return rows, skip


class AuditPDF:
    def __init__(self, font_path, title, author, keywords):
        from fpdf import FPDF

        disclaimer_short = "AI辅助生成，须经人工复核"

        class PDF(FPDF):
            def header(pdf):
                if pdf.page_no() == 1:
                    return  # 首页已有大标题，页眉从第2页起
                pdf.set_font("cjk", "", 9)
                pdf.set_text_color(120, 120, 120)
                pdf.cell(0, 8, title, align="L")
                pdf.cell(
                    0, 8, datetime.date.today().isoformat(), align="R",
                    new_x="LMARGIN", new_y="NEXT",
                )
                pdf.set_draw_color(200, 200, 200)
                pdf.line(pdf.l_margin, pdf.get_y(), pdf.w - pdf.r_margin, pdf.get_y())
                pdf.ln(2)
                pdf.set_text_color(0, 0, 0)

            def footer(pdf):
                pdf.set_y(-15)
                pdf.set_font("cjk", "", 9)
                pdf.set_text_color(120, 120, 120)
                pdf.cell(0, 10, f"第 {pdf.page_no()} 页 / " + "{nb}　" + disclaimer_short, align="C")

        self.pdf = PDF()
        self.pdf.alias_nb_pages("{nb}")
        self.pdf.set_auto_page_break(True, margin=20)
        self.pdf.add_font("cjk", "", font_path)
        # 归档元信息
        self.pdf.set_title(title)
        self.pdf.set_author(author or "review-eia-report")
        self.pdf.set_subject("建设项目环境影响报告表审核意见")
        self.pdf.set_keywords(keywords or "环评,报告表审核,AI辅助（须人工复核）")
        self.pdf.set_creator("review-eia-report scripts/export_pdf.py")
        self.title = title
        self.w = self.pdf.w - self.pdf.l_margin - self.pdf.r_margin

    def render(self, blocks, source_name, encoding, project_rows, skip_idx):
        pdf = self.pdf
        pdf.add_page()
        pdf.set_font("cjk", "", 20)
        pdf.multi_cell(0, 10, self.title, align="C")
        pdf.set_x(pdf.l_margin)
        pdf.set_font("cjk", "", 10)
        pdf.set_text_color(100, 100, 100)
        pdf.cell(
            0, 8,
            f"来源：{source_name}（{encoding}）  导出日期：{datetime.date.today().isoformat()}",
            align="C", new_x="LMARGIN", new_y="NEXT",
        )
        pdf.set_text_color(0, 0, 0)
        pdf.ln(2)
        if project_rows:
            pdf.set_font("cjk", "", 13)
            pdf.set_fill_color(242, 242, 242)
            pdf.multi_cell(0, 8, "项目信息（归档）", fill=True)
            pdf.set_font("cjk", "", 11)
            kwargs = {}
            if FontFace is not None:
                kwargs["headings_style"] = FontFace(fill_color=(240, 240, 240))
            with pdf.table(width=self.w, col_widths=(self.w * 0.3, self.w * 0.7),
                           text_align="LEFT", line_height=6.5, **kwargs) as table:
                for f, v in project_rows:
                    tr = table.row()
                    tr.cell(f)
                    tr.cell(v)
            pdf.ln(2)
        for idx, (kind, data) in enumerate(blocks):
            if idx in skip_idx:
                continue
            try:
                pdf.set_x(pdf.l_margin)
                self._block(pdf, kind, data)
            except Exception as exc:
                eprint(f"警告：第 {idx + 1} 个内容块渲染跳过（{exc}）")

    def _block(self, pdf, kind, data):
        if kind in ("h1", "h2", "h3", "h4"):
            size = {"h1": 17, "h2": 15, "h3": 13, "h4": 12}[kind]
            pdf.ln(2)
            pdf.set_fill_color(242, 242, 242)
            pdf.set_font("cjk", "", size)
            pdf.multi_cell(0, 8, data, fill=(kind in ("h1", "h2")))
            pdf.ln(1)
        elif kind == "para":
            pdf.set_font("cjk", "", 11)
            pdf.multi_cell(0, 6.5, data)
            pdf.ln(1)
        elif kind in ("ul", "ol"):
            pdf.set_font("cjk", "", 11)
            prefix = "• " if kind == "ul" else f"{data[0]}. "
            body = data if kind == "ul" else data[1]
            pdf.set_x(pdf.get_x() + 4)
            pdf.multi_cell(self.w - 4, 6.5, prefix + body)
        elif kind == "quote":
            pdf.set_font("cjk", "", 11)
            pdf.set_text_color(80, 80, 80)
            pdf.set_x(pdf.get_x() + 4)
            pdf.multi_cell(self.w - 4, 6.5, "| " + data)
            pdf.set_text_color(0, 0, 0)
        elif kind == "code":
            pdf.set_font("cjk", "", 10)
            pdf.set_fill_color(245, 245, 245)
            for cl in data or ["(空)"]:
                pdf.set_x(pdf.get_x() + 4)
                pdf.multi_cell(self.w - 4, 6, cl if cl else " ", fill=True)
            pdf.ln(1)
        elif kind == "hr":
            pdf.ln(2)
            pdf.set_draw_color(180, 180, 180)
            pdf.line(pdf.l_margin, pdf.get_y(), pdf.w - pdf.r_margin, pdf.get_y())
            pdf.ln(2)
        elif kind == "table":
            self._table(data)

    def _table(self, rows):
        pdf = self.pdf
        if not rows:
            return
        ncol = max(len(r) for r in rows)
        rows = [r + [""] * (ncol - len(r)) for r in rows]
        col_w = self.w / ncol
        pdf.set_font("cjk", "", 10)
        kwargs = {"first_row_as_headings": True}
        if FontFace is not None:
            kwargs["headings_style"] = FontFace(fill_color=(230, 230, 230))
        with pdf.table(
            width=self.w,
            col_widths=tuple([col_w] * ncol),
            text_align="LEFT",
            line_height=6,
            **kwargs,
        ) as table:
            for row in rows:
                tr = table.row()
                for cell in row:
                    tr.cell(cell)
        pdf.set_x(pdf.l_margin)
        pdf.ln(2)

    def render_disclaimer(self):
        pdf = self.pdf
        pdf.ln(2)
        pdf.set_x(pdf.l_margin)
        pdf.set_font("cjk", "", 13)
        pdf.set_fill_color(255, 243, 224)
        pdf.multi_cell(0, 8, "风险声明（必读）", fill=True)
        pdf.set_x(pdf.l_margin)  # multi_cell 后光标在行尾，复位后再写正文
        pdf.set_font("cjk", "", 11)
        pdf.multi_cell(0, 6.5, DISCLAIMER)
        pdf.ln(1)

    def save(self, out):
        outdir = os.path.dirname(os.path.abspath(out))
        if outdir and not os.path.exists(outdir):
            os.makedirs(outdir, exist_ok=True)
        try:
            self.pdf.output(out)
        except OSError as exc:
            raise RuntimeError(f"PDF 写入失败（检查路径/权限/文件是否被占用）：{exc}")


def parse_args(argv=None):
    ap = argparse.ArgumentParser(description="审核意见 Markdown 转 PDF（中文归档版）")
    ap.add_argument("input", help="输入 Markdown 文件（审核意见）")
    ap.add_argument("--out", default="", help="输出 PDF 路径（默认与输入同名 .pdf）")
    ap.add_argument("--title", default="环境影响报告表审核意见", help="PDF 标题")
    ap.add_argument("--author", default="", help="作者/审核人（写入PDF元信息）")
    ap.add_argument("--project", default="", help="项目名称（归档信息表首行，缺省从正文提取）")
    ap.add_argument("--company", default="", help="建设单位（归档信息表，缺省从正文提取）")
    ap.add_argument("--font", default="", help="中文字体文件路径（缺省自动发现）")
    return ap.parse_args(argv)


def main(argv=None):
    try:
        args = parse_args(argv)
        if not os.path.exists(args.input):
            eprint(f"错误：文件不存在 - {args.input}")
            return 2
        if args.input.lower().endswith(".json"):
            eprint("错误：暂不支持 JSON 输入，请先用 Markdown 版审核意见导出。")
            return 2
        try:
            from fpdf import FPDF  # noqa: F401
        except ImportError:
            eprint("错误：缺少 fpdf2，请执行 pip install fpdf2")
            return 2
        try:
            text, encoding = read_text_smart(args.input)
        except RuntimeError as exc:
            eprint(f"错误：{exc}")
            return 2
        if not text.strip():
            eprint("错误：输入文件为空。")
            return 2
        font_path = find_font(args.font)
        blocks = parse_blocks(text.splitlines())
        project_rows, skip_idx = extract_project_info(blocks)
        if args.project:
            project_rows = [("项目名称", args.project)] + [r for r in project_rows if r[0] != "项目名称"]
        if args.company:
            project_rows = [r for r in project_rows if r[0] != "建设单位"] + [("建设单位", args.company)]
        out = args.out or os.path.splitext(args.input)[0] + ".pdf"
        doc = AuditPDF(font_path, args.title, args.author, keywords="")
        doc.render(blocks, os.path.basename(args.input), encoding, project_rows, skip_idx)
        doc.render_disclaimer()
        try:
            doc.save(out)
        except RuntimeError as exc:
            eprint(f"错误：{exc}")
            return 2
        safe_print(f"已生成：{out}（编码{encoding}，字体{os.path.basename(font_path)}）")
        return 0
    except Exception:
        eprint("错误：导出过程异常：")
        eprint(traceback.format_exc(limit=3))
        return 3


if __name__ == "__main__":
    sys.exit(main())
