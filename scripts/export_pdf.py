"""
审核意见 Markdown 转 PDF（中文）：把 SKILL 第 5 步输出的审核意见导出为 PDF 存档/送审。

用法：
    python scripts/export_pdf.py <审核意见md文件> [--out 审核意见.pdf] [--title 标题] [--font 字体文件]

说明：
  - 支持标题/表格/列表/引用/代码块/分隔线的基础渲染（装饰性 Markdown 符号会被规整化）；
  - 中文字体自动发现：--font > EIA_PDF_FONT 环境变量 > 系统常见中文字体；
  - 依赖：pip install fpdf2（见 scripts/requirements.txt）。

依赖：fpdf2。
"""

import argparse
import datetime
import os
import re
import sys


def eprint(*args, **kwargs):
    print(*args, file=sys.stderr, **kwargs)


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
    """去装饰符号：加粗/行内代码/链接转纯文本；罕见符号转文字（PDF字体缺字形时保真）。"""
    text = text.replace("⛔", "[一票否决]").replace("★", "[必查]")
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
                buf.append(lines[i].rstrip().replace("⛔", "[一票否决]").replace("★", "[必查]"))
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
                if not re.match(r"^:?-{2,}:?$", cells[0].replace(" ", "")) or len(cells) > 1 and any(
                    not re.match(r"^:?-{2,}:?$", c.replace(" ", "")) for c in cells
                ):
                    # 分隔行（如 |---|---|）跳过
                    if all(re.match(r"^:?-{2,}:?$", c.replace(" ", "")) for c in cells):
                        j += 1
                        continue
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
        # 连续普通行合并为一段
        buf, j = [s], i + 1
        while j < n and lines[j].strip() and not re.match(
            r"^(#{1,4}\s+|\||\d+[.)]\s+|[-*+]\s+|>|```|\*{3,}|-{3,})", lines[j].strip()
        ):
            buf.append(lines[j].strip())
            j += 1
        blocks.append(("para", clean_inline(" ".join(buf))))
        i = j
    return blocks


class AuditPDF:
    def __init__(self, font_path, title):
        from fpdf import FPDF

        class PDF(FPDF):
            def footer(pdf):
                pdf.set_y(-15)
                pdf.set_font("cjk", "", 9)
                pdf.set_text_color(120, 120, 120)
                pdf.cell(0, 10, f"第 {pdf.page_no()} 页 / " + "{nb}", align="C")

        self.pdf = PDF()
        self.pdf.alias_nb_pages("{nb}")
        self.pdf.set_auto_page_break(True, margin=20)
        self.pdf.add_font("cjk", "", font_path)
        self.title = title
        self.w = self.pdf.w - self.pdf.l_margin - self.pdf.r_margin

    def render(self, blocks, source_name):
        pdf = self.pdf
        pdf.add_page()
        pdf.set_font("cjk", "", 20)
        pdf.multi_cell(0, 10, self.title, align="C")
        pdf.set_font("cjk", "", 10)
        pdf.set_text_color(100, 100, 100)
        pdf.cell(
            0, 8,
            f"来源：{source_name}  导出日期：{datetime.date.today().isoformat()}",
            align="C", new_x="LMARGIN", new_y="NEXT",
        )
        pdf.set_text_color(0, 0, 0)
        pdf.ln(4)
        for kind, data in blocks:
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
                x = pdf.get_x()
                pdf.set_x(x + 4)
                pdf.multi_cell(self.w - 4, 6.5, prefix + body)
            elif kind == "quote":
                pdf.set_font("cjk", "", 11)
                pdf.set_text_color(80, 80, 80)
                pdf.set_x(pdf.get_x() + 4)
                pdf.multi_cell(self.w - 4, 6.5, "▍ " + data)
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
        with pdf.table(
            width=self.w,
            col_widths=tuple([col_w] * ncol),
            first_row_as_headings=True,
            text_align="LEFT",
            line_height=6,
        ) as table:
            for idx, row in enumerate(rows):
                tr = table.row()
                for cell in row:
                    if idx == 0:
                        tr.cell(cell, fill_color=(230, 230, 230))
                    else:
                        tr.cell(cell)
        pdf.ln(2)

    def save(self, out):
        self.pdf.output(out)


def parse_args(argv=None):
    ap = argparse.ArgumentParser(description="审核意见 Markdown 转 PDF（中文）")
    ap.add_argument("input", help="输入 Markdown 文件（审核意见）")
    ap.add_argument("--out", default="", help="输出 PDF 路径（默认与输入同名 .pdf）")
    ap.add_argument("--title", default="环境影响报告表审核意见", help="PDF 标题")
    ap.add_argument("--font", default="", help="中文字体文件路径（缺省自动发现）")
    return ap.parse_args(argv)


def main(argv=None):
    args = parse_args(argv)
    if not os.path.exists(args.input):
        eprint(f"错误：文件不存在 - {args.input}")
        sys.exit(2)
    if args.input.lower().endswith(".json"):
        eprint("错误：暂不支持 JSON 输入，请先用 Markdown 版审核意见导出。")
        sys.exit(2)
    try:
        from fpdf import FPDF  # noqa: F401
    except ImportError:
        eprint("错误：缺少 fpdf2，请执行 pip install fpdf2")
        sys.exit(2)
    with open(args.input, "r", encoding="utf-8") as f:
        text = f.read()
    if not text.strip():
        eprint("错误：输入文件为空。")
        sys.exit(2)
    font_path = find_font(args.font)
    blocks = parse_blocks(text.splitlines())
    out = args.out or os.path.splitext(args.input)[0] + ".pdf"
    doc = AuditPDF(font_path, args.title)
    doc.render(blocks, os.path.basename(args.input))
    doc.save(out)
    print(f"已生成：{out}（字体：{os.path.basename(font_path)}）")


if __name__ == "__main__":
    main()
