"""
从 Word 文档(.docx) 中提取文本内容，用于环评报告表审核。

用法：python extract_docx.py <文件路径>

依赖：pip install python-docx
"""

import sys
import os

def extract_text(file_path):
    """提取 .docx 文件的全部文本内容，包括表格。"""
    try:
        from docx import Document
    except ImportError:
        print("错误：请先安装 python-docx 库：pip install python-docx", file=sys.stderr)
        sys.exit(1)

    if not os.path.exists(file_path):
        print(f"错误：文件不存在 - {file_path}", file=sys.stderr)
        sys.exit(1)

    doc = Document(file_path)
    content_parts = []

    # 提取正文段落
    content_parts.append("=== 正文内容 ===\n")
    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if text:
            # 标注段落样式（标题/正文）
            style_name = para.style.name if para.style else ""
            if "Heading" in style_name or "标题" in style_name:
                content_parts.append(f"\n【{text}】")
            else:
                content_parts.append(text)

    # 提取表格内容
    if doc.tables:
        content_parts.append("\n\n=== 表格内容 ===\n")
        for t_idx, table in enumerate(doc.tables):
            content_parts.append(f"\n--- 表格 {t_idx + 1} ---")
            for row in table.rows:
                cells = [cell.text.strip().replace("\n", " ") for cell in row.cells]
                content_parts.append(" | ".join(cells))

    return "\n".join(content_parts)


def main():
    if len(sys.argv) < 2:
        print("用法：python extract_docx.py <文件路径>", file=sys.stderr)
        print("示例：python extract_docx.py 建设项目环境影响报告表.docx", file=sys.stderr)
        sys.exit(1)

    file_path = sys.argv[1]
    text = extract_text(file_path)
    print(text)


if __name__ == "__main__":
    main()
