"""
将 Markdown 正文渲染为版式化 PDF（ReportLab Platypus），与 Word 共用 `markdown_structure` 解析结果。
中文使用系统中文字体（与 mock 其余 PDF 逻辑一致）。
"""

from __future__ import annotations

import html
import io
import re
from typing import Any, Dict, List

from .markdown_structure import parse_markdown_to_blocks

# A4 纵向可用宽度（与 SimpleDocTemplate 左右边距 56 一致）
_FRAME_W_PT = 595.27 - 56 * 2


def _md_inline_to_xml(text: str) -> str:
    """将 **加粗** 转为 ReportLab Paragraph 可用的伪 HTML。"""
    if not text:
        return ""
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    parts = re.split(r"(\*\*.+?\*\*)", text)
    out: List[str] = []
    for part in parts:
        if not part:
            continue
        if part.startswith("**") and part.endswith("**") and len(part) >= 4:
            inner = html.escape(part[2:-2])
            out.append(f"<b>{inner}</b>")
        else:
            out.append(html.escape(part))
    s = "".join(out)
    return s.replace("\n", "<br/>")


def _can_use_markdown_pdf() -> bool:
    try:
        import reportlab  # noqa: F401
    except ImportError:
        return False
    # 延迟导入，避免与 mock_placeholders 的懒加载形成环
    from .mock_placeholders import find_cjk_font_path

    return find_cjk_font_path() is not None


def render_markdown_to_pdf_bytes(document_title: str, md_text: str) -> bytes:
    """封面标题 + 与 docx 一致的 MD 结构。"""
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.platypus import (
        Paragraph,
        SimpleDocTemplate,
        Spacer,
        Table,
        TableStyle,
    )

    from .mock_placeholders import find_cjk_font_path

    font_path = find_cjk_font_path()
    if not font_path:
        raise OSError("未找到系统中文字体，无法生成 PDF")

    zh_name = "MdPdfZH"
    if font_path.lower().endswith(".ttc"):
        pdfmetrics.registerFont(TTFont(zh_name, font_path, subfontIndex=0))
    else:
        pdfmetrics.registerFont(TTFont(zh_name, font_path))

    base = ParagraphStyle(
        "md_base",
        fontName=zh_name,
        fontSize=11,
        leading=15,
        spaceAfter=6,
        textColor=colors.HexColor("#212121"),
    )
    title_style = ParagraphStyle(
        "md_doctitle",
        parent=base,
        fontSize=18,
        leading=24,
        spaceAfter=14,
        textColor=colors.HexColor("#1565c0"),
    )
    quote_style = ParagraphStyle(
        "md_quote",
        parent=base,
        leftIndent=18,
        spaceBefore=4,
        spaceAfter=6,
        borderColor=colors.HexColor("#90caf9"),
        borderWidth=0.5,
        borderPadding=8,
        backColor=colors.HexColor("#f5f9ff"),
    )

    def heading_style(level: int) -> ParagraphStyle:
        sizes = {1: 16, 2: 14, 3: 13, 4: 12, 5: 11, 6: 11}
        return ParagraphStyle(
            f"md_h{level}",
            parent=base,
            fontSize=sizes.get(min(level, 6), 11),
            leading=sizes.get(min(level, 6), 11) + 4,
            spaceBefore=10 if level <= 2 else 6,
            spaceAfter=6,
            textColor=colors.HexColor("#0d47a1") if level <= 2 else colors.HexColor("#37474f"),
        )

    bullet_style = ParagraphStyle(
        "md_bullet",
        parent=base,
        leftIndent=22,
        bulletIndent=12,
        fontName=zh_name,
    )
    number_style = ParagraphStyle(
        "md_num",
        parent=base,
        leftIndent=22,
        bulletIndent=0,
        fontName=zh_name,
    )

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=A4,
        rightMargin=56,
        leftMargin=56,
        topMargin=56,
        bottomMargin=56,
        title=document_title[:120],
    )

    story: List[Any] = []
    story.append(Paragraph(_md_inline_to_xml(document_title), title_style))
    story.append(Spacer(1, 8))

    blocks = parse_markdown_to_blocks(md_text)
    num_counter = 1

    for b in blocks:
        typ = b["type"]
        if typ == "hr":
            story.append(Spacer(1, 6))
            story.append(
                Table(
                    [[Paragraph("―" * 42, base)]],
                    colWidths=[_FRAME_W_PT],
                    style=TableStyle(
                        [
                            ("LINEBELOW", (0, 0), (-1, -1), 0.5, colors.HexColor("#bdbdbd")),
                            ("TOPPADDING", (0, 0), (-1, -1), 4),
                            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                        ]
                    ),
                )
            )
            story.append(Spacer(1, 6))
            num_counter = 1
        elif typ == "heading":
            lv = int(b["level"])
            story.append(Paragraph(_md_inline_to_xml(str(b["text"])), heading_style(lv)))
            num_counter = 1
        elif typ == "paragraph":
            story.append(Paragraph(_md_inline_to_xml(str(b["text"])), base))
            num_counter = 1
        elif typ == "bullet":
            story.append(
                Paragraph(
                    _md_inline_to_xml(str(b["text"])),
                    bullet_style,
                    bulletText="•",
                )
            )
            num_counter = 1
        elif typ == "numbered":
            txt = f"{num_counter}. {str(b['text'])}"
            story.append(Paragraph(_md_inline_to_xml(txt), number_style))
            num_counter += 1
        elif typ == "table":
            rows: List[List[str]] = b["rows"]
            if not rows:
                continue
            ncols = max(len(r) for r in rows)
            data: List[List[Any]] = []
            for ri, row in enumerate(rows):
                cells = []
                for ci in range(ncols):
                    cell_txt = row[ci] if ci < len(row) else ""
                    cells.append(Paragraph(_md_inline_to_xml(cell_txt), base))
                data.append(cells)
            tw = _FRAME_W_PT
            col_w = tw / ncols
            tbl = Table(data, colWidths=[col_w] * ncols, repeatRows=1)
            tbl.setStyle(
                TableStyle(
                    [
                        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#bdbdbd")),
                        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#e3f2fd")),
                        ("VALIGN", (0, 0), (-1, -1), "TOP"),
                        ("LEFTPADDING", (0, 0), (-1, -1), 6),
                        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                        ("TOPPADDING", (0, 0), (-1, -1), 4),
                        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                    ]
                )
            )
            story.append(Spacer(1, 6))
            story.append(tbl)
            story.append(Spacer(1, 8))
            num_counter = 1
        elif typ == "quote":
            story.append(Paragraph(_md_inline_to_xml(str(b["text"])), quote_style))
            num_counter = 1

    doc.build(story)
    return buf.getvalue()


def has_markdown_pdf() -> bool:
    return _can_use_markdown_pdf()
