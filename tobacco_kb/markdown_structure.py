"""
将 Markdown 正文解析为结构化块列表，供 Word / PDF 等后端共用。
支持的语法与 markdown_to_docx 一致：标题、列表、表格、引用、分隔线、段落、行内 **加粗**（正文内解析）。
"""

from __future__ import annotations

import re
from typing import Any, Dict, List


def _is_table_separator(line: str) -> bool:
    s = line.strip()
    if not s.startswith("|"):
        return False
    return bool(re.match(r"^\|[\s\-:|]+\|$", s))


def parse_markdown_to_blocks(md_text: str) -> List[Dict[str, Any]]:
    """
    返回块列表，每块至少含 \"type\" 键：
    hr, heading, paragraph, bullet, numbered, table, quote
    """
    lines = md_text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    blocks: List[Dict[str, Any]] = []
    i = 0
    n = len(lines)

    while i < n:
        raw = lines[i]
        stripped = raw.strip()

        if not stripped:
            i += 1
            continue

        if re.match(r"^(\-{3,}|\*{3,}|_{3,})$", stripped):
            blocks.append({"type": "hr"})
            i += 1
            continue

        hm = re.match(r"^(#{1,6})\s+(.+)$", stripped)
        if hm:
            depth = len(hm.group(1))
            blocks.append(
                {
                    "type": "heading",
                    "level": min(depth, 9),
                    "text": hm.group(2).strip(),
                }
            )
            i += 1
            continue

        if stripped.startswith("|"):
            table_rows: List[List[str]] = []
            while i < n:
                row_line = lines[i].strip()
                if not row_line.startswith("|"):
                    break
                if _is_table_separator(row_line):
                    i += 1
                    continue
                cells = [c.strip() for c in row_line.strip("|").split("|")]
                table_rows.append(cells)
                i += 1
            if table_rows:
                blocks.append({"type": "table", "rows": table_rows})
            continue

        if stripped.startswith(">"):
            quote_lines: List[str] = []
            while i < n:
                s = lines[i].strip()
                if not s.startswith(">"):
                    break
                quote_lines.append(s.lstrip(">").strip())
                i += 1
            blocks.append({"type": "quote", "text": "\n".join(quote_lines)})
            continue

        if re.match(r"^[\-\*]\s+", stripped):
            while i < n:
                s = lines[i].strip()
                m2 = re.match(r"^[\-\*]\s+(.+)$", s)
                if not m2:
                    break
                blocks.append({"type": "bullet", "text": m2.group(1).strip()})
                i += 1
            continue

        if re.match(r"^\d+\.\s+", stripped):
            while i < n:
                s = lines[i].strip()
                m2 = re.match(r"^\d+\.\s+(.+)$", s)
                if not m2:
                    break
                blocks.append({"type": "numbered", "text": m2.group(1).strip()})
                i += 1
            continue

        para_lines: List[str] = [stripped]
        i += 1
        while i < n and lines[i].strip():
            nxt = lines[i].strip()
            if (
                nxt.startswith("#")
                or nxt.startswith("|")
                or nxt.startswith(">")
                or re.match(r"^[\-\*]\s+", nxt)
                or re.match(r"^\d+\.\s+", nxt)
                or re.match(r"^(\-{3,}|\*{3,}|_{3,})$", nxt)
            ):
                break
            para_lines.append(nxt)
            i += 1
        body = " ".join(para_lines) if len(para_lines) > 1 else para_lines[0]
        blocks.append({"type": "paragraph", "text": body})

    return blocks
