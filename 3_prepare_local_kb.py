# -*- coding: utf-8 -*-
"""
本地知识库沉淀准备（Document-Centric 版）
================================================

核心思想：
  一个原始文件 = 一个 document
  输出整篇 Markdown：DOCX/PDF 含段落（标题样式）与表格（Markdown 表格语法）；manifest 不落 tables/rows。

目录结构：
  kb_local/
    manifest.json
    extracted_markdown/   # 每个 document 一份 .md（整篇正文）

用法：
  python 3_prepare_local_kb.py
  python 3_prepare_local_kb.py --extract-root extract_jsons

依赖：
  pip install python-docx pypdf pdfplumber

可选：
  .doc 文件建议安装 LibreOffice，并确保 soffice 在 PATH；
  Windows 可安装 Word + pywin32。
"""

from __future__ import annotations

import argparse
import hashlib
import json
import shutil
import subprocess
import sys
import tempfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

_REPO_ROOT = Path(__file__).resolve().parent
if str(_REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(_REPO_ROOT))

try:
    from tobacco_kb.acceptance_config import PHASE1_RULES
except Exception:
    # 兜底：如果单独拿脚本跑，没有 tobacco_kb 包，也不影响基础抽取
    PHASE1_RULES = []

KB_LOCAL_DIRNAME = "kb_local"
DOCS_METADATA_DIRNAME = "docs_metadata"
PROJECT_BASIC_JSON = "项目基础信息清单.json"

MARKDOWN_SUBDIR = "extracted_markdown"

TEXT_EXTRACT_EXTENSIONS = frozenset({".txt", ".md", ".docx", ".pdf", ".doc"})

KB_RISK_NAME_KEYWORDS: Tuple[str, ...] = (
    "接口文档",
    "数据字典",
    "部署手册",
    "安装配置",
    "报价",
    "合同",
    "账号",
    "密码",
    "服务器",
    "数据库",
    "样例数据",
    "巡检记录",
    "生产环境",
)

_STD_LEAVES_DESC = sorted(
    ({str(r.get("mock_std_dir", "")) for r in PHASE1_RULES if r.get("mock_std_dir")}),
    key=len,
    reverse=True,
)
_LEAF_TO_RULE: Dict[str, str] = {
    str(r.get("mock_std_dir")): str(r.get("id", ""))
    for r in PHASE1_RULES
    if r.get("mock_std_dir")
}
_LEAF_TO_LABEL: Dict[str, str] = {
    str(r.get("mock_std_dir")): str(r.get("label", ""))
    for r in PHASE1_RULES
    if r.get("mock_std_dir")
}


def _configure_stdio_utf8() -> None:
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8")
            sys.stderr.reconfigure(encoding="utf-8")
        except Exception:
            pass


def _now_iso() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")


def _sha256_file(path: Path, chunk: int = 1024 * 1024) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        while True:
            b = f.read(chunk)
            if not b:
                break
            h.update(b)
    return h.hexdigest()


def _doc_id_for_file(path: Path, sha256: Optional[str] = None) -> str:
    # 用文件 hash + 文件名生成稳定 doc_id，避免同名不同文件冲突
    base = (sha256 or _sha256_file(path)) + "|" + path.name
    return hashlib.sha256(base.encode("utf-8")).hexdigest()[:16]


def parse_bundle_filename(filename: str) -> Tuple[str, str, str, str]:
    """
    解析 docs_metadata 扁平文件名：<标准子目录>__<叶内路径用__代替斜杠>
    返回：standard_subdir, leaf_inner_posix, rule_id, rule_label
    """
    for leaf in _STD_LEAVES_DESC:
        prefix = leaf + "__"
        if filename.startswith(prefix):
            inner = filename[len(prefix) :]
            inner_posix = inner.replace("__", "/")
            rid = _LEAF_TO_RULE.get(leaf, "")
            lab = _LEAF_TO_LABEL.get(leaf, "")
            return leaf, inner_posix, rid, lab
    return "", filename, "", ""


def _risk_tags_from_name(name: str) -> List[str]:
    n = name.replace(" ", "")
    return [kw for kw in KB_RISK_NAME_KEYWORDS if kw in n]


def _clean_cell_text(s: Any) -> str:
    if s is None:
        return ""
    return " ".join(str(s).replace("\u3000", " ").split()).strip()


def _safe_markdown_cell(s: str) -> str:
    return _clean_cell_text(s).replace("|", "\\|")


def _table_to_markdown(rows: List[List[str]]) -> List[str]:
    if not rows:
        return []

    max_cols = max(len(r) for r in rows)
    normalized = [r + [""] * (max_cols - len(r)) for r in rows]

    header = normalized[0]
    if not any(c.strip() for c in header):
        header = [f"列{i + 1}" for i in range(max_cols)]
        body = normalized
    else:
        body = normalized[1:]

    lines = [
        "| " + " | ".join(_safe_markdown_cell(c) for c in header) + " |",
        "| " + " | ".join(["---"] * max_cols) + " |",
    ]
    for row in body:
        lines.append("| " + " | ".join(_safe_markdown_cell(c) for c in row) + " |")
    return lines


def _ensure_clean_dir(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)
    for child in path.iterdir():
        if child.is_file():
            child.unlink()
        elif child.is_dir():
            shutil.rmtree(child)


def _style_to_markdown_prefix(style_name: str) -> str:
    s = (style_name or "").strip()
    # 英文 Word 样式 + 中文 Word 样式
    mapping = {
        "Heading 1": "# ",
        "标题 1": "# ",
        "Heading 2": "## ",
        "标题 2": "## ",
        "Heading 3": "### ",
        "标题 3": "### ",
        "Heading 4": "#### ",
        "标题 4": "#### ",
        "Heading 5": "##### ",
        "标题 5": "##### ",
        "Heading 6": "###### ",
        "标题 6": "###### ",
    }
    if s in mapping:
        return mapping[s]
    for i in range(1, 7):
        if s.startswith(f"Heading {i}") or s.startswith(f"标题 {i}"):
            return "#" * i + " "
    return ""


def _iter_docx_blocks(parent: Any) -> Iterable[Any]:
    """按 Word 文档真实顺序遍历段落和表格。"""
    from docx.document import Document as DocxDocument
    from docx.oxml.table import CT_Tbl
    from docx.oxml.text.paragraph import CT_P
    from docx.table import Table, _Cell
    from docx.text.paragraph import Paragraph

    if isinstance(parent, DocxDocument):
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
    else:
        return

    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)


def _extract_cell_text_recursive(cell: Any) -> str:
    """单元格内可能包含多个段落/嵌套表，递归压成可读文本。"""
    from docx.table import Table
    from docx.text.paragraph import Paragraph

    parts: List[str] = []
    for block in _iter_docx_blocks(cell):
        if isinstance(block, Paragraph):
            t = _clean_cell_text(block.text)
            if t:
                parts.append(t)
        elif isinstance(block, Table):
            for row in block.rows:
                cells = [_extract_cell_text_recursive(c) for c in row.cells]
                line = " / ".join(c for c in cells if c)
                if line:
                    parts.append(line)
    return " ".join(parts).strip()


def extract_docx_as_document_views(
    path: Path,
    doc_id: str,
    markdown_dir: Path,
) -> Dict[str, Any]:
    """
    DOCX：按文档顺序输出整篇 Markdown（段落/标题 + 表格为 Markdown 表格）。
    """
    from docx import Document
    from docx.table import Table
    from docx.text.paragraph import Paragraph

    doc = Document(str(path))

    md_lines: List[str] = []
    text_lines: List[str] = []

    for block in _iter_docx_blocks(doc):
        if isinstance(block, Paragraph):
            txt = _clean_cell_text(block.text)
            if not txt:
                continue
            prefix = _style_to_markdown_prefix(block.style.name if block.style else "")
            md_lines.append(prefix + txt if prefix else txt)
            text_lines.append(txt)

        elif isinstance(block, Table):
            table_rows: List[List[str]] = []
            for row in block.rows:
                cells = [_extract_cell_text_recursive(cell) for cell in row.cells]
                table_rows.append(cells)

            md_lines.append("")
            md_lines.extend(_table_to_markdown(table_rows))
            md_lines.append("")

            for r in table_rows:
                text_lines.append("\t".join(r))

    markdown = "\n\n".join(md_lines).strip() + "\n"
    text = "\n".join(text_lines).strip() + "\n"

    md_name = f"{doc_id}.md"
    (markdown_dir / md_name).write_text(markdown, encoding="utf-8")

    return {
        "views": {
            "markdown": f"{MARKDOWN_SUBDIR}/{md_name}",
        },
        "text_chars": len(text.strip()),
        "warnings": [],
    }


def extract_text_as_document_views(
    path: Path,
    doc_id: str,
    markdown_dir: Path,
) -> Dict[str, Any]:
    raw = path.read_text(encoding="utf-8", errors="replace")
    md_name = f"{doc_id}.md"

    if path.suffix.lower() == ".md":
        markdown = raw
    else:
        markdown = "\n".join(line.rstrip() for line in raw.splitlines())

    (markdown_dir / md_name).write_text(markdown, encoding="utf-8")

    return {
        "views": {
            "markdown": f"{MARKDOWN_SUBDIR}/{md_name}",
        },
        "text_chars": len(raw.strip()),
        "warnings": [],
    }


def extract_pdf_as_document_views(
    path: Path,
    doc_id: str,
    markdown_dir: Path,
) -> Dict[str, Any]:
    """
    PDF：按页写入正文；若有 pdfplumber，额外抽取表格并写成 Markdown 表格（整页内容尽量完整）。
    注意：扫描版 PDF 需要 OCR，本脚本不做 OCR，只给 warning。
    """
    warnings: List[str] = []
    md_lines: List[str] = []
    text_lines: List[str] = []

    try:
        import pdfplumber
    except ImportError:
        # 兜底用 pypdf 只提文本
        try:
            from pypdf import PdfReader

            reader = PdfReader(str(path))
            for page_idx, page in enumerate(reader.pages, start=1):
                t = page.extract_text() or ""
                t = t.strip()
                if t:
                    md_lines.append(f"## 第 {page_idx} 页")
                    md_lines.append(t)
                    text_lines.append(f"[Page {page_idx}]")
                    text_lines.append(t)
            warnings.append("未安装 pdfplumber，PDF 表格未结构化写入 md，仅使用 pypdf 抽文本")
        except ImportError:
            raise RuntimeError("缺少 PDF 依赖：请安装 pip install pypdf pdfplumber")
    else:
        with pdfplumber.open(str(path)) as pdf:
            for page_idx, page in enumerate(pdf.pages, start=1):
                page_text = (page.extract_text() or "").strip()
                if page_text:
                    md_lines.append(f"## 第 {page_idx} 页")
                    md_lines.append(page_text)
                    text_lines.append(f"[Page {page_idx}]")
                    text_lines.append(page_text)

                try:
                    tables = page.extract_tables() or []
                except Exception as e:
                    warnings.append(f"第 {page_idx} 页表格提取失败：{e}")
                    tables = []

                for table in tables:
                    clean_rows = [[_clean_cell_text(c) for c in row] for row in table if row]
                    if not clean_rows:
                        continue
                    md_lines.append("")
                    md_lines.extend(_table_to_markdown(clean_rows))
                    md_lines.append("")
                    for r in clean_rows:
                        text_lines.append("\t".join(r))

    markdown = "\n\n".join(md_lines).strip() + "\n"
    text = "\n".join(text_lines).strip() + "\n"

    if not text.strip():
        warnings.append("PDF 抽取正文为空：可能是扫描件，需要 OCR")

    md_name = f"{doc_id}.md"
    (markdown_dir / md_name).write_text(markdown, encoding="utf-8")

    return {
        "views": {
            "markdown": f"{MARKDOWN_SUBDIR}/{md_name}",
        },
        "text_chars": len(text.strip()),
        "warnings": warnings,
    }


def _find_soffice() -> Optional[str]:
    for name in ("soffice", "libreoffice"):
        p = shutil.which(name)
        if p:
            return p
    if sys.platform == "win32":
        for candidate in (
            Path(r"C:\Program Files\LibreOffice\program\soffice.exe"),
            Path(r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"),
        ):
            if candidate.is_file():
                return str(candidate)
    return None


def _convert_doc_to_docx_with_soffice(path: Path) -> Optional[Path]:
    soffice = _find_soffice()
    if not soffice:
        return None
    td = Path(tempfile.mkdtemp(prefix="doc_to_docx_"))
    try:
        subprocess.run(
            [
                soffice,
                "--headless",
                "--convert-to",
                "docx",
                str(path.resolve()),
                "--outdir",
                str(td),
            ],
            check=True,
            timeout=180,
            capture_output=True,
        )
        cand = td / (path.stem + ".docx")
        if cand.is_file():
            return cand
        found = list(td.glob("*.docx"))
        return found[0] if found else None
    except Exception:
        shutil.rmtree(td, ignore_errors=True)
        return None


def _word_save_doc_to_docx(path: Path) -> Optional[Path]:
    """
    用本机 Word 将 .doc 另存为 .docx（OOXML UTF-8），避免 COM 直接读 Content.Text 时的中文乱码，
    且保留段落/表格结构供 python-docx 解析。
    """
    if sys.platform != "win32":
        return None
    try:
        import pythoncom  # type: ignore
        import win32com.client  # type: ignore
    except ImportError:
        return None

    td = Path(tempfile.mkdtemp(prefix="word_doc_to_docx_"))
    out_docx = td / f"{path.stem}_kb_convert.docx"
    word = None
    doc = None
    pythoncom.CoInitialize()
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        doc = word.Documents.Open(str(path.resolve()), ReadOnly=True, AddToRecentFiles=False)
        # 12 = wdFormatXMLDocument，Office Open XML (.docx)
        doc.SaveAs2(str(out_docx), FileFormat=12)
        doc.Close(SaveChanges=False)
        doc = None
        word.Quit()
        word = None
    except Exception:
        shutil.rmtree(td, ignore_errors=True)
        try:
            if doc is not None:
                doc.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if word is not None:
                word.Quit()
        except Exception:
            pass
        return None
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

    if not out_docx.is_file():
        shutil.rmtree(td, ignore_errors=True)
        return None
    return out_docx


def extract_doc_as_document_views(
    path: Path,
    doc_id: str,
    markdown_dir: Path,
) -> Dict[str, Any]:
    """
    老式 .doc：先转 docx 再按 docx 抽整篇 Markdown。
    Windows 已安装 Word 时优先用 Word 另存为 docx（编码/表格与 Office 一致，避免 LO 或 COM 读正文乱码）。
    否则尝试 LibreOffice；均失败则报错。
    """
    if sys.platform == "win32":
        word_docx = _word_save_doc_to_docx(path)
        if word_docx is not None:
            td = word_docx.parent
            try:
                result = extract_docx_as_document_views(word_docx, doc_id, markdown_dir)
                result.setdefault("warnings", []).append(".doc 已通过 Word 另存为 .docx 后抽取（推荐，修正中文编码）")
                return result
            finally:
                shutil.rmtree(td, ignore_errors=True)

    converted = _convert_doc_to_docx_with_soffice(path)
    if converted and converted.is_file():
        try:
            result = extract_docx_as_document_views(converted, doc_id, markdown_dir)
            result.setdefault("warnings", []).append(".doc 已通过 LibreOffice 转 .docx 后抽取")
            return result
        finally:
            shutil.rmtree(converted.parent, ignore_errors=True)

    raise RuntimeError(
        "无法抽取 .doc：Windows 请安装 Microsoft Word + pip install pywin32，"
        "或安装 LibreOffice 并确保 soffice 在 PATH，或先手动转为 .docx"
    )


def load_project_profile(project_dir: Path) -> Optional[Dict[str, Any]]:
    p = project_dir / PROJECT_BASIC_JSON
    if not p.is_file():
        return None
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
        return data.get("fields") if isinstance(data.get("fields"), dict) else None
    except Exception:
        return None


def build_document_record(fpath: Path, idx: int, kb_dir: Path) -> Dict[str, Any]:
    """
    一个原始文件 = 一个 document：manifest 记录来源与 views.markdown 路径，正文仅在 extracted_markdown。
    """
    name = fpath.name
    ext = fpath.suffix.lower()
    sha = _sha256_file(fpath)
    doc_id = _doc_id_for_file(fpath, sha)
    subdir, inner, rid, rlab = parse_bundle_filename(name)

    markdown_dir = kb_dir / MARKDOWN_SUBDIR

    document: Dict[str, Any] = {
        "doc_id": doc_id,
        "document_index": idx,
        "source_file": name,
        "source_path": str(fpath.resolve()),
        "standard_subdir": subdir,
        "leaf_inner_path_posix": inner,
        "rule_id": rid,
        "rule_label_zh": rlab,
        "extension": ext,
        "size_bytes": fpath.stat().st_size,
        "sha256": sha,
        "risk_keywords_hit": _risk_tags_from_name(name),
        "views": {
            "markdown": None,
        },
        "quality": {
            "status": "pending",
            "has_text": False,
            "text_chars": 0,
            "warnings": [],
        },
    }

    if ext not in TEXT_EXTRACT_EXTENSIONS:
        document["quality"].update(
            {
                "status": "skipped_format",
                "warnings": [f"extension not in extract list: {ext or '(no ext)'}"],
            }
        )
        return document

    try:
        if ext == ".docx":
            result = extract_docx_as_document_views(fpath, doc_id, markdown_dir)
        elif ext == ".pdf":
            result = extract_pdf_as_document_views(fpath, doc_id, markdown_dir)
        elif ext == ".doc":
            result = extract_doc_as_document_views(fpath, doc_id, markdown_dir)
        elif ext in {".txt", ".md"}:
            result = extract_text_as_document_views(fpath, doc_id, markdown_dir)
        else:
            result = {
                "views": {"markdown": None},
                "text_chars": 0,
                "warnings": [f"unhandled extension: {ext}"],
            }

        document["views"] = result.get("views", document["views"])

        text_chars = int(result.get("text_chars", 0) or 0)
        warnings = result.get("warnings", []) or []

        status = "ok" if text_chars > 0 else "empty"
        document["quality"].update(
            {
                "status": status,
                "has_text": text_chars > 0,
                "text_chars": text_chars,
                "warnings": warnings,
            }
        )

    except Exception as e:
        document["quality"].update(
            {
                "status": "failed",
                "warnings": [str(e)],
            }
        )

    return document


def discover_project_dirs(extract_root: Path) -> List[Path]:
    if not extract_root.is_dir():
        return []
    return [
        child
        for child in sorted(extract_root.iterdir())
        if child.is_dir() and (child / DOCS_METADATA_DIRNAME).is_dir()
    ]


def process_project(project_dir: Path) -> Dict[str, Any]:
    dm = project_dir / DOCS_METADATA_DIRNAME
    kb = project_dir / KB_LOCAL_DIRNAME

    if not dm.is_dir():
        return {"project": project_dir.name, "skipped": True, "reason": "no docs_metadata"}

    kb.mkdir(parents=True, exist_ok=True)
    _ensure_clean_dir(kb / MARKDOWN_SUBDIR)

    profile = load_project_profile(project_dir)

    bundle_paths = sorted(
        [x for x in dm.iterdir() if x.is_file() and not x.name.startswith("~$")]
    )

    documents: List[Dict[str, Any]] = []
    for idx, fpath in enumerate(bundle_paths, start=1):
        documents.append(build_document_record(fpath=fpath, idx=idx, kb_dir=kb))

    manifest: Dict[str, Any] = {
        "schema_version": "2.4",
        "purpose_zh": "本地知识库沉淀准备：一个原始文件对应一个 document；extracted_markdown 为整篇 Markdown（含表格）；仅 manifest 无 tables/rows 字段",
        "project_folder": project_dir.name,
        "generated_at": _now_iso(),
        "source_dir": DOCS_METADATA_DIRNAME,
        "project_fields": profile or {},
        "output_dirs": {
            "markdown": MARKDOWN_SUBDIR,
        },
        "total_documents": len(documents),
        "summary": {
            "documents_ok": sum(1 for d in documents if d.get("quality", {}).get("status") == "ok"),
            "documents_empty": sum(1 for d in documents if d.get("quality", {}).get("status") == "empty"),
            "documents_failed": sum(1 for d in documents if d.get("quality", {}).get("status") == "failed"),
            "documents_skipped": sum(1 for d in documents if d.get("quality", {}).get("status") == "skipped_format"),
        },
        "documents": documents,
    }

    manifest_path = kb / "manifest.json"
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")

    return {
        "project": project_dir.name,
        "skipped": False,
        "manifest": str(manifest_path.resolve()),
        "documents": len(documents),
        **manifest["summary"],
    }


def resolve_root(s: str) -> Path:
    p = Path(s)
    return p.resolve() if p.is_absolute() else (_REPO_ROOT / p).resolve()


def main(argv: Optional[List[str]] = None) -> int:
    _configure_stdio_utf8()
    parser = argparse.ArgumentParser(description="docs_metadata → kb_local document-centric manifest/views")
    parser.add_argument(
        "--extract-root",
        default="extract_jsons",
        help="2_export_confirmation_bundle 输出根目录，默认 extract_jsons",
    )
    args = parser.parse_args(argv)

    root = resolve_root(args.extract_root)
    projects = discover_project_dirs(root)
    if not projects:
        print(f"未在 {root} 下找到含「{DOCS_METADATA_DIRNAME}」的项目子目录。", file=sys.stderr)
        return 1

    for pd in projects:
        r = process_project(pd)
        if r.get("skipped"):
            print(f"跳过 {r.get('project')}: {r.get('reason')}", file=sys.stderr)
            continue
        print(
            f"{r['project']}: documents={r['documents']}, ok={r['documents_ok']}, "
            f"empty={r['documents_empty']}, failed={r['documents_failed']}, "
            f"skipped={r['documents_skipped']}",
            file=sys.stderr,
        )
        print(f"  {r['manifest']}", file=sys.stderr)

    print(f"共处理 {len(projects)} 个项目。", file=sys.stderr)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
