# -*- coding: utf-8 -*-
"""
仓库内关键路径扫描（仅供运维台展示，不修改任何业务文件）。
"""

from __future__ import annotations

import os
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple


# 运营关心的固定目录（相对仓库根）
KEY_DIRECTORIES: Tuple[Tuple[str, str], ...] = (
    (
        "extract_jsons",
        "阶段②③写入目录：每个项目一个子文件夹；内含三份清单 JSON、docs_metadata、"
        "kb_local/manifest 与 extracted_markdown（正文 Markdown）。点此下方可看列表。",
    ),
    ("markdown_quality_reports", "Markdown 质量检测报告"),
    ("extract_jsons_fixed", "Markdown 基础修复输出"),
    ("第一批", "原始资料包（验收扫描入口）"),
    ("mockdata", "Mock 脚本生成的模拟项目目录（默认输出根）"),
    ("检查结果", "自动验收 Excel/HTML/长图"),
    ("tobacco_kb/mock_templates", "Mock 模板（若存在，位于 tobacco_kb 包内）"),
)


def _safe_stat(path: Path) -> Optional[os.stat_result]:
    try:
        return path.stat()
    except OSError:
        return None


def summarize_directory(path: Path, *, list_preview: int = 12) -> Dict[str, Any]:
    """目录是否存在、规模概览、直接子项预览。"""
    out: Dict[str, Any] = {
        "exists": path.exists(),
        "is_dir": path.is_dir(),
        "is_file": path.is_file(),
    }
    if not path.exists():
        return out

    st = _safe_stat(path)
    if st:
        out["size_bytes"] = st.st_size if path.is_file() else None
        out["mtime"] = st.st_mtime

    if path.is_dir():
        try:
            children = sorted(path.iterdir(), key=lambda p: (not p.is_dir(), p.name.lower()))
        except OSError as e:
            out["read_error"] = str(e)
            return out
        out["child_count"] = len(children)
        preview = []
        for ch in children[:list_preview]:
            try:
                rel_name = ch.name
                is_dir = ch.is_dir()
                ent: Dict[str, Any] = {"name": rel_name, "type": "dir" if is_dir else "file"}
                if ch.is_file():
                    stc = _safe_stat(ch)
                    if stc:
                        ent["size_bytes"] = stc.st_size
                preview.append(ent)
            except OSError:
                continue
        out["children_preview"] = preview
        out["children_truncated"] = len(children) > list_preview

    return out


def count_glob(path: Path, pattern: str, limit: int = 5000) -> int:
    if not path.is_dir():
        return 0
    n = 0
    try:
        for _ in path.rglob(pattern):
            if _.is_file():
                n += 1
                if n >= limit:
                    break
    except OSError:
        return n
    return n


def extract_jsons_projects(root: Path) -> List[Dict[str, Any]]:
    """extract_jsons 下一层项目文件夹摘要。"""
    ej = root / "extract_jsons"
    if not ej.is_dir():
        return []
    rows: List[Dict[str, Any]] = []
    try:
        subs = sorted([p for p in ej.iterdir() if p.is_dir()], key=lambda x: x.name.lower())
    except OSError:
        return []

    for proj in subs:
        kb = proj / "kb_local"
        md_dir = kb / "extracted_markdown"
        manifest = kb / "manifest.json"
        row: Dict[str, Any] = {
            "name": proj.name,
            "relative": str(proj.relative_to(root)).replace("\\", "/"),
            "has_kb_local": kb.is_dir(),
            "has_manifest": manifest.is_file(),
            "extracted_md_count": 0,
            "json_in_root": 0,
        }
        if md_dir.is_dir():
            try:
                row["extracted_md_count"] = sum(1 for _ in md_dir.glob("*.md"))
            except OSError:
                pass
        try:
            row["json_in_root"] = sum(1 for f in proj.glob("*.json") if f.is_file())
        except OSError:
            pass
        rows.append(row)
    return rows


def recent_files_in_dir(
    root: Path,
    relative: str,
    *,
    patterns: Tuple[str, ...] = ("*.html", "*.xlsx", "*.xls", "*.png"),
    limit: int = 12,
) -> List[Dict[str, Any]]:
    """某目录下按修改时间倒序的近期文件（浅层 + 一层子目录内）。"""
    base = (root / relative).resolve()
    if not base.is_dir():
        return []

    found: List[Tuple[float, Path]] = []
    try:
        for p in patterns:
            for f in base.glob(p):
                if f.is_file():
                    st = _safe_stat(f)
                    if st:
                        found.append((st.st_mtime, f))
        for sub in base.iterdir():
            if sub.is_dir():
                for p in patterns:
                    for f in sub.glob(p):
                        if f.is_file():
                            st = _safe_stat(f)
                            if st:
                                found.append((st.st_mtime, f))
    except OSError:
        return []

    found.sort(key=lambda x: x[0], reverse=True)
    out: List[Dict[str, Any]] = []
    for mtime, fpath in found[:limit]:
        try:
            rel = fpath.relative_to(root)
        except ValueError:
            continue
        st = _safe_stat(fpath)
        out.append({
            "relative": str(rel).replace("\\", "/"),
            "name": fpath.name,
            "mtime": mtime,
            "size_bytes": st.st_size if st else None,
        })
    return out


def artifact_counts(root: Path) -> Dict[str, Any]:
    """跨目录产物计数（有上限，避免超大仓库卡死）。"""
    co = root / "markdown_quality_reports"
    fo = root / "extract_jsons_fixed"
    ej = root / "extract_jsons"

    def safe_count(base: Path, pat: str) -> int:
        return count_glob(base, pat, limit=10000) if base.is_dir() else 0

    return {
        "quality_report_json": safe_count(co, "*.quality_report.json"),
        "review_items_md": safe_count(co, "*.review_items.md"),
        "checker_summary_json": 1 if (co / "summary.json").is_file() else 0,
        "fix_summary_json": 1 if (fo / "fix_summary.json").is_file() else 0,
        "extract_jsons_total_md": safe_count(ej, "*.md"),
        "manifest_json_count": safe_count(ej, "manifest.json"),
    }


def list_root_scripts(root: Path) -> List[Dict[str, Any]]:
    """仓库根目录下可识别的脚本入口。"""
    rows: List[Dict[str, Any]] = []
    try:
        for p in sorted(root.glob("*.py"), key=lambda x: x.name.lower()):
            st = _safe_stat(p)
            rows.append({
                "relative": p.name,
                "name": p.name,
                "size_bytes": st.st_size if st else None,
                "mtime": st.st_mtime if st else None,
            })
    except OSError:
        pass
    return rows


def build_workspace_snapshot(root: Path) -> Dict[str, Any]:
    root = root.resolve()
    keys = []
    for rel, tip in KEY_DIRECTORIES:
        p = root / rel
        info = summarize_directory(p, list_preview=8)
        info["relative"] = rel
        info["tip"] = tip
        keys.append(info)

    return {
        "key_directories": keys,
        "artifacts": artifact_counts(root),
        "extract_projects": extract_jsons_projects(root),
        "recent_results": recent_files_in_dir(root, "检查结果", limit=14),
        "root_scripts": list_root_scripts(root),
    }
