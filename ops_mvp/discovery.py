# -*- coding: utf-8 -*-
"""
根据仓库根目录实际存在的文件，判断流水线脚本是否可在面板中展示/运行。
"""

from __future__ import annotations

from pathlib import Path
from typing import Any, Dict, List

from ops_mvp.pipeline_catalog import PIPELINE_SCRIPTS, PipelineScript, get_script_by_id


def resolve_script_path(
    repo_root: Path,
    script: PipelineScript,
    markdown_paths: Dict[str, Path],
) -> Path:
    """Markdown 类使用 markdown_paths，其余使用 script.filename。"""
    if script.runner == "markdown_check":
        return markdown_paths["check"]
    if script.runner == "markdown_fix":
        return markdown_paths["fix"]
    return repo_root / script.filename


def is_pipeline_script_available(
    repo_root: Path,
    script: PipelineScript,
    markdown_paths: Dict[str, Path],
) -> bool:
    return resolve_script_path(repo_root, script, markdown_paths).is_file()


def filter_pipeline_catalog(
    repo_root: Path,
    markdown_paths: Dict[str, Path],
) -> List[PipelineScript]:
    return [
        s
        for s in PIPELINE_SCRIPTS
        if is_pipeline_script_available(repo_root, s, markdown_paths)
    ]


def playbook_step_runnable(
    repo_root: Path,
    script_id: str,
    markdown_paths: Dict[str, Path],
) -> bool:
    meta = get_script_by_id(script_id)
    if not meta:
        return False
    return is_pipeline_script_available(repo_root, meta, markdown_paths)


def data_flow_step_visible(
    repo_root: Path,
    markdown_paths: Dict[str, Path],
    step: Dict[str, Any],
) -> bool:
    """
    流程卡片：related_script_ids 为「任选其一即可展示」
    （例如落地阶段可有验收或 mock；维护阶段多种脚本）。
    """
    ids = step.get("related_script_ids") or []
    if not ids:
        return True
    return any(
        playbook_step_runnable(repo_root, sid, markdown_paths) for sid in ids
    )


def filter_core_chain_lines(
    lines: List[str],
    markdown_paths: Dict[str, Path],
) -> List[str]:
    """去掉引用不存在 4_/5_ 脚本的条目。"""
    out: List[str] = []
    for line in lines:
        need4 = (
            "4_" in line
            or "④" in line
            or "markdown_quality_reports" in line
        )
        need5 = "5_" in line or "⑤" in line or "extract_jsons_fixed/" in line
        if need4 and not markdown_paths["check"].is_file():
            continue
        if need5 and not markdown_paths["fix"].is_file():
            continue
        out.append(line)
    return out


def filter_data_flow_steps(
    repo_root: Path,
    markdown_paths: Dict[str, Path],
    steps: List[Dict[str, Any]],
) -> List[Dict[str, Any]]:
    return [s for s in steps if data_flow_step_visible(repo_root, markdown_paths, s)]


def build_filename_order_hint(repo_root: Path, markdown_paths: Dict[str, Path]) -> str:
    parts: List[str] = []
    mapping = [
        ("`1_` 资料包目录验收", "1_check_tobacco_kb_required_files.py"),
        ("`2_` 确认单→清单", "2_export_confirmation_bundle.py"),
        ("`3_` 正文抽取", "3_prepare_local_kb.py"),
    ]
    for label, fname in mapping:
        if (repo_root / fname).is_file():
            parts.append(label)
    if not parts:
        return "未检测到根目录编号脚本；请将脚本置于仓库根后刷新。"
    return "本仓库已具备：" + " · ".join(parts) + "。"
