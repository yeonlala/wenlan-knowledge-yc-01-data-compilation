# -*- coding: utf-8 -*-
"""
流水线脚本清单（仅用于运维面板展示与调度，与仓库根脚本文件解耦）。

不向此处写入业务逻辑；脚本行为以后端 subprocess 调用为准。

字段说明：
- phase：在整体链路中所处阶段（便于运营排序理解）。
- when_use：什么时候该点 / 该跑这条脚本（运维话术）。
- inputs_desc / outputs_desc：典型输入、产出路径（相对仓库根描述）。
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import List, Literal, Optional

RunnerKind = Literal[
    "markdown_check",
    "markdown_fix",
    "subprocess_argv",
    "cli_hint_only",
]


@dataclass(frozen=True)
class PipelineScript:
    """描述仓库根目录下的一个可执行脚本。"""

    id: str
    filename: str
    title: str
    group: str
    description: str
    runner: RunnerKind
    phase: str
    when_use: str
    inputs_desc: str
    outputs_desc: str
    default_argv: tuple[str, ...] = ()
    cli_example: str = ""
    warn: Optional[str] = None


PIPELINE_SCRIPTS: tuple[PipelineScript, ...] = (
    PipelineScript(
        id="check_tobacco_kb",
        filename="1_check_tobacco_kb_required_files.py",
        title="资料包目录结构自动验收",
        group="资料包验收",
        description="按规则扫描「第一批」下各项目文件夹（标准二级目录、命名等），生成检查结果 Excel/HTML，并可导出验收长图。",
        runner="subprocess_argv",
        phase="资料合规检查",
        when_use="需要批量检查资料包是否齐套、命名是否符合约定时；可在整理前摸底，也可在入库前终检。运行时间随资料量变化。",
        inputs_desc="第一个参数为扫描根目录；面板默认传入相对路径 `第一批`（即仓库根下「第一批」文件夹）。",
        outputs_desc="仓库同级或指定工作区下的「检查结果」目录：时间戳命名的 `.xlsx`、`.html` 及长图 PNG 等。",
        default_argv=("第一批",),
        warn="运行时间较长；输出体积随项目数增加。",
    ),
    PipelineScript(
        id="export_confirmation_bundle",
        filename="2_export_confirmation_bundle.py",
        title="确认单 → 三份清单 + docs_metadata",
        group="清单与归集",
        description="解析《项目资料上报确认单》JSON，生成项目基础信息/材料提交/责任人确认清单；并把 02_核心资料 下文件扁平复制到 docs_metadata。",
        runner="subprocess_argv",
        phase="入库准备（结构化）",
        when_use="已完成确认单填写并保存为 JSON，要把「清单 + 素材路径」落到 extract_jsons 各项目目录时执行（通常先于正文抽取）。",
        inputs_desc="默认：自动在仓库「第一批」下查找各项目 `01_管理确认` 内确认单 `.json`；也可命令行传入单个确认单文件路径。",
        outputs_desc="`extract_jsons/<项目文件夹名>/` 下三份 `*.json`；同级 `docs_metadata/` 内为按规则重命名的素材副本。",
        default_argv=(),
    ),
    PipelineScript(
        id="prepare_local_kb",
        filename="3_prepare_local_kb.py",
        title="本地知识库抽取（kb_local）",
        group="正文抽取",
        description="读取各项目 docs_metadata 中的文档，抽取为整篇 Markdown；生成 manifest.json，不落表格明细 JSON。",
        runner="subprocess_argv",
        phase="入库准备（可检索正文）",
        when_use="清单与 docs_metadata 已就绪，需要把 Word/PDF 等转成逐文件 Markdown、供质检或下游系统使用时执行。",
        inputs_desc="默认 `--extract-root extract_jsons`，即处理该树下各项目的 docs_metadata（及清单路径约定）。依赖 python-docx、pypdf/pdfplumber 等。",
        outputs_desc="各项目下 `kb_local/manifest.json` 与 `kb_local/extracted_markdown/*.md`（哈希命名）。",
        default_argv=("--extract-root", "extract_jsons"),
    ),
    PipelineScript(
        id="markdown_check",
        filename="4_markdown_quality_checker.py",
        title="Markdown 质量检测",
        group="质量与安全",
        description="对 Markdown 做结构、噪声、重复行、疑似表格、敏感关键词等诊断；生成单文件报告与 summary.json，不修改原文。",
        runner="markdown_check",
        phase="入库质检 / AI 候选评估",
        when_use="抽取完成后，准备「是否适合入库或给 AI」判断前执行；亦可对修复后的副本目录再跑一轮复核。",
        inputs_desc="面板默认扫描目录 `extract_jsons`（可改为 `extract_jsons_fixed` 等）；递归 `.md`/`.txt`。",
        outputs_desc="`markdown_quality_reports/`（或自定义 `--out`）下 `*.quality_report.json`、`*.review_items.md`、`summary.json`。",
    ),
    PipelineScript(
        id="markdown_fix",
        filename="5_markdown_quality_fix.py",
        title="Markdown 基础修复（输出副本）",
        group="质量与安全",
        description="按与检测一致的规则删除噪声行、压缩连续空行、折叠连续重复行；始终写入独立目录，不覆盖原始 extract_jsons。",
        runner="markdown_fix",
        phase="入库前清洗（可选）",
        when_use="检测报告显示大量页码行、空行等问题，希望在不影响原件的情况下得到一份更干净的副本时执行；可先 dry-run。",
        inputs_desc="默认扫描 `extract_jsons`；与检测使用相同路径参数逻辑。",
        outputs_desc="默认 `extract_jsons_fixed/` 下镜像相对路径；附 `fix_summary.json`。",
    ),
    PipelineScript(
        id="generate_mock",
        filename="generate_mock_tobacco_project.py",
        title="生成模拟资料包（测试）",
        group="工具与测试",
        description="默认写入 **`mockdata/`**（`full` / `dirs_only` / `missing_required_partial`；`full` 含 **`01_管理确认`** 确认单）。",
        runner="cli_hint_only",
        phase="开发 / 演示",
        when_use="无真实资料联调或演示；运维台一键调用 **POST /api/run/generate-mock**。",
        inputs_desc="`--scenario`、`--batch`（可选）、`--ext`、`--customer`、`--project`、`--seed`（残缺场景）；见脚本 `--help`。",
        outputs_desc="**`mockdata/<项目文件夹>/`** 下占位目录与文件（`full` 含确认单 MD/JSON）。",
        cli_example="python generate_mock_tobacco_project.py --scenario full",
    ),
    PipelineScript(
        id="clean_workspace_preview",
        filename="clean_workspace.py",
        title="清理工作区（仅预览 dry-run）",
        group="维护",
        description="模拟 `--all --dry-run`：打印将删除的「检查结果」内条目与「第一批」下一级项目目录，不真正删除。",
        runner="subprocess_argv",
        phase="环境重置（安全预览）",
        when_use="准备大批量清理前，先看会删掉哪些路径；确认无误后再到终端带 `--yes` 执行真实删除。",
        inputs_desc="默认工作区为当前目录；面板在仓库根执行，删除范围相对于该根下的「检查结果」「第一批」。",
        outputs_desc="仅在终端输出预览日志；不写文件。",
        default_argv=("--all", "--dry-run"),
    ),
    PipelineScript(
        id="clean_workspace",
        filename="clean_workspace.py",
        title="清理工作区（终端命令参考）",
        group="维护",
        description="真实删除「检查结果」内容或「第一批」下项目子目录；必须通过终端附加 `--yes` 等参数确认。",
        runner="cli_hint_only",
        phase="环境重置（不可逆）",
        when_use="仅在明确需要清空模拟数据或重建第一批目录时使用；务必先 dry-run。",
        inputs_desc="`--results` / `--mock` / `--all`，可选 `--base` 指定工作区根；详见脚本说明。",
        outputs_desc="磁盘上对应目录被删除；不可逆。",
        cli_example="python clean_workspace.py --all --dry-run",
        warn="真实删除不可逆；确认前务必预览。",
    ),
)


def get_script_by_id(sid: str) -> Optional[PipelineScript]:
    for s in PIPELINE_SCRIPTS:
        if s.id == sid:
            return s
    return None


def scripts_by_group() -> dict[str, List[PipelineScript]]:
    g: dict[str, List[PipelineScript]] = {}
    for s in PIPELINE_SCRIPTS:
        g.setdefault(s.group, []).append(s)
    return g
