# -*- coding: utf-8 -*-
"""
项目说明与执行步骤（仅供运维台展示）。

主线仅三阶段：验收 → 提取三份清单 → 提取文字。
"""

from __future__ import annotations

from typing import Any, Dict, List

PROJECT_TITLE = "yc-checking · 资料入库执行指引"

PROJECT_SUMMARY = (
    "主线三阶段：验收 → 提取三份清单 → 提取文字。"
    "无真实资料时可先在 **mockdata/** 生成 Mock（资料准备区），整夹复制到 **第一批/** 再走验收。"
)

WHO_USES = ""

OPERATION_PRINCIPLE = (
    "路径均为相对仓库根目录；改输入框后，再点「运行」。"
)

CORE_CHAIN_LINE_ITEMS: List[str] = [
    "阶段一 · 验收：资料入「第一批」+（可选）① → 检查结果",
    "阶段二 · 提取三份清单：② → extract_jsons/",
    "阶段三 · 提取文字：③ → kb_local/extracted_markdown/",
]

DATA_FLOW_STEPS: List[Dict[str, Any]] = [
    {
        "order": 1,
        "title": "阶段一 · 验收",
        "path_hint": "第一批/<项目>/ → 检查结果/",
        "caption": "先把资料放进「第一批」。需要机检时运行①，在「检查结果」看目录与齐套情况。",
        "related_script_ids": [],
        "note": "",
    },
    {
        "order": 2,
        "title": "阶段二 · 提取三份清单",
        "path_hint": "extract_jsons/<项目>/ · docs_metadata/",
        "caption": "运行②：从确认单生成三份清单 JSON，并准备 docs_metadata 素材路径。",
        "related_script_ids": ["export_confirmation_bundle"],
        "note": "",
    },
    {
        "order": 3,
        "title": "阶段三 · 提取文字",
        "path_hint": "…/kb_local/extracted_markdown/",
        "caption": "运行③：把文档转为逐篇 Markdown，并写 manifest.json。",
        "related_script_ids": ["prepare_local_kb"],
        "note": "",
    },
]
