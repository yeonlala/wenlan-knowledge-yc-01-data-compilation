# -*- coding: utf-8 -*-
"""
================================================================================
                        资料包验收 —— 唯一配置入口
================================================================================

**只需编辑本文件**，即可同时影响：
  • `1_check_tobacco_kb_required_files.py`（自动验收 Excel）
  • `generate_mock_tobacco_project.py`（模拟资料生成）

以下可调：
  1. `PHASE1_RULES` — 第一阶段「必验材料」清单（关键词、建议目录、mock 嵌入词）
  2. `ACCEPTANCE_THRESHOLDS` — 验收建议档位（A/B/C/D）的数值门槛

**文件名格式**（版本号 `V1.0Final`、日期 `YYYYMMDD`、密级、扩展名白名单等）
不在本文件，而在同目录 `naming_convention.py`，避免单文件过长；
若单位只改「验哪些材料」，通常只改本文件即可。

------------------------------------------------------------------------
如何调整必验材料（PHASE1_RULES）
------------------------------------------------------------------------
• 每条规则是一个字典，维护字段含义：
    id               规则编号，建议 r01、r02… 唯一
    label            报表里显示的说明标题
    keywords         文件名中命中任一关键词即视为「有」这份材料（元组）
    mock_std_dir     建议标准**叶子**目录名，须与 `STANDARD_DIRS` 中某项一致；物理路径为 `02_核心资料/<mock_std_dir>`
    mock_embed_keyword  Mock 生成文件名时必须包含的词，且必须是 keywords 里的一项
    note             备注，可选

• **目录分层**：项目根下须具备 **01_管理确认**、**02_核心资料**（**03_安全与整改** 由本仓库 check / mock **跳过**，不作验收要求）。**第一阶段**各规则 **mock_std_dir** 均落在 **02_核心资料** 下（**r01→01_项目总览 … r07→07_项目复盘**，不混夹）。验收「缺目录」按此路径检查。
• **可选规则**：在规则 dict 中设 **`"optional": True`** 时，该条**不参与**必交完整率、亦不强制要求对应标准子目录存在（兼容旧项目）；报表仍展示该项，缺失时标为「缺（可选）」。

• **注释掉整条规则**：在该条首尾加 `#`，或直接从列表里删掉这一段字典。
• **改成单位自己的关键词**：只改 keywords / mock_embed_keyword，两者保持一致性
  （mock_embed_keyword 必须是 keywords 里的某一个）。
• **验收脚本**：若 **`02_核心资料/<mock_std_dir>`** 下已有任意文件，该条必交即视为「有」；目录为空时再按 **keywords** 扫全包文件名。（标记 **optional** 的规则不要求目录必存、不计入必交缺失数。）

• 规则列表可以为空：表示当前不验任何「必交材料」项（仅做目录与命名等检查）。

------------------------------------------------------------------------
"""

from __future__ import annotations

from typing import Any, Dict, List, Tuple, TypedDict


class _LevelThreshold(TypedDict, total=False):
    naming_pass_min: float
    required_complete_min: float
    max_missing_required: int
    require_zero_high_risk: bool


class AcceptanceThresholds(TypedDict):
    """验收建议档位阈值；可按单位制度微调数字。"""

    level_a: _LevelThreshold
    level_b: _LevelThreshold
    level_c: _LevelThreshold


# ----- 验收建议 A/B/C/D 档位（与 calculate_level 逻辑对应）-----
ACCEPTANCE_THRESHOLDS: AcceptanceThresholds = {
    "level_a": {
        "naming_pass_min": 90.0,
        "required_complete_min": 100.0,
        "max_missing_required": 0,
        "require_zero_high_risk": True,
    },
    "level_b": {
        "naming_pass_min": 80.0,
        "required_complete_min": 80.0,
        "max_missing_required": 1,
    },
    "level_c": {
        "naming_pass_min": 60.0,
        "required_complete_min": 60.0,
    },
}


# ----- 第一阶段必验材料（唯一数据源；说明见文件头「第一阶段目录口径」）-----
PHASE1_RULES: List[Dict[str, Any]] = [
    {
        "id": "r01",
        "label": "1.项目基本信息表",
        "match": "any",
        "keywords": ("项目基本信息表",),
        "note": "台账",
        "mock_std_dir": "01_项目总览",
        "mock_embed_keyword": "项目基本信息表",
    },
    {
        "id": "r02",
        "label": "2.建设方案/投标方案/汇报方案（三选一）",
        "match": "any",
        "keywords": ("建设方案", "投标方案", "汇报方案"),
        "note": "投标方案建议脱敏",
        "mock_std_dir": "02_售前方案",
        "mock_embed_keyword": "建设方案",
    },
    {
        "id": "r03",
        "label": "3.需求规格说明书/功能清单（二选一）",
        "match": "any",
        "keywords": ("需求规格说明书", "功能清单"),
        "mock_std_dir": "03_业务需求",
        "mock_embed_keyword": "需求规格说明书",
    },
    {
        "id": "r04",
        "label": "4.业务流程图",
        "match": "any",
        "keywords": ("业务流程图", "流程图"),
        "mock_std_dir": "04_产品设计",
        "mock_embed_keyword": "业务流程图",
    },
    {
        "id": "r05",
        "label": "5.用户手册/操作手册/培训资料（三选一）",
        "match": "any",
        "keywords": ("用户手册", "操作手册", "培训资料"),
        "note": "mock 示例文件名优先含「培训资料」",
        "mock_std_dir": "05_实施上线",
        "mock_embed_keyword": "培训资料",
    },
    {
        "id": "r06",
        "label": "6.验收报告/最终交付清单（二选一）",
        "match": "any",
        "keywords": ("验收报告", "最终交付清单"),
        "note": "已验收项目须具备其一",
        "mock_std_dir": "06_验收交付",
        "mock_embed_keyword": "验收报告",
    },
    {
        "id": "r07",
        "label": "7.项目知识沉淀表/项目复盘表（二选一）",
        "match": "any",
        "keywords": ("项目知识沉淀表", "项目复盘表", "项目复盘报告"),
        "mock_std_dir": "07_项目复盘",
        "mock_embed_keyword": "项目知识沉淀表",
        # 旧项目可无「07_项目复盘」及本条材料；有则仍按关键词/目录认定
        "optional": True,
    },
]

# ----- 顶层分区资料（当前仅 01 管理确认）：文件名关键词 + 必须在对应分区目录内 -----
# 「03_安全与整改」暂不纳入自动验收与 mock（见 check / mock 跳过逻辑）。
TOP_BUCKET_ALLOWED = frozenset({"01_管理确认"})

TOP_BUCKET_RULES: List[Dict[str, Any]] = [
    {
        "id": "t01",
        "label": "项目资料上报确认单（文件名须含其一）",
        "keywords": (
            "项目资料上报确认单",
            "资料上报确认单",
            "资料确认单",
            "移交",
            "签收",
            "确认书",
            "报送确认",
        ),
        "bucket_dir": "01_管理确认",
        "note": "须放置《项目资料上报确认单》导出的 .md 与 .json；文件名须含关键词之一",
    },
]


def validate_phase1_rules(rules: List[Dict[str, Any]]) -> None:
    """导入时校验规则结构；列表为空则跳过。"""
    if not rules:
        return
    seen: set[str] = set()
    for r in rules:
        rid = str(r["id"])
        if rid in seen:
            raise ValueError(f"acceptance_config：重复的规则 id：{rid}")
        seen.add(rid)
        emb = str(r["mock_embed_keyword"])
        kws: Tuple[str, ...] = tuple(r["keywords"])
        if r.get("optional") is not None and not isinstance(r.get("optional"), bool):
            raise ValueError(f"规则 {rid}: optional 须为 True/False，缺省表示必验")
        if emb not in kws:
            raise ValueError(
                f"规则 {rid}: mock_embed_keyword {emb!r} 必须是 keywords 之一，"
                "否则 mock 文件名与验收关键词不一致"
            )


validate_phase1_rules(PHASE1_RULES)


def validate_top_bucket_rules(rules: List[Dict[str, Any]]) -> None:
    """校验顶层分区规则；列表为空则跳过。"""
    if not rules:
        return
    seen: set[str] = set()
    for r in rules:
        rid = str(r["id"])
        if rid in seen:
            raise ValueError(f"acceptance_config：重复的顶层分区规则 id：{rid}")
        seen.add(rid)
        bd = str(r["bucket_dir"])
        if bd not in TOP_BUCKET_ALLOWED:
            raise ValueError(
                f"规则 {rid}: bucket_dir {bd!r} 须为 {sorted(TOP_BUCKET_ALLOWED)!r} 之一"
            )


validate_top_bucket_rules(TOP_BUCKET_RULES)
