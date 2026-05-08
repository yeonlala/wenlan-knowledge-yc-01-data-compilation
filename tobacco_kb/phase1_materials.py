"""
兼容旧导入路径。

第一阶段规则的唯一维护入口已迁至 **`acceptance_config.py`**（`PHASE1_RULES`、验收档位阈值）。
请编辑该文件；此处仅重新导出，勿在本文件重复维护规则。
"""

from __future__ import annotations

from tobacco_kb.acceptance_config import PHASE1_RULES

__all__ = ["PHASE1_RULES"]
