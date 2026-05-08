"""
资料包文件命名约定 —— **验收脚本与 mock 生成共用**，避免两端规则漂移。

「验哪些材料、关键词、档位阈值」在 **`acceptance_config.py`** 配置；本文件只管文件名格式。

1. **两段式** —— 仅 **`01_管理确认`**：`资料名称_YYYYMMDD`（末段为 8 位日期，**无密级段**）。

2. **三段式** —— **`02_核心资料`** 下**七类**标准子目录：`资料名称_日期_密级`（倒数第二段 YYYYMMDD，末段密级）。验收不扫描「03_安全与整改」。**格式不要求「资料名称」含必交关键词**（关键词见验收脚本整包扫描）。

3. **完整七段**（其余路径 / 历史规则）：  
   `客户单位_项目名称_资料类别_资料名称_版本号_日期_密级`

其中「版本号」「日期」「密级」的格式以本模块中的正则与默认值为准。
"""

from __future__ import annotations

import datetime
import re
from typing import Final, Optional, Tuple

# ----- 版本号：第五段（倒数第三段），须含 Final 后缀（与验收检查一致） -----
VERSION_SEGMENT_PATTERN: Final[str] = r"V\d+\.\d+Final"
VERSION_SEGMENT_RE: Final[re.Pattern[str]] = re.compile(VERSION_SEGMENT_PATTERN)

# Mock / 示例用的默认版本段（必须通过 VERSION_SEGMENT_RE）
DEFAULT_VERSION_SEGMENT: Final[str] = "V1.0Final"

# ----- 日期：第六段（倒数第二段），YYYYMMDD -----
DATE_SEGMENT_PATTERN: Final[str] = r"\d{8}"
DATE_SEGMENT_RE: Final[re.Pattern[str]] = re.compile(DATE_SEGMENT_PATTERN)

DEFAULT_DATE_SEGMENT: Final[str] = "20260101"

# ----- 密级：第七段（最后一段） -----
VALID_SECRET_LEVELS: Final[frozenset[str]] = frozenset({"公开", "内部", "敏感", "核心"})
DEFAULT_SECRET_SEGMENT: Final[str] = "内部"

# ----- 资料类别：第三段；须与标准目录「NN_类别名」中的类别名一致 -----
VALID_CATEGORIES: Final[frozenset[str]] = frozenset(
    {
        "项目总览",
        "售前方案",
        "业务需求",
        "产品设计",
        "技术设计",
        "开发测试",
        "实施上线",
        "验收交付",
        "运维售后",
        "项目复盘",
        "数据指标",
    }
)

# ----- 允许的文件扩展名 -----
VALID_EXTENSIONS: Final[frozenset[str]] = frozenset(
    {
        ".doc",
        ".docx",
        ".xls",
        ".xlsx",
        ".ppt",
        ".pptx",
        ".pdf",
        ".png",
        ".jpg",
        ".jpeg",
        ".zip",
        ".rar",
    }
)


def try_parse_confirmation_bucket_stem(stem: str) -> Optional[Tuple[str, str]]:
    """
    解析「01_管理确认」两段式：资料名称_YYYYMMDD（末段为日期，无密级）。
    资料名称可含下划线；合法则返回 (资料名称, YYYYMMDD)。
    """
    parts = stem.split("_")
    if len(parts) < 2:
        return None
    date = parts[-1]
    doc = "_".join(parts[:-1])
    if not doc.strip():
        return None
    if not DATE_SEGMENT_RE.fullmatch(date):
        return None
    try:
        datetime.datetime.strptime(date, "%Y%m%d")
    except ValueError:
        return None
    return (doc, date)


def try_parse_simplified_stem(stem: str) -> Optional[Tuple[str, str, str]]:
    """
    解析「资料名称_日期_密级」三段式主干（资料名称可含下划线，取倒数两段为日期与密级）。
    用于「02_核心资料」下七类标准子目录；合法则返回 (资料名称, YYYYMMDD, 密级)。
    仅做格式解析，不校验资料名称是否含业务关键词。
    """
    parts = stem.split("_")
    if len(parts) < 3:
        return None
    secret = parts[-1]
    date = parts[-2]
    doc = "_".join(parts[:-2])
    if not doc.strip():
        return None
    if secret not in VALID_SECRET_LEVELS:
        return None
    if not DATE_SEGMENT_RE.fullmatch(date):
        return None
    try:
        datetime.datetime.strptime(date, "%Y%m%d")
    except ValueError:
        return None
    return (doc, date, secret)


def stem_is_legacy_seven_segment(stem: str) -> bool:
    """是否为完整七段命名（以版本段特征识别，避免与三段式混淆）。"""
    parts = stem.split("_")
    if len(parts) < 7:
        return False
    return bool(VERSION_SEGMENT_RE.fullmatch(parts[-3]))


def build_standard_filename_stem(
    customer: str,
    project: str,
    category_label: str,
    doc_label: str,
    *,
    version_segment: str = DEFAULT_VERSION_SEGMENT,
    date_segment: str = DEFAULT_DATE_SEGMENT,
    secret_segment: str = DEFAULT_SECRET_SEGMENT,
) -> str:
    """
    生成符合验收规则的七段文件名主干（不含后缀）。
    category_label 一般为标准目录名的下划线后半段，例如「项目总览」「售前方案」。
    """
    return (
        f"{customer}_{project}_{category_label}_{doc_label}_"
        f"{version_segment}_{date_segment}_{secret_segment}"
    )
