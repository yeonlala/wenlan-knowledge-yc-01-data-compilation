# -*- coding: utf-8 -*-
"""
与「模板示例/确认单生成器/index.html」导出结构一致的确认单 JSON / Markdown 生成（供 Mock 自动化，无需浏览器）。

schema 与 collectData() / buildMarkdown() 保持一致。
"""

from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, Tuple


def _yes_no(val: Any) -> str:
    if val is True:
        return "是"
    if val is False:
        return "否"
    return "否"


def build_confirmation_payload_full(
    *,
    branch_name: str,
    customer_name: str,
    project_name: str,
    project_year: str,
    project_status: str,
    package_name: str,
    submit_date: str,
    business_domain: str = "信息化",
    is_sensitive: str = "否",
    ai_candidate_status: str = "否",
    remark: str = "",
) -> Dict[str, Any]:
    """full Mock：清单项全部勾选为已提交，便于联调 2_export_confirmation_bundle。"""
    now = datetime.now(timezone.utc).replace(microsecond=0)
    gen_at = now.isoformat().replace("+00:00", "Z")

    def row() -> Dict[str, Any]:
        return {"submitted": True, "note": ""}

    return {
        "schema_version": "1.0",
        "document_type": "项目资料上报确认单",
        "project": {
            "branch_name": branch_name,
            "customer_name": customer_name,
            "project_name": project_name,
            "project_year": project_year,
            "project_status": project_status,
            "business_domain": business_domain,
            "package_name": package_name,
            "submit_date": submit_date,
            "is_sensitive": is_sensitive,
            "ai_candidate_status": ai_candidate_status,
        },
        "checklist": {
            "project_basic_info": row(),
            "solution_plan": row(),
            "requirement_or_function_list": row(),
            "business_flow": row(),
            "manual_or_training": row(),
            "acceptance_or_delivery": row(),
            "knowledge_summary_or_exemption": row(),
        },
        "security": {
            "identified_sensitive_materials": "否",
            "sanitized_or_marked": "是",
            "no_cloud_attachment": "是",
        },
        "responsibility": {
            "material_owner": "资料整理人(Mock)",
            "material_owner_role": "信息中心",
            "project_owner": "项目负责人(Mock)",
            "project_owner_role": "项目组",
            "branch_owner": "分公司负责人(Mock)",
            "branch_owner_role": "管理部门",
            "confirmation_method": "姓名确认",
        },
        "remark": remark
        or "本确认单由 generate_mock_tobacco_project.py（scenario=full）按确认单生成器模板自动生成。",
        "generated_at": gen_at,
    }


def build_yaml_frontmatter(data: Dict[str, Any]) -> str:
    p = data["project"]
    r = data["responsibility"]
    return f"""---
schema_version: "1.0"
document_type: 项目资料上报确认单
branch_name: {p.get("branch_name") or ""}
customer_name: {p.get("customer_name") or ""}
project_name: {p.get("project_name") or ""}
project_year: "{p.get('project_year') or ''}"
project_status: {p.get("project_status") or ""}
business_domain: {p.get("business_domain") or ""}
package_name: {p.get("package_name") or ""}
submit_date: "{p.get('submit_date') or ''}"
is_sensitive: {p.get("is_sensitive") or ""}
ai_candidate_status: {p.get("ai_candidate_status") or ""}
material_owner: {r.get("material_owner") or ""}
project_owner: {r.get("project_owner") or ""}
branch_owner: {r.get("branch_owner") or ""}
confirmation_method: {r.get("confirmation_method") or ""}
---
"""


def render_confirmation_markdown(data: Dict[str, Any]) -> str:
    """对齐 index.html 中 buildMarkdown（不含浏览器差异）。"""
    p = data["project"]
    c = data["checklist"]
    s = data["security"]
    r = data["responsibility"]

    def yn(key: str, field: str = "submitted") -> str:
        item = c.get(key) or {}
        return _yes_no(item.get(field))

    body = f"""{build_yaml_frontmatter(data)}

# 项目资料上报确认单

> 说明：本确认单用于确认项目资料包的提交责任、资料真实性、资料完整性、密级与脱敏情况。  
> 项目资料包不得作为钉钉、微信、个人网盘、个人邮箱等云端附件上传。  
> 如需走钉钉 / OA 审批，可在清单中填写资料包名称及责任人、确认意见等，不上传真实项目资料附件。

---

## 一、项目基本信息

| 字段 | 内容 |
|---|---|
| 分公司名称 | {p.get("branch_name") or ""} |
| 客户单位 | {p.get("customer_name") or ""} |
| 项目名称 | {p.get("project_name") or ""} |
| 项目年份 | {p.get("project_year") or ""} |
| 项目状态 | {p.get("project_status") or ""} |
| 业务板块 | {p.get("business_domain") or ""} |
| 提交日期 | {p.get("submit_date") or ""} |
| 是否涉敏 | {p.get("is_sensitive") or ""} |
| 是否作为 AI 入库候选 | {p.get("ai_candidate_status") or ""} |

---

## 二、资料提交清单

| 序号 | 材料名称 | 是否提交 | 备注 |
|---:|---|---|---|
| 1 | 资料包名称 | — | {p.get("package_name") or ""} |
| 2 | 项目基本信息表 | {yn("project_basic_info")} | {(c.get("project_basic_info") or {}).get("note") or ""} |
| 3 | 建设方案 / 投标方案 / 汇报方案 | {yn("solution_plan")} | {(c.get("solution_plan") or {}).get("note") or ""} |
| 4 | 需求规格说明书 / 功能清单 | {yn("requirement_or_function_list")} | {(c.get("requirement_or_function_list") or {}).get("note") or ""} |
| 5 | 业务流程图 | {yn("business_flow")} | {(c.get("business_flow") or {}).get("note") or ""} |
| 6 | 用户手册 / 操作手册 / 培训资料 | {yn("manual_or_training")} | {(c.get("manual_or_training") or {}).get("note") or ""} |
| 7 | 验收报告 / 最终交付清单 | {yn("acceptance_or_delivery")} | {(c.get("acceptance_or_delivery") or {}).get("note") or ""} |
| 8 | 项目知识沉淀表 / 项目复盘表 / 项目知识简表 / 无复盘资料说明 | {yn("knowledge_summary_or_exemption")} | {(c.get("knowledge_summary_or_exemption") or {}).get("note") or ""} |

---

## 三、安全确认

| 确认项 | 结果 |
|---|---|
| 是否已识别涉敏资料 | {s.get("identified_sensitive_materials") or ""} |
| 是否已脱敏或标记敏感 | {s.get("sanitized_or_marked") or ""} |
| 是否未上传云端附件 | {s.get("no_cloud_attachment") or ""} |

---

## 四、责任人确认

| 责任类型 | 姓名 | 部门 / 岗位 |
|---|---|---|
| 资料整理人 | {r.get("material_owner") or ""} | {r.get("material_owner_role") or ""} |
| 项目负责人 | {r.get("project_owner") or ""} | {r.get("project_owner_role") or ""} |
| 分公司负责人 | {r.get("branch_owner") or ""} | {r.get("branch_owner_role") or ""} |

确认方式：{r.get("confirmation_method") or ""}

---

## 五、补充说明

{data.get("remark") or ""}

---

## 六、承诺说明

1. 本次提交的项目资料真实、有效，来源可追溯；
2. 已按照公司资料整理要求进行目录归档和文件命名；
3. 已对资料进行初步密级标记；
4. 对涉及客户数据、合同报价、接口信息、账号密码、服务器信息、数据库信息等敏感内容，已按要求进行标记或脱敏；
5. 未故意隐瞒、篡改、伪造、混入其他项目资料；
6. 未通过钉钉、微信、个人网盘、个人邮箱等云端渠道上传真实项目资料附件；
7. 后续如总部验收发现资料缺失、命名不规范、密级标记错误、脱敏不到位或内容失实，将按要求配合整改；
8. 如因资料虚假、遗漏、违规提交或涉敏资料处理不当造成风险，由相关责任人按公司制度承担相应责任。
"""
    return body


def filename_date_compact(submit_date: str) -> str:
    """submit_date 为 YYYY-MM-DD 时转为 YYYYMMDD，用于文件名。"""
    s = (submit_date or "").strip().replace("-", "")
    if len(s) >= 8 and s[:8].isdigit():
        return s[:8]
    return datetime.now().strftime("%Y%m%d")


def write_confirmation_pair(
    admin_confirm_dir: Path,
    data: Dict[str, Any],
    *,
    submit_date: str,
) -> Tuple[Path, Path]:
    """写入 .md + .json，与生成器默认命名一致。"""
    admin_confirm_dir.mkdir(parents=True, exist_ok=True)
    dcompact = filename_date_compact(submit_date)
    base = f"项目资料上报确认单_{dcompact}"
    md_path = admin_confirm_dir / f"{base}.md"
    json_path = admin_confirm_dir / f"{base}.json"
    md_path.write_text(render_confirmation_markdown(data), encoding="utf-8")
    json_path.write_text(
        json.dumps(data, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    return md_path, json_path
