# -*- coding: utf-8 -*-
"""
《项目资料上报确认单》→ 固定扁平 KV（入库用）。

• 同一份 schema、同一套键名、缺省一律空字符串。
• 正文不规范时：尽量按标签模糊匹配 + 按表格顺序兜底；异常写入 parse_warnings（仍为字符串）。
"""

from __future__ import annotations

import re
from typing import Any, Dict, List, Mapping, Optional, Sequence, Tuple

FLAT_SCHEMA_VERSION = 4

# ---------------------------------------------------------------------------
# 全量键表（顺序即输出顺序；值一律 str）
# ---------------------------------------------------------------------------

_METAVARS = (
    "file_path",
    "file_relative_path",
    "file_name",
    "file_format",
    "stem_doc_name",
    "stem_date",
    "parse_warnings",
)

_BASIC = (
    "basic_branch_company",
    "basic_customer_unit",
    "basic_project_name",
    "basic_project_year",
    "basic_project_status",
    "basic_business_segment",
    "basic_submit_date",
    "basic_sensitive_flag",
    "basic_ai_candidate",
)


def _subm_keys() -> Tuple[str, ...]:
    keys: List[str] = []
    for i in range(1, 10):
        p = f"{i:02d}"
        keys.extend((f"subm_{p}_material", f"subm_{p}_submitted", f"subm_{p}_remark"))
    return tuple(keys)


def _norm_keys() -> Tuple[str, ...]:
    keys: List[str] = []
    for i in range(1, 7):
        p = f"{i:02d}"
        keys.extend((f"norm_{p}_confirm", f"norm_{p}_remark"))
    return tuple(keys)


def _sec_keys() -> Tuple[str, ...]:
    keys: List[str] = []
    for i in range(1, 8):
        p = f"{i:02d}"
        keys.extend((f"sec_{p}_confirm", f"sec_{p}_remark"))
    return tuple(keys)


def _resp_keys() -> Tuple[str, ...]:
    keys: List[str] = []
    for i in range(1, 4):
        p = f"{i:02d}"
        keys.extend(
            (
                f"resp_{p}_role",
                f"resp_{p}_name",
                f"resp_{p}_dept",
                f"resp_{p}_contact",
                f"resp_{p}_date",
            )
        )
    return tuple(keys)


def _duty_keys() -> Tuple[str, ...]:
    keys: List[str] = []
    for i in range(1, 4):
        p = f"{i:02d}"
        keys.extend((f"duty_{p}_role", f"duty_{p}_content"))
    return tuple(keys)


def _meth_keys() -> Tuple[str, ...]:
    return tuple(f"meth_{i:02d}_value" for i in range(1, 7))


def _commit_keys() -> Tuple[str, ...]:
    return tuple(f"commit_{i:02d}" for i in range(1, 9))


def _sign_keys() -> Tuple[str, ...]:
    keys: List[str] = []
    for i in range(1, 4):
        p = f"{i:02d}"
        keys.extend((f"sign_{p}_label", f"sign_{p}_signer", f"sign_{p}_date", f"sign_{p}_remark"))
    return tuple(keys)


ALL_FLAT_KEYS: Tuple[str, ...] = (
    _METAVARS
    + _BASIC
    + _subm_keys()
    + _norm_keys()
    + _sec_keys()
    + _resp_keys()
    + _duty_keys()
    + _meth_keys()
    + _commit_keys()
    + _sign_keys()
)


def empty_flat_document() -> Dict[str, str]:
    return {k: "" for k in ALL_FLAT_KEYS}


_ws_re = re.compile(r"\s+")


def _norm_label(s: str) -> str:
    if not s:
        return ""
    t = str(s).replace("\u3000", " ").strip()
    t = _ws_re.sub("", t)
    return t


def _contains_any(hay: str, needles: Tuple[str, ...]) -> bool:
    h = _norm_label(hay)
    return any(n in h for n in needles if n)


# ----- 标签 → basic_*（先精确匹配表头，再子串兜底，避免「资料包*」互串） -----
_BASIC_EXACT: Tuple[Tuple[str, str], ...] = (
    ("分公司名称", "basic_branch_company"),
    ("客户单位", "basic_customer_unit"),
    ("项目名称", "basic_project_name"),
    ("项目年份", "basic_project_year"),
    ("项目状态", "basic_project_status"),
    ("业务板块", "basic_business_segment"),
    ("提交日期", "basic_submit_date"),
    ("是否涉敏", "basic_sensitive_flag"),
    ("是否作为AI入库候选", "basic_ai_candidate"),
)

_BASIC_FUZZY: Tuple[Tuple[str, Tuple[str, ...]], ...] = (
    ("basic_ai_candidate", ("AI入库候选", "是否作为AI", "入库候选")),
)


def _fill_basic_from_kv(src: Mapping[str, str], flat: Dict[str, str]) -> None:
    if not src:
        return
    exact_norm = {_norm_label(k): v for k, v in _BASIC_EXACT}

    for label, val in src.items():
        nl = _norm_label(label)
        if not nl:
            continue
        canon = exact_norm.get(nl)
        if canon:
            flat[canon] = str(val).strip()

    for canon, hints in _BASIC_FUZZY:
        if flat.get(canon):
            continue
        for label, val in src.items():
            nl = _norm_label(label)
            if not nl:
                continue
            for h in hints:
                if h and _norm_label(h) in nl:
                    flat[canon] = str(val).strip()
                    break
            if flat.get(canon):
                break


# ----- 资料提交清单：按模板顺序号或材料关键词 -----
_SUBM_HINTS: Tuple[Tuple[str, ...], ...] = (
    ("项目基本信息表",),
    ("建设方案", "投标方案", "汇报方案"),
    ("需求规格", "功能清单"),
    ("业务流程", "流程图"),
    ("用户手册", "操作手册", "培训资料"),
    ("验收报告", "最终交付清单", "交付清单"),
    ("知识沉淀", "复盘"),
    ("密级", "脱敏确认表"),
    ("其他补充",),
)


def _fill_submission(recs: List[Dict[str, str]], flat: Dict[str, str]) -> None:
    if not recs:
        return
    for r in recs:
        # 列名兼容
        seq = ""
        mat = ""
        sub = ""
        rem = ""
        for lk, lv in r.items():
            lkn = _norm_label(lk)
            if "序号" in lkn or lkn == "no":
                seq = lv.strip()
            elif "材料" in lkn or "名称" in lkn:
                mat = lv.strip()
            elif "提交" in lkn or "是否" in lkn:
                sub = lv.strip()
            elif "备注" in lkn:
                rem = lv.strip()

        idx: Optional[int] = None
        if seq.isdigit():
            n = int(seq)
            if 1 <= n <= 9:
                idx = n
        if idx is None and mat:
            for i, hints in enumerate(_SUBM_HINTS, start=1):
                if _contains_any(mat, hints):
                    idx = i
                    break
        if idx is None:
            continue
        p = f"{idx:02d}"
        if not flat[f"subm_{p}_material"]:
            flat[f"subm_{p}_material"] = mat
        flat[f"subm_{p}_submitted"] = sub or flat[f"subm_{p}_submitted"]
        flat[f"subm_{p}_remark"] = rem or flat[f"subm_{p}_remark"]

    # 无表头：按行序填满 1..N
    if not any(flat.get(f"subm_{i:02d}_material") for i in range(1, 10)):
        for i, r in enumerate(recs[:9], start=1):
            vals = [str(x).strip() for x in r.values() if str(x).strip()]
            if len(vals) >= 4:
                _, mat, sub, rem = vals[0], vals[1], vals[2], vals[3]
            elif len(vals) >= 3:
                mat, sub, rem = vals[0], vals[1], vals[2]
            elif len(vals) >= 2:
                mat, sub = vals[0], vals[1]
                rem = ""
            else:
                continue
            p = f"{i:02d}"
            flat[f"subm_{p}_material"] = mat
            flat[f"subm_{p}_submitted"] = sub
            flat[f"subm_{p}_remark"] = rem


_NORM_HINTS: Tuple[Tuple[str, ...], ...] = (
    ("已按项目资料包目录整理", "目录整理"),
    ("文件命名", "命名"),
    ("资料属于本项目", "属于本项目"),
    ("可归档", "有效资料"),
    ("剔除", "重复"),
    ("知识沉淀表",),
)

_SEC_HINTS: Tuple[Tuple[str, ...], ...] = (
    ("已识别涉敏",),
    ("客户真实数据", "脱敏"),
    ("合同金额", "报价"),
    ("接口", "账号", "密码", "服务器", "数据库"),
    ("高风险", "账号密码", "生产环境"),
    ("未通过钉钉", "微信", "个人网盘", "个人邮箱"),
    ("如走钉钉", "仅填写内网", "不上传项目资料"),
)


def _fill_norm_like(
    recs: List[Dict[str, str]],
    hints: Tuple[Tuple[str, ...], ...],
    prefix: str,
    flat: Dict[str, str],
) -> None:
    if not recs:
        return
    filled_row = [False] * len(hints)

    for r in recs:
        vals = [str(v).strip() for v in r.values()]
        item = conf = rem = ""
        if len(vals) >= 3:
            item, conf, rem = vals[0], vals[1], vals[2]
        elif len(vals) == 2:
            item, conf = vals[0], vals[1]
        else:
            continue

        if not item:
            continue
        for idx, group in enumerate(hints):
            if filled_row[idx]:
                continue
            if _contains_any(item, group):
                p = f"{idx + 1:02d}"
                flat[f"{prefix}_{p}_confirm"] = conf
                flat[f"{prefix}_{p}_remark"] = rem
                filled_row[idx] = True
                break

    if not any(filled_row):
        for i, r in enumerate(recs[: len(hints)], start=1):
            vals = [str(x).strip() for x in r.values()]
            if len(vals) >= 3:
                conf, rem = vals[1], vals[2]
            elif len(vals) >= 2:
                conf, rem = vals[1], ""
            else:
                continue
            p = f"{i:02d}"
            flat[f"{prefix}_{p}_confirm"] = conf
            flat[f"{prefix}_{p}_remark"] = rem


def _fill_responsible(recs: List[Dict[str, str]], flat: Dict[str, str]) -> None:
    if not recs:
        return
    role_hints = (
        ("资料整理人", "整理人"),
        ("项目负责人", "项目经理"),
        ("分公司负责人", "负责人"),
    )
    for r in recs:
        role = name = dept = contact = date_s = ""
        for lk, lv in r.items():
            n = lk.strip()
            if "责任类型" in n or "类型" in n or "角色" in n:
                role = str(lv).strip()
            elif "姓名" in n:
                name = str(lv).strip()
            elif "部门" in n or "岗位" in n:
                dept = str(lv).strip()
            elif "联系" in n:
                contact = str(lv).strip()
            elif "日期" in n:
                date_s = str(lv).strip()
        if not role and r:
            vals = list(r.values())
            if len(vals) >= 5:
                role, name, dept, contact, date_s = [str(x).strip() for x in vals[:5]]

        row_idx: Optional[int] = None
        if role:
            for i, hints in enumerate(role_hints, start=1):
                if _contains_any(role, hints):
                    row_idx = i
                    break
        if row_idx is None:
            continue
        p = f"{row_idx:02d}"
        flat[f"resp_{p}_role"] = role
        flat[f"resp_{p}_name"] = name
        flat[f"resp_{p}_dept"] = dept
        flat[f"resp_{p}_contact"] = contact
        flat[f"resp_{p}_date"] = date_s

    if not any(flat.get(f"resp_{i:02d}_name") for i in range(1, 4)):
        for i, r in enumerate(recs[:3], start=1):
            vals = [str(x).strip() for x in r.values()]
            if len(vals) >= 5:
                role, name, dept, contact, date_s = vals[:5]
                p = f"{i:02d}"
                flat[f"resp_{p}_role"] = role
                flat[f"resp_{p}_name"] = name
                flat[f"resp_{p}_dept"] = dept
                flat[f"resp_{p}_contact"] = contact
                flat[f"resp_{p}_date"] = date_s


def _fill_duty(recs: List[Dict[str, str]], flat: Dict[str, str]) -> None:
    if not recs:
        return
    for i, r in enumerate(recs[:3], start=1):
        keys = list(r.keys())
        vals = list(r.values())
        if len(vals) >= 2:
            flat[f"duty_{i:02d}_role"] = str(vals[0]).strip()
            flat[f"duty_{i:02d}_content"] = str(vals[1]).strip()


def _fill_method(recs: List[Dict[str, str]], flat: Dict[str, str]) -> None:
    if not recs:
        return
    for i, r in enumerate(recs[:6], start=1):
        vals = list(r.values())
        if len(vals) >= 2:
            flat[f"meth_{i:02d}_value"] = str(vals[1]).strip()
        elif len(vals) == 1:
            flat[f"meth_{i:02d}_value"] = str(vals[0]).strip()


def _fill_sign(recs: List[Dict[str, str]], flat: Dict[str, str]) -> None:
    if not recs:
        return
    for r in recs:
        lab = signer = date_s = rem = ""
        for lk, lv in r.items():
            n = lk.strip()
            if "签署" in n or "项" in n:
                lab = str(lv).strip()
            elif "姓名" in n or "确认人" in n:
                signer = str(lv).strip()
            elif "日期" in n:
                date_s = str(lv).strip()
            elif "备注" in n:
                rem = str(lv).strip()
        row_idx: Optional[int] = None
        if lab:
            if "资料整理人" in lab:
                row_idx = 1
            elif "项目负责人" in lab:
                row_idx = 2
            elif "分公司负责人" in lab:
                row_idx = 3
        if row_idx is None:
            continue
        p = f"{row_idx:02d}"
        flat[f"sign_{p}_label"] = lab
        flat[f"sign_{p}_signer"] = signer
        flat[f"sign_{p}_date"] = date_s
        flat[f"sign_{p}_remark"] = rem

    if not any(flat.get(f"sign_{i:02d}_signer") for i in range(1, 4)):
        for i, r in enumerate(recs[:3], start=1):
            vals = [str(x).strip() for x in r.values()]
            if len(vals) >= 4:
                lab, signer, date_s, rem = vals[:4]
                p = f"{i:02d}"
                flat[f"sign_{p}_label"] = lab
                flat[f"sign_{p}_signer"] = signer
                flat[f"sign_{p}_date"] = date_s
                flat[f"sign_{p}_remark"] = rem


def _fill_commitment(paragraphs: List[str], flat: Dict[str, str]) -> None:
    for i, line in enumerate(paragraphs[:8], start=1):
        flat[f"commit_{i:02d}"] = line.strip()


def _gather_kv_tables_from_sections(sections: Any) -> List[Dict[str, str]]:
    """从分区原文里收集所有 2 列表格，合并为标签→值（用于不规范正文兜底）。"""
    acc: List[Dict[str, str]] = []
    if not isinstance(sections, list):
        return acc
    for sec in sections:
        if not isinstance(sec, dict):
            continue
        for mat in sec.get("tables") or []:
            if not isinstance(mat, list):
                continue
            for row in mat:
                if not isinstance(row, list) or len(row) < 2:
                    continue
                k, v = str(row[0]).strip(), str(row[1]).strip()
                if k and v:
                    acc.append({k: v})
    return acc


def _merge_kv(rows: List[Dict[str, str]]) -> Dict[str, str]:
    m: Dict[str, str] = {}
    for r in rows:
        for k, v in r.items():
            if k not in m:
                m[k] = v
    return m


def build_flat_record(
    meta: Mapping[str, Any],
    form: Mapping[str, Any],
    sections: Any,
    warnings: Sequence[str],
) -> Dict[str, str]:
    """
    将抽取结果压成单一 dict[str,str]，键全集见 ALL_FLAT_KEYS。
    """
    flat = empty_flat_document()

    flat["file_path"] = str(meta.get("path") or "").strip()
    flat["file_relative_path"] = str(meta.get("relative_path") or "").strip()
    flat["file_name"] = str(meta.get("filename") or "").strip()
    flat["file_format"] = str(meta.get("format") or "").strip()
    sp = meta.get("stem_parsed")
    if isinstance(sp, dict):
        flat["stem_doc_name"] = str(sp.get("document_name") or "").strip()
        flat["stem_date"] = str(sp.get("date_suffix") or "").strip()

    form = form or {}
    pb = form.get("project_basic")
    if isinstance(pb, dict):
        _fill_basic_from_kv(pb, flat)

    _fill_submission(form.get("submission_checklist") or [], flat)
    _fill_norm_like(form.get("norm_confirmation") or [], _NORM_HINTS, "norm", flat)
    _fill_norm_like(form.get("security_confirmation") or [], _SEC_HINTS, "sec", flat)
    _fill_responsible(form.get("responsible_persons") or [], flat)
    _fill_duty(form.get("duty_description") or [], flat)
    _fill_method(form.get("confirmation_method") or [], flat)
    _fill_sign(form.get("sign_off") or [], flat)
    _fill_commitment(list(form.get("commitment_paragraphs") or []), flat)

    # 分区丢失时：从全部 2 列表格再尝试灌 basic
    merged_kv = _merge_kv(_gather_kv_tables_from_sections(sections))
    _fill_basic_from_kv(merged_kv, flat)

    msgs = [m for m in warnings if m]
    flat["parse_warnings"] = "; ".join(msgs)

    # 保证键全集
    for k in ALL_FLAT_KEYS:
        flat.setdefault(k, "")

    out: Dict[str, str] = {}
    for k in ALL_FLAT_KEYS:
        out[k] = str(flat.get(k) or "")
    return out


def ordered_flat_for_json(flat: Mapping[str, str]) -> Dict[str, str]:
    """输出时固定键顺序。"""
    return {k: str(flat.get(k) or "") for k in ALL_FLAT_KEYS}
