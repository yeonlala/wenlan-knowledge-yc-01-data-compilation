# -*- coding: utf-8 -*-
"""
从《项目资料上报确认单》结构化 JSON 拆分为 **3 份**入库用清单 JSON：

  1. 项目基础信息清单
  2. 材料提交清单（含 02_核心资料 七类目录文件盘点）
  3. 责任人确认清单

「安全脱敏确认清单」「AI 入库候选清单」不在此脚本生成，须后续按实际文件单独处理。

默认输出目录：
  extract_jsons/<项目文件夹名>/*.json  
  同级目录 **docs_metadata/**：自资料包 **02_核心资料** 扁平归集的副本（与清单 JSON 并列）。

归集文件夹名为 ``docs_metadata``：**无子文件夹**，复制各叶目录下**任意扩展名**的文件（仅排除
``__pycache__``、隐藏文件等无关项）；文件名用 ``<标准子目录>__…`` 扁平化避免重名。

项目文件夹名**动态取自路径**：确认单位于 ``…/<项目文件夹名>/01_管理确认/xxx.json`` 时，
子目录即为 ``<项目文件夹名>``（例如 ``福州市局_民主管理网升级改造_2025_已验收``）。
也可用 --folder-name 覆盖（自动去掉误写的【】）。

用法：
  python extract_confirmation_checklists.py \\
    "第一批/福州市局_民主管理网升级改造_2025_已验收/01_管理确认/项目资料上报确认单_20260508.json"

  python extract_confirmation_checklists.py --folder-name "福州市局_民主管理网升级改造_2025_已验收" \\
    "path/to/确认单.json"

  # 省略路径时默认：自动扫描仓库下「第一批」各项目的 01_管理确认 内确认单 .json
  python extract_confirmation_checklists.py
"""

from __future__ import annotations

import argparse
import json
import shutil
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, Iterator, List, Optional, Tuple

_REPO_ROOT = Path(__file__).resolve().parent
if str(_REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(_REPO_ROOT))

from tobacco_kb.acceptance_config import PHASE1_RULES

FIRST_BATCH = "第一批"
CORE_MATERIALS_DIR = "02_核心资料"
# 与 extract_jsons/<项目>/ 下各 *.json 同级：02_核心资料 文件扁平归集目录名
EXTRACT_CORE_BUNDLE_DIRNAME = "docs_metadata"

# checklist 内英文键 → 材料类别中文说明（与确认单 schema 一致时可对照补充）
CHECKLIST_ITEM_LABELS: Dict[str, str] = {
    "project_basic_info": "项目基本信息表",
    "solution_plan": "建设方案/总体设计与技术架构",
    "requirement_or_function_list": "需求规格说明书或功能清单",
    "business_flow": "业务流程与流程图",
    "manual_or_training": "用户手册、培训资料或操作指南",
    "acceptance_or_delivery": "验收报告或最终交付清单",
    "knowledge_summary_or_exemption": "知识沉淀总结或免入库说明",
}

# project 字段 → 中文列名
PROJECT_FIELD_LABELS: Dict[str, str] = {
    "branch_name": "分公司名称",
    "customer_name": "客户单位",
    "project_name": "项目名称",
    "project_year": "项目年份",
    "project_status": "项目状态",
    "business_domain": "业务板块",
    "package_name": "资料包名称",
    "submit_date": "提交日期",
    "is_sensitive": "是否涉敏",
    "ai_candidate_status": "是否作为AI入库候选",
}

RESPONSIBILITY_FIELD_LABELS: Dict[str, str] = {
    "material_owner": "资料整理责任人",
    "material_owner_role": "资料整理人角色",
    "project_owner": "项目负责人",
    "project_owner_role": "项目负责人角色",
    "branch_owner": "分公司/部门负责人",
    "branch_owner_role": "分公司负责人角色",
    "confirmation_method": "确认方式",
}


def resolve_project_root_from_confirmation(json_path: Path) -> Optional[Path]:
    """确认单位于 ``<项目根>/01_管理确认/*.json`` 时返回项目根目录。"""
    p = json_path.resolve()
    if p.parent.name == "01_管理确认":
        return p.parent.parent
    return None


def safe_relative_project_root(project_root: Path) -> str:
    """
    相对仓库根的路径（POSIX），通常形如 ``第一批/客户_项目_年_状态``；
    不入库到仓库内时退化为仅文件夹名，避免写入本机绝对路径。
    """
    try:
        return project_root.resolve().relative_to(_REPO_ROOT.resolve()).as_posix()
    except ValueError:
        return project_root.name


def _flatten_bundle_filename(leaf_std_dir: str, file_path: Path, leaf_root: Path) -> str:
    """叶目录内相对路径扁平化为单一文件名（不含子目录）。"""
    rel = file_path.relative_to(leaf_root).as_posix()
    flat = rel.replace("/", "__").replace("\\", "__")
    return f"{leaf_std_dir}__{flat}"


def iter_phase1_leaf_files(
    core_dir: Path,
) -> Iterator[Tuple[str, Path, Path]]:
    """
    遍历各标准子目录下**所有层级**的文件，**扩展名不限**（含无扩展名、压缩包、图片等），
    仅跳过 ``_skip_inventory_path`` 中的无关项；产出 ``(standard_subdir, leaf_root, file_path)``。
    """
    for rule in PHASE1_RULES:
        leaf = str(rule["mock_std_dir"])
        leaf_path = core_dir / leaf
        if not leaf_path.is_dir():
            continue
        for f in sorted(leaf_path.rglob("*")):
            if not f.is_file() or _skip_inventory_path(f):
                continue
            yield leaf, leaf_path, f


def copy_docs_metadata_bundle(project_root: Path, extract_dir: Path) -> bool:
    """
    将 ``<项目根>/02_核心资料`` 下各标准目录中的**全部文件**复制到 ``extract_dir/docs_metadata/``，
    **不创建子文件夹**；文件名经 ``_flatten_bundle_filename`` 扁平化；与同目录清单 JSON 并列。
    """
    src = project_root / CORE_MATERIALS_DIR
    dst = extract_dir / EXTRACT_CORE_BUNDLE_DIRNAME
    if not src.is_dir():
        return False
    if dst.exists():
        shutil.rmtree(dst)
    dst.mkdir(parents=True)

    used: set[str] = set()
    for leaf, leaf_root, f in iter_phase1_leaf_files(src):
        base_name = _flatten_bundle_filename(leaf, f, leaf_root)
        out_name = base_name
        n = 0
        while out_name in used:
            n += 1
            stem = Path(base_name).stem
            suf = Path(base_name).suffix
            out_name = f"{stem}_dup{n}{suf}"
        used.add(out_name)
        shutil.copy2(f, dst / out_name)
    return True


def _skip_inventory_path(path: Path) -> bool:
    if "__pycache__" in path.parts:
        return True
    name = path.name
    if name.startswith("."):
        return True
    if name in ("Thumbs.db", "desktop.ini"):
        return True
    return False


def scan_phase1_core_inventory(project_root: Path) -> Dict[str, Any]:
    """
    扫描 ``02_核心资料`` 下第一阶段七类标准目录（与 PHASE1_RULES 一致），
    列出 **任意扩展名** 的文件（与 ``docs_metadata`` 归集规则一致）；条目含扁平文件名（无子目录输出）。
    """
    core = project_root / CORE_MATERIALS_DIR
    proj_ref = safe_relative_project_root(project_root)
    if not core.is_dir():
        return {
            "resolved_project_root": proj_ref,
            "core_materials_dir": CORE_MATERIALS_DIR,
            "core_materials_exists": False,
            "total_document_count": 0,
            "hint": (
                f"项目根下不存在「{CORE_MATERIALS_DIR}」目录，无法列出实际文档。"
            ),
            "phase1_directories": [],
        }

    by_leaf: Dict[str, List[Dict[str, str]]] = {str(r["mock_std_dir"]): [] for r in PHASE1_RULES}
    for leaf, leaf_root, f in iter_phase1_leaf_files(core):
        try:
            rel_proj = f.relative_to(project_root).as_posix()
        except ValueError:
            rel_proj = str(f)
        flat_name = _flatten_bundle_filename(leaf, f, leaf_root)
        ext = f.suffix.lower()
        by_leaf[leaf].append(
            {
                "filename": f.name,
                "file_extension": ext if ext else "",
                "flattened_bundle_filename": flat_name,
                "source_relative_path": rel_proj,
            }
        )

    phase_rows: List[Dict[str, Any]] = []
    for rule in PHASE1_RULES:
        leaf = str(rule["mock_std_dir"])
        leaf_path = core / leaf
        docs = by_leaf.get(leaf, [])
        phase_rows.append(
            {
                "rule_id": str(rule["id"]),
                "rule_label_zh": str(rule["label"]),
                "standard_subdir": leaf,
                "directory_exists": leaf_path.is_dir(),
                "rule_optional": bool(rule.get("optional")),
                "document_count": len(docs),
                "documents": docs,
            }
        )

    total_document_count = sum(int(row["document_count"]) for row in phase_rows)

    return {
        "resolved_project_root": proj_ref,
        "core_materials_dir": CORE_MATERIALS_DIR,
        "core_materials_exists": True,
        "total_document_count": total_document_count,
        "hint": "",
        "inventory_note_zh": (
            "列出各叶目录下全部文件（扩展名不限）；与 extract 目录 docs_metadata 扁平归集一致；"
            "排除 __pycache__、以「.」开头的隐藏文件及 Thumbs.db 等。"
        ),
        "phase1_directories": phase_rows,
    }


def _configure_stdio_utf8() -> None:
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8")
        except Exception:
            pass


def _strip_bookends(name: str) -> str:
    return name.strip().strip("【】").strip()


def resolve_project_folder_name(input_path: Path, override: Optional[str]) -> str:
    """
    输出子目录名 = 资料包项目根文件夹名（四段式命名的那一层），不含【】。

    解析规则：路径中出现 ``01_管理确认`` 时，取其**上一级**目录名；
    否则若确认单直接在 ``01_管理确认`` 下，仍取该文件夹的上一级。
    """
    if override:
        return _strip_bookends(override) or "unknown_project"
    parts = input_path.resolve().parts
    for i, p in enumerate(parts):
        if p == "01_管理确认" and i > 0:
            return parts[i - 1]
    parent = input_path.parent
    if parent.name == "01_管理确认":
        return parent.parent.name
    return parent.name


def _now_iso() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")


def load_confirmation_json(path: Path) -> Dict[str, Any]:
    data = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(data, dict):
        raise ValueError(f"根节点须为 JSON 对象：{path}")
    return data


def build_project_basic(doc: Dict[str, Any], source_file: str) -> Dict[str, Any]:
    proj = doc.get("project") if isinstance(doc.get("project"), dict) else {}
    return {
        "list_type": "project_basic_info",
        "schema_version": doc.get("schema_version", ""),
        "source_document_type": doc.get("document_type", ""),
        "fields": proj,
        "remarks": {
            "title_zh": "项目基础信息清单",
            "column_labels_zh": PROJECT_FIELD_LABELS,
        },
        "meta": {
            "source_file": source_file,
            "extracted_at": _now_iso(),
            "generated_at": doc.get("generated_at", ""),
        },
    }


def build_material_submission(
    doc: Dict[str, Any],
    source_file: str,
    confirmation_json_path: Optional[Path] = None,
) -> Dict[str, Any]:
    raw = doc.get("checklist") if isinstance(doc.get("checklist"), dict) else {}
    rows: List[Dict[str, Any]] = []
    # 按模板顺序；其余键按字典序接在后面
    ordered_keys = list(CHECKLIST_ITEM_LABELS.keys()) + sorted(
        k for k in raw.keys() if k not in CHECKLIST_ITEM_LABELS
    )
    for key in ordered_keys:
        if key not in raw:
            continue
        cell = raw[key]
        if not isinstance(cell, dict):
            continue
        rows.append(
            {
                "item_key": key,
                "category_label_zh": CHECKLIST_ITEM_LABELS.get(key, key),
                "submitted": cell.get("submitted"),
                "note": cell.get("note", ""),
            }
        )

    core_inventory: Dict[str, Any]
    if confirmation_json_path is not None:
        proj_root = resolve_project_root_from_confirmation(confirmation_json_path)
        if proj_root is not None:
            core_inventory = scan_phase1_core_inventory(proj_root)
        else:
            core_inventory = {
                "resolved_project_root": None,
                "core_materials_dir": CORE_MATERIALS_DIR,
                "core_materials_exists": False,
                "total_document_count": 0,
                "hint": (
                    "确认单不在「01_管理确认」下，无法关联扫描「02_核心资料」七类目录。"
                ),
                "phase1_directories": [],
            }
    else:
        core_inventory = {
            "resolved_project_root": None,
            "total_document_count": 0,
            "hint": "未提供确认单路径，跳过磁盘扫描。",
            "phase1_directories": [],
        }

    return {
        "list_type": "material_submission",
        "schema_version": doc.get("schema_version", ""),
        "items": rows,
        "core_phase1_inventory": core_inventory,
        "remarks": {
            "title_zh": "材料提交清单",
            "core_inventory_scope_zh": (
                "对应 PHASE1_RULES：02_核心资料 下七个标准子目录内的文件列表（扩展名不限，与 docs_metadata 归集一致）"
            ),
        },
        "meta": {
            "source_file": source_file,
            "extracted_at": _now_iso(),
        },
    }


def build_responsibility(doc: Dict[str, Any], source_file: str) -> Dict[str, Any]:
    resp = doc.get("responsibility") if isinstance(doc.get("responsibility"), dict) else {}
    return {
        "list_type": "responsibility",
        "schema_version": doc.get("schema_version", ""),
        "fields": resp,
        "remarks": {
            "title_zh": "责任人确认清单",
            "column_labels_zh": RESPONSIBILITY_FIELD_LABELS,
        },
        "meta": {
            "source_file": source_file,
            "extracted_at": _now_iso(),
        },
    }


OUTPUT_FILENAMES = (
    "项目基础信息清单.json",
    "材料提交清单.json",
    "责任人确认清单.json",
)


def write_extract_files(out_dir: Path, payloads: Tuple[Dict[str, Any], ...], indent: int) -> List[Path]:
    out_dir.mkdir(parents=True, exist_ok=True)
    written: List[Path] = []
    for name, payload in zip(OUTPUT_FILENAMES, payloads):
        p = out_dir / name
        p.write_text(json.dumps(payload, ensure_ascii=False, indent=indent), encoding="utf-8")
        written.append(p)
    return written


def process_one_file(
    json_path: Path,
    output_root: Path,
    folder_name_override: Optional[str],
    indent: int,
) -> Tuple[str, List[str]]:
    doc = load_confirmation_json(json_path)
    folder = resolve_project_folder_name(json_path, folder_name_override)
    out_dir = output_root / folder
    try:
        rel_source = str(json_path.resolve().relative_to(_REPO_ROOT))
    except ValueError:
        rel_source = str(json_path.resolve())

    payloads = (
        build_project_basic(doc, rel_source),
        build_material_submission(doc, rel_source, json_path),
        build_responsibility(doc, rel_source),
    )
    paths = write_extract_files(out_dir, payloads, indent)
    written = [str(x.resolve()) for x in paths]
    proj_root = resolve_project_root_from_confirmation(json_path)
    if proj_root is not None and copy_docs_metadata_bundle(proj_root, out_dir):
        written.append(str((out_dir / EXTRACT_CORE_BUNDLE_DIRNAME).resolve()))
    return folder, written


def iter_batch_confirmation_jsons(batch_root: Path) -> List[Path]:
    """第一批/<项目>/01_管理确认/*确认单*.json"""
    found: List[Path] = []
    if not batch_root.is_dir():
        return found
    for project_dir in sorted(batch_root.iterdir()):
        if not project_dir.is_dir():
            continue
        bucket = project_dir / "01_管理确认"
        if not bucket.is_dir():
            continue
        for f in sorted(bucket.iterdir()):
            if not f.is_file() or f.suffix.lower() != ".json":
                continue
            if "确认" in f.name:
                found.append(f)
    return found


def resolve_output_root(s: str) -> Path:
    p = Path(s)
    return p.resolve() if p.is_absolute() else (_REPO_ROOT / p).resolve()


def run_batch_scan(batch_root: Path, output_root: Path, indent: int) -> int:
    """扫描 batch_root 下一层项目目录中的 01_管理确认/*.json（文件名含「确认」）。"""
    files = iter_batch_confirmation_jsons(batch_root)
    if not files:
        print(f"未在 {batch_root} 下找到确认单 JSON（期望 …/<项目>/01_管理确认/*确认*.json）。", file=sys.stderr)
        return 1
    for fp in files:
        folder, written = process_one_file(fp, output_root, None, indent)
        print(f"{folder} ← {fp.name}", file=sys.stderr)
        for w in written:
            print(f"  {w}", file=sys.stderr)
    print(f"共处理 {len(files)} 个确认单。", file=sys.stderr)
    return 0


def main(argv: Optional[List[str]] = None) -> int:
    _configure_stdio_utf8()
    parser = argparse.ArgumentParser(description="确认单 JSON → 3 份分类清单 JSON（不含安全脱敏 / AI 入库候选）")
    parser.add_argument(
        "input_json",
        nargs="?",
        default=None,
        help="可选：《项目资料上报确认单》.json；省略则自动扫描仓库「第一批」下各项目",
    )
    parser.add_argument(
        "--output-root",
        default="extract_jsons",
        help="输出根目录（默认 extract_jsons，相对仓库根）",
    )
    parser.add_argument(
        "--folder-name",
        default=None,
        help="覆盖输出子目录名（与资料包项目文件夹同名，如 福州市局_民主管理网升级改造_2025_已验收）",
    )
    parser.add_argument("--indent", type=int, default=2, help="JSON 缩进，默认 2")
    parser.add_argument(
        "--batch",
        action="store_true",
        help=f"显式仅批量模式（与省略 input_json 相同：扫描「{FIRST_BATCH}」）",
    )
    args = parser.parse_args(argv)

    output_root = resolve_output_root(args.output_root)
    batch_path = _REPO_ROOT / FIRST_BATCH

    # 未指定文件 → 默认扫描「第一批」下各项目
    if not args.input_json:
        if not batch_path.is_dir():
            print(
                f"未指定确认单文件，且仓库下不存在「{FIRST_BATCH}」目录：{batch_path}",
                file=sys.stderr,
            )
            return 1
        return run_batch_scan(batch_path, output_root, args.indent)

    if args.batch:
        print("已忽略 --batch（已指定单个确认单文件）。", file=sys.stderr)

    path = Path(args.input_json)
    if not path.is_file():
        print(f"文件不存在：{path}", file=sys.stderr)
        return 1

    folder, written = process_one_file(path, output_root, args.folder_name, args.indent)
    print(f"已写入目录：{output_root / folder}", file=sys.stderr)
    for w in written:
        print(f"  {w}", file=sys.stderr)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
