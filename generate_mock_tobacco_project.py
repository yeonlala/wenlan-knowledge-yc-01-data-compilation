"""
在仓库 **mockdata/** 下生成模拟项目资料目录（仅三种场景：完整、仅空目录、部分缺失）。

项目根下创建 **01_管理确认**、**02_核心资料**（不创建「03_安全与整改」）。
标准叶子名 **01_～07_** 与 `1_check_tobacco_kb_required_files.py` 中 `STANDARD_DIRS` 一致；**`missing_required_partial`** 时仅在 **02** 下 **随机建其中若干叶目录**（不先建齐 7 个）。

**scenario=full** 时在 **01_管理确认** 下生成与「模板示例/确认单生成器」同结构的确认单 **.md / .json**。

未传入 --project-name 时，四段式目录名「客户单位_项目简称_四位年份_待验收」，与包内七段文件名客户/项目段对齐。
"""

from __future__ import annotations

import argparse
import importlib.util
import random
import shutil
import sys
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Set, Tuple

_REPO_ROOT = Path(__file__).resolve().parent
_spec = importlib.util.spec_from_file_location(
    "check_tobacco_kb_required_files",
    _REPO_ROOT / "1_check_tobacco_kb_required_files.py",
)
assert _spec is not None and _spec.loader is not None
_check_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_check_mod)

CORE_MATERIALS_DIR = _check_mod.CORE_MATERIALS_DIR
PROJECT_FOLDER_NAME_SPEC = _check_mod.PROJECT_FOLDER_NAME_SPEC
PROJECT_TOP_LEVEL_DIRS = _check_mod.PROJECT_TOP_LEVEL_DIRS
STANDARD_DIRS = _check_mod.STANDARD_DIRS
check_project_folder_name = _check_mod.check_project_folder_name
from tobacco_kb.naming_convention import VALID_EXTENSIONS, DEFAULT_DATE_SEGMENT, DEFAULT_SECRET_SEGMENT
from tobacco_kb.mock_placeholders import (
    bytes_for_mock_file,
    find_cjk_font_path,
)
from tobacco_kb.acceptance_config import PHASE1_RULES

ALL_RULE_IDS: Tuple[str, ...] = tuple(str(r["id"]) for r in PHASE1_RULES)

DEFAULT_MOCK_BATCH = _REPO_ROOT / "mockdata"


def _resolve_batch_dir(raw: Optional[Path]) -> Path:
    """默认仓库根下 mockdata；相对路径相对仓库根解析。"""
    if raw is None:
        return DEFAULT_MOCK_BATCH.resolve()
    p = raw.expanduser()
    if not p.is_absolute():
        return (_REPO_ROOT / p).resolve()
    return p.resolve()


def _year_from_folder_name(folder_name: str) -> str:
    parts = folder_name.split("_")
    if len(parts) >= 2 and len(parts[-2]) == 4 and parts[-2].isdigit():
        return parts[-2]
    return str(datetime.now().year)


def ensure_project_buckets(project_root: Path) -> None:
    project_root.mkdir(parents=True, exist_ok=True)
    for name in PROJECT_TOP_LEVEL_DIRS:
        (project_root / name).mkdir(parents=True, exist_ok=True)


def under_core(project_root: Path, leaf_dir: str) -> Path:
    return project_root / CORE_MATERIALS_DIR / leaf_dir


def normalize_ext(raw: str) -> str:
    raw = raw.strip()
    if not raw.startswith("."):
        raw = "." + raw
    return raw.lower()


def write_mock_file(
    path: Path,
    rule: Dict[str, Any],
    customer: str,
    project: str,
) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    data = bytes_for_mock_file(path.suffix, rule, customer, project)
    path.write_bytes(data)


def parse_csv_dirs(s: Optional[str]) -> Set[str]:
    if not s or not s.strip():
        return set()
    return {x.strip() for x in s.split(",") if x.strip()}


def scenario_dirs_only(project_root: Path, dirs: Sequence[str]) -> None:
    ensure_project_buckets(project_root)
    for d in dirs:
        under_core(project_root, d).mkdir(parents=True, exist_ok=True)


def write_required_files(
    project_root: Path,
    dirs_present: Set[str],
    rule_ids: Iterable[str],
    customer: str,
    project: str,
    file_ext: str,
) -> None:
    for rid in rule_ids:
        rid = rid.strip().lower()
        rule = next((r for r in PHASE1_RULES if str(r["id"]) == rid), None)
        if not rule:
            continue
        std_dir = str(rule["mock_std_dir"])
        kw = str(rule["mock_embed_keyword"])
        if std_dir not in dirs_present:
            continue
        name = f"{kw}_{DEFAULT_DATE_SEGMENT}_{DEFAULT_SECRET_SEGMENT}{file_ext}"
        out_file = under_core(project_root, std_dir) / name
        write_mock_file(out_file, rule, customer, project)


def scenario_full(
    project_root: Path,
    customer: str,
    project: str,
    file_ext: str,
    *,
    package_folder_name: str,
    year_str: str,
) -> None:
    """7 个标准叶子目录齐全，第一阶段规则材料均有占位文件；并写入确认单 MD/JSON。"""
    scenario_dirs_only(project_root, list(STANDARD_DIRS))
    write_required_files(
        project_root,
        set(STANDARD_DIRS),
        ALL_RULE_IDS,
        customer,
        project,
        file_ext,
    )
    submit = datetime.now().strftime("%Y-%m-%d")
    from tobacco_kb.mock_confirmation_form import (
        build_confirmation_payload_full,
        write_confirmation_pair,
    )

    data = build_confirmation_payload_full(
        branch_name=customer,
        customer_name=customer,
        project_name=project,
        project_year=year_str,
        project_status="待验收",
        package_name=package_folder_name,
        submit_date=submit,
    )
    admin = project_root / PROJECT_TOP_LEVEL_DIRS[0]
    md_p, js_p = write_confirmation_pair(admin, data, submit_date=submit)
    print(f"  已写入确认单（与生成器模板一致）：{md_p.name} · {js_p.name}")


def scenario_missing_required_partial(
    project_root: Path,
    customer: str,
    project: str,
    file_ext: str,
    *,
    seed: Optional[int],
) -> None:
    """
    「02_核心资料」下 **随机建 1～n 个标准叶目录**（不预先建齐全部），目录 **是哪几片也随机**；
    再在 **已建的叶子** 中随机抽若干写入占位文件（已建叶子里可有目录无文件）。
    模拟缺目录 + 缺材料的杂乱资料包（区别于 full / dirs_only）。
    """
    rng = random.Random(seed if seed is not None else random.randrange(1 << 30))
    n_total = len(STANDARD_DIRS)
    n_leaf = rng.randint(1, n_total)
    created_leaf_dirs = rng.sample(list(STANDARD_DIRS), k=n_leaf)
    scenario_dirs_only(project_root, created_leaf_dirs)

    n_fill = rng.randint(1, n_leaf)
    dirs_to_fill = rng.sample(created_leaf_dirs, k=n_fill)

    dir_to_rule_id = {str(r["mock_std_dir"]): str(r["id"]) for r in PHASE1_RULES}
    ids_write = [dir_to_rule_id[d] for d in dirs_to_fill if d in dir_to_rule_id]
    write_required_files(
        project_root,
        set(created_leaf_dirs),
        ids_write,
        customer,
        project,
        file_ext,
    )
    empty_created = n_leaf - n_fill
    not_created = n_total - n_leaf
    print(
        f"  missing_required_partial：「02_核心资料」下随机建了 {n_leaf}/{n_total} 个标准叶目录；"
        f"其中 {len(ids_write)} 个写有占位材料（已建叶子里 {empty_created} 个尚无占位文件；"
        f"另有 {not_created} 个标准叶目录本次未建）；相同 --seed 可复现。"
    )


def clean_project(project_root: Path) -> None:
    if project_root.exists():
        shutil.rmtree(project_root)


def _configure_stdio_utf8() -> None:
    if sys.platform != "win32":
        return
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")
        except Exception:
            pass


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = argparse.ArgumentParser(
        description="在 mockdata/ 下生成模拟项目：full | dirs_only | missing_required_partial",
    )
    parser.add_argument(
        "--batch",
        type=Path,
        default=None,
        help="模拟资料根目录，默认：仓库根目录/mockdata（可用 --batch 指定其它路径）",
    )
    parser.add_argument(
        "--project-name",
        default=None,
        metavar="NAME",
        help=(
            "mockdata（或 --batch）下项目文件夹名（客户单位_项目名称_年份_项目状态）。"
            "省略则自动生成四段式名称；反复覆盖请显式命名并加 --clean。"
        ),
    )
    parser.add_argument(
        "--scenario",
        choices=("full", "dirs_only", "missing_required_partial"),
        default="full",
        help=(
            "full=目录与材料齐全；dirs_only=02 下齐建 01～07 空目录；"
            "missing_required_partial=02 下随机建若干叶目录并随机写入占位材料"
        ),
    )
    parser.add_argument(
        "--omit-dirs",
        default="",
        help="仅 dirs_only：逗号分隔，不创建的叶子目录名（如 06_验收交付,07_项目复盘）",
    )
    parser.add_argument(
        "--customer",
        default="测试市局",
        help="七段文件名中的客户段",
    )
    parser.add_argument(
        "--project",
        default="知识库验收Mock",
        help="七段文件名中的项目段（勿含下划线）",
    )
    parser.add_argument(
        "--ext",
        default=".docx",
        metavar="EXT",
        help="占位文件扩展名，如 .docx、.pdf、.xlsx",
    )
    parser.add_argument(
        "--seed",
        type=int,
        default=None,
        help="missing_required_partial 随机种子（省略则每次不同）",
    )
    parser.add_argument(
        "--clean",
        action="store_true",
        help="若项目目录已存在则先删除再生成",
    )

    args = parser.parse_args(argv)

    _configure_stdio_utf8()

    file_ext = normalize_ext(args.ext)
    if file_ext not in VALID_EXTENSIONS:
        print(f"扩展名不在允许范围：{file_ext}")
        print(f"允许：{', '.join(sorted(VALID_EXTENSIONS))}")
        return 1

    if file_ext == ".pdf":
        try:
            import reportlab  # noqa: F401
        except ImportError:
            print(
                "错误：生成中文 PDF 需要先安装依赖：pip install reportlab\n"
                "也可改用：--ext .docx\n"
                "或：pip install -r requirements.txt"
            )
            return 1
        if find_cjk_font_path() is None:
            print(
                "错误：当前系统未找到中文字体文件，无法用 PDF 正确显示中文。\n"
                "请改用：--ext .docx"
            )
            return 1

    batch = _resolve_batch_dir(args.batch)
    batch.mkdir(parents=True, exist_ok=True)

    if args.project_name is None:
        cust = (args.customer or "").strip() or "测试市局"
        proj_base = (args.project or "").strip() or "知识库验收Mock"
        year_seg = str(datetime.now().year)
        project_dir_name = ""
        chosen_proj_part = ""
        for i in range(10000):
            proj_part = proj_base if i == 0 else f"{proj_base}{i}"
            candidate = f"{cust}_{proj_part}_{year_seg}_待验收"
            folder_ok, folder_issue = check_project_folder_name(Path(candidate))
            if not folder_ok:
                print(f"错误：自动生成的项目文件夹名不符合验收规则：{folder_issue}")
                print(
                    f"请调整 --customer / --project，"
                    f"或使用 --project-name 显式指定「{PROJECT_FOLDER_NAME_SPEC}」。"
                )
                return 1
            if not (batch / candidate).exists():
                project_dir_name = candidate
                chosen_proj_part = proj_part
                if i > 0:
                    print(
                        f"提示：批次目录下已有同名项目文件夹，已自动使用项目段「{proj_part}」。"
                    )
                break
        else:
            print(
                "错误：无法在批次目录下分配不重名的四段式文件夹名；"
                "请删除旧目录或使用 --project-name。"
            )
            return 1
        args.customer = cust
        args.project = chosen_proj_part
    else:
        project_dir_name = args.project_name
        folder_ok, folder_issue = check_project_folder_name(Path(project_dir_name))
        if not folder_ok:
            print(f"错误：--project-name 不符合验收命名规则：{folder_issue}")
            print(
                f"示例：测试市局_知识库试点_2026_待验收（须「{PROJECT_FOLDER_NAME_SPEC}」）。"
            )
            return 1

    year_str = _year_from_folder_name(project_dir_name)

    project_root = (batch / project_dir_name).resolve()
    if args.clean and project_root.exists():
        clean_project(project_root)

    omit_dirs = parse_csv_dirs(args.omit_dirs)

    if args.scenario == "full":
        scenario_full(
            project_root,
            args.customer,
            args.project,
            file_ext,
            package_folder_name=project_dir_name,
            year_str=year_str,
        )
    elif args.scenario == "dirs_only":
        pool = [d for d in STANDARD_DIRS if d not in omit_dirs]
        scenario_dirs_only(project_root, pool)
    elif args.scenario == "missing_required_partial":
        scenario_missing_required_partial(
            project_root,
            args.customer,
            args.project,
            file_ext,
            seed=args.seed,
        )
    else:
        parser.error("未知 scenario")

    print(f"已生成：{project_root}")
    if args.scenario == "dirs_only":
        n_leaf = len([d for d in STANDARD_DIRS if d not in omit_dirs])
        print(
            f"  dirs_only：「02_核心资料」下 {n_leaf} 个空目录（01_～07_；可用 --omit-dirs 排除）。"
        )
    if args.scenario == "missing_required_partial":
        if args.seed is not None:
            print(f"  随机种子：{args.seed}")
    if args.project_name is None:
        print("（本轮使用自动目录名）")
    print("可用：python 1_check_tobacco_kb_required_files.py")
    return 0


if __name__ == "__main__":
    sys.exit(main())
