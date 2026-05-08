"""
在「第一批」下生成用于测试的模拟项目资料目录。

项目根下创建 **01_管理确认**、**02_核心资料**（**不创建「03_安全与整改」**，与验收脚本跳过口径一致）；原标准目录 **01_～11_**（叶子名）均在 **02_核心资料** 下创建。

默认 **--phase 1**：使用各规则 `mock_std_dir` 并集（**不含** `acceptance_config` 中标记 **optional** 的叶子；当前仅 **r07/07_项目复盘** 可选，故默认可建 **6** 个必验叶子 + **r07** 仍可按规则生成文件）。

**--phase 2** 时，可对**全套 11 个**叶子目录名在 **02_核心资料** 下操作。

通过 --scenario / --omit-dirs / --omit-required 等可生成残缺样例。生成文件为可读占位，支持 .docx / .pdf / .xlsx 等（见 --ext）。

未传入 --project-name 时，每轮在「第一批」下新建带日期时间的独立示例目录，便于多项目对比；需覆盖同一路径时显式写全名并加 --clean。
"""

from __future__ import annotations

import argparse
import random
import shutil
import sys
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Set, Tuple

from check_tobacco_kb_required_files import (
    CORE_MATERIALS_DIR,
    PROJECT_FOLDER_NAME_SPEC,
    PROJECT_TOP_LEVEL_DIRS,
    STANDARD_DIRS,
    check_project_folder_name,
    normalize_name,
)
from tobacco_kb.naming_convention import (
    VALID_EXTENSIONS,
    build_standard_filename_stem,
    DEFAULT_DATE_SEGMENT,
    DEFAULT_SECRET_SEGMENT,
)
from tobacco_kb.mock_placeholders import (
    bytes_for_mock_file,
    find_cjk_font_path,
    minimal_zip_bytes,
)
from tobacco_kb.acceptance_config import PHASE1_RULES

ALL_RULE_IDS: Tuple[str, ...] = tuple(str(r["id"]) for r in PHASE1_RULES)

# 仅放必验材料所需的最小标准子目录（不含 optional 规则；与验收「须存在目录」一致）
PHASE1_MINIMAL_DIRS: Tuple[str, ...] = tuple(
    sorted(
        {str(r["mock_std_dir"]) for r in PHASE1_RULES if not r.get("optional")},
        key=lambda d: STANDARD_DIRS.index(d),
    )
)


def ensure_project_buckets(project_root: Path) -> None:
    """项目根下创建 01_管理确认 / 02_核心资料（不创建 03_安全与整改）。"""
    project_root.mkdir(parents=True, exist_ok=True)
    for name in PROJECT_TOP_LEVEL_DIRS:
        (project_root / name).mkdir(parents=True, exist_ok=True)


def under_core(project_root: Path, leaf_dir: str) -> Path:
    """标准叶子目录物理路径：02_核心资料/<leaf>。"""
    return project_root / CORE_MATERIALS_DIR / leaf_dir


def _dir_category(std_dir: str) -> str:
    return std_dir.split("_", 1)[1]


def normalize_ext(raw: str) -> str:
    raw = raw.strip()
    if not raw.startswith("."):
        raw = "." + raw
    return raw.lower()


def good_filename(
    customer: str,
    project: str,
    std_dir: str,
    doc_label: str,
    ext: str = ".docx",
    date: str = DEFAULT_DATE_SEGMENT,
    secret: str = DEFAULT_SECRET_SEGMENT,
) -> str:
    """符合 tobacco_kb.naming_convention 的 7 段下划线命名（与验收脚本共用规则）。"""
    cat = _dir_category(std_dir)
    stem = build_standard_filename_stem(
        customer, project, cat, doc_label, date_segment=date, secret_segment=secret
    )
    return stem + ext


def write_mock_file(
    path: Path,
    rule: Dict[str, Any],
    customer: str,
    project: str,
) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    data = bytes_for_mock_file(path.suffix, rule, customer, project)
    path.write_bytes(data)


def write_simple_mock(
    path: Path,
    customer: str,
    project: str,
    display_label: str,
) -> None:
    """无 PHASE1 规则时的占位（如 bad_names 场景）。"""
    fake: Dict[str, Any] = {
        "id": "aux",
        "label": display_label,
        "note": "辅助测试文件",
        "keywords": (),
    }
    write_mock_file(path, fake, customer, project)


def parse_rule_ids(s: Optional[str]) -> Set[str]:
    if not s or not s.strip():
        return set()
    return {x.strip().lower() for x in s.split(",") if x.strip()}


def parse_csv_dirs(s: Optional[str]) -> Set[str]:
    if not s or not s.strip():
        return set()
    return {x.strip() for x in s.split(",") if x.strip()}


def remove_files_for_rule(project_root: Path, rid: str) -> None:
    rid = rid.strip().lower()
    rule = next((r for r in PHASE1_RULES if str(r["id"]) == rid), None)
    if not rule:
        return
    std_dir = str(rule["mock_std_dir"])
    folder = under_core(project_root, std_dir)
    if not folder.is_dir():
        return
    keywords: Tuple[str, ...] = tuple(rule["keywords"])
    for f in list(folder.iterdir()):
        if not f.is_file():
            continue
        nn = normalize_name(f.name)
        if any(kw in nn for kw in keywords):
            f.unlink()


def scenario_empty(project_root: Path) -> None:
    """仅空项目目录，无子目录、无文件。"""
    project_root.mkdir(parents=True, exist_ok=True)


def scenario_dirs_only(project_root: Path, dirs: Sequence[str]) -> None:
    """只有标准目录壳子，无文件（叶子目录建在 02_核心资料 下）。"""
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
) -> None:
    # full：必验叶子 + optional 叶子（如 07_项目复盘），七条材料样例齐全
    dirs = sorted(
        set(PHASE1_MINIMAL_DIRS)
        | {str(r["mock_std_dir"]) for r in PHASE1_RULES if r.get("optional")},
        key=lambda d: STANDARD_DIRS.index(d),
    )
    scenario_dirs_only(project_root, dirs)
    write_required_files(
        project_root, set(dirs), ALL_RULE_IDS, customer, project, file_ext
    )


def scenario_missing_dirs(
    project_root: Path,
    omit: Set[str],
    customer: str,
    project: str,
    file_ext: str,
    *,
    phase: int,
) -> None:
    pool = list(STANDARD_DIRS) if phase == 2 else list(PHASE1_MINIMAL_DIRS)
    omit_eff = set(omit)
    if not omit_eff:
        if phase == 2:
            omit_eff = {"08_技术设计", "09_开发测试", "10_运维售后"}
        elif pool:
            omit_eff = set(pool[-2:])
    dirs = [d for d in pool if d not in omit_eff]
    scenario_dirs_only(project_root, dirs)
    write_required_files(
        project_root, set(dirs), ALL_RULE_IDS, customer, project, file_ext
    )


def scenario_missing_required_partial(
    project_root: Path,
    omit_rule_ids: Set[str],
    customer: str,
    project: str,
    file_ext: str,
    *,
    phase: int,
) -> None:
    dirs = list(STANDARD_DIRS) if phase == 2 else list(PHASE1_MINIMAL_DIRS)
    scenario_dirs_only(project_root, dirs)
    ids = [rid for rid in ALL_RULE_IDS if rid not in omit_rule_ids]
    write_required_files(project_root, set(dirs), ids, customer, project, file_ext)


def scenario_bad_names(
    project_root: Path,
    customer: str,
    project: str,
    file_ext: str,
    *,
    phase: int,
) -> None:
    """目录齐全（相对当前 phase），混入不合格/禁用命名与风险词。"""
    dirs = list(STANDARD_DIRS) if phase == 2 else list(PHASE1_MINIMAL_DIRS)
    scenario_dirs_only(project_root, dirs)
    ext = file_ext
    d0 = dirs[0]
    write_simple_mock(under_core(project_root, d0) / f"资料{ext}", customer, project, "资料")
    write_simple_mock(
        under_core(project_root, d0) / f"最终版{ext}", customer, project, "最终版"
    )
    if phase == 2:
        risk_dir = "08_技术设计"
    else:
        risk_dir = "04_产品设计"
    write_simple_mock(
        under_core(project_root, risk_dir) / f"接口文档{ext}",
        customer,
        project,
        "接口文档",
    )
    good_path = (
        under_core(project_root, risk_dir)
        / f"接口文档整理_{DEFAULT_DATE_SEGMENT}_{DEFAULT_SECRET_SEGMENT}{ext}"
    )
    write_simple_mock(good_path, customer, project, "接口文档整理")


def scenario_random_sparse(
    project_root: Path,
    customer: str,
    project: str,
    seed: Optional[int],
    file_ext: str,
    *,
    phase: int,
) -> None:
    rng = random.Random(seed)
    # 第一阶段：只在 PHASE1_MINIMAL_DIRS 抽样。第二阶段：可在全套 STANDARD_DIRS 中抽样。
    pool = list(STANDARD_DIRS) if phase == 2 else list(PHASE1_MINIMAL_DIRS)
    if not pool:
        scenario_empty(project_root)
        return
    k_max = len(pool)
    k_min = min(3, k_max) if k_max >= 3 else 1
    k = rng.randint(k_min, k_max)
    dirs = sorted(rng.sample(pool, k=k))
    scenario_dirs_only(project_root, dirs)
    n_req = rng.randint(1, max(1, len(ALL_RULE_IDS) - 1))
    chosen = set(rng.sample(list(ALL_RULE_IDS), k=n_req))
    write_required_files(project_root, set(dirs), chosen, customer, project, file_ext)
    if rng.random() < 0.4:
        d = rng.choice(dirs)
        zp = under_core(project_root, d) / "备份.zip"
        zp.parent.mkdir(parents=True, exist_ok=True)
        zp.write_bytes(
            minimal_zip_bytes(
                f"模拟备份压缩包\n客户：{customer}\n项目：{project}\n"
                f"用途：random_sparse 场景随机生成的占位 zip。"
            )
        )


def clean_project(project_root: Path) -> None:
    if project_root.exists():
        shutil.rmtree(project_root)


def _configure_stdio_utf8() -> None:
    import sys

    if sys.platform != "win32":
        return
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")
        except Exception:
            pass


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = argparse.ArgumentParser(
        description="在「第一批」下生成模拟项目目录（第一阶段 7 条规则），可残缺，用于验收脚本测试。"
    )
    parser.add_argument(
        "--batch",
        type=Path,
        default=None,
        help="资料批次根目录，默认：当前工作目录/第一批",
    )
    parser.add_argument(
        "--project-name",
        default=None,
        metavar="NAME",
        help=(
            "「第一批」下项目文件夹名，须与验收一致："
            "客户单位_项目名称_年份_项目状态（至少 4 段下划线分段，倒数第二段为四位年份）。"
            "省略时每轮自动生成唯一名称；固定路径反复覆盖请加显式名称并配合 --clean。"
        ),
    )
    parser.add_argument(
        "--scenario",
        choices=[
            "full",
            "empty",
            "dirs_only",
            "missing_dirs",
            "missing_required_partial",
            "bad_names",
            "random_sparse",
        ],
        default="random_sparse",
        help="预设场景；可与 --omit-dirs / --omit-required 叠加",
    )
    parser.add_argument(
        "--omit-dirs",
        default="",
        help="逗号分隔，不创建的目录名（--phase 1 时仅作用于第一阶段目录池；二阶段示例：08_技术设计,09_开发测试）",
    )
    parser.add_argument(
        "--omit-required",
        default="",
        help="逗号分隔，不生成/事后删除的规则 id（与 tobacco_kb.acceptance_config.PHASE1_RULES 中 id 一致），"
        "如：r04,r07",
    )
    parser.add_argument(
        "--customer",
        default="测试市局",
        help="用于规范文件名的客户段",
    )
    parser.add_argument(
        "--project",
        default="知识库验收Mock",
        help="用于规范文件名的项目段（勿含下划线，否则破坏 7 段解析）",
    )
    parser.add_argument(
        "--ext",
        default=".docx",
        metavar="EXT",
        help="生成文件扩展名，如 .docx、.pdf、.xlsx（须为验收允许的类型）；ppt/pptx 建议安装 python-pptx",
    )
    parser.add_argument("--seed", type=int, default=None, help="random_sparse 用随机种子")
    parser.add_argument(
        "--phase",
        type=int,
        choices=(1, 2),
        default=None,
        help="1=仅第一阶段目录池（7 个标准叶子）；2=含 08～11 共 11 个标准叶子。"
        "省略时：dirs_only 默认按 2（建齐 11 个空目录，与资料包标准结构一致）；"
        "其余场景默认 1。",
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
                "也可改用默认 Word 占位（完整中文、无需额外库）：--ext .docx\n"
                "或一次安装全部：pip install -r requirements.txt"
            )
            return 1
        if find_cjk_font_path() is None:
            print(
                "错误：当前系统未找到中文字体文件（如微软雅黑 msyh.ttc），"
                "无法用 PDF 正确显示中文。\n"
                "请安装中文字体后重试，或改用：--ext .docx"
            )
            return 1

    batch = (args.batch or (Path.cwd() / "第一批")).resolve()
    if not batch.is_dir():
        print(f"批次目录不存在：{batch}")
        return 1

    if args.project_name is None:
        now = datetime.now()
        # 保持「…_四位年份_项目状态」为末两段，满足验收；中间用日期+时间做唯一性，避免再插入微秒段
        project_dir_name = (
            f"测试市局_知识库验收Mock_{now:%Y%m%d}_{now:%H%M%S}_2026_待验收"
        )
    else:
        project_dir_name = args.project_name
        folder_ok, folder_issue = check_project_folder_name(Path(project_dir_name))
        if not folder_ok:
            print(f"错误：--project-name 不符合验收命名规则：{folder_issue}")
            print(
                f"正确示例：测试市局_知识库试点_2026_待验收（须含「{PROJECT_FOLDER_NAME_SPEC}」四段结构，"
                "不可写成 测试市局_2026_待验收）。"
            )
            return 1

    project_root = (batch / project_dir_name).resolve()
    if args.clean and project_root.exists():
        clean_project(project_root)

    omit_dirs = parse_csv_dirs(args.omit_dirs)
    omit_req = parse_rule_ids(args.omit_required)

    if args.phase is None:
        phase_eff = 2 if args.scenario == "dirs_only" else 1
    else:
        phase_eff = args.phase

    if args.scenario == "full":
        scenario_full(project_root, args.customer, args.project, file_ext)
    elif args.scenario == "empty":
        scenario_empty(project_root)
    elif args.scenario == "dirs_only":
        pool = (
            list(STANDARD_DIRS) if phase_eff == 2 else list(PHASE1_MINIMAL_DIRS)
        )
        dirs = [d for d in pool if d not in omit_dirs]
        scenario_dirs_only(project_root, dirs)
    elif args.scenario == "missing_dirs":
        od = omit_dirs or set()
        scenario_missing_dirs(
            project_root,
            od,
            args.customer,
            args.project,
            file_ext,
            phase=phase_eff,
        )
    elif args.scenario == "missing_required_partial":
        oq = omit_req or {"r04", "r07"}
        scenario_missing_required_partial(
            project_root,
            oq,
            args.customer,
            args.project,
            file_ext,
            phase=phase_eff,
        )
    elif args.scenario == "bad_names":
        scenario_bad_names(
            project_root, args.customer, args.project, file_ext, phase=phase_eff
        )
    elif args.scenario == "random_sparse":
        scenario_random_sparse(
            project_root,
            args.customer,
            args.project,
            args.seed,
            file_ext,
            phase=phase_eff,
        )
    else:
        parser.error("未知 scenario")

    if omit_dirs and args.scenario not in ("missing_dirs", "dirs_only"):
        for d in omit_dirs:
            p = under_core(project_root, d)
            if p.is_dir():
                shutil.rmtree(p)
    if omit_req and args.scenario not in ("missing_required_partial",):
        for rid in omit_req:
            remove_files_for_rule(project_root, rid)

    print(f"已生成：{project_root}")
    if args.scenario == "dirs_only":
        pool_d = (
            list(STANDARD_DIRS) if phase_eff == 2 else list(PHASE1_MINIMAL_DIRS)
        )
        n_leaf = len([d for d in pool_d if d not in omit_dirs])
        ph_desc = "11 项（08～11 含）" if phase_eff == 2 else "7 项（仅阶段一）"
        print(
            f"  dirs_only：本次在「02_核心资料」下创建 {n_leaf} 个空目录"
            f"（--phase 省略时默认 2；当前 phase={phase_eff}，{ph_desc}）。"
        )
    if args.project_name is None:
        print("（本轮使用自动目录名；「第一批」内可同时保留多套示例以便对比）")
    print("可用：python check_tobacco_kb_required_files.py")
    return 0


if __name__ == "__main__":
    sys.exit(main())
