"""
一键清理工作区输出：

  • 「检查结果」目录内的验收报告等文件（默认仅删该目录下内容，不删目录本身）
  • 「第一批」目录下的各项目子文件夹（模拟数据通常在此，**不可逆**）

用法（在仓库根目录 `yc-checking` 下执行）：

  # 仅清空检查结果
  python clean_workspace.py --results

  # 仅删除第一批下全部项目子目录（需确认）
  python clean_workspace.py --mock --yes

  # 两项一起做
  python clean_workspace.py --all --yes

  # 先预览不真正删除
  python clean_workspace.py --all --dry-run

  # 指定工作区根目录（该目录下应有「第一批」「检查结果」）
  python clean_workspace.py --base "D:\\资料上报" --results

Windows 双击 .bat「没反应」或闪退：多半是未找到 python。批处理已优先使用 py -3；
若仍失败，请在「设置 → 应用 → 应用执行别名」中关闭 python.exe 占位，或改用 CMD 手动执行上面命令。
"""

from __future__ import annotations

import argparse
import shutil
import sys
from pathlib import Path
from typing import List

RESULT_DIR_NAME = "检查结果"
BATCH_DIR_NAME = "第一批"


def _rm_tree(path: Path) -> None:
    if path.is_file():
        path.unlink()
    elif path.is_dir():
        shutil.rmtree(path)


def clean_results_dir(base: Path, dry_run: bool) -> List[str]:
    """删除「检查结果」内全部条目；目录本身保留。"""
    rd = base / RESULT_DIR_NAME
    msgs: List[str] = []
    if not rd.is_dir():
        msgs.append(f"跳过（不存在）：{rd}")
        return msgs
    for child in sorted(rd.iterdir()):
        if dry_run:
            msgs.append(f"[预览] 将删除：{child}")
        else:
            _rm_tree(child)
            msgs.append(f"已删除：{child}")
    return msgs


def clean_batch_projects(batch_dir: Path, dry_run: bool) -> List[str]:
    """删除「第一批」下所有一级子目录（每个视为一个项目资料包）。"""
    msgs: List[str] = []
    if not batch_dir.is_dir():
        msgs.append(f"跳过（不存在）：{batch_dir}")
        return msgs
    subs = sorted([p for p in batch_dir.iterdir() if p.is_dir()])
    if not subs:
        msgs.append(f"{batch_dir} 下无项目子目录。")
        return msgs
    for p in subs:
        if dry_run:
            msgs.append(f"[预览] 将删除目录：{p}")
        else:
            shutil.rmtree(p)
            msgs.append(f"已删除目录：{p}")
    return msgs


def _configure_stdio_utf8() -> None:
    if sys.platform != "win32":
        return
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")
        except Exception:
            pass


def main() -> int:
    _configure_stdio_utf8()
    parser = argparse.ArgumentParser(
        description="清理「检查结果」与/或「第一批」下的项目目录（模拟资料）。"
    )
    parser.add_argument(
        "--base",
        type=Path,
        default=None,
        help="工作区根目录（其下含「第一批」「检查结果」），默认：当前工作目录",
    )
    parser.add_argument(
        "--results",
        action="store_true",
        help=f"清空「{RESULT_DIR_NAME}」内的全部文件与子目录",
    )
    parser.add_argument(
        "--mock",
        action="store_true",
        help=f"删除「{BATCH_DIR_NAME}」下全部一级项目子目录",
    )
    parser.add_argument(
        "--all",
        action="store_true",
        help="等价于同时指定 --results 与 --mock",
    )
    parser.add_argument(
        "--yes",
        action="store_true",
        help="确认执行对「第一批」下项目的删除（mock 场景必填，除非使用 --dry-run）",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="仅打印将要删除的路径，不真正删除",
    )

    args = parser.parse_args()
    base = (args.base or Path.cwd()).resolve()

    do_results = args.results or args.all
    do_mock = args.mock or args.all

    if not do_results and not do_mock:
        parser.error("请指定 --results、--mock 之一，或使用 --all")

    if do_mock and not args.dry_run and not args.yes:
        parser.error(
            f"删除「{BATCH_DIR_NAME}」下项目目录具有风险，请追加 --yes 确认；"
            "或先使用 --dry-run 预览。"
        )

    print(f"工作区根目录：{base}")

    if do_results:
        print(f"\n—— {RESULT_DIR_NAME} ——")
        for line in clean_results_dir(base, args.dry_run):
            print(line)

    if do_mock:
        print(f"\n—— {BATCH_DIR_NAME}（项目子目录）——")
        for line in clean_batch_projects(base / BATCH_DIR_NAME, args.dry_run):
            print(line)

    if args.dry_run:
        print("\n（以上为预览，未执行删除。去掉 --dry-run 并按要求加 --yes 后才会删除。）")

    return 0


if __name__ == "__main__":
    sys.exit(main())
