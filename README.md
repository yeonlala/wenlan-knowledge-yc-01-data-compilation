# yc-checking

资料包**目录结构、必交材料关键词、文件命名**的机检与辅助工具：对 **`第一批/<项目资料夹>/`** 跑验收，在 **`检查结果/`** 输出 Excel 与 HTML 报告；可选管线把《项目资料上报确认单》与业务文档整理为 **JSON + 本地知识库 Markdown**。

## 快速开始

在仓库根目录执行：

```bash
python -m pip install -r requirements.txt
python 1_check_tobacco_kb_required_files.py
```

第二个命令默认扫描当前目录下的 **`第一批`**。报告生成在同级 **`检查结果`** 中。

## 文档

| 文档 | 内容 |
|------|------|
| [常用命令示例.md](./常用命令示例.md) | 各脚本的完整命令行与联调步骤 |
| [资料整理与命名说明（非技术版）.md](./资料整理与命名说明（非技术版）.md) | 给整理同事：资料放哪、怎么起名 |
| [模板示例/使用说明.md](./模板示例/使用说明.md) | 模板项目夹与确认单生成器 |

## 根目录脚本一览

- **`1_check_tobacco_kb_required_files.py`** — 自动验收（**`第一批`** → **`检查结果`**）
- **`2_export_confirmation_bundle.py`** — 确认单 JSON → 多份清单 JSON，输出 **`extract_jsons/`**
- **`3_prepare_local_kb.py`** — 基于 `extract_jsons` 生成各项目 **`kb_local/`**（`manifest.json`、`extracted_markdown/`）
- **`generate_mock_tobacco_project.py`** — 默认在 **`mockdata/`** 生成模拟项目（`full` 含确认单）
- **`clean_workspace.py`** — 清理 **`检查结果`** 和/或 **`第一批`** 下项目子目录（同目录有 `clean_results.bat`、`clean_mock_projects.bat`）

## 配置

规则集中在 **`tobacco_kb/acceptance_config.py`**；文件名格式在 **`tobacco_kb/naming_convention.py`**。修改后机检与 Mock 会同步受影响。

## 其他目录

- **`mockdata/`** — Mock 脚本默认输出（可复制整夹到 **`第一批`** 再跑机检）
- **`模板示例/`** — 模板项目夹与 **`确认单生成器/index.html`**
- **`tobacco_kb/mock_templates/`** — 阶段一材料占位正文 `r01.md`～`r07.md`

更细的说明与边界见 [常用命令示例.md](./常用命令示例.md)。
