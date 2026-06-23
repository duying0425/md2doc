# md2doc

[简体中文](README_zh.md) | [English](README.md)

支持三类项目格式的本地文档转换工具：
1. **Markdown 转 Word 文档** (`md2doc`)：使用 Pandoc 和 `mermaid-filter` 将 Markdown (`.md`, `.markdown`) 转换为 DOCX。
2. **Office 文档转 Markdown** (`doc2md`)：使用 MarkItDown 将 Word、PowerPoint 和 Excel (`.docx`, `.pptx`, `.xlsx` 等) 转换为 Markdown。
3. **Quarto 转 PowerPoint** (`qmd2ppt`)：使用 Quarto CLI 将 Quarto Markdown (`.qmd`) 转换为 PowerPoint 演示文稿 (`.pptx`)。

## 功能特性

- **从文件夹创建项目**，支持选择三类转换项目类型。
- **递归扫描源文件**（Markdown 项目扫描 `.md`/`.markdown`，Office 项目扫描 `.docx`/`.pptx`/等，Quarto 项目扫描 `.qmd`）。
- **转换单个选中文件或批量转换文件**。
- **集成外部工具进行渲染**（Markdown 项目结合 Pandoc 与 `mermaid-filter` 渲染 Mermaid 并导出 DOCX，Quarto 项目使用 Quarto CLI 渲染 PPTX 幻灯片）。
- **智能跳过未修改的源文件**（当已有历史输出且源文件未更改时）。
- **个性化项目配置参数**。
- **在 `.md2doc/project.json` 中存储项目元数据**。
- **在 `.md2doc/manifest.json` 中存储转换历史记录**。

默认情况下，生成的文件会写入源 Markdown 文件同级目录下：

```text
README.md       -> README.docx
docs/guide.md   -> docs/guide.docx
```

如果需要单独的输出文件夹，可设置 Output（输出）字段或 `--output-dir` 参数。

## 格式设置

在桌面应用中打开 **Settings（设置）** 可以配置：

- **文档（Document）**：目录、目录深度、章节编号、标题、副标题、作者和日期。
- **Word**：`reference.docx`、默认字体、默认字号和表格边框样式。
- **Mermaid**：图片格式、主题和背景。
- **高级（Advanced）**：额外的 Pandoc 参数。

对于 DOCX 格式，选中的 `reference.docx` 在 Word 特征样式上具有最高优先级。如果未选择引用文件，md2doc 可以根据配置的字体、字号和表格边框选项自动生成 `.md2doc/generated-reference.docx`。

转换完成后，DOCX 中的图片段落会自动居中对齐。

## 环境要求

根据您使用的项目类型安装对应的外部转换工具：

- **Markdown 转文档**：安装 Pandoc 和 `mermaid-filter`：
  ```powershell
  winget install JohnMacFarlane.Pandoc
  npm install -g mermaid-filter
  ```
- **Office 文档转 Markdown**：必需的 Python 库 `markitdown` 会在安装项目包时作为依赖自动安装。
- **Quarto 转 PowerPoint**：从 [quarto.org](https://quarto.org/docs/get-started/) 安装 Quarto CLI。

即使未安装这些工具，应用依然可以打开并扫描项目。

## 运行方式

从本仓库直接运行：

```powershell
python -m md2doc
```

如果直接从源码运行且不安装包，请使用：

```powershell
$env:PYTHONPATH = "$PWD\src"
python -m md2doc
```

安装包之后，可以通过以下方式运行该应用：

```powershell
python -m pip install -e .
md2doc
```

不带子命令直接运行会打开桌面 GUI 应用。使用 `md2doc --help` 或 `python -m md2doc --help` 可以查看 CLI 帮助信息。

## 命令行界面 (CLI)

CLI 可以执行初始化项目、扫描项目源文件、预览转换计划、批量转换项目、直接转换单个源文件以及检查外部转换工具等操作。

```powershell
$env:PYTHONPATH = "$PWD\src"
python -m md2doc init C:\docs --name "Docs" --format docx
python -m md2doc scan C:\docs --no-recursive
python -m md2doc plan C:\docs
python -m md2doc convert C:\docs --format docx
python -m md2doc convert C:\docs README.md docs\guide.md --format docx --force
python -m md2doc convert C:\docs\README.md --format docx --output-dir C:\docs\build
python -m md2doc deps
```

默认的输出路径是源文件所在的文件夹。

当包安装完成后，请将 `python -m md2doc` 替换为 `md2doc`。

### 命令列表

- `md2doc` 或 `md2doc gui`：打开桌面应用。
- `md2doc init <folder>`：创建 `.md2doc/project.json`，可通过 `--kind md2doc|doc2md|qmd2ppt` 参数指定类型。
- `md2doc scan <folder>`：列出项目中的源文件。
- `md2doc plan <folder-or-file> [files...]`：打印转换计划。
- `md2doc convert <folder-or-file> [files...]`：执行文档转换。
- `md2doc deps`：检查 Markdown 转换工具（Pandoc 和 `mermaid-filter`）的安装状态。

`convert` 和 `plan` 既可以接收项目文件夹，也可以接收单个 Markdown 文件。当指定文件夹时，可选的文件参数会解析为相对于项目文件夹的相对路径：

```powershell
md2doc plan C:\docs README.md docs\guide.md
md2doc convert C:\docs\README.md --format docx
```

### 转换选项

`plan` 和 `convert` 的通用选项：

- `--format docx`：覆盖 Markdown 项目的输出格式。
- `--output-dir <folder>`：将输出文件写入单独的文件夹。
- `--recursive` / `--no-recursive`：控制项目扫描是否递归。
- `--force`：强制转换，即使输出文件看起来已是最新。
- `--no-skip`：对未修改的文件禁用智能跳过。
- `--dry-run`：从 `convert` 打印计划而不运行 Pandoc。
- `--toc`、`--toc-depth <n>`、`--number-sections`：文档结构选项（目录、目录深度、章节编号）。
- `--title-page`、`--title`、`--subtitle`、`--author`、`--date`：文档元数据选项。
- `--reference-docx <file>`、`--default-font <name>`、`--font-size <n>`、`--table-borders template|bordered|plain`：DOCX 样式选项。
- `--mermaid-format png|svg|pdf`、`--mermaid-theme <name>`、`--mermaid-background <value>`、`--mermaid-scale <n>`、`--mermaid-min-dpi <n>`：Mermaid 渲染和尺寸选项。
- `--pandoc <command>`、`--mermaid-filter <command>`：覆盖工具的执行命令或路径。
- `--pandoc-arg=<arg>`：追加原始 Pandoc 参数。如需多个参数请重复使用该选项。

示例：

```powershell
md2doc convert C:\docs --toc --toc-depth 2 --number-sections --title "Team Handbook"
md2doc convert C:\docs --reference-docx C:\templates\reference.docx
md2doc convert C:\docs --pandoc "C:\Tools\Pandoc\pandoc.exe" --pandoc-arg=--embed-resources
```

### 退出状态码

- `0`：命令成功执行。
- `1`：至少有一个转换失败，或者 `deps` 检测到缺失外部工具。
- `2`：用法错误、转换配置无效或缺少所需的外部工具。
