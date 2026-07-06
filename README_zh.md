# md2doc

[简体中文](README_zh.md) | [English](README.md)

支持四类项目格式的本地文档转换工具：
1. **Markdown 转 Word 文档** (`md2doc`)：使用 Pandoc 和 `mermaid-filter` 将 Markdown (`.md`, `.markdown`) 转换为 DOCX。
2. **Office 文档转 Markdown** (`doc2md`)：使用 MarkItDown 将 Word、PowerPoint 和 Excel (`.docx`, `.pptx`, `.xlsx` 等) 转换为 Markdown。
3. **Quarto 转 PowerPoint** (`qmd2ppt`)：使用 Quarto CLI 将 Quarto Markdown (`.qmd`) 转换为 PowerPoint 演示文稿 (`.pptx`)。
4. **HTML 转单页 PDF** (`html2pdf`)：在 Chromium 中渲染 HTML (`.html`, `.htm`) 并根据渲染后的 HTML 实际尺寸导出自定义大小的单页 PDF。

## 功能特性

- **从文件夹创建项目**，支持选择四类转换项目类型。
- **递归扫描源文件**（Markdown 项目扫描 `.md`/`.markdown`，Office 项目扫描 `.docx`/`.pptx`/等，Quarto 项目扫描 `.qmd`，HTML 项目扫描 `.html`/`.htm`）。
- **转换单个选中文件或批量转换文件**。
- **集成外部工具进行渲染**：
  - Markdown 项目结合 Pandoc 与 `mermaid-filter` 渲染 Mermaid 并导出 DOCX。
  - Quarto 项目使用 Quarto CLI 渲染 PPTX 幻灯片。
  - HTML 项目使用 Playwright / Chromium 将页面渲染并导出为自适应尺寸的单页 PDF。
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
- **Word**：`reference.docx`、默认字体（提供预设常用字体下拉框）、表格边框样式以及是否将水平分割线 (`---`) 转换为分页符。
- **Mermaid**：图片格式、主题和背景。
- **HTML 转 PDF**：浏览器视口宽高、设备缩放因子、渲染等待延迟以及是否打印背景图形。
- **高级（Advanced）**：额外的 Pandoc 参数。

对于 DOCX 格式，选中的 `reference.docx` 在 Word 特征样式上具有最高优先级。新建项目（或打开未配置的旧项目）时，系统会自动在 `.md2doc/` 目录下生成一个默认的 Word 模版 `reference.docx` 并在配置中默认引用它，用户可以直接打开该模版进行样式调整。如果用户清除了该模版引用，md2doc 仍然会根据配置的默认字体和表格边框选项自动生成 `.md2doc/generated-reference.docx`。

转换完成后，DOCX 中的图片段落会自动居中对齐。

## 环境要求

根据您使用的项目类型安装对应的外部转换工具：

- **Markdown 转文档**：安装 Pandoc 和 `mermaid-filter`。Mermaid 图表渲染还需要受控的 Playwright Chromium 运行时：
  ```powershell
  winget install JohnMacFarlane.Pandoc
  npm install -g mermaid-filter
  python -m playwright install chromium
  ```
- **Office 文档转 Markdown**：必需的 Python 库 `markitdown` 会在安装项目包时作为依赖自动安装。
- **Quarto 转 PowerPoint**：从 [quarto.org](https://quarto.org/docs/get-started/) 安装 Quarto CLI。
- **HTML 转 PDF**：安装 Python `playwright` 库。md2doc 会优先使用已安装的 Microsoft Edge 或 Google Chrome；如果均不可用，请安装 Playwright Chromium：
  ```powershell
  python -m pip install playwright
  python -m playwright install chromium
  ```

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
python -m md2doc init C:\pages --kind html2pdf
python -m md2doc convert C:\pages\poster.html
python -m md2doc deps
```

默认的输出路径是源文件所在的文件夹。

当包安装完成后，请将 `python -m md2doc` 替换为 `md2doc`。

### 命令列表

- `md2doc` 或 `md2doc gui`：打开桌面应用。
- `md2doc init <folder>`：创建 `.md2doc/project.json`，可通过 `--kind md2doc|doc2md|qmd2ppt|html2pdf` 参数指定类型。
- `md2doc scan <folder>`：列出项目中的源文件。
- `md2doc plan <folder-or-file> [files...]`：打印转换计划。
- `md2doc convert <folder-or-file> [files...]`：执行文档转换。
- `md2doc deps`：检查转换工具的安装状态。若使用 `--kind html2pdf` 参数，可检查 Playwright 及浏览器可用性。支持可选的 `--install` 参数（如 `md2doc deps --install --kind html2pdf`）自动下载和安装缺失的外部工具。

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
- `--reference-docx <file>`、`--default-font <name>`、`--font-size <n>`（仅限命令行）、`--table-borders template|bordered|plain`、`--hr-to-pagebreak` / `--no-hr-to-pagebreak`：DOCX 样式与布局选项。
- `--mermaid-format png|svg|pdf`、`--mermaid-theme <name>`、`--mermaid-background <value>`、`--mermaid-scale <n>`、`--mermaid-min-dpi <n>`：Mermaid 渲染和尺寸选项。
- `--figure-numbering` / `--no-figure-numbering`、`--figure-prefix <label>`、`--figure-caption-position below|above`：使用 Word 原生 `SEQ` 域为图片题注编号。
- `--pandoc <command>`、`--mermaid-filter <command>`：覆盖工具的执行命令或路径。
- `--pandoc-arg=<arg>`：追加原始 Pandoc 参数。如需多个参数请重复使用该选项。

启用图片编号后，普通图片可以直接使用 Markdown 图片说明作为 Word 题注：

```markdown
![系统架构](assets/arch.png){#fig:arch}
```

Mermaid 图的 Word 题注请写在 fenced-code 属性里：

````markdown
```{.mermaid #fig:login caption="登录流程"}
flowchart TD
  A[打开登录页] --> B[登录]
```
````

示例：

```powershell
md2doc convert C:\docs --toc --toc-depth 2 --number-sections --title "Team Handbook"
md2doc convert C:\docs --reference-docx C:\templates\reference.docx
md2doc convert C:\docs --figure-numbering --figure-prefix "图"
md2doc convert C:\docs --pandoc "C:\Tools\Pandoc\pandoc.exe" --pandoc-arg=--embed-resources
```

### 退出状态码

- `0`：命令成功执行。
- `1`：至少有一个转换失败，或者 `deps` 检测到缺失外部工具。
- `2`：用法错误、转换配置无效或缺少所需的外部工具。

## 开发与发布

本项目的构建与发布流程已经封装在 `scripts/` 目录下：

### 1. 本地打包编译

使用 PyInstaller 将项目打包为单文件绿色版可执行程序（输出至 `bin/md2doc.exe`）：

```powershell
powershell -File scripts/build_exe.ps1 -Python .venv\Scripts\python.exe
```

### 2. 版本升级与 Git 发布

使用 `scripts/release.ps1` 脚本可以自动修改版本号、提交修改、打 Tag 并推送至 GitHub 仓库：

```powershell
# 方式 A：自动递增小版本号（例如从 0.4.1 自动递增到 0.4.2）
powershell -File scripts/release.ps1

# 方式 B：手动指定具体的目标版本号
powershell -File scripts/release.ps1 -Version 0.5.0
```

> [!NOTE]
> 推送版本 Tag 到 GitHub 仓库后，GitHub Actions 工作流会自动编译 Windows 版本的 Standalone 可执行文件并发布到 Release 中。
