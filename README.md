# md2doc

[简体中文](README_zh.md) | English

Local document converter supporting four types of project formats:
1. **Markdown to Word documents** (`md2doc`): Convert Markdown (`.md`, `.markdown`) to DOCX using Pandoc and `mermaid-filter`.
2. **Office documents to Markdown** (`doc2md`): Convert Word, PowerPoint, and Excel (`.docx`, `.pptx`, `.xlsx`, etc.) to Markdown using MarkItDown.
3. **Quarto Markdown to PowerPoint** (`qmd2ppt`): Convert Quarto Markdown (`.qmd`) to PowerPoint presentations (`.pptx`) using Quarto CLI.
4. **HTML to single-page PDF** (`html2pdf`): Render HTML (`.html`, `.htm`) in Chromium and export one custom-sized PDF page based on the rendered HTML dimensions.

## Features

- Create projects from folders, choosing from four conversion formats.
- Scan source files recursively (handles `.md`/`.markdown` for Markdown projects, `.docx`/`.pptx`/etc. for Office projects, `.qmd` for Quarto projects, and `.html`/`.htm` for HTML projects).
- Convert one selected file or a batch of files.
- Use Pandoc with `mermaid-filter` so Mermaid diagrams are rendered during DOCX export (for Markdown projects).
- Use Quarto CLI to compile `.qmd` documents into `.pptx` slides (for Quarto projects).
- Use Chromium through Playwright to render HTML into one PDF page whose size follows the rendered HTML content instead of a default paper template.
- Skip unchanged source files when a previous output already exists.
- Configure document output per project.
- Store project metadata in `.md2doc/project.json`.
- Store conversion history in `.md2doc/manifest.json`.

By default, generated files are written next to the source Markdown file:

```text
README.md       -> README.docx
docs/guide.md   -> docs/guide.docx
```

Set the Output field or `--output-dir` if you want a separate output folder.

## Formatting

Open **Settings** in the desktop app to configure:

- Document: table of contents, TOC depth, section numbering, title, subtitle, author, and date.
- Word: `reference.docx`, default font (with preset choices), table border style, and an option to convert horizontal rules to page breaks.
- Mermaid: format, theme, and background.
- HTML to PDF: viewport width/height, device scale factor, render delay, and print background graphics.
- Advanced: extra Pandoc arguments.

For DOCX, a selected `reference.docx` has priority for Word-specific styling. When creating a new project (or opening an unconfigured legacy project), the tool automatically generates a default Word template at `.md2doc/reference.docx` and references it in the project configuration by default, allowing you to edit the styles in Word directly. If you clear this template reference, md2doc can still generate `.md2doc/generated-reference.docx` for the configured font and table border options.
DOCX image paragraphs are centered automatically after conversion.

## Requirements

Install the external conversion tools depending on your project type:

- **Markdown to documents**: Install Pandoc and `mermaid-filter`. Mermaid diagram rendering also needs the managed Playwright Chromium runtime:
  ```powershell
  winget install JohnMacFarlane.Pandoc
  npm install -g mermaid-filter
  python -m playwright install chromium
  ```
- **Office documents to Markdown**: The required `markitdown` Python package is automatically installed as a dependency.
- **Quarto Markdown to PowerPoint**: Install Quarto CLI from [quarto.org](https://quarto.org/docs/get-started/).
- **HTML to PDF**: Install the Python `playwright` package. md2doc uses installed Microsoft Edge or Google Chrome when available; otherwise install Playwright Chromium:
  ```powershell
  python -m pip install playwright
  python -m playwright install chromium
  ```

The app can still open and scan projects without those tools installed.

## Run

From this repository:

```powershell
python -m md2doc
```

If running directly from source without installing the package, use:

```powershell
$env:PYTHONPATH = "$PWD\src"
python -m md2doc
```

After installing the package, the same app is available as:

```powershell
python -m pip install -e .
md2doc
```

Running without a subcommand opens the desktop app. Use `md2doc --help` or
`python -m md2doc --help` to inspect the CLI.

## CLI

The CLI can initialize projects, scan project source files, preview conversion
plans, convert project batches, convert a single source file directly, and check
the external conversion tools.

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

The default output location is the source file's folder.

When the package is installed, replace `python -m md2doc` with `md2doc`.

### Commands

- `md2doc` or `md2doc gui`: open the desktop app.
- `md2doc init <folder>`: create `.md2doc/project.json`. Accepts optional `--kind md2doc|doc2md|qmd2ppt|html2pdf`.
- `md2doc scan <folder>`: list source files in a project.
- `md2doc plan <folder-or-file> [files...]`: print the conversion plan.
- `md2doc convert <folder-or-file> [files...]`: run conversions.
- `md2doc deps`: check installed conversion tools. Use `--kind html2pdf` to check Playwright and browser availability. Accepts an optional `--install` flag (e.g. `md2doc deps --install --kind html2pdf`) to automatically download and configure missing dependencies.

`convert` and `plan` accept either a project folder or one source file. When a
folder is used, optional file arguments are resolved relative to the project
folder:

```powershell
md2doc plan C:\docs README.md docs\guide.md
md2doc convert C:\docs\README.md --format docx
```

### Conversion Options

Common `plan` and `convert` options:

- `--format docx`: override the Markdown project output format.
- `--output-dir <folder>`: write outputs under a separate folder.
- `--recursive` / `--no-recursive`: control project scanning.
- `--force`: convert even when outputs look up to date.
- `--no-skip`: disable smart skipping for unchanged files.
- `--dry-run`: print the plan from `convert` without running Pandoc.
- `--toc`, `--toc-depth <n>`, `--number-sections`: document structure options.
- `--title-page`, `--title`, `--subtitle`, `--author`, `--date`: metadata options.
- `--reference-docx <file>`, `--default-font <name>`, `--font-size <n>` (CLI only), `--table-borders template|bordered|plain`, `--hr-to-pagebreak` / `--no-hr-to-pagebreak`: DOCX styling and layout options.
- `--mermaid-format png|svg|pdf`, `--mermaid-theme <name>`, `--mermaid-background <value>`, `--mermaid-scale <n>`, `--mermaid-min-dpi <n>`: Mermaid rendering and sizing options.
- `--figure-numbering` / `--no-figure-numbering`, `--figure-prefix <label>`, `--figure-caption-position below|above`: number image captions with Word `SEQ` fields.
- `--pandoc <command>`, `--mermaid-filter <command>`: override tool commands or paths.
- `--pandoc-arg=<arg>`: append a raw Pandoc argument. Repeat for multiple arguments.

Image captions can be numbered with Word-native fields when figure numbering is enabled:

```markdown
![System architecture](assets/arch.png){#fig:arch}
```

For Mermaid diagrams, put the Word caption in the fenced-code attributes:

````markdown
```{.mermaid #fig:login caption="Login flow"}
flowchart TD
  A[Open] --> B[Login]
```
````

Examples:

```powershell
md2doc convert C:\docs --toc --toc-depth 2 --number-sections --title "Team Handbook"
md2doc convert C:\docs --reference-docx C:\templates\reference.docx
md2doc convert C:\docs --figure-numbering --figure-prefix "Figure"
md2doc convert C:\docs --pandoc "C:\Tools\Pandoc\pandoc.exe" --pandoc-arg=--embed-resources
```

### Exit Codes

- `0`: command completed successfully.
- `1`: at least one conversion failed, or `deps` found a missing tool.
- `2`: usage error, invalid conversion settings, or missing required external tools.

## Development & Release

The build and release automation scripts are located in the `scripts/` folder:

### 1. Build Standalone Executable Locally

Package the application as a standalone executable (saved to `bin/md2doc.exe`) using PyInstaller:

```powershell
powershell -File scripts/build_exe.ps1 -Python .venv\Scripts\python.exe
```

### 2. Version Bump and Git Release

Use `scripts/release.ps1` to update file versions, commit, tag, and push to the remote repository:

```powershell
# Option A: Automatically increment the patch version (e.g., 0.4.1 -> 0.4.2)
powershell -File scripts/release.ps1

# Option B: Manually specify a target version
powershell -File scripts/release.ps1 -Version 0.5.0
```

> [!NOTE]
> Pushing a new version tag to GitHub triggers a GitHub Actions workflow that automatically compiles the executable and creates a GitHub Release.
