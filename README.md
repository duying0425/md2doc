# md2doc

Local Markdown document converter powered by Pandoc and `mermaid-filter`.

## Features

- Create projects from folders.
- Scan Markdown files recursively.
- Convert one selected file or a batch of files.
- Use Pandoc with `mermaid-filter` so Mermaid diagrams are rendered during export.
- Skip unchanged Markdown files when a previous output already exists.
- Configure document output per project: table of contents, section numbering, title metadata, Word reference template, font, table borders, Mermaid image defaults, and extra Pandoc arguments.
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
- Word: `reference.docx`, default font, default font size, and table border style.
- Mermaid: format, theme, and background.
- Advanced: extra Pandoc arguments.

For DOCX, a selected `reference.docx` has priority for Word-specific styling. If no reference file is selected, md2doc can generate `.md2doc/generated-reference.docx` for the configured font, font size, and table border options.
DOCX image paragraphs are centered automatically after conversion.

## Requirements

Install the external conversion tools before converting:

```powershell
winget install JohnMacFarlane.Pandoc
npm install -g mermaid-filter
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

The CLI can initialize projects, scan Markdown files, preview conversion plans,
convert project batches, convert a single Markdown file directly, and check the
external conversion tools.

```powershell
$env:PYTHONPATH = "$PWD\src"
python -m md2doc init C:\docs --name "Docs" --format docx
python -m md2doc scan C:\docs --no-recursive
python -m md2doc plan C:\docs
python -m md2doc convert C:\docs --format docx
python -m md2doc convert C:\docs README.md docs\guide.md --format pdf --force
python -m md2doc convert C:\docs\README.md --format html --output-dir C:\docs\build
python -m md2doc deps
```

The default output location is the source file's folder.

When the package is installed, replace `python -m md2doc` with `md2doc`.

### Commands

- `md2doc` or `md2doc gui`: open the desktop app.
- `md2doc init <folder>`: create `.md2doc/project.json`.
- `md2doc scan <folder>`: list Markdown files in a project.
- `md2doc plan <folder-or-file> [files...]`: print the conversion plan without running Pandoc.
- `md2doc convert <folder-or-file> [files...]`: run conversions.
- `md2doc deps`: check Pandoc and `mermaid-filter`.

`convert` and `plan` accept either a project folder or one Markdown file. When a
folder is used, optional file arguments are resolved relative to the project
folder:

```powershell
md2doc plan C:\docs README.md docs\guide.md
md2doc convert C:\docs\README.md --format docx
```

### Conversion Options

Common `plan` and `convert` options:

- `--format docx|html|pdf`: override the project output format.
- `--output-dir <folder>`: write outputs under a separate folder.
- `--recursive` / `--no-recursive`: control project scanning.
- `--force`: convert even when outputs look up to date.
- `--no-skip`: disable smart skipping for unchanged files.
- `--dry-run`: print the plan from `convert` without running Pandoc.
- `--toc`, `--toc-depth <n>`, `--number-sections`: document structure options.
- `--title-page`, `--title`, `--subtitle`, `--author`, `--date`: metadata options.
- `--reference-docx <file>`, `--default-font <name>`, `--font-size <n>`, `--table-borders template|bordered|plain`: DOCX styling options.
- `--mermaid-format png|svg|pdf`, `--mermaid-theme <name>`, `--mermaid-background <value>`: Mermaid rendering options.
- `--pandoc <command>`, `--mermaid-filter <command>`: override tool commands or paths.
- `--pandoc-arg=<arg>`: append a raw Pandoc argument. Repeat for multiple arguments.

Examples:

```powershell
md2doc convert C:\docs --toc --toc-depth 2 --number-sections --title "Team Handbook"
md2doc convert C:\docs --reference-docx C:\templates\reference.docx
md2doc convert C:\docs --pandoc "C:\Tools\Pandoc\pandoc.exe" --pandoc-arg=--embed-resources
```

### Exit Codes

- `0`: command completed successfully.
- `1`: at least one conversion failed, or `deps` found a missing tool.
- `2`: usage error, invalid conversion settings, or missing required external tools.
