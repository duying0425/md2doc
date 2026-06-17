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

## CLI

```powershell
$env:PYTHONPATH = "$PWD\src"
python -m md2doc init C:\path\to\project
python -m md2doc scan C:\path\to\project
python -m md2doc convert C:\path\to\project --format docx
```

The default output location is the source file's folder.
