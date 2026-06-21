from __future__ import annotations

import argparse
from dataclasses import replace
import os
from pathlib import Path
import sys

from . import __version__
from .app import run_app
from .converter import (
    ConvertSettings,
    PlanItem,
    check_dependencies,
    plan_conversions,
    run_conversions,
    scan_source_files,
    settings_from_project,
)
from .project import KIND_DOC2MD, KIND_MD2DOC, KIND_QMD2PPT, VALID_KINDS, ProjectConfig, create_project, load_project


MARKDOWN_SUFFIXES = {".md", ".markdown"}
OFFICE_SUFFIXES = {".docx", ".doc", ".pptx", ".ppt", ".xlsx", ".xls"}
QMD_SUFFIXES = {".qmd"}
SOURCE_SUFFIXES = MARKDOWN_SUFFIXES | OFFICE_SUFFIXES | QMD_SUFFIXES
OUTPUT_FORMATS = ("docx",)


class CliUsageError(ValueError):
    pass


def main(argv: list[str] | None = None) -> int:
    _ensure_standard_streams()
    parser = argparse.ArgumentParser(
        prog="md2doc",
        description=(
            "Convert Markdown to DOCX with Pandoc, or convert Word/PPT/Excel "
            "to Markdown with MarkItDown."
        ),
    )
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    subparsers = parser.add_subparsers(dest="command")

    subparsers.add_parser("gui", help="Open the desktop app")

    init_parser = subparsers.add_parser("init", help="Create a project from a folder")
    init_parser.add_argument("folder")
    init_parser.add_argument("--name")
    init_parser.add_argument(
        "--kind",
        choices=sorted(VALID_KINDS),
        default=KIND_MD2DOC,
        help="md2doc converts Markdown to documents; doc2md converts Word/PPT/Excel to Markdown; qmd2ppt converts Quarto Markdown to PowerPoint.",
    )
    init_parser.add_argument("--format", choices=OUTPUT_FORMATS, dest="output_format")
    init_parser.add_argument("--output-dir")
    init_parser.add_argument("--recursive", action=argparse.BooleanOptionalAction, default=None)

    scan_parser = subparsers.add_parser("scan", help="Scan Markdown files in a project")
    scan_parser.add_argument("folder")
    scan_parser.add_argument("--output-dir", default=None)
    scan_parser.add_argument("--recursive", action=argparse.BooleanOptionalAction, default=None)

    plan_parser = subparsers.add_parser("plan", help="Show planned conversions without running Pandoc")
    _add_conversion_arguments(plan_parser, dry_run=False)

    convert_parser = subparsers.add_parser("convert", help="Convert Markdown files")
    _add_conversion_arguments(convert_parser, dry_run=True)

    deps_parser = subparsers.add_parser("deps", help="Check conversion tools")
    deps_parser.add_argument("--format", default="docx", choices=OUTPUT_FORMATS)
    deps_parser.add_argument("--pandoc", dest="pandoc_cmd", default="pandoc")
    deps_parser.add_argument("--mermaid-filter", dest="mermaid_filter_cmd", default="mermaid-filter")

    args = parser.parse_args(argv)

    try:
        if args.command in (None, "gui"):
            run_app()
            return 0
        if args.command == "init":
            config = create_project(args.folder, args.name, kind=args.kind)
            if args.output_format and config.kind == KIND_MD2DOC:
                config.output_format = args.output_format
            if args.output_dir:
                config.output_dir = args.output_dir
            if args.recursive is not None:
                config.recursive = args.recursive
            config.save()
            print(f"Created project: {config.name} ({config.kind})")
            print(config.root)
            return 0
        if args.command == "scan":
            return _scan(args)
        if args.command == "deps":
            checks = check_dependencies(
                ConvertSettings(
                    output_format=args.format,
                    pandoc_cmd=args.pandoc_cmd,
                    mermaid_filter_cmd=args.mermaid_filter_cmd,
                )
            )
            for check in checks:
                state = "ok" if check.available else "missing"
                print(f"{check.name}: {state} - {check.detail}")
            return 0 if all(check.available for check in checks) else 1
        if args.command == "plan":
            args.dry_run = True
            return _convert(args)
        if args.command == "convert":
            return _convert(args)
    except CliUsageError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2

    parser.print_help()
    return 1


def _ensure_standard_streams() -> None:
    if sys.stdout is None:
        sys.stdout = open(os.devnull, "w", encoding="utf-8")
    elif hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(errors="replace")
        except Exception:
            pass

    if sys.stderr is None:
        sys.stderr = open(os.devnull, "w", encoding="utf-8")
    elif hasattr(sys.stderr, "reconfigure"):
        try:
            sys.stderr.reconfigure(errors="replace")
        except Exception:
            pass


def _add_conversion_arguments(parser: argparse.ArgumentParser, *, dry_run: bool) -> None:
    parser.add_argument(
        "target",
        help="Project folder, or a single Markdown/Office file to convert directly.",
    )
    parser.add_argument("files", nargs="*")
    parser.add_argument("--format", default=None, choices=OUTPUT_FORMATS)
    parser.add_argument("--output-dir", default=None)
    parser.add_argument("--recursive", action=argparse.BooleanOptionalAction, default=None)
    parser.add_argument("--force", action="store_true")
    parser.add_argument("--no-skip", action="store_true")
    parser.add_argument("--toc", action=argparse.BooleanOptionalAction, default=None)
    parser.add_argument("--toc-depth", type=int, default=None)
    parser.add_argument("--title-page", action=argparse.BooleanOptionalAction, default=None)
    parser.add_argument("--title", default=None)
    parser.add_argument("--subtitle", default=None)
    parser.add_argument("--author", default=None)
    parser.add_argument("--date", default=None)
    parser.add_argument("--number-sections", action=argparse.BooleanOptionalAction, default=None)
    parser.add_argument("--reference-docx", default=None)
    parser.add_argument("--default-font", default=None)
    parser.add_argument("--font-size", type=int, dest="default_font_size", default=None)
    parser.add_argument("--table-borders", choices=("template", "bordered", "plain"), default=None)
    parser.add_argument("--mermaid-format", choices=("png", "svg", "pdf"), default=None)
    parser.add_argument("--mermaid-theme", default=None)
    parser.add_argument("--mermaid-background", default=None)
    parser.add_argument("--mermaid-scale", type=float, default=None)
    parser.add_argument("--pandoc", dest="pandoc_cmd", default=None)
    parser.add_argument("--mermaid-filter", dest="mermaid_filter_cmd", default=None)
    parser.add_argument(
        "--pandoc-arg",
        action="append",
        default=[],
        help="Append one raw Pandoc argument. Repeat for multiple arguments.",
    )
    if dry_run:
        parser.add_argument("--dry-run", action="store_true", help="Print the conversion plan without running Pandoc.")


def _convert(args: argparse.Namespace) -> int:
    config, explicit_sources = _load_conversion_target(args.target, args.files)
    settings = _settings_from_args(config, args)
    sources = explicit_sources or scan_source_files(
        config.root,
        kind=settings.kind,
        recursive=settings.recursive,
        output_dir=settings.output_dir,
    )

    planned = plan_conversions(config.root, sources, settings)
    _print_plan(planned)
    if args.dry_run:
        return 0
    try:
        results = run_conversions(config.root, sources, settings)
    except RuntimeError as exc:
        print(str(exc), file=sys.stderr)
        return 2
    failed = [result for result in results if result.status == "failed"]
    for result in results:
        print(f"{result.status:9} {result.item.relative_source}: {result.message}")
    return 1 if failed else 0


def _scan(args: argparse.Namespace) -> int:
    config = load_project(args.folder)
    output_dir = Path(args.output_dir).expanduser().resolve() if args.output_dir else config.output_path
    recursive = config.recursive if args.recursive is None else args.recursive
    sources = scan_source_files(
        config.root,
        kind=config.kind,
        recursive=recursive,
        output_dir=output_dir,
    )
    if not sources:
        print(f"No {_input_label(config.kind)} files found.")
        return 0
    for source in sources:
        print(source.relative_to(config.root).as_posix())
    return 0


def _load_conversion_target(target: str, files: list[str]) -> tuple[ProjectConfig, list[Path] | None]:
    target_path = Path(target).expanduser().resolve()
    if _looks_like_source(target_path):
        if files:
            raise CliUsageError("A single file target cannot be combined with additional file arguments.")
        if not target_path.exists():
            raise CliUsageError(f"Input file not found: {target_path}")
        config = load_project(target_path.parent)
        return config, [target_path]

    config = load_project(target_path)
    if not files:
        return config, None
    return config, _resolve_project_files(config, files)


def _resolve_project_files(config: ProjectConfig, files: list[str]) -> list[Path]:
    project_root = config.root
    if config.kind == KIND_DOC2MD:
        accepted = OFFICE_SUFFIXES
    elif config.kind == KIND_QMD2PPT:
        accepted = QMD_SUFFIXES
    else:
        accepted = MARKDOWN_SUFFIXES
    label = _input_label(config.kind)
    sources: list[Path] = []
    for value in files:
        path = Path(value).expanduser()
        source = path.resolve() if path.is_absolute() else (project_root / path).resolve()
        if not source.exists():
            raise CliUsageError(f"Input file not found: {source}")
        if source.suffix.lower() not in accepted:
            raise CliUsageError(f"Not a {label} file: {source}")
        try:
            source.relative_to(project_root)
        except ValueError as exc:
            raise CliUsageError(f"File is outside the project folder: {source}") from exc
        sources.append(source)
    return sources


def _settings_from_args(config: ProjectConfig, args: argparse.Namespace) -> ConvertSettings:
    output_dir = Path(args.output_dir).expanduser().resolve() if args.output_dir else config.output_path
    settings = replace(
        settings_from_project(config, force=args.force),
        output_format=args.format or config.output_format,
        output_dir=output_dir,
        skip_unchanged=not args.no_skip,
    )

    overrides = {}
    for name in (
        "recursive",
        "toc",
        "toc_depth",
        "title_page",
        "title",
        "subtitle",
        "author",
        "date",
        "number_sections",
        "reference_docx",
        "default_font",
        "default_font_size",
        "table_borders",
        "mermaid_format",
        "mermaid_theme",
        "mermaid_background",
        "mermaid_scale",
        "pandoc_cmd",
        "mermaid_filter_cmd",
    ):
        value = getattr(args, name, None)
        if value is not None:
            overrides[name] = value
    if args.pandoc_arg:
        overrides["extra_pandoc_args"] = (*settings.extra_pandoc_args, *args.pandoc_arg)
    if overrides:
        settings = replace(settings, **overrides)
    return settings


def _print_plan(planned: list[PlanItem]) -> None:
    if not planned:
        print("No input files found.")
        return
    for item in planned:
        print(f"{item.action:7} {item.relative_source} -> {item.output} ({item.reason})")


def _looks_like_source(path: Path) -> bool:
    return path.suffix.lower() in SOURCE_SUFFIXES


def _input_label(kind: str) -> str:
    if kind == KIND_DOC2MD:
        return "Office"
    if kind == KIND_QMD2PPT:
        return "Quarto"
    return "Markdown"


if __name__ == "__main__":
    raise SystemExit(main())
