from __future__ import annotations

import argparse
from dataclasses import replace
from pathlib import Path
import sys

from .app import run_app
from .converter import (
    ConvertSettings,
    check_dependencies,
    missing_dependency_message,
    plan_conversions,
    run_conversions,
    scan_markdown_files,
    settings_from_project,
)
from .project import create_project, load_project


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="md2doc")
    subparsers = parser.add_subparsers(dest="command")

    subparsers.add_parser("gui", help="Open the desktop app")

    init_parser = subparsers.add_parser("init", help="Create a project from a folder")
    init_parser.add_argument("folder")
    init_parser.add_argument("--name")

    scan_parser = subparsers.add_parser("scan", help="Scan Markdown files in a project")
    scan_parser.add_argument("folder")

    convert_parser = subparsers.add_parser("convert", help="Convert Markdown files in a project")
    convert_parser.add_argument("folder")
    convert_parser.add_argument("files", nargs="*")
    convert_parser.add_argument("--format", default=None, choices=["docx", "html", "pdf"])
    convert_parser.add_argument("--output-dir", default=None)
    convert_parser.add_argument("--force", action="store_true")
    convert_parser.add_argument("--no-skip", action="store_true")

    deps_parser = subparsers.add_parser("deps", help="Check conversion tools")
    deps_parser.add_argument("--format", default="docx", choices=["docx", "html", "pdf"])

    args = parser.parse_args(argv)

    if args.command in (None, "gui"):
        run_app()
        return 0
    if args.command == "init":
        config = create_project(args.folder, args.name)
        print(f"Created project: {config.name}")
        print(config.root)
        return 0
    if args.command == "scan":
        config = load_project(args.folder)
        sources = scan_markdown_files(config.root, recursive=config.recursive, output_dir=config.output_path)
        if not sources:
            print("No Markdown files found.")
            return 0
        for source in sources:
            print(source.relative_to(config.root).as_posix())
        return 0
    if args.command == "deps":
        checks = check_dependencies(ConvertSettings(output_format=args.format))
        for check in checks:
            state = "ok" if check.available else "missing"
            print(f"{check.name}: {state} - {check.detail}")
        return 0 if all(check.available for check in checks) else 1
    if args.command == "convert":
        return _convert(args)

    parser.print_help()
    return 1


def _convert(args: argparse.Namespace) -> int:
    config = load_project(args.folder)
    output_format = args.format or config.output_format
    output_dir = Path(args.output_dir).expanduser().resolve() if args.output_dir else config.output_path
    settings = replace(
        settings_from_project(config, force=args.force),
        output_format=output_format,
        output_dir=output_dir,
        skip_unchanged=not args.no_skip,
    )
    if args.files:
        sources = [(config.root / file).resolve() for file in args.files]
    else:
        sources = scan_markdown_files(config.root, recursive=config.recursive, output_dir=output_dir)

    planned = plan_conversions(config.root, sources, settings)
    for item in planned:
        print(f"{item.action:7} {item.relative_source} -> {item.output.name} ({item.reason})")
    try:
        results = run_conversions(config.root, sources, settings)
    except RuntimeError as exc:
        print(str(exc), file=sys.stderr)
        return 2
    failed = [result for result in results if result.status == "failed"]
    for result in results:
        print(f"{result.status:9} {result.item.relative_source}: {result.message}")
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
