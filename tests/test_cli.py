from __future__ import annotations

from contextlib import redirect_stderr, redirect_stdout
import io
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from md2doc import cli
from md2doc.project import KIND_HTML2PDF, load_project


class CliTests(unittest.TestCase):
    def test_plan_accepts_single_markdown_file_without_converting(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "notes.md"
            source.write_text("# Notes", encoding="utf-8")
            stdout = io.StringIO()

            with (
                redirect_stdout(stdout),
                patch("md2doc.cli.run_conversions", side_effect=AssertionError("should not convert")),
            ):
                code = cli.main(["plan", str(source)])

            self.assertEqual(code, 0)
            self.assertIn("convert", stdout.getvalue())
            self.assertIn("notes.md", stdout.getvalue())

    def test_convert_accepts_single_file_and_applies_overrides(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "guide.md"
            source.write_text("# Guide", encoding="utf-8")

            with (
                redirect_stdout(io.StringIO()),
                patch("md2doc.cli.run_conversions", return_value=[]) as run_conversions,
            ):
                code = cli.main(
                    [
                        "convert",
                        str(source),
                        "--format",
                        "docx",
                        "--toc",
                        "--toc-depth",
                        "2",
                        "--title",
                        "Guide",
                        "--number-sections",
                        "--pandoc",
                        "custom-pandoc",
                        "--mermaid-filter",
                        "custom-filter",
                        "--mermaid-min-dpi",
                        "360",
                        "--figure-numbering",
                        "--figure-prefix",
                        "图",
                        "--figure-caption-position",
                        "above",
                        "--pandoc-arg=--standalone",
                        "--hr-to-pagebreak",
                    ]
                )

            self.assertEqual(code, 0)
            root, sources, settings = run_conversions.call_args.args
            self.assertEqual(root, source.parent.resolve())
            self.assertEqual(sources, [source.resolve()])
            self.assertEqual(settings.output_format, "docx")
            self.assertTrue(settings.toc)
            self.assertEqual(settings.toc_depth, 2)
            self.assertEqual(settings.title, "Guide")
            self.assertTrue(settings.number_sections)
            self.assertEqual(settings.pandoc_cmd, "custom-pandoc")
            self.assertEqual(settings.mermaid_filter_cmd, "custom-filter")
            self.assertEqual(settings.mermaid_min_dpi, 360.0)
            self.assertTrue(settings.figure_numbering)
            self.assertEqual(settings.figure_prefix, "图")
            self.assertEqual(settings.figure_caption_position, "above")
            self.assertIn("--standalone", settings.extra_pandoc_args)
            self.assertTrue(settings.hr_to_pagebreak)

    def test_scan_can_disable_recursive_search(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            (root / "a.md").write_text("# A", encoding="utf-8")
            (root / "sub").mkdir()
            (root / "sub" / "b.md").write_text("# B", encoding="utf-8")
            stdout = io.StringIO()

            with redirect_stdout(stdout):
                code = cli.main(["scan", str(root), "--no-recursive"])

            self.assertEqual(code, 0)
            self.assertEqual(stdout.getvalue().splitlines(), ["a.md"])

    def test_init_accepts_html2pdf_kind(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp) / "pages"

            with redirect_stdout(io.StringIO()):
                code = cli.main(["init", str(root), "--kind", "html2pdf"])

            self.assertEqual(code, 0)
            config = load_project(root)
            self.assertEqual(config.kind, KIND_HTML2PDF)
            self.assertEqual(config.output_format, "pdf")

    def test_plan_accepts_single_html_file_as_html2pdf_project(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "poster.html"
            source.write_text("<main></main>", encoding="utf-8")
            stdout = io.StringIO()

            with (
                redirect_stdout(stdout),
                patch("md2doc.cli.run_conversions", side_effect=AssertionError("should not convert")),
            ):
                code = cli.main(["plan", str(source)])

            self.assertEqual(code, 0)
            self.assertIn("poster.html", stdout.getvalue())
            self.assertIn("poster.pdf", stdout.getvalue())
            self.assertEqual(load_project(source.parent).kind, KIND_HTML2PDF)

    def test_missing_markdown_file_target_returns_usage_error(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            stderr = io.StringIO()

            with redirect_stderr(stderr):
                code = cli.main(["convert", str(Path(tmp) / "missing.md")])

            self.assertEqual(code, 2)
            self.assertIn("Input file not found", stderr.getvalue())

    def test_deps_command_with_install_flag(self) -> None:
        with patch("md2doc.dependencies.ensure_startup_dependencies") as mock_ensure, patch("md2doc.cli.check_dependencies", return_value=[]) as mock_check:
            code = cli.main(["deps", "--install", "--kind", "qmd2ppt"])
            self.assertEqual(code, 0)
            mock_ensure.assert_called_once_with(kind="qmd2ppt", on_progress=unittest.mock.ANY)
            mock_check.assert_called_once()

    def test_version_flag(self) -> None:
        # argparse version flag exits with SystemExit (or raises standard output/SystemExit)
        with patch("sys.stdout", new_callable=io.StringIO) as stdout:
            with self.assertRaises(SystemExit) as cm:
                cli.main(["--version"])
            self.assertEqual(cm.exception.code, 0)
            self.assertIn("md2doc", stdout.getvalue())

    def test_gui_subcommand_invokes_run_app(self) -> None:
        with patch("md2doc.cli.run_app") as mock_run_app:
            # 1. gui command
            code1 = cli.main(["gui"])
            self.assertEqual(code1, 0)
            
            # 2. empty/None command
            code2 = cli.main([])
            self.assertEqual(code2, 0)
            
            self.assertEqual(mock_run_app.call_count, 2)

    def test_cli_usage_error_combining_single_file_with_additional_args(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "a.md"
            source.write_text("# A", encoding="utf-8")
            
            stderr = io.StringIO()
            with redirect_stderr(stderr):
                code = cli.main(["convert", str(source), "extra_file.md"])
            
            self.assertEqual(code, 2)
            self.assertIn("A single file target cannot be combined", stderr.getvalue())

    def test_cli_usage_error_mismatched_file_suffix_for_project_kind(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            # create doc2md project config in the directory
            from md2doc.project import create_project, KIND_DOC2MD
            create_project(tmp, kind=KIND_DOC2MD)
            
            # target file is markdown (.md) but doc2md project expects Office suffixes (.docx/etc.)
            source = Path(tmp) / "invalid.md"
            source.write_text("# markdown", encoding="utf-8")
            
            stderr = io.StringIO()
            with redirect_stderr(stderr):
                code = cli.main(["convert", str(tmp), "invalid.md"])
            
            self.assertEqual(code, 2)
            self.assertIn("Not a Office file", stderr.getvalue())

    def test_cli_usage_error_file_outside_project_folder(self) -> None:
        with tempfile.TemporaryDirectory() as tmp1, tempfile.TemporaryDirectory() as tmp2:
            from md2doc.project import create_project
            create_project(tmp1)
            
            # File is in tmp2 (outside project root tmp1)
            outside_file = Path(tmp2) / "outside.md"
            outside_file.write_text("# Outside", encoding="utf-8")
            
            stderr = io.StringIO()
            with redirect_stderr(stderr):
                code = cli.main(["convert", tmp1, str(outside_file)])
            
            self.assertEqual(code, 2)
            self.assertIn("File is outside the project folder", stderr.getvalue())

    def test_scan_empty_directory_prints_no_files_found(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            # No files present in empty project
            stdout = io.StringIO()
            with redirect_stdout(stdout):
                code = cli.main(["scan", tmp])
            
            self.assertEqual(code, 0)
            self.assertIn("No Markdown files found", stdout.getvalue())


if __name__ == "__main__":
    unittest.main()

