from __future__ import annotations

from contextlib import redirect_stderr, redirect_stdout
import io
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from md2doc import cli


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
                        "html",
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
                        "--pandoc-arg=--standalone",
                    ]
                )

            self.assertEqual(code, 0)
            root, sources, settings = run_conversions.call_args.args
            self.assertEqual(root, source.parent.resolve())
            self.assertEqual(sources, [source.resolve()])
            self.assertEqual(settings.output_format, "html")
            self.assertTrue(settings.toc)
            self.assertEqual(settings.toc_depth, 2)
            self.assertEqual(settings.title, "Guide")
            self.assertTrue(settings.number_sections)
            self.assertEqual(settings.pandoc_cmd, "custom-pandoc")
            self.assertEqual(settings.mermaid_filter_cmd, "custom-filter")
            self.assertIn("--standalone", settings.extra_pandoc_args)

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

    def test_missing_markdown_file_target_returns_usage_error(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            stderr = io.StringIO()

            with redirect_stderr(stderr):
                code = cli.main(["convert", str(Path(tmp) / "missing.md")])

            self.assertEqual(code, 2)
            self.assertIn("Markdown file not found", stderr.getvalue())


if __name__ == "__main__":
    unittest.main()
