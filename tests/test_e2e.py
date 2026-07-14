"""End-to-end conversion tests that exercise the real external toolchain.

These tests are skipped automatically when the corresponding tool is missing, so
they remain green on minimal CI runners while providing genuine conversion
coverage on developer machines where Pandoc / Quarto / Playwright / MarkItDown
are installed.
"""

from __future__ import annotations

import dataclasses
import shutil
import tempfile
import unittest
import zipfile
from pathlib import Path

from md2doc.converter import run_conversions, scan_source_files, settings_from_project
from md2doc.project import (
    KIND_DOC2MD,
    KIND_HTML2PDF,
    KIND_MD2DOC,
    KIND_QMD2PPT,
    ProjectConfig,
    create_project,
)


def _has_pandoc() -> bool:
    return shutil.which("pandoc") is not None


def _has_quarto() -> bool:
    return shutil.which("quarto") is not None


def _has_markitdown() -> bool:
    try:
        import importlib.util

        return importlib.util.find_spec("markitdown") is not None
    except Exception:
        return False


def _has_playwright() -> bool:
    try:
        import importlib.util

        if importlib.util.find_spec("playwright") is None:
            return False
        from playwright.sync_api import sync_playwright

        with sync_playwright() as pw:
            return bool(pw.chromium.executable_path)
    except Exception:
        return False


def _write(path: Path, content: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content, encoding="utf-8")


def _make_project(root: Path, kind: str) -> ProjectConfig:
    return create_project(root, name=root.name, kind=kind)


def _override(settings, **changes):
    """Return a copy of *settings* with the given fields overridden."""
    return dataclasses.replace(settings, **changes)


class Md2DocE2ETests(unittest.TestCase):
    """Real Markdown -> DOCX conversions through Pandoc + mermaid-filter."""

    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.root = Path(self._tmp.name)
        self.config = _make_project(self.root, KIND_MD2DOC)
        self.settings = settings_from_project(self.config)

    def tearDown(self) -> None:
        self._tmp.cleanup()

    @unittest.skipUnless(_has_pandoc(), "Pandoc is required for md2doc e2e tests")
    def test_basic_markdown_to_docx(self) -> None:
        source = self.root / "hello.md"
        _write(source, "# Hello World\n\nThis is a smoke test.\n\n## Section\n\nText.\n")
        results = run_conversions(self.root, [source], self.settings)
        self.assertEqual(len(results), 1)
        result = results[0]
        self.assertTrue(result.item.output.exists(), "output DOCX must be created")
        self.assertGreater(result.item.output.stat().st_size, 0)
        # DOCX is a zip archive starting with the PK signature.
        with open(result.item.output, "rb") as fp:
            self.assertEqual(fp.read(2), b"PK")
        self.assertEqual(result.status, "converted")

    @unittest.skipUnless(_has_pandoc(), "Pandoc is required for md2doc e2e tests")
    def test_table_and_toc_and_number_sections(self) -> None:
        source = self.root / "doc.md"
        _write(
            source,
            "# Report\n\n"
            "| Name | Value |\n|------|-------|\n| Alpha | 1 |\n| Beta | 2 |\n\n"
            "## Analysis\n\nSome analysis text.\n",
        )
        settings = _override(self.settings, toc=True, toc_depth=3, number_sections=True)
        results = run_conversions(self.root, [source], settings)
        self.assertEqual(len(results), 1)
        self.assertTrue(results[0].item.output.exists())
        with zipfile.ZipFile(results[0].item.output) as z:
            doc = z.read("word/document.xml").decode("utf-8")
        self.assertIn("TOC", doc)
        self.assertIn("Alpha", doc)

    @unittest.skipUnless(_has_pandoc(), "Pandoc is required for md2doc e2e tests")
    def test_horizontal_rule_to_page_break(self) -> None:
        source = self.root / "paged.md"
        _write(
            source,
            "# First Page\n\nContent one.\n\n---\n\n# Second Page\n\nContent two.\n",
        )
        settings = _override(self.settings, hr_to_pagebreak=True)
        results = run_conversions(self.root, [source], settings)
        self.assertTrue(results[0].item.output.exists())
        with zipfile.ZipFile(results[0].item.output) as z:
            doc = z.read("word/document.xml").decode("utf-8")
        self.assertIn('w:type="page"', doc)

    @unittest.skipUnless(_has_pandoc(), "Pandoc is required for md2doc e2e tests")
    def test_title_page_metadata(self) -> None:
        source = self.root / "titled.md"
        _write(source, "# Body\n\nContent.\n")
        settings = _override(
            self.settings,
            title_page=True,
            title="My Title",
            author="Tester",
            date="2026-07-14",
        )
        results = run_conversions(self.root, [source], settings)
        self.assertTrue(results[0].item.output.exists())
        with zipfile.ZipFile(results[0].item.output) as z:
            doc = z.read("word/document.xml").decode("utf-8")
        self.assertIn("My Title", doc)
        self.assertIn("Tester", doc)
        self.assertIn("2026-07-14", doc)

    @unittest.skipUnless(_has_pandoc(), "Pandoc is required for md2doc e2e tests")
    def test_skip_unchanged_on_second_run(self) -> None:
        source = self.root / "skip.md"
        _write(source, "# Skip\n\nBody.\n")
        first = run_conversions(self.root, [source], self.settings)
        self.assertEqual(first[0].status, "converted")
        # Re-run without modifying the source -> should be skipped.
        second = run_conversions(self.root, [source], self.settings)
        self.assertEqual(second[0].status, "skipped")

    @unittest.skipUnless(_has_pandoc(), "Pandoc is required for md2doc e2e tests")
    def test_force_overrides_skip(self) -> None:
        source = self.root / "force.md"
        _write(source, "# Force\n\nBody.\n")
        run_conversions(self.root, [source], self.settings)
        settings = _override(self.settings, force=True)
        second = run_conversions(self.root, [source], settings)
        self.assertEqual(second[0].status, "converted")

    @unittest.skipUnless(_has_pandoc(), "Pandoc is required for md2doc e2e tests")
    def test_default_font_generates_reference_docx(self) -> None:
        source = self.root / "styled.md"
        _write(source, "# Styled\n\nBody.\n")
        # Clear the explicit reference_docx so the generated-reference path is used.
        settings = _override(
            self.settings,
            reference_docx="",
            default_font="Courier New",
            default_font_size=12,
        )
        results = run_conversions(self.root, [source], settings)
        self.assertTrue(results[0].item.output.exists())
        generated = self.root / ".md2doc" / "generated-reference.docx"
        self.assertTrue(generated.exists())
        with zipfile.ZipFile(results[0].item.output) as z:
            styles = z.read("word/styles.xml").decode("utf-8")
        self.assertIn("Courier New", styles)


class Doc2MdE2ETests(unittest.TestCase):
    """Real DOCX -> Markdown conversions through MarkItDown."""

    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.root = Path(self._tmp.name)
        self.config = _make_project(self.root, KIND_DOC2MD)
        self.settings = settings_from_project(self.config)

    def tearDown(self) -> None:
        self._tmp.cleanup()

    @unittest.skipUnless(
        _has_pandoc() and _has_markitdown(),
        "Pandoc and MarkItDown are required for doc2md e2e tests",
    )
    def test_docx_roundtrip_to_markdown(self) -> None:
        # First produce a DOCX via md2doc so we have a known-good input.
        md_root = self.root / "src_md"
        md_root.mkdir()
        md_config = _make_project(md_root, KIND_MD2DOC)
        md_settings = settings_from_project(md_config)
        source_md = md_root / "hello.md"
        _write(source_md, "# Hello World\n\nThis is a smoke test.\n\n## Section\n\nText.\n")
        md_results = run_conversions(md_root, [source_md], md_settings)
        self.assertTrue(md_results[0].item.output.exists())

        # Then convert that DOCX back to Markdown via doc2md.
        docx_input = self.root / "hello.docx"
        shutil.copyfile(md_results[0].item.output, docx_input)
        sources = scan_source_files(
            self.root,
            kind=self.settings.kind,
            recursive=self.settings.recursive,
        )
        self.assertIn(docx_input, sources)
        results = run_conversions(self.root, [docx_input], self.settings)
        self.assertEqual(len(results), 1)
        self.assertTrue(results[0].item.output.exists())
        markdown = results[0].item.output.read_text(encoding="utf-8")
        self.assertIn("Hello World", markdown)
        self.assertIn("Section", markdown)


class Qmd2PptE2ETests(unittest.TestCase):
    """Real QMD -> PPTX conversions through Quarto."""

    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.root = Path(self._tmp.name)
        self.config = _make_project(self.root, KIND_QMD2PPT)
        self.settings = settings_from_project(self.config)

    def tearDown(self) -> None:
        self._tmp.cleanup()

    @unittest.skipUnless(_has_quarto(), "Quarto is required for qmd2ppt e2e tests")
    def test_qmd_to_pptx(self) -> None:
        source = self.root / "slides.qmd"
        _write(
            source,
            "---\ntitle: Smoke Test\nformat: pptx\n---\n\n"
            "# Hello\n\nSlide one content\n\n"
            "## Second\n\nMore content\n",
        )
        results = run_conversions(self.root, [source], self.settings)
        self.assertEqual(len(results), 1)
        self.assertTrue(results[0].item.output.exists())
        self.assertGreater(results[0].item.output.stat().st_size, 0)
        # PPTX is a zip archive starting with the PK signature.
        with open(results[0].item.output, "rb") as fp:
            self.assertEqual(fp.read(2), b"PK")
        self.assertEqual(results[0].status, "converted")


class Html2PdfE2ETests(unittest.TestCase):
    """Real HTML -> PDF conversions through Playwright/Chromium."""

    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.root = Path(self._tmp.name)
        self.config = _make_project(self.root, KIND_HTML2PDF)
        self.settings = settings_from_project(self.config)

    def tearDown(self) -> None:
        self._tmp.cleanup()

    @unittest.skipUnless(_has_playwright(), "Playwright/Chromium is required for html2pdf e2e tests")
    def test_html_to_pdf(self) -> None:
        source = self.root / "poster.html"
        _write(
            source,
            "<!DOCTYPE html><html><head><style>"
            "body{font-family:sans-serif;padding:40px;}"
            "h1{color:#333}"
            "</style></head><body><h1>Poster</h1>"
            "<p>Hello HTML to PDF.</p></body></html>",
        )
        results = run_conversions(self.root, [source], self.settings)
        self.assertEqual(len(results), 1)
        self.assertTrue(results[0].item.output.exists())
        self.assertGreater(results[0].item.output.stat().st_size, 0)
        with open(results[0].item.output, "rb") as fp:
            self.assertEqual(fp.read(4), b"%PDF")
        self.assertEqual(results[0].status, "converted")


if __name__ == "__main__":
    unittest.main()
