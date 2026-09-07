from __future__ import annotations

from pathlib import Path
import tempfile
import unittest
from unittest.mock import MagicMock, patch
import zipfile

from md2doc.converter import ConvertSettings, _ensure_diagram_filter_lua, _pandoc_command, run_conversions
from md2doc.diagram_renderer import (
    compute_diagram_hash,
    normalize_drawio_xml,
    render_d2,
    render_drawio,
    render_svg_code,
)
from md2doc.project import ProjectConfig


class TestDiagramRenderer(unittest.TestCase):
    def test_normalize_drawio_xml_full_mxfile(self) -> None:
        full_xml = '<mxfile host="Electron"><diagram><mxGraphModel/></diagram></mxfile>'
        self.assertEqual(normalize_drawio_xml(full_xml), full_xml)

    def test_normalize_drawio_xml_bare_graph_model(self) -> None:
        bare_xml = '<mxGraphModel dx="100" dy="100"><root><mxCell id="0"/></root></mxGraphModel>'
        normalized = normalize_drawio_xml(bare_xml)
        self.assertIn("<mxfile", normalized)
        self.assertIn("<diagram", normalized)
        self.assertIn(bare_xml, normalized)

    def test_normalize_drawio_xml_strips_fences(self) -> None:
        fenced = "```drawio\n<mxfile><diagram/></mxfile>\n```"
        normalized = normalize_drawio_xml(fenced)
        self.assertEqual(normalized, "<mxfile><diagram/></mxfile>")

    def test_render_svg_code(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            out_file = Path(tmpdir) / "test.svg"
            code = '<svg width="100" height="50"><rect/></svg>'
            render_svg_code(code, out_file)
            self.assertTrue(out_file.exists())
            content = out_file.read_text(encoding="utf-8")
            self.assertIn('xmlns="http://www.w3.org/2000/svg"', content)
            self.assertIn("<rect/>", content)

    def test_compute_diagram_hash_deterministic(self) -> None:
        h1 = compute_diagram_hash("d2", "x -> y", theme="default", sketch=False)
        h2 = compute_diagram_hash("d2", "x -> y", theme="default", sketch=False)
        h3 = compute_diagram_hash("d2", "x -> y", theme="dark", sketch=False)
        self.assertEqual(h1, h2)
        self.assertNotEqual(h1, h3)


class TestDiagramConversion(unittest.TestCase):
    def test_ensure_diagram_filter_lua_creates_file(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            root = Path(tmpdir)
            settings = ConvertSettings(d2_theme="neutral-default", d2_sketch=True)
            lua_path = _ensure_diagram_filter_lua(root, settings)
            self.assertTrue(lua_path.exists())
            content = lua_path.read_text(encoding="utf-8")
            self.assertIn("diagram-filter.lua", content)
            self.assertIn("neutral-default", content)
            self.assertIn('"true"', content)

    def test_pandoc_command_includes_diagram_filter(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            root = Path(tmpdir)
            src = root / "test.md"
            src.write_text("# Hello", encoding="utf-8")
            out = root / "test.docx"
            from md2doc.converter import scan_source_files, plan_conversions
            planned = plan_conversions(root, [src], ConvertSettings())
            cmd = _pandoc_command(root, planned[0], ConvertSettings())
            has_diagram_filter = any("diagram-filter.lua" in arg for arg in cmd)
            self.assertTrue(has_diagram_filter)

    def test_convert_markdown_with_svg_codeblock_and_raw_svg(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            root = Path(tmpdir)
            md_content = """# Diagram Test Document

Here is an inline raw HTML SVG:

<svg xmlns="http://www.w3.org/2000/svg" width="200" height="100">
  <rect width="200" height="100" fill="#2563eb"/>
  <text x="50" y="55" fill="white">Raw SVG</text>
</svg>

Here is an SVG codeblock with caption:

```{.svg caption="Architecture State Chart"}
<svg xmlns="http://www.w3.org/2000/svg" width="200" height="100">
  <circle cx="50" cy="50" r="40" fill="#16a34a"/>
  <text x="100" y="55" fill="black">Codeblock SVG</text>
</svg>
```
"""
            src = root / "diagrams.md"
            src.write_text(md_content, encoding="utf-8")
            settings = ConvertSettings(output_dir=root, output_format="docx", force=True)

            results = run_conversions(root, [src], settings)
            self.assertEqual(len(results), 1)
            self.assertEqual(results[0].status, "converted", msg=results[0].message)

            docx_path = root / "diagrams.docx"
            self.assertTrue(docx_path.exists())

            with zipfile.ZipFile(docx_path) as z:
                media_files = [n for n in z.namelist() if n.startswith("word/media/")]
                # Both raw SVG and codeblock SVG should have generated media entries
                self.assertGreaterEqual(len(media_files), 2, f"Found media files: {media_files}")
                svg_files = [n for n in media_files if n.endswith(".svg")]
                self.assertGreaterEqual(len(svg_files), 2, f"Expected SVG media files, got: {media_files}")

    def test_convert_markdown_with_plain_svg_fence(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            root = Path(tmpdir)
            md_content = """# Plain SVG Test

```svg
<!-- caption: Plain Fence Diagram -->
<svg xmlns="http://www.w3.org/2000/svg" width="150" height="80">
  <rect width="150" height="80" fill="#9333ea"/>
</svg>
```
"""
            src = root / "plain_svg.md"
            src.write_text(md_content, encoding="utf-8")
            settings = ConvertSettings(output_dir=root, output_format="docx", force=True)

            results = run_conversions(root, [src], settings)
            self.assertEqual(results[0].status, "converted", msg=results[0].message)

            docx_path = root / "plain_svg.docx"
            self.assertTrue(docx_path.exists())
            with zipfile.ZipFile(docx_path) as z:
                media_files = [n for n in z.namelist() if n.startswith("word/media/") and n.endswith(".svg")]
                self.assertGreaterEqual(len(media_files), 1)

    def test_dependency_checks_d2_and_drawio(self) -> None:
        from md2doc.converter import check_d2, check_drawio
        c_d2 = check_d2("d2_nonexistent_executable")
        self.assertFalse(c_d2.available)
        self.assertEqual(c_d2.name, "D2 CLI")

        c_drawio = check_drawio("drawio_nonexistent_executable")
        # May be True if Playwright is available, or False if not, but name should be Draw.io
        self.assertEqual(c_drawio.name, "Draw.io")



if __name__ == "__main__":
    unittest.main()
