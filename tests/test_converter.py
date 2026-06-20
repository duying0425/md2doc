from __future__ import annotations

import json
import os
from pathlib import Path
import subprocess
import tempfile
import unittest
from unittest.mock import patch
import xml.etree.ElementTree as ET
import zipfile

from md2doc.converter import (
    BuildManifest,
    ConvertSettings,
    _center_docx_images,
    _ensure_generated_reference_docx,
    _markitdown_command,
    _mermaid_environment,
    _pandoc_command,
    _resolve_command,
    check_dependencies,
    file_fingerprint,
    plan_conversions,
    run_conversions,
    scan_markdown_files,
    scan_source_files,
    settings_from_project,
    settings_signature,
)
from md2doc.project import KIND_DOC2MD, KIND_QMD2PPT, ProjectConfig


class ConverterTests(unittest.TestCase):
    def test_scan_markdown_files_excludes_metadata_and_output(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            (root / "a.md").write_text("# A", encoding="utf-8")
            (root / "sub").mkdir()
            (root / "sub" / "b.markdown").write_text("# B", encoding="utf-8")
            (root / ".md2doc").mkdir()
            (root / ".md2doc" / "hidden.md").write_text("# Hidden", encoding="utf-8")
            (root / "output").mkdir()
            (root / "output" / "old.md").write_text("# Old", encoding="utf-8")

            files = scan_markdown_files(root, output_dir=root / "output")

            self.assertEqual(
                [file.relative_to(root).as_posix() for file in files],
                ["a.md", "sub/b.markdown"],
            )

    def test_scan_markdown_files_keeps_subdirs_when_output_is_project_root(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            (root / "a.md").write_text("# A", encoding="utf-8")
            (root / "sub").mkdir()
            (root / "sub" / "b.md").write_text("# B", encoding="utf-8")

            files = scan_markdown_files(root, output_dir=root)

            self.assertEqual(
                [file.relative_to(root).as_posix() for file in files],
                ["a.md", "sub/b.md"],
            )

    def test_plan_outputs_next_to_source_when_output_dir_is_project_root(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / "sub" / "a.md"
            source.parent.mkdir()
            source.write_text("# A", encoding="utf-8")

            item = plan_conversions(root, [source], ConvertSettings(output_dir=root))[0]

            self.assertEqual(item.output, root / "sub" / "a.docx")

    def test_plan_skips_existing_output_without_manifest_when_output_is_newer(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / "a.md"
            output = root / "output" / "a.docx"
            source.write_text("# A", encoding="utf-8")
            output.parent.mkdir()
            output.write_text("generated", encoding="utf-8")
            os.utime(source, (100, 100))
            os.utime(output, (200, 200))

            item = plan_conversions(
                root,
                [source],
                ConvertSettings(output_dir=root / "output"),
                BuildManifest(path=root / ".md2doc" / "manifest.json"),
            )[0]

            self.assertEqual(item.action, "skip")
            self.assertEqual(item.reason, "output is newer than source")

    def test_plan_converts_when_manifest_hash_changes(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / "a.md"
            output = root / "output" / "a.docx"
            source.write_text("# A", encoding="utf-8")
            output.parent.mkdir()
            output.write_text("generated", encoding="utf-8")
            old_fingerprint = file_fingerprint(source)
            source.write_text("# A changed", encoding="utf-8")

            manifest = BuildManifest(path=root / ".md2doc" / "manifest.json")
            manifest.records["a.md"] = {
                "source_sha256": old_fingerprint.sha256,
                "settings_signature": "anything",
                "output": str(output),
            }

            item = plan_conversions(root, [source], ConvertSettings(output_dir=root / "output"), manifest)[0]

            self.assertEqual(item.action, "convert")
            self.assertEqual(item.reason, "source changed")

    def test_manifest_record_success_serializes_metadata(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / "a.md"
            source.write_text("# A", encoding="utf-8")
            settings = ConvertSettings(output_dir=root / "output")
            item = plan_conversions(root, [source], settings, BuildManifest(path=root / ".md2doc" / "manifest.json"))[0]
            manifest = BuildManifest(path=root / ".md2doc" / "manifest.json")

            manifest.record_success(item)
            manifest.save()

            payload = json.loads((root / ".md2doc" / "manifest.json").read_text(encoding="utf-8"))
            self.assertEqual(payload["records"]["a.md"]["source_sha256"], item.fingerprint.sha256)

    def test_plan_can_reuse_cached_fingerprint_when_source_metadata_matches(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / "a.md"
            source.write_text("# A", encoding="utf-8")
            settings = ConvertSettings(output_dir=root / "output")
            item = plan_conversions(root, [source], settings)[0]
            item.output.parent.mkdir()
            item.output.write_text("generated", encoding="utf-8")
            manifest = BuildManifest(path=root / ".md2doc" / "manifest.json")
            manifest.record_success(item)

            with patch(
                "md2doc.converter._file_fingerprint_from_stat",
                side_effect=AssertionError("full fingerprint should not run"),
            ):
                planned = plan_conversions(
                    root,
                    [source],
                    settings,
                    manifest,
                    use_cached_fingerprints=True,
                )

            self.assertEqual(planned[0].action, "skip")
            self.assertEqual(planned[0].fingerprint.sha256, item.fingerprint.sha256)

    def test_cached_plan_uses_stat_only_when_no_manifest_record_exists(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / "a.md"
            source.write_text("# A", encoding="utf-8")

            with patch(
                "md2doc.converter._file_fingerprint_from_stat",
                side_effect=AssertionError("full fingerprint should not run"),
            ):
                planned = plan_conversions(
                    root,
                    [source],
                    ConvertSettings(output_dir=root / "output"),
                    use_cached_fingerprints=True,
                )

            self.assertEqual(planned[0].action, "convert")
            self.assertEqual(planned[0].reason, "output missing")
            self.assertEqual(planned[0].fingerprint.sha256, "")

    def test_cached_plan_uses_stat_only_when_source_metadata_changes(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / "a.md"
            source.write_text("# A", encoding="utf-8")
            settings = ConvertSettings(output_dir=root / "output")
            item = plan_conversions(root, [source], settings)[0]
            item.output.parent.mkdir()
            item.output.write_text("generated", encoding="utf-8")
            manifest = BuildManifest(path=root / ".md2doc" / "manifest.json")
            manifest.record_success(item)
            source.write_text("# A changed", encoding="utf-8")

            with patch(
                "md2doc.converter._file_fingerprint_from_stat",
                side_effect=AssertionError("full fingerprint should not run"),
            ):
                planned = plan_conversions(
                    root,
                    [source],
                    settings,
                    manifest,
                    use_cached_fingerprints=True,
                )

            self.assertEqual(planned[0].action, "convert")
            self.assertEqual(planned[0].reason, "source changed")
            self.assertEqual(planned[0].fingerprint.sha256, "")

    def test_project_format_options_round_trip_into_settings(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            project = ProjectConfig(
                name="Docs",
                root=root,
                toc=True,
                toc_depth=2,
                title_page=True,
                title="Handbook",
                subtitle="Internal",
                author="Team",
                date="2026-06-17",
                number_sections=True,
                reference_docx=str(root / "reference.docx"),
                default_font="Aptos",
                default_font_size=11,
                table_borders="bordered",
                mermaid_format="svg",
                mermaid_theme="forest",
                mermaid_background="transparent",
            )

            loaded = ProjectConfig.from_dict(project.to_dict())
            settings = settings_from_project(loaded)

            self.assertTrue(settings.toc)
            self.assertEqual(settings.toc_depth, 2)
            self.assertTrue(settings.title_page)
            self.assertEqual(settings.reference_docx, str(root / "reference.docx"))
            self.assertEqual(settings.table_borders, "bordered")
            self.assertEqual(settings.mermaid_format, "svg")

    def test_settings_signature_resolves_reference_docx_relative_to_project_root(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            ref_path = root / "template.docx"
            ref_path.write_text("dummy", encoding="utf-8")
            settings = ConvertSettings(reference_docx="template.docx")
            sig_with_root = settings_signature(settings, root)
            sig_no_root = settings_signature(settings)
            if not Path("template.docx").exists():
                self.assertNotEqual(sig_with_root, sig_no_root)

    def test_pandoc_command_includes_document_options(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / "a.md"
            reference = root / "reference.docx"
            source.write_text("# A", encoding="utf-8")
            reference.write_text("placeholder", encoding="utf-8")
            settings = ConvertSettings(
                output_dir=root,
                toc=True,
                toc_depth=2,
                title_page=True,
                title="Handbook",
                author="Team",
                number_sections=True,
                reference_docx=str(reference),
            )
            item = plan_conversions(root, [source], settings)[0]

            cmd = _pandoc_command(root, item, settings)

            self.assertIn("--toc", cmd)
            self.assertIn("--toc-depth=2", cmd)
            self.assertIn("--number-sections", cmd)
            self.assertIn("title=Handbook", cmd)
            self.assertIn("author=Team", cmd)
            self.assertIn("--reference-doc", cmd)
            self.assertIn(str(reference), cmd)

    def test_mermaid_environment_uses_rendering_options(self) -> None:
        env = _mermaid_environment(
            ConvertSettings(
                mermaid_format="svg",
                mermaid_theme="dark",
                mermaid_background="transparent",
            )
        )

        self.assertEqual(env["MERMAID_FILTER_FORMAT"], "svg")
        self.assertEqual(env["MERMAID_FILTER_THEME"], "dark")
        self.assertEqual(env["MERMAID_FILTER_BACKGROUND"], "transparent")

    def test_mermaid_environment_omits_size_options_by_default(self) -> None:
        env = _mermaid_environment(ConvertSettings())

        self.assertNotIn("MERMAID_FILTER_WIDTH", env)
        self.assertNotIn("MERMAID_FILTER_SCALE", env)
        self.assertEqual(env["MERMAID_FILTER_FORMAT"], "png")

    def test_docx_image_paragraphs_are_centered(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            docx_path = Path(tmp) / "image.docx"
            with zipfile.ZipFile(docx_path, "w", zipfile.ZIP_DEFLATED) as docx:
                docx.writestr(
                    "word/document.xml",
                    (
                        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                        "<w:body>"
                        "<w:p><w:r><w:t>Text</w:t></w:r></w:p>"
                        '<w:p><w:pPr><w:jc w:val="left"/></w:pPr><w:r><w:drawing/></w:r></w:p>'
                        "</w:body>"
                        "</w:document>"
                    ),
                )

            _center_docx_images(docx_path)

            with zipfile.ZipFile(docx_path, "r") as docx:
                root = ET.fromstring(docx.read("word/document.xml"))

            namespace = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
            paragraphs = root.findall(".//w:p", namespace)
            self.assertIsNone(paragraphs[0].find("w:pPr/w:jc", namespace))
            self.assertEqual(paragraphs[1].find("w:pPr/w:jc", namespace).get(f"{{{namespace['w']}}}val"), "center")

    def test_generated_reference_docx_patches_font_and_table_borders(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            reference_bytes = _minimal_reference_docx()
            completed = subprocess.CompletedProcess(
                args=["pandoc"],
                returncode=0,
                stdout=reference_bytes,
                stderr=b"",
            )
            settings = ConvertSettings(default_font="Aptos", default_font_size=11, table_borders="bordered")

            with patch("md2doc.converter.subprocess.run", return_value=completed):
                reference = _ensure_generated_reference_docx(root, settings)

            with zipfile.ZipFile(reference, "r") as docx:
                styles = docx.read("word/styles.xml").decode("utf-8")
            self.assertIn('w:ascii="Aptos"', styles)
            self.assertIn('w:val="22"', styles)
            self.assertIn("<w:tblBorders>", styles)

    def test_resolve_pandoc_from_winget_package_when_path_is_stale(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            local_app_data = Path(tmp)
            pandoc = (
                local_app_data
                / "Microsoft"
                / "WinGet"
                / "Packages"
                / "JohnMacFarlane.Pandoc_Microsoft.Winget.Source_8wekyb3d8bbwe"
                / "pandoc-3.10"
                / "pandoc.exe"
            )
            pandoc.parent.mkdir(parents=True)
            pandoc.write_text("", encoding="utf-8")

            with (
                patch.dict(os.environ, {"LOCALAPPDATA": str(local_app_data)}, clear=False),
                patch("md2doc.converter.shutil.which", return_value=None),
                patch("md2doc.converter._windows_registry_tool_locations", return_value=[]),
            ):
                self.assertEqual(_resolve_command("pandoc")[0], str(pandoc))

    def test_resolve_mermaid_filter_from_npm_global_dir_when_path_is_stale(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            app_data = Path(tmp)
            command = app_data / "npm" / "mermaid-filter.cmd"
            command.parent.mkdir(parents=True)
            command.write_text("", encoding="utf-8")

            with (
                patch.dict(os.environ, {"APPDATA": str(app_data)}, clear=False),
                patch("md2doc.converter.shutil.which", return_value=None),
            ):
                self.assertEqual(_resolve_command("mermaid-filter")[0], str(command))

    def test_run_conversions_emits_start_before_result(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / "a.md"
            output_dir = root / "output"
            fake_pandoc = root / "fake_pandoc.py"
            source.write_text("# A", encoding="utf-8")
            fake_pandoc.write_text(
                "\n".join(
                    [
                        "from pathlib import Path",
                        "import sys",
                        "import zipfile",
                        "if '--version' in sys.argv:",
                        "    print('fake pandoc 1.0')",
                        "    raise SystemExit(0)",
                        "output = Path(sys.argv[sys.argv.index('-o') + 1])",
                        "output.parent.mkdir(parents=True, exist_ok=True)",
                        "with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as docx:",
                        "    docx.writestr('word/document.xml', '<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body/></w:document>')",
                    ]
                ),
                encoding="utf-8",
            )
            events: list[str] = []

            results = run_conversions(
                root,
                [source],
                ConvertSettings(
                    output_dir=output_dir,
                    pandoc_cmd=f"python {fake_pandoc}",
                    mermaid_filter_cmd="python",
                ),
                on_start=lambda item: events.append(f"start:{item.relative_source}"),
                on_event=lambda result: events.append(f"{result.status}:{result.item.relative_source}"),
            )

            self.assertEqual([result.status for result in results], ["converted"])
            self.assertEqual(events, ["start:a.md", "converted:a.md"])

    def test_run_conversions_removes_empty_mermaid_filter_error_log_on_success(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / "a.md"
            output_dir = root / "output"
            fake_pandoc = root / "fake_pandoc.py"
            source.write_text("# A", encoding="utf-8")
            fake_pandoc.write_text(
                "\n".join(
                    [
                        "from pathlib import Path",
                        "import sys",
                        "import zipfile",
                        "if '--version' in sys.argv:",
                        "    print('fake pandoc 1.0')",
                        "    raise SystemExit(0)",
                        "Path('mermaid-filter.err').write_text('', encoding='utf-8')",
                        "output = Path(sys.argv[sys.argv.index('-o') + 1])",
                        "output.parent.mkdir(parents=True, exist_ok=True)",
                        "with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as docx:",
                        "    docx.writestr('word/document.xml', '<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body/></w:document>')",
                    ]
                ),
                encoding="utf-8",
            )

            results = run_conversions(
                root,
                [source],
                ConvertSettings(
                    output_dir=output_dir,
                    pandoc_cmd=f"python {fake_pandoc}",
                    mermaid_filter_cmd="python",
                ),
            )

            self.assertEqual([result.status for result in results], ["converted"])
            self.assertFalse((root / "mermaid-filter.err").exists())

    def test_run_conversions_keeps_nonempty_mermaid_filter_error_log_on_failure(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / "a.md"
            output_dir = root / "output"
            fake_pandoc = root / "fake_pandoc.py"
            source.write_text("# A", encoding="utf-8")
            fake_pandoc.write_text(
                "\n".join(
                    [
                        "from pathlib import Path",
                        "import sys",
                        "if '--version' in sys.argv:",
                        "    print('fake pandoc 1.0')",
                        "    raise SystemExit(0)",
                        "Path('mermaid-filter.err').write_text('render failed\\n', encoding='utf-8')",
                        "raise SystemExit(2)",
                    ]
                ),
                encoding="utf-8",
            )

            results = run_conversions(
                root,
                [source],
                ConvertSettings(
                    output_dir=output_dir,
                    pandoc_cmd=f"python {fake_pandoc}",
                    mermaid_filter_cmd="python",
                ),
            )

            err_path = root / "mermaid-filter.err"
            self.assertEqual([result.status for result in results], ["failed"])
            self.assertTrue(err_path.exists())
            self.assertEqual(err_path.read_text(encoding="utf-8"), "render failed\n")
            self.assertIn("mermaid-filter.err:\nrender failed", results[0].message)

class Doc2MdConverterTests(unittest.TestCase):
    def test_scan_source_files_picks_office_documents_for_doc2md(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            (root / "report.docx").write_text("doc", encoding="utf-8")
            (root / "deck.pptx").write_text("ppt", encoding="utf-8")
            (root / "sheet.xlsx").write_text("xls", encoding="utf-8")
            (root / "notes.md").write_text("# Notes", encoding="utf-8")

            files = scan_source_files(root, kind=KIND_DOC2MD)

            self.assertEqual(
                [file.relative_to(root).as_posix() for file in files],
                ["deck.pptx", "report.docx", "sheet.xlsx"],
            )

    def test_plan_emits_markdown_output_for_doc2md(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / "report.docx"
            source.write_text("doc", encoding="utf-8")

            settings = ConvertSettings(kind=KIND_DOC2MD, output_dir=root)
            item = plan_conversions(root, [source], settings)[0]

            self.assertEqual(item.output, root / "report.md")

    def test_markitdown_command_uses_output_flag(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / "report.docx"
            source.write_text("doc", encoding="utf-8")
            settings = ConvertSettings(kind=KIND_DOC2MD, output_dir=root, markitdown_cmd="markitdown")
            item = plan_conversions(root, [source], settings)[0]

            cmd = _markitdown_command(item, settings)

            self.assertEqual(cmd[-3:], [str(item.source), "-o", str(item.output)])

    def test_check_dependencies_uses_markitdown_for_doc2md(self) -> None:
        checks = check_dependencies(ConvertSettings(kind=KIND_DOC2MD))

        self.assertEqual([check.name for check in checks], ["MarkItDown"])

    def test_settings_from_doc2md_project_round_trips_kind(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            project = ProjectConfig(name="Docs", root=Path(tmp), kind=KIND_DOC2MD, output_format="md")
            loaded = ProjectConfig.from_dict(project.to_dict())

            settings = settings_from_project(loaded)

            self.assertEqual(settings.kind, KIND_DOC2MD)
            self.assertEqual(settings.output_suffix(), ".md")


def _minimal_reference_docx() -> bytes:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        path = Path(tmp.name)
    try:
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as docx:
            docx.writestr(
                "word/styles.xml",
                (
                    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                    '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                    '<w:style w:type="paragraph" w:styleId="Normal"><w:rPr/></w:style>'
                    '<w:style w:type="table" w:styleId="Table"><w:tblPr/></w:style>'
                    "</w:styles>"
                ),
            )
        return path.read_bytes()
    finally:
        if path.exists():
            path.unlink()


class Qmd2PptConverterTests(unittest.TestCase):
    def test_scan_source_files_picks_qmd_documents_for_qmd2ppt(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            (root / "presentation.qmd").write_text("content", encoding="utf-8")
            (root / "notes.md").write_text("# Notes", encoding="utf-8")

            files = scan_source_files(root, kind=KIND_QMD2PPT)

            self.assertEqual(
                [file.relative_to(root).as_posix() for file in files],
                ["presentation.qmd"],
            )

    def test_plan_emits_pptx_output_for_qmd2ppt(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / "presentation.qmd"
            source.write_text("content", encoding="utf-8")

            settings = ConvertSettings(kind=KIND_QMD2PPT, output_dir=root)
            item = plan_conversions(root, [source], settings)[0]

            self.assertEqual(item.output, root / "presentation.pptx")

    def test_check_dependencies_uses_quarto_for_qmd2ppt(self) -> None:
        checks = check_dependencies(ConvertSettings(kind=KIND_QMD2PPT))

        self.assertEqual([check.name for check in checks], ["Quarto"])


if __name__ == "__main__":
    unittest.main()
