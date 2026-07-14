from __future__ import annotations

import base64
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import unittest
from unittest.mock import patch
import xml.etree.ElementTree as ET
import zipfile

from md2doc.converter import (
    BuildManifest,
    ConvertSettings,
    ConversionResult,
    DependencyCheck,
    FileFingerprint,
    PlanItem,
    _center_docx_images,
    _check_mermaid_browser_runtime,
    _clean_yaml_scalar,
    _decide_action,
    _effective_reference_docx,
    _ensure_figure_caption_lua,
    _ensure_generated_reference_docx,
    _figure_caption_lua_content,
    _generated_reference_signature,
    _is_same_or_child,
    _looks_like_external_reference,
    _markitdown_command,
    _mermaid_environment,
    _missing_quarto_reference_doc_message,
    _needs_generated_reference_docx,
    _pandoc_command,
    _pandoc_format_args,
    _quarto_front_matter_lines,
    _requires_recorded_settings_to_skip,
    _resolve_project_path,
    _shorten_output_path,
    _shorten_windows_path,
    _patch_reference_docx,
    _resolve_command,
    _validate_settings,
    _yaml_comment_index,
    check_dependencies,
    file_fingerprint,
    plan_conversions,
    run_conversions,
    scan_markdown_files,
    scan_source_files,
    settings_from_project,
    settings_signature,
)
from md2doc.project import KIND_DOC2MD, KIND_HTML2PDF, KIND_QMD2PPT, ProjectConfig


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

    def test_scan_markdown_files_excludes_subprojects(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            (root / "a.md").write_text("# A", encoding="utf-8")
            
            # Independent subproject with its own .md2doc folder
            (root / "sub_project").mkdir()
            (root / "sub_project" / ".md2doc").mkdir()
            (root / "sub_project" / "sub.md").write_text("# Sub", encoding="utf-8")
            
            # Normal subfolder (should be scanned)
            (root / "normal_sub").mkdir()
            (root / "normal_sub" / "b.md").write_text("# B", encoding="utf-8")

            files = scan_markdown_files(root, output_dir=root / "output")

            self.assertEqual(
                [file.relative_to(root).as_posix() for file in files],
                ["a.md", "normal_sub/b.md"],
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

    def test_plan_converts_existing_output_without_manifest_when_reference_docx_is_set(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / "a.md"
            output = root / "output" / "a.docx"
            reference = root / "reference.docx"
            source.write_text("# A", encoding="utf-8")
            reference.write_text("template", encoding="utf-8")
            output.parent.mkdir()
            output.write_text("generated", encoding="utf-8")
            os.utime(source, (100, 100))
            os.utime(output, (200, 200))

            item = plan_conversions(
                root,
                [source],
                ConvertSettings(output_dir=root / "output", reference_docx=str(reference)),
                BuildManifest(path=root / ".md2doc" / "manifest.json"),
            )[0]

            self.assertEqual(item.action, "convert")
            self.assertEqual(item.reason, "conversion settings untracked")

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
                mermaid_scale=2.5,
                mermaid_min_dpi=360.0,
                figure_numbering=True,
                figure_prefix="图",
                figure_caption_position="above",
            )

            loaded = ProjectConfig.from_dict(project.to_dict())
            settings = settings_from_project(loaded)

            self.assertTrue(settings.toc)
            self.assertEqual(settings.toc_depth, 2)
            self.assertTrue(settings.title_page)
            self.assertEqual(settings.reference_docx, str(root / "reference.docx"))
            self.assertEqual(settings.table_borders, "bordered")
            self.assertEqual(settings.mermaid_format, "svg")
            self.assertEqual(settings.mermaid_scale, 2.5)
            self.assertEqual(settings.mermaid_min_dpi, 360.0)
            self.assertTrue(settings.figure_numbering)
            self.assertEqual(settings.figure_prefix, "图")
            self.assertEqual(settings.figure_caption_position, "above")

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
                figure_numbering=True,
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
            self.assertTrue(any(arg.startswith("--lua-filter=") for arg in cmd))
            self.assertTrue(any("figure-caption.lua" in arg for arg in cmd))

    def test_shorten_windows_path_leaves_short_paths_untouched(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "short.lua"
            path.write_text("-- x", encoding="utf-8")
            self.assertEqual(_shorten_windows_path(path), str(path))
            self.assertEqual(_shorten_output_path(Path(tmp) / "out.docx"), str(Path(tmp) / "out.docx"))

    @unittest.skipUnless(os.name == "nt", "8.3 short paths are Windows-only")
    def test_pandoc_command_shortens_filter_paths_past_max_path(self) -> None:
        # Pandoc is not long-path aware and fails to open Lua filters/reference
        # docs whose path exceeds MAX_PATH (260). The command must hand Pandoc
        # short (8.3) aliases so a deeply nested project root still converts.
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            while len(str(root / ".md2doc" / "figure-caption.lua")) < 265:
                root = root / "padding_segment_0123456789abcdef"
            root.mkdir(parents=True, exist_ok=True)
            source = root / "doc.md"
            source.write_text("# A", encoding="utf-8")
            settings = ConvertSettings(output_dir=root, figure_numbering=True, hr_to_pagebreak=True)
            item = plan_conversions(root, [source], settings)[0]

            cmd = _pandoc_command(root, item, settings)

            lua_args = [arg[len("--lua-filter="):] for arg in cmd if arg.startswith("--lua-filter=")]
            self.assertTrue(lua_args)
            for filter_path in lua_args:
                self.assertLess(len(filter_path), 260)
                self.assertTrue(Path(filter_path).exists())

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
        self.assertEqual(env["MERMAID_FILTER_SCALE"], "3.0")
        self.assertEqual(env["MERMAID_FILTER_MIN_DPI"], "450.0")
        self.assertEqual(env["MERMAID_FILTER_FORMAT"], "png")

    def test_mermaid_environment_uses_custom_scale(self) -> None:
        env = _mermaid_environment(
            ConvertSettings(
                mermaid_scale=2.5,
                mermaid_min_dpi=360.0,
            )
        )

        self.assertEqual(env["MERMAID_FILTER_SCALE"], "2.5")
        self.assertEqual(env["MERMAID_FILTER_MIN_DPI"], "360.0")

    def test_mermaid_environment_sets_puppeteer_executable_path(self) -> None:
        with (
            patch("md2doc.converter._known_mermaid_managed_browser_path") as mock_browser,
            patch("md2doc.converter._known_system_chromium_browser_path", return_value=Path("/mocked/path/to/edge")),
        ):
            mock_browser.return_value = Path("/mocked/path/to/chrome")
            with patch.dict(os.environ, {}, clear=True):
                env = _mermaid_environment(ConvertSettings())
                self.assertEqual(env.get("PUPPETEER_EXECUTABLE_PATH"), str(Path("/mocked/path/to/chrome")))
                self.assertNotIn("MERMAID_FILTER_PUPPETEER_CONFIG", env)

            with tempfile.TemporaryDirectory() as tmp:
                env_browser = Path(tmp) / "chrome.exe"
                env_browser.write_text("", encoding="utf-8")
                with patch.dict(os.environ, {"PUPPETEER_EXECUTABLE_PATH": str(env_browser)}, clear=True):
                    env = _mermaid_environment(ConvertSettings())
                    self.assertNotIn("PUPPETEER_EXECUTABLE_PATH", env)

            mock_browser.return_value = None
            with patch.dict(os.environ, {}, clear=True):
                env = _mermaid_environment(ConvertSettings())
                self.assertEqual(env.get("PUPPETEER_EXECUTABLE_PATH"), str(Path("/mocked/path/to/edge")))

    def test_check_mermaid_browser_requires_managed_chromium_before_system_browser(self) -> None:
        with (
            patch.dict(os.environ, {}, clear=True),
            patch("md2doc.converter._known_mermaid_managed_browser_path", return_value=None),
            patch("md2doc.converter._known_system_chromium_browser_path", return_value=Path("/mocked/path/to/edge")),
        ):
            check = _check_mermaid_browser_runtime()

        self.assertFalse(check.available)
        self.assertIn("managed Chromium is not installed", check.detail)

    def test_check_mermaid_browser_accepts_managed_chromium(self) -> None:
        with (
            patch.dict(os.environ, {}, clear=True),
            patch("md2doc.converter._known_mermaid_managed_browser_path", return_value=Path("/mocked/path/to/chromium")),
            patch("md2doc.converter._known_system_chromium_browser_path", return_value=Path("/mocked/path/to/edge")),
        ):
            check = _check_mermaid_browser_runtime()

        self.assertTrue(check.available)
        self.assertIn("managed Chromium", check.detail)

    def test_run_conversions_allows_missing_mermaid_browser_without_mermaid_block(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / "plain.md"
            source.write_text("# Plain\n\nNo diagrams here.", encoding="utf-8")
            checks = [
                DependencyCheck("Pandoc", "pandoc", True, "ready"),
                DependencyCheck("mermaid-filter", "mermaid-filter", True, "ready"),
                DependencyCheck("Mermaid browser", "chromium", False, "missing"),
            ]

            def fake_run(_root: Path, item, _settings, *, cancel_event=None) -> ConversionResult:
                return ConversionResult(item=item, status="converted", message="converted", returncode=0)

            with (
                patch("md2doc.converter.check_dependencies", return_value=checks),
                patch("md2doc.converter._run_one", side_effect=fake_run),
            ):
                results = run_conversions(root, [source], ConvertSettings(output_dir=root))

            self.assertEqual(results[0].status, "converted")

    def test_run_conversions_requires_mermaid_browser_for_mermaid_block(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / "diagram.md"
            source.write_text("```{.mermaid}\ngraph TD; A-->B;\n```", encoding="utf-8")
            checks = [
                DependencyCheck("Pandoc", "pandoc", True, "ready"),
                DependencyCheck("mermaid-filter", "mermaid-filter", True, "ready"),
                DependencyCheck("Mermaid browser", "chromium", False, "missing"),
            ]

            with patch("md2doc.converter.check_dependencies", return_value=checks):
                with self.assertRaisesRegex(RuntimeError, "Mermaid browser"):
                    run_conversions(root, [source], ConvertSettings(output_dir=root))

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

    @unittest.skipUnless(shutil.which("pandoc"), "Pandoc is required for DOCX caption integration test")
    def test_figure_caption_lua_emits_word_seq_field(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            png_path = root / "arch.png"
            png_path.write_bytes(
                base64.b64decode(
                    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII="
                )
            )
            md_path = root / "input.md"
            md_path.write_text("![System architecture](arch.png){#fig:arch}", encoding="utf-8")
            output_path = root / "output.docx"
            lua_path = _ensure_figure_caption_lua(root, ConvertSettings(figure_prefix="图"))

            completed = subprocess.run(
                [
                    "pandoc",
                    str(md_path),
                    "-o",
                    str(output_path),
                    f"--lua-filter={lua_path}",
                    f"--resource-path={root}",
                ],
                capture_output=True,
                text=True,
                cwd=str(root),
            )

            self.assertEqual(completed.returncode, 0, completed.stderr)
            with zipfile.ZipFile(output_path) as docx:
                document = docx.read("word/document.xml").decode("utf-8")
            self.assertIn("w:fldSimple", document)
            self.assertIn("SEQ 图", document)
            self.assertIn("System architecture", document)
            self.assertIn('w:pStyle w:val="ImageCaption"', document)

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

    @unittest.skipUnless(shutil.which("pandoc"), "Pandoc is required for DOCX style integration test")
    def test_run_conversions_applies_reference_docx_styles(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / "a.md"
            reference = root / "reference.docx"
            filter_script = root / "pass_filter.py"
            filter_command = root / "pass_filter.cmd"
            source.write_text(
                "\n".join(
                    [
                        "# Heading",
                        "",
                        "Body text.",
                        "",
                        "| A | B |",
                        "|---|---|",
                        "| 1 | 2 |",
                    ]
                ),
                encoding="utf-8",
            )
            filter_script.write_text("import sys\nsys.stdout.write(sys.stdin.read())\n", encoding="utf-8")
            filter_command.write_text(
                f'@echo off\n"{sys.executable}" "{filter_script}" %*\n',
                encoding="utf-8",
            )

            completed = subprocess.run(
                ["pandoc", "--print-default-data-file", "reference.docx"],
                capture_output=True,
                check=True,
            )
            reference.write_bytes(completed.stdout)
            _patch_reference_docx(
                reference,
                ConvertSettings(default_font="Aptos", default_font_size=11, table_borders="bordered"),
            )
            _add_normal_style_color(reference, "C00000")

            results = run_conversions(
                root,
                [source],
                ConvertSettings(
                    output_dir=root,
                    reference_docx=str(reference),
                    mermaid_filter_cmd=str(filter_command),
                ),
            )

            self.assertEqual(results[0].status, "converted")
            with zipfile.ZipFile(root / "a.docx", "r") as docx:
                styles = docx.read("word/styles.xml").decode("utf-8")
            self.assertIn('w:ascii="Aptos"', styles)
            self.assertIn('w:val="22"', styles)
            self.assertIn('w:color w:val="C00000"', styles)
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
                    pandoc_cmd=f'"{sys.executable}" "{fake_pandoc}"',
                    mermaid_filter_cmd=f'"{sys.executable}"',
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
                    pandoc_cmd=f'"{sys.executable}" "{fake_pandoc}"',
                    mermaid_filter_cmd=f'"{sys.executable}"',
                ),
            )

            self.assertEqual([result.status for result in results], ["converted"])
            self.assertFalse((root / "mermaid-filter.err").exists())

    def test_run_conversions_sets_mermaid_filter_loc_to_local_image_dir(self) -> None:
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
                        "import os",
                        "import sys",
                        "import zipfile",
                        "if '--version' in sys.argv:",
                        "    print('fake pandoc 1.0')",
                        "    raise SystemExit(0)",
                        "Path('loc.txt').write_text(os.environ.get('MERMAID_FILTER_LOC', ''), encoding='utf-8')",
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
                    pandoc_cmd=f'"{sys.executable}" "{fake_pandoc}"',
                    mermaid_filter_cmd=f'"{sys.executable}"',
                ),
            )

            self.assertEqual([result.status for result in results], ["converted"])
            mermaid_loc = Path((root / "loc.txt").read_text(encoding="utf-8"))
            self.assertEqual(mermaid_loc.parent, root / ".md2doc" / "mermaid-images")
            self.assertTrue(mermaid_loc.is_dir())

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
                    pandoc_cmd=f'"{sys.executable}" "{fake_pandoc}"',
                    mermaid_filter_cmd=f'"{sys.executable}"',
                ),
            )

            err_path = root / "mermaid-filter.err"
            self.assertEqual([result.status for result in results], ["failed"])
            self.assertTrue(err_path.exists())
            self.assertEqual(err_path.read_text(encoding="utf-8"), "render failed\n")
            self.assertIn("mermaid-filter.err:\nrender failed", results[0].message)

    def test_run_conversions_cancellation(self) -> None:
        from md2doc.converter import ConversionCancelledError
        import threading
        import time

        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / "a.md"
            output_dir = root / "output"
            fake_pandoc = root / "fake_pandoc.py"
            source.write_text("# A", encoding="utf-8")
            
            fake_pandoc.write_text(
                "\n".join(
                    [
                        "import time",
                        "import sys",
                        "if '--version' in sys.argv:",
                        "    print('fake pandoc 1.0')",
                        "    sys.exit(0)",
                        "time.sleep(2.0)",
                    ]
                ),
                encoding="utf-8",
            )
            
            cancel_event = threading.Event()
            
            def trigger_cancel():
                time.sleep(0.3)
                cancel_event.set()
                
            threading.Thread(target=trigger_cancel, daemon=True).start()
            
            start_time = time.time()
            with self.assertRaises(ConversionCancelledError):
                run_conversions(
                    root,
                    [source],
                    ConvertSettings(
                        output_dir=output_dir,
                        pandoc_cmd=f'"{sys.executable}" "{fake_pandoc}"',
                        mermaid_filter_cmd=f'"{sys.executable}"',
                    ),
                    cancel_event=cancel_event,
                )
            duration = time.time() - start_time
            self.assertLess(duration, 1.5)

    def test_run_conversions_updates_manifest_metadata_on_skip_unchanged_with_outdated_mtime(self) -> None:
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
                        "    docx.writestr('word/document.xml', '<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body></w:body></w:document>')",
                    ]
                ),
                encoding="utf-8",
            )

            settings = ConvertSettings(
                output_dir=output_dir,
                pandoc_cmd=f'"{sys.executable}" "{fake_pandoc}"',
                mermaid_filter_cmd=f'"{sys.executable}"',
            )

            results = run_conversions(root, [source], settings)
            self.assertEqual(results[0].status, "converted")

            manifest_path = root / ".md2doc" / "manifest.json"
            self.assertTrue(manifest_path.exists())

            manifest = BuildManifest.load(root)
            old_mtime = manifest.records["a.md"]["source_mtime_ns"]

            # Change the source file's mtime but keep content unchanged
            stat = source.stat()
            os.utime(source, (stat.st_atime + 100, stat.st_mtime + 100))
            new_actual_mtime = source.stat().st_mtime_ns
            self.assertNotEqual(old_mtime, new_actual_mtime)

            results2 = run_conversions(root, [source], settings)
            self.assertEqual(results2[0].status, "skipped")
            self.assertEqual(results2[0].message, "unchanged")

            updated_manifest = BuildManifest.load(root)
            self.assertEqual(updated_manifest.records["a.md"]["source_mtime_ns"], new_actual_mtime)

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


def _add_normal_style_color(docx_path: Path, color: str) -> None:
    namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    w = f"{{{namespace}}}"
    with zipfile.ZipFile(docx_path, "r") as source:
        entries = {name: source.read(name) for name in source.namelist()}

    root = ET.fromstring(entries["word/styles.xml"])
    normal = None
    for style in root.findall(f"{w}style"):
        if style.get(f"{w}type") == "paragraph" and style.get(f"{w}styleId") == "Normal":
            normal = style
            break
    if normal is None:
        normal = ET.SubElement(root, f"{w}style")
        normal.set(f"{w}type", "paragraph")
        normal.set(f"{w}styleId", "Normal")
    rpr = normal.find(f"{w}rPr")
    if rpr is None:
        rpr = ET.SubElement(normal, f"{w}rPr")
    color_el = rpr.find(f"{w}color")
    if color_el is None:
        color_el = ET.SubElement(rpr, f"{w}color")
    color_el.set(f"{w}val", color)
    entries["word/styles.xml"] = ET.tostring(root, encoding="utf-8", xml_declaration=True)

    with zipfile.ZipFile(docx_path, "w", zipfile.ZIP_DEFLATED) as target:
        for name, data in entries.items():
            target.writestr(name, data)


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

    def test_qmd2ppt_reports_missing_reference_doc_before_render(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / "presentation.qmd"
            source.write_text(
                "\n".join(
                    [
                        "---",
                        "format:",
                        "  pptx:",
                        "    reference-doc: missing-template.pptx",
                        "---",
                        "",
                        "## Slide",
                    ]
                ),
                encoding="utf-8",
            )
            fake_quarto = root / "fake_quarto.py"
            fake_quarto.write_text(
                "\n".join(
                    [
                        "from pathlib import Path",
                        "import sys",
                        "if '--version' in sys.argv:",
                        "    print('fake quarto 1.0')",
                        "    raise SystemExit(0)",
                        "Path('render-called.txt').write_text('called', encoding='utf-8')",
                        "raise SystemExit(0)",
                    ]
                ),
                encoding="utf-8",
            )

            results = run_conversions(
                root,
                [source],
                ConvertSettings(
                    kind=KIND_QMD2PPT,
                    output_dir=root,
                    quarto_cmd=f'"{sys.executable}" "{fake_quarto}"',
                ),
            )

            self.assertEqual(results[0].status, "failed")
            self.assertIn("Reference PPTX not found", results[0].message)
            self.assertIn(str(root / "missing-template.pptx"), results[0].message)
            self.assertFalse((root / "render-called.txt").exists())


class Html2PdfConverterTests(unittest.TestCase):
    def test_scan_source_files_picks_html_documents_for_html2pdf(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            (root / "index.html").write_text("<h1>Index</h1>", encoding="utf-8")
            (root / "page.htm").write_text("<h1>Page</h1>", encoding="utf-8")
            (root / "notes.md").write_text("# Notes", encoding="utf-8")

            files = scan_source_files(root, kind=KIND_HTML2PDF)

            self.assertEqual(
                [file.relative_to(root).as_posix() for file in files],
                ["index.html", "page.htm"],
            )

    def test_plan_emits_pdf_output_for_html2pdf(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / "poster.html"
            source.write_text("<main></main>", encoding="utf-8")

            settings = ConvertSettings(kind=KIND_HTML2PDF, output_dir=root)
            item = plan_conversions(root, [source], settings)[0]

            self.assertEqual(item.output, root / "poster.pdf")

    def test_check_dependencies_uses_playwright_for_html2pdf(self) -> None:
        expected = DependencyCheck(
            name="Playwright/Chromium",
            command="playwright",
            available=True,
            detail="ready",
        )

        with patch("md2doc.converter._check_html_pdf_runtime", return_value=expected):
            checks = check_dependencies(ConvertSettings(kind=KIND_HTML2PDF))

        self.assertEqual(checks, [expected])

    def test_run_conversions_uses_html_pdf_renderer(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / "poster.html"
            source.write_text("<main></main>", encoding="utf-8")

            def fake_render(_source: Path, output: Path, *args, **kwargs) -> None:
                output.write_bytes(b"%PDF-1.4\n")

            with (
                patch("md2doc.converter._check_html_pdf_runtime", return_value=DependencyCheck("Playwright/Chromium", "playwright", True, "ready")),
                patch("md2doc.converter._render_html_to_single_page_pdf", side_effect=fake_render) as render,
            ):
                results = run_conversions(
                    root,
                    [source],
                    ConvertSettings(kind=KIND_HTML2PDF, output_dir=root),
                )

            self.assertEqual([result.status for result in results], ["converted"])
            self.assertEqual(results[0].item.output, root / "poster.pdf")
            self.assertTrue((root / "poster.pdf").exists())
            render.assert_called_once()

    def test_run_conversions_passes_custom_html_settings(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / "poster.html"
            source.write_text("<main></main>", encoding="utf-8")

            settings = ConvertSettings(
                kind=KIND_HTML2PDF,
                output_dir=root,
                html_viewport_width=1920,
                html_viewport_height=1080,
                html_device_scale_factor=2.5,
                html_print_background=False,
                html_render_delay=2.0,
            )

            def fake_render(_source: Path, output: Path, s: ConvertSettings, *, cancel_event=None) -> None:
                self.assertEqual(s.html_viewport_width, 1920)
                self.assertEqual(s.html_viewport_height, 1080)
                self.assertEqual(s.html_device_scale_factor, 2.5)
                self.assertFalse(s.html_print_background)
                self.assertEqual(s.html_render_delay, 2.0)
                output.write_bytes(b"%PDF-1.4\n")

            with (
                patch("md2doc.converter._check_html_pdf_runtime", return_value=DependencyCheck("Playwright/Chromium", "playwright", True, "ready")),
                patch("md2doc.converter._render_html_to_single_page_pdf", side_effect=fake_render) as render,
            ):
                results = run_conversions(root, [source], settings)

            self.assertEqual([result.status for result in results], ["converted"])
            render.assert_called_once()

@unittest.skipUnless(shutil.which("pandoc"), "Pandoc is required for Lua filter tests")
class LuaFilterTests(unittest.TestCase):
    def test_png_scaling(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            # Write a PNG with width 1920, height 1080 (aspect ratio 16:9)
            # At 96 DPI, display size is 1920/96 = 20 in, 1080/96 = 11.25 in.
            # Max width is 6.0 in, max height is 8.5 in.
            # 20 in exceeds 6.0 in. Scale factor = 6.0 / 20.0 = 0.3.
            # Height = 11.25 * 0.3 = 3.375 in (which fits <= 8.5 in).
            # So width should be scaled to 6.00in, height to 3.38in.
            png_path = root / "large.png"
            png_bytes = b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR" + (1920).to_bytes(4, 'big') + (1080).to_bytes(4, 'big')
            png_path.write_bytes(png_bytes)

            from md2doc.converter import _ensure_mermaid_fit_lua
            lua_path = _ensure_mermaid_fit_lua(root)

            md_path = root / "input.md"
            md_path.write_text("![image](large.png)", encoding="utf-8")

            cmd = [
                "pandoc",
                str(md_path),
                "--lua-filter",
                str(lua_path),
                "-t",
                "html"
            ]
            res = subprocess.run(cmd, capture_output=True, text=True, cwd=str(root))
            self.assertEqual(res.returncode, 0)
            self.assertIn('style="width:6in;height:3.38in"', res.stdout)

    def test_svg_viewbox_scaling(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            # SVG with viewBox 0 0 1000 2000 (aspect ratio 1:2, very tall image)
            # Display width: 1000/96 = 10.42 in. Display height: 2000/96 = 20.83 in.
            # Height exceeds 8.5 in. Scale factor = 8.5 / 20.83 = 0.408.
            # Width = 10.42 * 0.408 = 4.25 in.
            # So width should be scaled to 4.25in, height to 8.50in.
            svg_path = root / "tall.svg"
            svg_path.write_text('<svg viewBox="0 0 1000 2000"></svg>', encoding="utf-8")

            from md2doc.converter import _ensure_mermaid_fit_lua
            lua_path = _ensure_mermaid_fit_lua(root)

            md_path = root / "input.md"
            md_path.write_text("![image](tall.svg)", encoding="utf-8")

            cmd = [
                "pandoc",
                str(md_path),
                "--lua-filter",
                str(lua_path),
                "-t",
                "html"
            ]
            res = subprocess.run(cmd, capture_output=True, text=True, cwd=str(root))
            self.assertEqual(res.returncode, 0)
            self.assertIn('style="width:4.25in;height:8.5in"', res.stdout)

    def test_url_decoding_and_resource_paths(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            
            sub = root / "sub dir"
            sub.mkdir()
            
            png_bytes = b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR" + (96).to_bytes(4, 'big') + (96).to_bytes(4, 'big')
            (sub / "image file.png").write_bytes(png_bytes)

            from md2doc.converter import _ensure_mermaid_fit_lua
            lua_path = _ensure_mermaid_fit_lua(root)

            md_path = root / "input.md"
            md_path.write_text("![image](image%20file.png)", encoding="utf-8")

            cmd = [
                "pandoc",
                str(md_path),
                "--lua-filter",
                str(lua_path),
                "-t",
                "html"
            ]
            
            env = os.environ.copy()
            env["MD2DOC_RESOURCE_PATHS"] = str(sub)
            
            res = subprocess.run(cmd, capture_output=True, text=True, cwd=str(root), env=env)
            self.assertEqual(res.returncode, 0)
            self.assertIn('style="width:1in;height:1in"', res.stdout)

    def test_mermaid_png_min_dpi_limits_a4_like_image_size(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            image_dir = root / ".md2doc" / "mermaid-images" / "abc"
            image_dir.mkdir(parents=True)
            png_path = image_dir / "a4.png"
            png_bytes = b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR" + (2400).to_bytes(4, 'big') + (3396).to_bytes(4, 'big')
            png_path.write_bytes(png_bytes)

            from md2doc.converter import _ensure_mermaid_fit_lua
            lua_path = _ensure_mermaid_fit_lua(root)

            md_path = root / "input.md"
            md_path.write_text("![image](.md2doc/mermaid-images/abc/a4.png)", encoding="utf-8")

            env = os.environ.copy()
            env["MERMAID_FILTER_SCALE"] = "3"
            env["MERMAID_FILTER_MIN_DPI"] = "600"

            cmd = [
                "pandoc",
                str(md_path),
                "--lua-filter",
                str(lua_path),
                "-t",
                "html"
            ]
            res = subprocess.run(cmd, capture_output=True, text=True, cwd=str(root), env=env)
            self.assertEqual(res.returncode, 0)
            self.assertIn('style="width:4in;height:5.66in"', res.stdout)

    def test_mermaid_svg_ignores_min_dpi(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            image_dir = root / ".md2doc" / "mermaid-images" / "abc"
            image_dir.mkdir(parents=True)
            svg_path = image_dir / "tall.svg"
            svg_path.write_text('<svg viewBox="0 0 1000 2000"></svg>', encoding="utf-8")

            from md2doc.converter import _ensure_mermaid_fit_lua
            lua_path = _ensure_mermaid_fit_lua(root)

            md_path = root / "input.md"
            md_path.write_text("![image](.md2doc/mermaid-images/abc/tall.svg)", encoding="utf-8")

            env = os.environ.copy()
            env["MERMAID_FILTER_SCALE"] = "1"
            env["MERMAID_FILTER_MIN_DPI"] = "1000"

            cmd = [
                "pandoc",
                str(md_path),
                "--lua-filter",
                str(lua_path),
                "-t",
                "html"
            ]
            res = subprocess.run(cmd, capture_output=True, text=True, cwd=str(root), env=env)
            self.assertEqual(res.returncode, 0)
            self.assertIn('style="width:4.25in;height:8.5in"', res.stdout)

    @unittest.skipUnless(shutil.which("pandoc"), "Pandoc is required for hr-to-pagebreak tests")
    def test_hr_to_pagebreak_filter(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            md_path = root / "input.md"
            md_path.write_text("Hello\n\n---\n\nWorld", encoding="utf-8")

            from md2doc.converter import _ensure_hr_to_pagebreak_lua
            lua_path = _ensure_hr_to_pagebreak_lua(root)

            # Test HTML output
            cmd_html = [
                "pandoc",
                str(md_path),
                "--lua-filter",
                str(lua_path),
                "-t",
                "html"
            ]
            res_html = subprocess.run(cmd_html, capture_output=True, text=True, cwd=str(root))
            self.assertEqual(res_html.returncode, 0)
            self.assertIn('<div style="page-break-after: always;"></div>', res_html.stdout)
            self.assertNotIn("<hr", res_html.stdout)

            # Test LaTeX output
            cmd_latex = [
                "pandoc",
                str(md_path),
                "--lua-filter",
                str(lua_path),
                "-t",
                "latex"
            ]
            res_latex = subprocess.run(cmd_latex, capture_output=True, text=True, cwd=str(root))
            self.assertEqual(res_latex.returncode, 0)
            self.assertIn('\\newpage', res_latex.stdout)

    @unittest.skipUnless(shutil.which("pandoc"), "Pandoc is required for DOCX integration test")
    def test_run_conversions_with_hr_to_pagebreak(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / "a.md"
            source.write_text("Hello\n\n---\n\nWorld", encoding="utf-8")

            # Run conversion with hr_to_pagebreak=True to docx
            results = run_conversions(
                root,
                [source],
                ConvertSettings(
                    output_dir=root,
                    hr_to_pagebreak=True,
                ),
            )
            self.assertEqual(results[0].status, "converted")
            
            output_file = results[0].item.output
            self.assertTrue(output_file.exists())
            
            # Verify page break exists in docx zip container
            with zipfile.ZipFile(output_file) as docx:
                document = docx.read("word/document.xml").decode("utf-8")
            self.assertIn('<w:br w:type="page"/>', document)

            # Run conversion with hr_to_pagebreak=False to docx
            results_no_pb = run_conversions(
                root,
                [source],
                ConvertSettings(
                    output_dir=root,
                    hr_to_pagebreak=False,
                ),
            )
            self.assertEqual(results_no_pb[0].status, "converted")
            output_file_no_pb = results_no_pb[0].item.output
            with zipfile.ZipFile(output_file_no_pb) as docx:
                document_no_pb = docx.read("word/document.xml").decode("utf-8")
            self.assertNotIn('<w:br w:type="page"/>', document_no_pb)


class ConverterInternalHelperTests(unittest.TestCase):
    def test_windows_short_path_handles_errors_and_fallbacks(self) -> None:
        from md2doc.converter import _windows_short_path, _shorten_windows_path
        
        # Test non-windows / ctypes import failure fallback
        with patch("sys.platform", "posix"), patch("builtins.__import__", side_effect=ImportError):
            self.assertIsNone(_windows_short_path("/some/long/path"))

        # Test failure where GetShortPathNameW is not found
        class FakeKernel32:
            @property
            def GetShortPathNameW(self):
                raise AttributeError("not found")

        with patch("ctypes.windll.kernel32", FakeKernel32()):
            self.assertIsNone(_windows_short_path("/some/long/path"))

        # Test shorten_windows_path falls back to input string if _windows_short_path returns None
        with patch("md2doc.converter._windows_short_path", return_value=None):
            path = Path("some_file.txt")
            self.assertEqual(_shorten_windows_path(path), str(path))

    def test_pandoc_failure_message_formatting(self) -> None:
        from md2doc.converter import _pandoc_failure_message
        import subprocess

        # Standard stderr and stdout empty, no mermaid error
        proc1 = subprocess.CompletedProcess(args=[], returncode=1, stdout="", stderr="")
        self.assertEqual(_pandoc_failure_message(proc1, ""), "Pandoc failed")

        # Standard stderr set
        proc2 = subprocess.CompletedProcess(args=[], returncode=1, stdout="", stderr="Error output")
        self.assertEqual(_pandoc_failure_message(proc2, ""), "Error output")

        # Stderr, stdout, and mermaid error all set
        proc3 = subprocess.CompletedProcess(args=[], returncode=1, stdout="Stdout output", stderr="Stderr output")
        self.assertEqual(
            _pandoc_failure_message(proc3, "Mermaid drawing failed"),
            "Stderr output\n\nStdout output\n\nmermaid-filter.err:\nMermaid drawing failed"
        )

    def test_lua_string_literal_serialization(self) -> None:
        from md2doc.converter import _lua_string_literal

        self.assertEqual(_lua_string_literal("hello"), '"hello"')
        self.assertEqual(_lua_string_literal("hello \"world\""), '"hello \\"world\\""')
        # Check unicode handling (ensure_ascii=False)
        self.assertEqual(_lua_string_literal("中文"), '"中文"')

    def test_remove_file_suppresses_exceptions(self) -> None:
        from md2doc.converter import _remove_file

        # Removing non-existent file should not raise FileNotFoundError
        _remove_file(Path("non_existent_file_xyz_123.txt"))

        # Check OSError handling
        mock_path = unittest.mock.MagicMock(spec=Path)
        mock_path.unlink.side_effect = OSError("Permission denied")
        _remove_file(mock_path)  # Should not raise

    def test_should_use_markitdown_api(self) -> None:
        from md2doc.converter import _should_use_markitdown_api

        # Not markitdown command
        self.assertFalse(_should_use_markitdown_api("pandoc"))
        self.assertFalse(_should_use_markitdown_api(""))

        # Test command exists in PATH -> should return False because command is available directly
        with patch("md2doc.converter._command_exists", return_value=True):
            self.assertFalse(_should_use_markitdown_api("markitdown"))

        # Test command does not exist but API available -> should return True
        with (
            patch("md2doc.converter._command_exists", return_value=False),
            patch("md2doc.converter._markitdown_api_available", return_value=True)
        ):
            self.assertTrue(_should_use_markitdown_api("markitdown"))

        # Test command does not exist and API not available -> should return False
        with (
            patch("md2doc.converter._command_exists", return_value=False),
            patch("md2doc.converter._markitdown_api_available", return_value=False)
        ):
            self.assertFalse(_should_use_markitdown_api("markitdown"))

    def test_check_markitdown_dependency(self) -> None:
        from md2doc.converter import _check_markitdown

        # 1. Command exists
        with patch("md2doc.converter._command_exists", return_value=True):
            check = _check_markitdown("markitdown")
            self.assertTrue(check.available)
            self.assertIn("found at", check.detail)

        # 2. Command does not exist, but API is available
        with (
            patch("md2doc.converter._command_exists", return_value=False),
            patch("md2doc.converter._should_use_markitdown_api", return_value=True)
        ):
            check = _check_markitdown("markitdown")
            self.assertTrue(check.available)
            self.assertIn("available through bundled Python package", check.detail)

        # 3. Neither command nor API is available
        with (
            patch("md2doc.converter._command_exists", return_value=False),
            patch("md2doc.converter._should_use_markitdown_api", return_value=False)
        ):
            check = _check_markitdown("markitdown")
            self.assertFalse(check.available)
            self.assertIn("was not found", check.detail)


class DecideActionTests(unittest.TestCase):
    """Cover the _decide_action planner decision table and its reason strings."""

    def _fingerprint(self) -> FileFingerprint:
        return FileFingerprint(size=10, mtime_ns=100, sha256="abc")

    def _output(self, root: Path, *, exists: bool = False, mtime_ns: int = 50) -> Path:
        output = root / "out.docx"
        output.parent.mkdir(parents=True, exist_ok=True)
        if exists:
            output.write_text("x", encoding="utf-8")
            os.utime(output, (100, mtime_ns / 1e9))
        return output

    def test_force_overrides_everything(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            output = self._output(Path(tmp), exists=True, mtime_ns=200)
            action, reason = _decide_action(
                settings=ConvertSettings(force=True),
                record={"source_sha256": "abc", "settings_signature": "sig"},
                fingerprint=self._fingerprint(),
                output=output,
                signature="sig",
            )
            self.assertEqual((action, reason), ("convert", "forced"))

    def test_skip_disabled_converts_even_when_output_matches(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            output = self._output(Path(tmp), exists=True, mtime_ns=200)
            action, reason = _decide_action(
                settings=ConvertSettings(skip_unchanged=False),
                record={"source_sha256": "abc", "settings_signature": "sig"},
                fingerprint=self._fingerprint(),
                output=output,
                signature="sig",
            )
            self.assertEqual((action, reason), ("convert", "skip disabled"))

    def test_output_missing_converts(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            output = self._output(Path(tmp), exists=False)
            action, reason = _decide_action(
                settings=ConvertSettings(),
                record=None,
                fingerprint=self._fingerprint(),
                output=output,
                signature="sig",
            )
            self.assertEqual((action, reason), ("convert", "output missing"))

    def test_no_record_output_newer_skips(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            output = self._output(Path(tmp), exists=True, mtime_ns=200)
            action, reason = _decide_action(
                settings=ConvertSettings(),
                record=None,
                fingerprint=FileFingerprint(size=10, mtime_ns=100, sha256="abc"),
                output=output,
                signature="sig",
            )
            self.assertEqual((action, reason), ("skip", "output is newer than source"))

    def test_no_record_source_newer_converts(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            output = self._output(Path(tmp), exists=True, mtime_ns=50)
            action, reason = _decide_action(
                settings=ConvertSettings(),
                record=None,
                fingerprint=FileFingerprint(size=10, mtime_ns=100, sha256="abc"),
                output=output,
                signature="sig",
            )
            self.assertEqual((action, reason), ("convert", "no history and source is newer"))

    def test_record_settings_change_converts(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            output = self._output(Path(tmp), exists=True, mtime_ns=200)
            action, reason = _decide_action(
                settings=ConvertSettings(),
                record={"source_sha256": "abc", "settings_signature": "old"},
                fingerprint=self._fingerprint(),
                output=output,
                signature="new",
            )
            self.assertEqual((action, reason), ("convert", "conversion settings changed"))

    def test_unchanged_skips(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            output = self._output(Path(tmp), exists=True, mtime_ns=200)
            action, reason = _decide_action(
                settings=ConvertSettings(),
                record={"source_sha256": "abc", "settings_signature": "sig"},
                fingerprint=self._fingerprint(),
                output=output,
                signature="sig",
            )
            self.assertEqual((action, reason), ("skip", "unchanged"))

    def test_source_hash_change_converts(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            output = self._output(Path(tmp), exists=True, mtime_ns=200)
            action, reason = _decide_action(
                settings=ConvertSettings(),
                record={"source_sha256": "old", "settings_signature": "sig"},
                fingerprint=self._fingerprint(),
                output=output,
                signature="sig",
            )
            self.assertEqual((action, reason), ("convert", "source changed"))

    def test_requires_recorded_settings_to_skip_detects_styled_settings(self) -> None:
        self.assertTrue(_requires_recorded_settings_to_skip(ConvertSettings(toc=True)))
        self.assertTrue(_requires_recorded_settings_to_skip(ConvertSettings(reference_docx="r.docx")))
        self.assertTrue(_requires_recorded_settings_to_skip(ConvertSettings(table_borders="bordered")))
        self.assertTrue(_requires_recorded_settings_to_skip(ConvertSettings(mermaid_format="svg")))
        self.assertFalse(_requires_recorded_settings_to_skip(ConvertSettings()))


class ValidateSettingsTests(unittest.TestCase):
    def test_validate_settings_passes_for_doc2md_without_reference(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            _validate_settings(Path(tmp), ConvertSettings(kind=KIND_DOC2MD))

    def test_validate_settings_rejects_bad_table_borders(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            with self.assertRaisesRegex(RuntimeError, "Table borders"):
                _validate_settings(Path(tmp), ConvertSettings(table_borders="weird"))

    def test_validate_settings_rejects_bad_mermaid_format(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            with self.assertRaisesRegex(RuntimeError, "Mermaid format"):
                _validate_settings(Path(tmp), ConvertSettings(mermaid_format="gif"))

    def test_validate_settings_rejects_bad_figure_caption_position(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            with self.assertRaisesRegex(RuntimeError, "Figure caption position"):
                _validate_settings(Path(tmp), ConvertSettings(figure_caption_position="side"))

    def test_validate_settings_rejects_missing_reference_docx(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            with self.assertRaisesRegex(RuntimeError, "Reference DOCX not found"):
                _validate_settings(Path(tmp), ConvertSettings(reference_docx="missing.docx"))

    def test_validate_settings_regenerates_default_reference_docx(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            # The default .md2doc/reference.docx template is regenerated on demand.
            project_root = Path(tmp)
            settings = ConvertSettings(reference_docx=".md2doc/reference.docx")

            def _fake_generate(root: Path, cfg: ConvertSettings) -> None:
                reference = root / ".md2doc" / "reference.docx"
                reference.parent.mkdir(parents=True, exist_ok=True)
                reference.write_bytes(b"PK\x03\x04")

            with patch(
                "md2doc.converter._generate_default_reference_docx_if_needed",
                side_effect=_fake_generate,
            ) as gen:
                _validate_settings(project_root, settings)
            gen.assert_called_once()
            self.assertTrue((project_root / ".md2doc" / "reference.docx").exists())


class EffectiveReferenceDocxTests(unittest.TestCase):
    def test_returns_none_for_non_docx_output(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            result = _effective_reference_docx(Path(tmp), ConvertSettings(output_format="pptx"))
            self.assertIsNone(result)

    def test_returns_none_when_no_reference_and_no_styling(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            result = _effective_reference_docx(Path(tmp), ConvertSettings(reference_docx=""))
            self.assertIsNone(result)

    def test_returns_explicit_reference_path(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            ref = Path(tmp) / "ref.docx"
            ref.write_text("x", encoding="utf-8")
            result = _effective_reference_docx(Path(tmp), ConvertSettings(reference_docx=str(ref)))
            self.assertEqual(result, ref.resolve())

    def test_generates_reference_when_font_or_borders_configured(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            with patch("md2doc.converter._ensure_generated_reference_docx", return_value=Path(tmp) / "gen.docx") as gen:
                result = _effective_reference_docx(
                    Path(tmp),
                    ConvertSettings(default_font="Aptos", table_borders="bordered"),
                )
            self.assertEqual(result, Path(tmp) / "gen.docx")
            gen.assert_called_once()

    def test_needs_generated_reference_docx_detects_styling(self) -> None:
        self.assertTrue(_needs_generated_reference_docx(ConvertSettings(default_font="Aptos")))
        self.assertTrue(_needs_generated_reference_docx(ConvertSettings(default_font_size=12)))
        self.assertTrue(_needs_generated_reference_docx(ConvertSettings(table_borders="plain")))
        self.assertFalse(_needs_generated_reference_docx(ConvertSettings()))

    def test_generated_reference_signature_changes_with_styling(self) -> None:
        sig_a = _generated_reference_signature(ConvertSettings(default_font="Aptos", default_font_size=11, table_borders="bordered"))
        sig_b = _generated_reference_signature(ConvertSettings(default_font="Calibri", default_font_size=11, table_borders="bordered"))
        sig_c = _generated_reference_signature(ConvertSettings(default_font="Aptos", default_font_size=11, table_borders="bordered"))
        self.assertNotEqual(sig_a, sig_b)
        self.assertEqual(sig_a, sig_c)


class QuartoFrontMatterTests(unittest.TestCase):
    def _write(self, path: Path, content: str) -> None:
        path.write_text(content, encoding="utf-8")

    def test_parses_yaml_front_matter(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            qmd = Path(tmp) / "doc.qmd"
            self._write(qmd, "---\ntitle: Hello\nreference-doc: template.pptx\n---\n\n# Slide\n")
            lines = _quarto_front_matter_lines(qmd)
            self.assertEqual(lines, ["title: Hello", "reference-doc: template.pptx"])

    def test_returns_empty_when_no_front_matter(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            qmd = Path(tmp) / "doc.qmd"
            self._write(qmd, "# No front matter\n\nBody\n")
            self.assertEqual(_quarto_front_matter_lines(qmd), [])

    def test_returns_empty_for_unclosed_front_matter(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            qmd = Path(tmp) / "doc.qmd"
            self._write(qmd, "---\ntitle: Unclosed\n\n# Body\n")
            self.assertEqual(_quarto_front_matter_lines(qmd), [])

    def test_stops_at_three_dots_close(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            qmd = Path(tmp) / "doc.qmd"
            self._write(qmd, "---\ntitle: Dots\n...\n\n# Body\n")
            self.assertEqual(_quarto_front_matter_lines(qmd), ["title: Dots"])

    def test_handles_utf8_bom(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            qmd = Path(tmp) / "doc.qmd"
            qmd.write_bytes(b"\xef\xbb\xbf---\ntitle: Bom\n---\n\n# Body\n")
            self.assertEqual(_quarto_front_matter_lines(qmd), ["title: Bom"])


class YamlScalarTests(unittest.TestCase):
    def test_strips_unquoted_value(self) -> None:
        self.assertEqual(_clean_yaml_scalar("template.pptx"), "template.pptx")

    def test_strips_whitespace_around_value(self) -> None:
        self.assertEqual(_clean_yaml_scalar("  template.pptx  "), "template.pptx")

    def test_strips_single_quotes(self) -> None:
        self.assertEqual(_clean_yaml_scalar("'my template.pptx'"), "my template.pptx")

    def test_strips_double_quotes(self) -> None:
        self.assertEqual(_clean_yaml_scalar('"my template.pptx"'), "my template.pptx")

    def test_handles_unclosed_quote(self) -> None:
        self.assertEqual(_clean_yaml_scalar("'unclosed"), "unclosed")

    def test_strips_trailing_comment(self) -> None:
        self.assertEqual(_clean_yaml_scalar("template.pptx # inline comment"), "template.pptx")

    def test_preserves_hash_inside_value(self) -> None:
        self.assertEqual(_clean_yaml_scalar("a#b.pptx"), "a#b.pptx")

    def test_returns_empty_for_empty_input(self) -> None:
        self.assertEqual(_clean_yaml_scalar(""), "")
        self.assertEqual(_clean_yaml_scalar("   "), "")


class YamlCommentIndexTests(unittest.TestCase):
    def test_returns_negative_one_when_no_comment(self) -> None:
        self.assertEqual(_yaml_comment_index("template.pptx"), -1)

    def test_finds_leading_hash(self) -> None:
        self.assertEqual(_yaml_comment_index("# comment"), 0)

    def test_finds_hash_after_whitespace(self) -> None:
        self.assertEqual(_yaml_comment_index("value # comment"), 6)

    def test_ignores_hash_inside_word(self) -> None:
        self.assertEqual(_yaml_comment_index("a#b"), -1)


class ExternalReferenceTests(unittest.TestCase):
    def test_http_url_is_external(self) -> None:
        self.assertTrue(_looks_like_external_reference("http://example.com/template.pptx"))

    def test_https_url_is_external(self) -> None:
        self.assertTrue(_looks_like_external_reference("https://example.com/template.pptx"))

    def test_data_uri_is_external(self) -> None:
        self.assertTrue(_looks_like_external_reference("data:application/vnd.openxmlformats-officedocument.presentationml.presentation;base64,..."))

    def test_local_path_is_not_external(self) -> None:
        self.assertFalse(_looks_like_external_reference("template.pptx"))
        self.assertFalse(_looks_like_external_reference("./templates/x.pptx"))
        self.assertFalse(_looks_like_external_reference("C:\\templates\\x.pptx"))


class MissingQuartoReferenceTests(unittest.TestCase):
    def test_returns_empty_string_when_reference_exists(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            qmd = Path(tmp) / "doc.qmd"
            ref = Path(tmp) / "template.pptx"
            ref.write_text("x", encoding="utf-8")
            qmd.write_text(f"---\nreference-doc: template.pptx\n---\n\n# Slide\n", encoding="utf-8")
            self.assertEqual(_missing_quarto_reference_doc_message(qmd), "")

    def test_returns_empty_string_when_no_reference_doc(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            qmd = Path(tmp) / "doc.qmd"
            qmd.write_text("---\ntitle: Hello\n---\n\n# Slide\n", encoding="utf-8")
            self.assertEqual(_missing_quarto_reference_doc_message(qmd), "")

    def test_reports_missing_reference_with_expected_path(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            qmd = Path(tmp) / "doc.qmd"
            qmd.write_text("---\nreference-doc: missing.pptx\n---\n\n# Slide\n", encoding="utf-8")
            message = _missing_quarto_reference_doc_message(qmd)
            self.assertIn("Reference PPTX not found", message)
            self.assertIn("missing.pptx", message)
            self.assertIn(str((qmd.parent / "missing.pptx").resolve()), message)

    def test_ignores_external_urls(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            qmd = Path(tmp) / "doc.qmd"
            qmd.write_text("---\nreference-doc: https://example.com/x.pptx\n---\n\n# Slide\n", encoding="utf-8")
            self.assertEqual(_missing_quarto_reference_doc_message(qmd), "")


class ResolveProjectPathTests(unittest.TestCase):
    def test_resolves_relative_path_against_project_root(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            result = _resolve_project_path(Path(tmp), "templates/ref.docx")
            self.assertEqual(result, (Path(tmp) / "templates" / "ref.docx").resolve())

    def test_resolves_absolute_path_unchanged(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            absolute = Path(tmp) / "ref.docx"
            result = _resolve_project_path(Path(tmp) / "other", str(absolute))
            self.assertEqual(result, absolute.resolve())

    def test_expands_user_path(self) -> None:
        result = _resolve_project_path(Path("/tmp"), "~/ref.docx")
        self.assertTrue(str(result).endswith("ref.docx"))


class IsSameOrChildTests(unittest.TestCase):
    def test_child_path_returns_true(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            parent = Path(tmp).resolve()
            child = parent / "sub"
            child.mkdir()
            self.assertTrue(_is_same_or_child(child.resolve(), parent))

    def test_same_path_returns_true(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            parent = Path(tmp).resolve()
            self.assertTrue(_is_same_or_child(parent, parent))

    def test_unrelated_path_returns_false(self) -> None:
        with tempfile.TemporaryDirectory() as tmp1, tempfile.TemporaryDirectory() as tmp2:
            self.assertFalse(_is_same_or_child(Path(tmp1).resolve(), Path(tmp2).resolve()))

    def test_none_parent_returns_false(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            self.assertFalse(_is_same_or_child(Path(tmp).resolve(), None))


class CenterDocxImagesTests(unittest.TestCase):
    def _make_docx(self, path: Path, *, with_image: bool) -> None:
        drawing = (
            '<w:p><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:extent cx="100" cy="100"/></wp:inline></w:drawing></w:r></w:p>'
            if with_image
            else '<w:p><w:r><w:t>plain text</w:t></w:r></w:p>'
        )
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as docx:
            docx.writestr(
                "word/document.xml",
                f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                f'<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                f'<w:body>{drawing}</w:body></w:document>',
            )

    def test_centers_image_paragraph(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            docx_path = Path(tmp) / "out.docx"
            self._make_docx(docx_path, with_image=True)
            _center_docx_images(docx_path)
            with zipfile.ZipFile(docx_path) as z:
                doc = z.read("word/document.xml").decode("utf-8")
            self.assertIn('w:jc w:val="center"', doc)

    def test_leaves_text_paragraph_untouched(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            docx_path = Path(tmp) / "out.docx"
            self._make_docx(docx_path, with_image=False)
            _center_docx_images(docx_path)
            with zipfile.ZipFile(docx_path) as z:
                doc = z.read("word/document.xml").decode("utf-8")
            self.assertNotIn("w:jc", doc)

    def test_no_document_xml_is_noop(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            docx_path = Path(tmp) / "out.docx"
            with zipfile.ZipFile(docx_path, "w", zipfile.ZIP_DEFLATED) as z:
                z.writestr("foo.txt", "bar")
            _center_docx_images(docx_path)  # should not raise


class PandocFormatArgsTests(unittest.TestCase):
    def _item(self, root: Path) -> PlanItem:
        source = root / "a.md"
        source.write_text("# A", encoding="utf-8")
        return plan_conversions(root, [source], ConvertSettings(output_dir=root))[0]

    def test_toc_and_toc_depth_emitted(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            args = _pandoc_format_args(Path(tmp), self._item(Path(tmp)), ConvertSettings(toc=True, toc_depth=4))
            self.assertIn("--toc", args)
            self.assertIn("--toc-depth=4", args)

    def test_number_sections_emitted(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            args = _pandoc_format_args(Path(tmp), self._item(Path(tmp)), ConvertSettings(number_sections=True))
            self.assertIn("--number-sections", args)

    def test_title_page_emits_metadata_with_fallback_to_filename(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            args = _pandoc_format_args(
                Path(tmp),
                self._item(Path(tmp)),
                ConvertSettings(title_page=True, title="", subtitle="Sub", author="Author", date="2026-07-14"),
            )
            self.assertIn("title=a", args)
            self.assertIn("subtitle=Sub", args)
            self.assertIn("author=Author", args)
            self.assertIn("date=2026-07-14", args)

    def test_reference_docx_emitted_when_set(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            ref = Path(tmp) / "ref.docx"
            ref.write_text("x", encoding="utf-8")
            args = _pandoc_format_args(Path(tmp), self._item(Path(tmp)), ConvertSettings(reference_docx=str(ref)))
            self.assertIn("--reference-doc", args)
            self.assertIn(str(ref), args)

    def test_no_args_for_plain_settings(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            args = _pandoc_format_args(Path(tmp), self._item(Path(tmp)), ConvertSettings())
            self.assertEqual(args, [])


class FigureCaptionLuaContentTests(unittest.TestCase):
    def test_substitutes_prefix_and_position(self) -> None:
        content = _figure_caption_lua_content(ConvertSettings(figure_prefix="Figure", figure_caption_position="above"))
        self.assertIn('local configured_prefix = "Figure"', content)
        self.assertIn('local configured_caption_position = "above"', content)

    def test_falls_back_to_default_when_empty(self) -> None:
        content = _figure_caption_lua_content(ConvertSettings(figure_prefix="", figure_caption_position=""))
        self.assertIn('local configured_prefix = "Figure"', content)
        self.assertIn('local configured_caption_position = "below"', content)

    def test_escapes_quotes_in_prefix(self) -> None:
        content = _figure_caption_lua_content(ConvertSettings(figure_prefix='hello "world"'))
        self.assertIn('local configured_prefix = "hello \\"world\\""', content)


if __name__ == "__main__":
    unittest.main()
