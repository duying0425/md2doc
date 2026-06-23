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
    DependencyCheck,
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
            self.assertTrue(any(arg.startswith("--lua-filter=") for arg in cmd))

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
                    pandoc_cmd=f"python {fake_pandoc}",
                    mermaid_filter_cmd="python",
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
                    pandoc_cmd=f"python {fake_pandoc}",
                    mermaid_filter_cmd="python",
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
                        pandoc_cmd=f"python {fake_pandoc}",
                        mermaid_filter_cmd="python",
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
                pandoc_cmd=f"python {fake_pandoc}",
                mermaid_filter_cmd="python",
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
                    quarto_cmd=f"python {fake_quarto}",
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

            def fake_render(_source: Path, output: Path, *, cancel_event=None) -> None:
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


if __name__ == "__main__":
    unittest.main()
