from __future__ import annotations

import json
from pathlib import Path
import tempfile
import unittest

from md2doc.project import (
    CURRENT_PROJECT_CONFIG_VERSION,
    KIND_DOC2MD,
    KIND_HTML2PDF,
    KIND_MD2DOC,
    KIND_QMD2PPT,
    PROJECT_CONFIG_NAME,
    PROJECT_DIR_NAME,
    ProjectConfig,
    create_project,
    load_project,
)


class ProjectKindTests(unittest.TestCase):
    def test_hr_to_pagebreak_serialization(self) -> None:
        config = ProjectConfig.from_dict({"name": "Docs", "root": "/tmp/docs", "hr_to_pagebreak": True})
        self.assertTrue(config.hr_to_pagebreak)
        self.assertTrue(config.to_dict()["hr_to_pagebreak"])

        config_default = ProjectConfig.from_dict({"name": "Docs", "root": "/tmp/docs"})
        self.assertTrue(config_default.hr_to_pagebreak)
        self.assertTrue(config_default.figure_numbering)

    def test_figure_numbering_serialization(self) -> None:
        config = ProjectConfig.from_dict(
            {
                "name": "Docs",
                "root": "/tmp/docs",
                "config_version": CURRENT_PROJECT_CONFIG_VERSION,
                "figure_numbering": True,
                "figure_prefix": "图",
                "figure_caption_position": "above",
            }
        )

        self.assertTrue(config.figure_numbering)
        self.assertEqual(config.figure_prefix, "图")
        self.assertEqual(config.figure_caption_position, "above")
        self.assertEqual(config.config_version, CURRENT_PROJECT_CONFIG_VERSION)
        self.assertTrue(config.to_dict()["figure_numbering"])
        self.assertEqual(config.to_dict()["config_version"], CURRENT_PROJECT_CONFIG_VERSION)

    def test_legacy_config_without_kind_defaults_to_md2doc(self) -> None:
        config = ProjectConfig.from_dict({"name": "Docs", "root": "/tmp/docs", "output_format": "docx"})

        self.assertEqual(config.kind, KIND_MD2DOC)
        self.assertEqual(config.output_format, "docx")

    def test_doc2md_config_forces_markdown_output_format(self) -> None:
        config = ProjectConfig.from_dict(
            {"name": "Docs", "root": "/tmp/docs", "kind": KIND_DOC2MD, "output_format": "docx"}
        )

        self.assertEqual(config.kind, KIND_DOC2MD)
        self.assertEqual(config.output_format, "md")

    def test_unknown_kind_falls_back_to_md2doc(self) -> None:
        config = ProjectConfig.from_dict({"name": "Docs", "root": "/tmp/docs", "kind": "bogus"})

        self.assertEqual(config.kind, KIND_MD2DOC)

    def test_qmd2ppt_config_forces_pptx_output_format(self) -> None:
        config = ProjectConfig.from_dict(
            {"name": "Docs", "root": "/tmp/docs", "kind": KIND_QMD2PPT, "output_format": "docx"}
        )

        self.assertEqual(config.kind, KIND_QMD2PPT)
        self.assertEqual(config.output_format, "pptx")

    def test_html2pdf_config_forces_pdf_output_format(self) -> None:
        config = ProjectConfig.from_dict(
            {"name": "Pages", "root": "/tmp/pages", "kind": KIND_HTML2PDF, "output_format": "docx"}
        )

        self.assertEqual(config.kind, KIND_HTML2PDF)
        self.assertEqual(config.output_format, "pdf")

    def test_html2pdf_settings_serialization(self) -> None:
        config = ProjectConfig.from_dict(
            {
                "name": "Pages",
                "root": "/tmp/pages",
                "kind": KIND_HTML2PDF,
                "html_viewport_width": 1920,
                "html_viewport_height": 1080,
                "html_device_scale_factor": 2.0,
                "html_print_background": False,
                "html_render_delay": 1.5,
            }
        )
        self.assertEqual(config.html_viewport_width, 1920)
        self.assertEqual(config.html_viewport_height, 1080)
        self.assertEqual(config.html_device_scale_factor, 2.0)
        self.assertFalse(config.html_print_background)
        self.assertEqual(config.html_render_delay, 1.5)

        serialized = config.to_dict()
        self.assertEqual(serialized["html_viewport_width"], 1920)
        self.assertEqual(serialized["html_viewport_height"], 1080)
        self.assertEqual(serialized["html_device_scale_factor"], 2.0)
        self.assertFalse(serialized["html_print_background"])
        self.assertEqual(serialized["html_render_delay"], 1.5)

    def test_create_doc2md_project_emits_markdown(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            config = create_project(Path(tmp) / "proj", kind=KIND_DOC2MD)

            self.assertEqual(config.kind, KIND_DOC2MD)
            self.assertEqual(config.output_format, "md")

    def test_create_html2pdf_project_emits_pdf(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            config = create_project(Path(tmp) / "proj", kind=KIND_HTML2PDF)

            self.assertEqual(config.kind, KIND_HTML2PDF)
            self.assertEqual(config.output_format, "pdf")

    def test_create_md2doc_project_sets_default_reference_docx(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            config = create_project(Path(tmp) / "proj", kind=KIND_MD2DOC)

            self.assertEqual(config.kind, KIND_MD2DOC)
            self.assertEqual(config.reference_docx, ".md2doc/reference.docx")
            self.assertTrue((config.root / ".md2doc").exists())

    def test_create_md2doc_project_runs_pandoc_to_generate_template(self) -> None:
        from unittest.mock import patch, MagicMock
        with tempfile.TemporaryDirectory() as tmp:
            with patch("subprocess.run") as mock_run:
                mock_run.return_value = MagicMock(returncode=0, stdout=b"dummy docx content")
                config = create_project(Path(tmp) / "proj", kind=KIND_MD2DOC)
                
                self.assertEqual(config.reference_docx, ".md2doc/reference.docx")
                expected_template_path = config.root / ".md2doc" / "reference.docx"
                self.assertTrue(expected_template_path.exists())
                self.assertEqual(expected_template_path.read_bytes(), b"dummy docx content")
                
                mock_run.assert_called_once()
                args_called = mock_run.call_args[0][0]
                self.assertIn("--print-default-data-file", args_called)
                self.assertIn("reference.docx", args_called)

    def test_load_project_cleans_legacy_config_on_disk(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            meta_dir = root / PROJECT_DIR_NAME
            meta_dir.mkdir(parents=True)
            config_path = meta_dir / PROJECT_CONFIG_NAME
            config_path.write_text(
                json.dumps({"name": "Legacy", "root": str(root), "output_format": "docx"}),
                encoding="utf-8",
            )

            load_project(root)

            cleaned = json.loads(config_path.read_text(encoding="utf-8"))
            self.assertEqual(cleaned["kind"], KIND_MD2DOC)
            self.assertEqual(cleaned["config_version"], CURRENT_PROJECT_CONFIG_VERSION)

    def test_load_project_marks_legacy_config_migration(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            meta_dir = root / PROJECT_DIR_NAME
            meta_dir.mkdir(parents=True)
            config_path = meta_dir / PROJECT_CONFIG_NAME
            config_path.write_text(
                json.dumps({"name": "Legacy", "root": str(root), "output_format": "docx"}),
                encoding="utf-8",
            )

            config = load_project(root)

            self.assertTrue(config.config_was_migrated)
            self.assertEqual(config.loaded_config_version, 1)
            self.assertEqual(config.config_version, CURRENT_PROJECT_CONFIG_VERSION)

    def test_load_legacy_project_migrates_missing_reference_docx(self) -> None:
        from unittest.mock import patch, MagicMock
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            meta_dir = root / PROJECT_DIR_NAME
            meta_dir.mkdir(parents=True)
            config_path = meta_dir / PROJECT_CONFIG_NAME
            config_path.write_text(
                json.dumps({
                    "name": "Legacy",
                    "root": str(root),
                    "kind": KIND_MD2DOC,
                    "output_format": "docx",
                    "config_version": CURRENT_PROJECT_CONFIG_VERSION
                }),
                encoding="utf-8",
            )

            with patch("subprocess.run") as mock_run:
                mock_run.return_value = MagicMock(returncode=0, stdout=b"dummy docx")
                config = load_project(root)

                self.assertEqual(config.reference_docx, ".md2doc/reference.docx")
                expected_template_path = root / ".md2doc" / "reference.docx"
                self.assertTrue(expected_template_path.exists())
                self.assertEqual(expected_template_path.read_bytes(), b"dummy docx")
                
                saved_config = json.loads(config_path.read_text(encoding="utf-8"))
                self.assertEqual(saved_config["reference_docx"], ".md2doc/reference.docx")


class ProjectRegistryTests(unittest.TestCase):
    def test_list_returns_sorted_projects_alphabetically(self) -> None:
        from md2doc.project import ProjectRegistry, ProjectConfig
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            registry_file = tmp_path / "projects.json"
            registry = ProjectRegistry(registry_file)
            
            p_c = ProjectConfig(name="C Project", root=tmp_path / "c")
            p_a = ProjectConfig(name="a Project", root=tmp_path / "a")
            p_b = ProjectConfig(name="B Project", root=tmp_path / "b")
            
            # Create subdirectories to satisfy root.exists() check in registry.list()
            (tmp_path / "c").mkdir()
            (tmp_path / "a").mkdir()
            (tmp_path / "b").mkdir()
            
            # Save in custom order
            registry._save([p_c, p_a, p_b])
            
            # Fetch list, should be sorted alphabetically case-insensitively: a Project -> B Project -> C Project
            listed = registry.list()
            self.assertEqual(len(listed), 3)
            self.assertEqual(listed[0].name, "a Project")
            self.assertEqual(listed[1].name, "B Project")
            self.assertEqual(listed[2].name, "C Project")


if __name__ == "__main__":
    unittest.main()
