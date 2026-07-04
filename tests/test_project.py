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

    def test_load_legacy_project_migrates_empty_reference_docx(self) -> None:
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
                    "reference_docx": "",
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

    def test_create_project_kinds(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root_dir = Path(tmp)
            
            # doc2md
            proj_doc2md = create_project(root_dir / "doc2md_proj", kind="doc2md")
            self.assertEqual(proj_doc2md.kind, "doc2md")
            self.assertEqual(proj_doc2md.output_format, "md")
            
            # qmd2ppt
            proj_qmd = create_project(root_dir / "qmd_proj", kind="qmd2ppt")
            self.assertEqual(proj_qmd.kind, "qmd2ppt")
            self.assertEqual(proj_qmd.output_format, "pptx")

    def test_load_project_migrations_for_mermaid_defaults_and_version(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            meta_dir = root / ".md2doc"
            meta_dir.mkdir(parents=True)
            config_path = meta_dir / "project.json"
            
            # Save a config with mermaid_scale = 0.0 and mermaid_min_dpi = -1.0,
            # which should trigger migration.
            config_data = {
                "name": "MigrateMe",
                "root": str(root),
                "kind": "md2doc",
                "config_version": 1,
                "mermaid_scale": 0.0,
                "mermaid_min_dpi": -1.0
            }
            config_path.write_text(json.dumps(config_data), encoding="utf-8")
            
            # load it
            from unittest.mock import patch, MagicMock
            with patch("subprocess.run") as mock_run:
                mock_run.return_value = MagicMock(returncode=0, stdout=b"dummy docx")
                config = load_project(root)
                
                # Check that config version is upgraded, and defaults are migrated
                self.assertEqual(config.config_version, CURRENT_PROJECT_CONFIG_VERSION)
                self.assertEqual(config.mermaid_scale, 3.0)
                self.assertEqual(config.mermaid_min_dpi, 450.0)
                self.assertTrue(config.config_was_migrated)


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

    def test_remove_project(self) -> None:
        from md2doc.project import ProjectRegistry, ProjectConfig
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            registry_file = tmp_path / "projects.json"
            registry = ProjectRegistry(registry_file)
            
            p_a = ProjectConfig(name="A", root=tmp_path / "a")
            p_b = ProjectConfig(name="B", root=tmp_path / "b")
            (tmp_path / "a").mkdir()
            (tmp_path / "b").mkdir()
            
            registry._save([p_a, p_b])
            self.assertEqual(len(registry.list()), 2)
            
            # Remove B
            registry.remove(tmp_path / "b")
            listed = registry.list()
            self.assertEqual(len(listed), 1)
            self.assertEqual(listed[0].name, "A")

    def test_list_with_invalid_or_missing_json(self) -> None:
        from md2doc.project import ProjectRegistry
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            registry_file = tmp_path / "projects.json"
            registry = ProjectRegistry(registry_file)
            
            # Missing file returns empty list
            self.assertEqual(registry.list(), [])
            
            # Malformed JSON file returns empty list
            registry_file.write_text("invalid json", encoding="utf-8")
            self.assertEqual(registry.list(), [])

    def test_list_excludes_nonexistent_roots(self) -> None:
        from md2doc.project import ProjectRegistry, ProjectConfig
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            registry_file = tmp_path / "projects.json"
            registry = ProjectRegistry(registry_file)
            
            p_exists = ProjectConfig(name="Exists", root=tmp_path / "a")
            p_missing = ProjectConfig(name="Missing", root=tmp_path / "missing")
            (tmp_path / "a").mkdir()
            
            registry._save([p_exists, p_missing])
            listed = registry.list()
            self.assertEqual(len(listed), 1)
            self.assertEqual(listed[0].name, "Exists")

    def test_list_skips_malformed_entries(self) -> None:
        from md2doc.project import ProjectRegistry
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            registry_file = tmp_path / "projects.json"
            registry = ProjectRegistry(registry_file)
            
            # Entry lacks "root" which causes KeyError/ValueError inside from_dict
            payload = {
                "projects": [
                    {"name": "Malformed (no root)"},
                    {"name": "Valid", "root": str(tmp_path / "valid")}
                ]
            }
            registry_file.write_text(json.dumps(payload), encoding="utf-8")
            (tmp_path / "valid").mkdir()
            
            listed = registry.list()
            self.assertEqual(len(listed), 1)
            self.assertEqual(listed[0].name, "Valid")


if __name__ == "__main__":
    unittest.main()
