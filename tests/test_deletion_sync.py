from __future__ import annotations

import json
from pathlib import Path
import tempfile
import unittest

from md2doc.cli import main as cli_main
from md2doc.converter import (
    BuildManifest,
    ConvertSettings,
    PlanItem,
    clean_orphans,
    plan_conversions,
    run_conversions,
)


class DeletionSyncBehaviorTests(unittest.TestCase):
    def test_current_behavior_default_sync_deletes_false(self) -> None:
        """测试默认 sync_deletes=False 时，MD 文件被删除后 Word 文件及 Manifest 记录保留不被清理。"""
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            md_file = root / "doc.md"
            docx_file = root / "doc.docx"
            
            md_file.write_text("# Test Document", encoding="utf-8")
            docx_file.write_text("dummy docx content", encoding="utf-8")

            manifest = BuildManifest.load(root)
            manifest.records["doc.md"] = {
                "source_sha256": "abc123hash",
                "source_size": 15,
                "source_mtime_ns": 1000,
                "output": str(docx_file),
                "output_format": "docx",
                "settings_signature": "sig123",
                "converted_at": "2026-08-03T15:00:00Z",
            }
            manifest.save()

            # 删除 MD 文件
            md_file.unlink()

            # 使用默认配置 (sync_deletes=False) 规划转换
            settings = ConvertSettings(sync_deletes=False)
            sources = [p for p in root.iterdir() if p.suffix == ".md"]
            plans = plan_conversions(root, sources, settings, manifest)

            self.assertEqual(len(plans), 0)
            self.assertTrue(docx_file.exists())

            reloaded_manifest = BuildManifest.load(root)
            self.assertIn("doc.md", reloaded_manifest.records)

    def test_sync_deletes_enabled_deletes_orphan_docx_and_manifest_record(self) -> None:
        """测试开启 sync_deletes=True 时，自动生成 delete 计划并清理孤儿 Word 和 Manifest。"""
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            md_file = root / "doc.md"
            docx_file = root / "doc.docx"
            
            docx_file.write_text("dummy docx content", encoding="utf-8")

            manifest = BuildManifest.load(root)
            manifest.records["doc.md"] = {
                "source_sha256": "abc123hash",
                "output": str(docx_file),
                "output_format": "docx",
            }
            manifest.save()

            settings = ConvertSettings(sync_deletes=True)
            sources = [p for p in root.iterdir() if p.suffix == ".md"]
            plans = plan_conversions(root, sources, settings, manifest)

            self.assertEqual(len(plans), 1)
            self.assertEqual(plans[0].action, "delete")
            self.assertEqual(plans[0].relative_source, "doc.md")

            results = run_conversions(root, sources, settings)
            self.assertEqual(len(results), 1)
            self.assertEqual(results[0].status, "converted")

            self.assertFalse(docx_file.exists())
            reloaded_manifest = BuildManifest.load(root)
            self.assertNotIn("doc.md", reloaded_manifest.records)

    def test_clean_orphans_function_and_cli(self) -> None:
        """测试 clean_orphans 函数及 md2doc clean 命令行接口。"""
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            # 初始化项目
            cli_main(["init", str(root)])

            docx_file = root / "orphan.docx"
            docx_file.write_text("orphan docx content", encoding="utf-8")

            manifest = BuildManifest.load(root)
            manifest.records["orphan.md"] = {
                "source_sha256": "abc123hash",
                "output": str(docx_file),
                "output_format": "docx",
            }
            manifest.save()

            # 测试 dry-run
            dry_results = clean_orphans(root, dry_run=True)
            self.assertEqual(len(dry_results), 1)
            self.assertTrue(docx_file.exists())

            # CLI clean 测试
            exit_code = cli_main(["clean", str(root)])
            self.assertEqual(exit_code, 0)
            self.assertFalse(docx_file.exists())

            reloaded_manifest = BuildManifest.load(root)
            self.assertNotIn("orphan.md", reloaded_manifest.records)

    def test_sync_deletes_when_deleted_source_in_sources_list(self) -> None:
        """测试当显式传入包含已删除 md 文件的 sources 列表时，plan_conversions 和 run_conversions 不抛出 FileNotFoundError 并正确执行清理。"""
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            md_file = root / "deleted.md"
            docx_file = root / "deleted.docx"
            docx_file.write_text("dummy content", encoding="utf-8")

            manifest = BuildManifest.load(root)
            manifest.records["deleted.md"] = {
                "source_sha256": "abc123hash",
                "output": str(docx_file),
                "output_format": "docx",
            }
            manifest.save()

            settings = ConvertSettings(sync_deletes=True)
            # 模拟 UI 或全量转换时，sources 列表中包含非空 Path（如已删除的 md_file）
            sources = [md_file]

            plans = plan_conversions(root, sources, settings, manifest)
            self.assertEqual(len(plans), 1)
            self.assertEqual(plans[0].action, "delete")
            self.assertEqual(plans[0].relative_source, "deleted.md")

            results = run_conversions(root, sources, settings)
            self.assertEqual(len(results), 1)
            self.assertEqual(results[0].status, "converted")
            self.assertFalse(docx_file.exists())


if __name__ == "__main__":
    unittest.main()

