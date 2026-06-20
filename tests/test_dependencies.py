from __future__ import annotations

import subprocess
import unittest
from unittest.mock import patch

from md2doc.converter import DependencyCheck
from md2doc.dependencies import ensure_startup_dependencies


def _check(name: str, available: bool) -> DependencyCheck:
    return DependencyCheck(name=name, command=name.lower(), available=available, detail=name)


class DependencySetupTests(unittest.TestCase):
    def test_ready_dependencies_do_not_run_installers(self) -> None:
        calls: list[list[str]] = []

        with (
            patch(
                "md2doc.dependencies.check_dependencies",
                return_value=[_check("Pandoc", True), _check("mermaid-filter", True)],
            ),
            patch("md2doc.dependencies._refresh_windows_path"),
        ):
            ensure_startup_dependencies(runner=lambda args: calls.append(list(args)) or _ok(args))

        self.assertEqual(calls, [])

    def test_missing_default_toolchain_installs_with_winget_and_npm(self) -> None:
        calls: list[list[str]] = []
        progress: list[str] = []
        missing = [_check("Pandoc", False), _check("mermaid-filter", False)]
        ready = [_check("Pandoc", True), _check("mermaid-filter", True)]

        with (
            patch("md2doc.dependencies.os.name", "nt"),
            patch("md2doc.dependencies.check_dependencies", side_effect=[missing, ready]),
            patch("md2doc.dependencies._tool_available", side_effect=lambda command: False),
            patch("md2doc.dependencies._resolve_winget", return_value="winget"),
            patch("md2doc.dependencies._resolve_npm", return_value="npm"),
            patch("md2doc.dependencies._refresh_windows_path"),
        ):
            ensure_startup_dependencies(
                on_progress=progress.append,
                runner=lambda args: calls.append(list(args)) or _ok(args),
            )

        self.assertEqual(calls[0][:5], ["winget", "install", "--id", "JohnMacFarlane.Pandoc", "--exact"])
        self.assertEqual(calls[1][:5], ["winget", "install", "--id", "OpenJS.NodeJS.LTS", "--exact"])
        self.assertEqual(calls[2], ["npm", "install", "-g", "mermaid-filter"])
        self.assertIn("Dependency setup completed.", progress)

    def test_winget_missing_reports_actionable_error(self) -> None:
        missing = [_check("Pandoc", False), _check("mermaid-filter", True)]

        with (
            patch("md2doc.dependencies.os.name", "nt"),
            patch("md2doc.dependencies.check_dependencies", return_value=missing),
            patch("md2doc.dependencies._tool_available", return_value=False),
            patch("md2doc.dependencies._resolve_winget", return_value=None),
            patch("md2doc.dependencies._refresh_windows_path"),
        ):
            with self.assertRaisesRegex(RuntimeError, "winget was not found"):
                ensure_startup_dependencies(runner=_ok)


def _ok(args: object) -> subprocess.CompletedProcess[str]:
    return subprocess.CompletedProcess(args=args, returncode=0, stdout="", stderr="")


if __name__ == "__main__":
    unittest.main()
