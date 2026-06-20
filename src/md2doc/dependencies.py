from __future__ import annotations

from collections.abc import Callable, Sequence
import os
from pathlib import Path
import shutil
import subprocess

from .converter import ConvertSettings, check_dependencies, missing_dependency_message


InstallerRunner = Callable[[Sequence[str]], subprocess.CompletedProcess[str]]
ProgressCallback = Callable[[str], None]


def ensure_startup_dependencies(
    *,
    on_progress: ProgressCallback | None = None,
    runner: InstallerRunner | None = None,
) -> None:
    """Install external conversion tools when the default GUI toolchain is missing."""

    _refresh_windows_path()
    settings = ConvertSettings()
    checks = check_dependencies(settings)
    if all(check.available for check in checks):
        _emit(on_progress, "Conversion tools are ready.")
        return

    if os.name != "nt":
        raise RuntimeError(missing_dependency_message(checks))

    run = runner or _run_command
    _emit(on_progress, "Installing missing conversion tools...")

    if not _tool_available("pandoc"):
        _install_with_winget(
            "Pandoc",
            "JohnMacFarlane.Pandoc",
            run,
            on_progress,
        )
        _refresh_windows_path()

    if not _tool_available("npm"):
        _install_with_winget(
            "Node.js LTS",
            "OpenJS.NodeJS.LTS",
            run,
            on_progress,
        )
        _refresh_windows_path()

    if not _tool_available("mermaid-filter"):
        npm = _resolve_npm()
        if not npm:
            raise RuntimeError("npm was not found after installing Node.js.")
        _emit(on_progress, "Installing mermaid-filter with npm...")
        _run_or_raise([npm, "install", "-g", "mermaid-filter"], run)
        _refresh_windows_path()

    final_checks = check_dependencies(settings)
    if not all(check.available for check in final_checks):
        raise RuntimeError(missing_dependency_message(final_checks))
    _emit(on_progress, "Dependency setup completed.")


def _install_with_winget(
    label: str,
    package_id: str,
    runner: InstallerRunner,
    on_progress: ProgressCallback | None,
) -> None:
    winget = _resolve_winget()
    if not winget:
        raise RuntimeError("winget was not found. Install App Installer from Microsoft Store, then reopen md2doc.")

    _emit(on_progress, f"Installing {label} with winget...")
    _run_or_raise(
        [
            winget,
            "install",
            "--id",
            package_id,
            "--exact",
            "--silent",
            "--accept-package-agreements",
            "--accept-source-agreements",
        ],
        runner,
    )


def _run_command(args: Sequence[str]) -> subprocess.CompletedProcess[str]:
    return subprocess.run(args, capture_output=True, text=True, check=False)


def _run_or_raise(args: Sequence[str], runner: InstallerRunner) -> None:
    completed = runner(args)
    if completed.returncode == 0:
        return
    output = (completed.stderr or completed.stdout or "").strip()
    command = " ".join(args)
    raise RuntimeError(output or f"Command failed: {command}")


def _tool_available(command: str) -> bool:
    return _resolve_command_path(command) is not None


def _resolve_winget() -> str | None:
    return _resolve_command_path("winget") or _first_existing(
        [
            _env_path("LOCALAPPDATA") / "Microsoft" / "WindowsApps" / "winget.exe",
        ]
    )


def _resolve_npm() -> str | None:
    return _resolve_command_path("npm") or _first_existing(
        [
            _env_path("ProgramFiles") / "nodejs" / "npm.cmd",
            _env_path("ProgramFiles(x86)") / "nodejs" / "npm.cmd",
        ]
    )


def _resolve_command_path(command: str) -> str | None:
    path = shutil.which(command)
    if path:
        return path
    if os.name != "nt":
        return None
    for suffix in (".cmd", ".exe", ".bat"):
        path = shutil.which(f"{command}{suffix}")
        if path:
            return path
    return None


def _refresh_windows_path() -> None:
    if os.name != "nt":
        return
    paths = [
        _env_path("ProgramFiles") / "nodejs",
        _env_path("ProgramFiles(x86)") / "nodejs",
        _env_path("APPDATA") / "npm",
        _env_path("LOCALAPPDATA") / "Pandoc",
        _env_path("LOCALAPPDATA") / "Microsoft" / "WinGet" / "Links",
    ]
    for path in paths:
        if path.exists():
            _prepend_path(path)


def _prepend_path(path: Path) -> None:
    value = str(path)
    current = os.environ.get("PATH", "")
    parts = [part for part in current.split(os.pathsep) if part]
    if any(part.lower() == value.lower() for part in parts):
        return
    os.environ["PATH"] = value + os.pathsep + current if current else value


def _first_existing(paths: Sequence[Path]) -> str | None:
    for path in paths:
        if path.exists():
            return str(path)
    return None


def _env_path(name: str) -> Path:
    value = os.environ.get(name)
    return Path(value) if value else Path()


def _emit(callback: ProgressCallback | None, message: str) -> None:
    if callback:
        callback(message)
