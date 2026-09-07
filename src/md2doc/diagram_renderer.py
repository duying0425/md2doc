from __future__ import annotations

import argparse
import hashlib
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys
import tempfile
from typing import Sequence

from ._process import hidden_subprocess_kwargs

DRAWIO_DESKTOP_NAMES = ("draw.io.exe", "drawio.exe", "draw.io", "drawio")
D2_NAMES = ("d2.exe", "d2")


def normalize_drawio_xml(raw_text: str) -> str:
    """Ensure drawio content is a valid <mxfile> XML document.
    
    Draw.io users frequently copy just <mxGraphModel>...</mxGraphModel> from the UI clipboard.
    This helper wraps it into a complete <mxfile><diagram> structure.
    """
    text = raw_text.strip()
    if not text:
        return text

    # Remove markdown code block fences if accidentally passed in
    text = re.sub(r"^```[a-zA-Z0-9_\-]*\s*\n", "", text)
    text = re.sub(r"\n```\s*$", "", text).strip()

    if "<mxfile" in text:
        return text

    if "<mxGraphModel" in text:
        # Extract mxGraphModel block if surrounded by other text
        match = re.search(r"(<mxGraphModel[\s\S]*?</mxGraphModel>)", text)
        if match:
            inner_xml = match.group(1)
        else:
            inner_xml = text
        return (
            '<mxfile host="md2doc" modified="2026-01-01T00:00:00.000Z" agent="md2doc" version="20.0.0" type="device">\n'
            '  <diagram id="diagram-1" name="Page-1">\n'
            f"    {inner_xml}\n"
            "  </diagram>\n"
            "</mxfile>"
        )

    return text


def find_windows_d2_candidates() -> list[Path]:
    candidates: list[Path] = []
    local_app_data = _env_path("LOCALAPPDATA")
    program_files = _env_path("ProgramFiles")
    program_files_x86 = _env_path("ProgramFiles(x86)")

    candidates.extend(
        path
        for path in [
            local_app_data / "Microsoft" / "WinGet" / "Links" / "d2.exe" if local_app_data else None,
            local_app_data / "Programs" / "D2" / "d2.exe" if local_app_data else None,
            local_app_data / "Programs" / "d2" / "d2.exe" if local_app_data else None,
            local_app_data / "Programs" / "d2" / "bin" / "d2.exe" if local_app_data else None,
            program_files / "D2" / "d2.exe" if program_files else None,
            program_files / "d2" / "d2.exe" if program_files else None,
            program_files / "D2" / "bin" / "d2.exe" if program_files else None,
            program_files / "d2" / "bin" / "d2.exe" if program_files else None,
            program_files_x86 / "D2" / "d2.exe" if program_files_x86 else None,
            program_files_x86 / "d2" / "d2.exe" if program_files_x86 else None,
            program_files_x86 / "d2" / "bin" / "d2.exe" if program_files_x86 else None,
        ]
        if path is not None
    )

    if local_app_data:
        winget_packages = local_app_data / "Microsoft" / "WinGet" / "Packages"
        if winget_packages.exists():
            for pkg in winget_packages.glob("Terrastruct.D2*"):
                d2_exe = pkg / "d2.exe"
                if d2_exe.exists():
                    candidates.append(d2_exe)
                for found in pkg.rglob("d2.exe"):
                    candidates.append(found)

    return _dedupe_paths(candidates)


def find_windows_drawio_candidates() -> list[Path]:
    candidates: list[Path] = []
    local_app_data = _env_path("LOCALAPPDATA")
    program_files = _env_path("ProgramFiles")
    program_files_x86 = _env_path("ProgramFiles(x86)")

    for name in ("draw.io.exe", "drawio.exe"):
        candidates.extend(
            path
            for path in [
                program_files / "draw.io" / name if program_files else None,
                program_files_x86 / "draw.io" / name if program_files_x86 else None,
                local_app_data / "Programs" / "draw.io" / name if local_app_data else None,
                local_app_data / "Microsoft" / "WinGet" / "Links" / name if local_app_data else None,
            ]
            if path is not None
        )

    if local_app_data:
        winget_packages = local_app_data / "Microsoft" / "WinGet" / "Packages"
        if winget_packages.exists():
            for pkg in winget_packages.glob("JGraph.Draw*"):
                for name in ("draw.io.exe", "drawio.exe"):
                    d_exe = pkg / name
                    if d_exe.exists():
                        candidates.append(d_exe)
                    for found in pkg.rglob(name):
                        candidates.append(found)

    return _dedupe_paths(candidates)


def resolve_d2_command(command: str = "d2") -> list[str]:
    cmd_path = shutil.which(command)
    if cmd_path:
        return [cmd_path]
    if os.name == "nt" and command.lower() in ("d2", "d2.exe"):
        for cand in find_windows_d2_candidates():
            if cand.exists():
                return [str(cand)]
    return [command]


def resolve_drawio_command(command: str = "drawio") -> list[str] | None:
    cmd_path = shutil.which(command)
    if cmd_path:
        return [cmd_path]
    if os.name == "nt" and command.lower() in ("drawio", "draw.io", "draw.io.exe", "drawio.exe"):
        for cand in find_windows_drawio_candidates():
            if cand.exists():
                return [str(cand)]
    # also check if "draw.io" works
    cand2 = shutil.which("draw.io")
    if cand2:
        return [cand2]
    return None


def render_svg_code(code: str, output_path: Path) -> Path:
    """Save inline SVG code to output_path."""
    clean_code = code.strip()
    # Strip markdown fences if any
    clean_code = re.sub(r"^```[a-zA-Z0-9_\-]*\s*\n", "", clean_code)
    clean_code = re.sub(r"\n```\s*$", "", clean_code).strip()

    # Ensure xmlns if missing
    if "<svg" in clean_code and 'xmlns="http://www.w3.org/2000/svg"' not in clean_code:
        clean_code = clean_code.replace("<svg", '<svg xmlns="http://www.w3.org/2000/svg"', 1)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(clean_code, encoding="utf-8")
    return output_path


def render_d2(
    source_text_or_path: str | Path,
    output_path: Path,
    *,
    d2_cmd: str = "d2",
    theme: str = "default",
    layout: str = "dagre",
    sketch: bool = False,
    pad: int = 10,
) -> Path:
    """Render D2 code or file to SVG using the D2 CLI."""
    resolved_cmd = resolve_d2_command(d2_cmd)
    if not (Path(resolved_cmd[0]).exists() or shutil.which(resolved_cmd[0])):
        raise RuntimeError(
            f"D2 CLI command '{d2_cmd}' was not found. "
            "Please install D2 via: winget install Terrastruct.D2 or see https://d2lang.com"
        )

    output_path.parent.mkdir(parents=True, exist_ok=True)
    cmd = [resolved_cmd[0], *resolved_cmd[1:]]

    if theme and theme != "default":
        cmd.extend(["--theme", theme])
    if layout and layout != "dagre":
        cmd.extend(["--layout", layout])
    if sketch:
        cmd.append("--sketch")
    if pad != 100:
        cmd.extend(["--pad", str(pad)])

    if isinstance(source_text_or_path, Path) and source_text_or_path.exists():
        cmd.extend([str(source_text_or_path), str(output_path)])
        input_data = None
    else:
        # Pass via stdin
        cmd.extend(["-", str(output_path)])
        input_data = str(source_text_or_path)

    completed = subprocess.run(
        cmd,
        input=input_data,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        check=False,
        **hidden_subprocess_kwargs(),
    )
    if completed.returncode != 0 or not output_path.exists():
        err_msg = (completed.stderr or completed.stdout or "D2 execution failed").strip()
        raise RuntimeError(f"D2 rendering failed:\n{err_msg}")

    return output_path


def render_drawio(
    source_text_or_path: str | Path,
    output_path: Path,
    *,
    drawio_cmd: str = "drawio",
    theme: str = "light",
) -> Path:
    """Render Draw.io diagram to SVG.
    
    1. Try desktop Draw.io CLI (drawio.exe -x -f svg)
    2. Fallback to Playwright Chromium offline/headless viewer.
    """
    if isinstance(source_text_or_path, Path) and source_text_or_path.exists():
        raw_xml = source_text_or_path.read_text(encoding="utf-8", errors="replace")
    else:
        raw_xml = str(source_text_or_path)

    xml_content = normalize_drawio_xml(raw_xml)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    # 1. Try desktop drawio CLI
    resolved = resolve_drawio_command(drawio_cmd)
    if resolved:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".drawio", mode="wb") as tmp:
            tmp.write(xml_content.encode("utf-8", errors="replace"))
            temp_drawio_path = Path(tmp.name)
        try:
            cmd = [
                resolved[0],
                *resolved[1:],
                "--export",
                "--format",
                "svg",
                "--output",
                str(output_path),
                str(temp_drawio_path),
            ]
            completed = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                check=False,
                **hidden_subprocess_kwargs(),
            )
            if completed.returncode == 0 and output_path.exists() and output_path.stat().st_size > 0:
                return output_path
        finally:
            if temp_drawio_path.exists():
                try:
                    temp_drawio_path.unlink()
                except OSError:
                    pass

    # 2. Try Playwright headless export
    try:
        _render_drawio_with_playwright(xml_content, output_path)
        if output_path.exists() and output_path.stat().st_size > 0:
            return output_path
    except Exception as exc:
        raise RuntimeError(
            "Draw.io export failed. Neither Draw.io Desktop CLI nor Playwright headless viewer succeeded.\n"
            f"Error details: {exc}\n"
            "Tip: You can install Draw.io Desktop with: winget install JGraph.Draw"
        ) from exc

    raise RuntimeError("Draw.io export failed to produce an output file.")


def _render_drawio_with_playwright(xml_content: str, output_path: Path) -> None:
    """Render Draw.io diagram using Playwright and draw.io viewer js."""
    from playwright.sync_api import sync_playwright

    data_obj = {
        "highlight": "#0000ff",
        "nav": False,
        "resize": True,
        "xml": xml_content,
    }

    local_viewer = Path(__file__).resolve().parent / "assets" / "viewer-static.min.js"
    if local_viewer.exists():
        script_src = local_viewer.as_uri()
    else:
        script_src = "https://cdn.jsdelivr.net/gh/jgraph/drawio@dev/src/main/webapp/js/viewer-static.min.js"

    html_content = f"""<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <script type="text/javascript" src="{script_src}"></script>
</head>
<body style="margin:0; padding:0; background:white;">
  <div id="container" class="mxgraph" style="max-width:100%;border:1px solid transparent;" data-mxgraph='{json.dumps(data_obj)}'></div>
</body>
</html>
"""
    with tempfile.NamedTemporaryFile(delete=False, suffix=".html", mode="w", encoding="utf-8") as tmp:
        tmp.write(html_content)
        temp_html = Path(tmp.name)

    try:
        with sync_playwright() as p:
            launch_args = ["--allow-file-access-from-files", "--disable-web-security"]
            browser = None
            for launch_opts in [
                {"channel": "msedge", "args": launch_args},
                {"args": launch_args},
            ]:
                try:
                    browser = p.chromium.launch(headless=True, **launch_opts)
                    break
                except Exception:
                    continue
            if not browser:
                browser = p.chromium.launch(headless=True, args=launch_args)

            try:
                page = browser.new_page()
                page.on("console", lambda msg: print("[BROWSER LOG]:", msg.text))
                page.on("pageerror", lambda err: print("[BROWSER ERROR]:", err))
                page.goto(temp_html.as_uri(), wait_until="load", timeout=15000)
                page.wait_for_selector("svg", timeout=10000)
                svg_content = page.evaluate("() => document.querySelector('svg').outerHTML")
                output_path.write_text(svg_content, encoding="utf-8")
            finally:
                browser.close()
    finally:
        if temp_html.exists():
            try:
                temp_html.unlink()
            except OSError:
                pass


def compute_diagram_hash(kind: str, content: str, **kwargs) -> str:
    """Compute deterministic SHA-256 hash for caching."""
    h = hashlib.sha256()
    h.update(kind.encode("utf-8"))
    h.update(content.strip().encode("utf-8"))
    for k, v in sorted(kwargs.items()):
        h.update(f"{k}={v}".encode("utf-8"))
    return h.hexdigest()[:24]


def _env_path(name: str) -> Path | None:
    value = os.environ.get(name)
    return Path(value) if value else None


def _dedupe_paths(paths: Sequence[Path | None]) -> list[Path]:
    seen: set[str] = set()
    deduped: list[Path] = []
    for path in paths:
        if path is None:
            continue
        key = str(path).lower()
        if key in seen:
            continue
        seen.add(key)
        deduped.append(path)
    return deduped


def main() -> None:
    parser = argparse.ArgumentParser(description="Render diagram (SVG, D2, Draw.io) to SVG file")
    parser.add_argument("--kind", choices=("svg", "d2", "drawio"), required=True)
    parser.add_argument("--output", "-o", required=True, help="Target output SVG path")
    parser.add_argument("--input", "-i", default=None, help="Input source file (if omitted, reads stdin)")
    parser.add_argument("--theme", default="default")
    parser.add_argument("--layout", default="dagre")
    parser.add_argument("--sketch", action="store_true")
    parser.add_argument("--d2-cmd", default="d2")
    parser.add_argument("--drawio-cmd", default="drawio")

    args = parser.parse_args()
    output_path = Path(args.output).resolve()

    if args.input:
        content = Path(args.input).read_text(encoding="utf-8", errors="replace")
    else:
        raw_bytes = sys.stdin.buffer.read()
        try:
            content = raw_bytes.decode("utf-8")
        except UnicodeDecodeError:
            content = raw_bytes.decode("utf-8", errors="replace")

    if args.kind == "svg":
        render_svg_code(content, output_path)
    elif args.kind == "d2":
        render_d2(
            content,
            output_path,
            d2_cmd=args.d2_cmd,
            theme=args.theme,
            layout=args.layout,
            sketch=args.sketch,
        )
    elif args.kind == "drawio":
        render_drawio(
            content,
            output_path,
            drawio_cmd=args.drawio_cmd,
            theme=args.theme,
        )


if __name__ == "__main__":
    main()
