from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime, timezone
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import shlex
import shutil
import subprocess
import tempfile
import threading
from typing import Callable, Iterable, Literal
import xml.etree.ElementTree as ET
import zipfile

from ._process import hidden_subprocess_kwargs
from .project import (
    KIND_DOC2MD,
    KIND_HTML2PDF,
    KIND_MD2DOC,
    KIND_QMD2PPT,
    PROJECT_DIR_NAME,
    ProjectConfig,
)


SUPPORTED_FORMATS = {"docx": ".docx", "pptx": ".pptx"}
MARKDOWN_SUFFIXES = {".md", ".markdown"}
# Office documents that MarkItDown can turn into Markdown (Word/PowerPoint/Excel).
OFFICE_SUFFIXES = {".docx", ".doc", ".pptx", ".ppt", ".xlsx", ".xls"}
HTML_SUFFIXES = {".html", ".htm"}
DOC2MD_OUTPUT_SUFFIX = ".md"
HTML2PDF_OUTPUT_SUFFIX = ".pdf"
DEFAULT_EXCLUDED_DIRS = {
    ".git",
    ".hg",
    ".svn",
    ".md2doc",
    ".venv",
    "node_modules",
    "dist",
    "build",
    "bin",
}
MANIFEST_NAME = "manifest.json"
GENERATED_REFERENCE_DOCX = "generated-reference.docx"
GENERATED_REFERENCE_META = "generated-reference.json"
MERMAID_FILTER_ERR_NAME = "mermaid-filter.err"
MERMAID_IMAGE_DIR_NAME = "mermaid-images"
MERMAID_IMAGE_LOCATION_SIGNATURE = "metadata-mermaid-images-v1"
# Pandoc (a Haskell binary) is not long-path aware, so it cannot open Lua
# filters, reference docs, or resource files whose Windows path exceeds
# MAX_PATH (260) even when the OS has long paths enabled. We hand it 8.3 short
# aliases for any path at or beyond this threshold; the margin keeps us clear of
# the limit once Pandoc appends its own suffixes.
WINDOWS_MAX_PATH = 260
LONG_PATH_THRESHOLD = 200
WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{WORD_NS}}}"
ET.register_namespace("w", WORD_NS)


class ConversionCancelledError(Exception):
    """Raised when conversion is cancelled by the user."""
    pass


PlanAction = Literal["convert", "skip"]


@dataclass(frozen=True)
class DependencyCheck:
    name: str
    command: str
    available: bool
    detail: str


@dataclass(frozen=True)
class ConvertSettings:
    kind: str = KIND_MD2DOC
    output_format: str = "docx"
    output_dir: Path | None = None
    recursive: bool = True
    pandoc_cmd: str = "pandoc"
    mermaid_filter_cmd: str = "mermaid-filter"
    markitdown_cmd: str = "markitdown"
    quarto_cmd: str = "quarto"
    extra_pandoc_args: tuple[str, ...] = ()
    hr_to_pagebreak: bool = True
    force: bool = False
    skip_unchanged: bool = True
    toc: bool = False
    toc_depth: int = 3
    title_page: bool = False
    title: str = ""
    subtitle: str = ""
    author: str = ""
    date: str = ""
    number_sections: bool = False
    reference_docx: str = ""
    default_font: str = ""
    default_font_size: int = 0
    table_borders: str = "template"
    mermaid_format: str = "png"
    mermaid_theme: str = "default"
    mermaid_background: str = "white"
    mermaid_scale: float = 3.0
    mermaid_min_dpi: float = 450.0
    figure_numbering: bool = True
    figure_prefix: str = "图表"
    figure_caption_position: str = "below"
    html_viewport_width: int = 1280
    html_viewport_height: int = 720
    html_device_scale_factor: float = 1.0
    html_print_background: bool = True
    html_render_delay: float = 0.0

    def output_suffix(self) -> str:
        if self.kind == KIND_DOC2MD:
            return DOC2MD_OUTPUT_SUFFIX
        if self.kind == KIND_QMD2PPT:
            return ".pptx"
        if self.kind == KIND_HTML2PDF:
            return HTML2PDF_OUTPUT_SUFFIX
        try:
            return SUPPORTED_FORMATS[self.output_format]
        except KeyError as exc:
            supported = ", ".join(sorted(SUPPORTED_FORMATS))
            raise ValueError(f"Unsupported output format: {self.output_format}. Use: {supported}") from exc

    def input_suffixes(self) -> set[str]:
        if self.kind == KIND_DOC2MD:
            return OFFICE_SUFFIXES
        if self.kind == KIND_QMD2PPT:
            return {".qmd"}
        if self.kind == KIND_HTML2PDF:
            return HTML_SUFFIXES
        return MARKDOWN_SUFFIXES


@dataclass(frozen=True)
class FileFingerprint:
    size: int
    mtime_ns: int
    sha256: str


@dataclass(frozen=True)
class PlanItem:
    source: Path
    relative_source: str
    output: Path
    action: PlanAction
    reason: str
    fingerprint: FileFingerprint
    settings_signature: str


@dataclass(frozen=True)
class ConversionResult:
    item: PlanItem
    status: Literal["converted", "skipped", "failed"]
    message: str
    returncode: int | None = None


@dataclass
class BuildManifest:
    path: Path
    records: dict[str, dict] = field(default_factory=dict)

    @classmethod
    def load(cls, project_root: Path) -> "BuildManifest":
        path = project_root / PROJECT_DIR_NAME / MANIFEST_NAME
        if not path.exists():
            return cls(path=path)
        try:
            payload = json.loads(path.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            return cls(path=path)
        return cls(path=path, records=dict(payload.get("records", {})))

    def save(self) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        payload = {
            "version": 1,
            "records": self.records,
        }
        self.path.write_text(json.dumps(payload, indent=2, ensure_ascii=True), encoding="utf-8")

    def record_success(self, item: PlanItem) -> None:
        self.records[item.relative_source] = {
            "source_sha256": item.fingerprint.sha256,
            "source_size": item.fingerprint.size,
            "source_mtime_ns": item.fingerprint.mtime_ns,
            "output": str(item.output),
            "output_format": item.output.suffix.lstrip("."),
            "settings_signature": item.settings_signature,
            "converted_at": datetime.now(timezone.utc).isoformat(),
        }


def settings_from_project(config: ProjectConfig, *, force: bool = False) -> ConvertSettings:
    return ConvertSettings(
        kind=config.kind,
        output_format=config.output_format,
        output_dir=config.output_path,
        recursive=config.recursive,
        extra_pandoc_args=tuple(config.extra_pandoc_args),
        hr_to_pagebreak=config.hr_to_pagebreak,
        force=force,
        toc=config.toc,
        toc_depth=config.toc_depth,
        title_page=config.title_page,
        title=config.title,
        subtitle=config.subtitle,
        author=config.author,
        date=config.date,
        number_sections=config.number_sections,
        reference_docx=config.reference_docx,
        default_font=config.default_font,
        default_font_size=config.default_font_size,
        table_borders=config.table_borders,
        mermaid_format=config.mermaid_format,
        mermaid_theme=config.mermaid_theme,
        mermaid_background=config.mermaid_background,
        mermaid_scale=config.mermaid_scale,
        mermaid_min_dpi=config.mermaid_min_dpi,
        figure_numbering=config.figure_numbering,
        figure_prefix=config.figure_prefix,
        figure_caption_position=config.figure_caption_position,
        html_viewport_width=config.html_viewport_width,
        html_viewport_height=config.html_viewport_height,
        html_device_scale_factor=config.html_device_scale_factor,
        html_print_background=config.html_print_background,
        html_render_delay=config.html_render_delay,
    )


def check_dependencies(settings: ConvertSettings) -> list[DependencyCheck]:
    if settings.kind == KIND_DOC2MD:
        return [_check_markitdown(settings.markitdown_cmd)]
    if settings.kind == KIND_QMD2PPT:
        return [_check_command("Quarto", settings.quarto_cmd)]
    if settings.kind == KIND_HTML2PDF:
        return [_check_html_pdf_runtime()]
    return [
        _check_command("Pandoc", settings.pandoc_cmd),
        _check_command("mermaid-filter", settings.mermaid_filter_cmd, allow_version_failure=True),
        _check_mermaid_browser_runtime(),
    ]


def missing_dependency_message(checks: Iterable[DependencyCheck]) -> str:
    missing = [check for check in checks if not check.available]
    if not missing:
        return ""
    lines = ["Missing required conversion tools:"]
    for check in missing:
        lines.append(f"- {check.name}: {check.detail}")
    names = {check.name for check in missing}
    if names & {"Pandoc", "mermaid-filter"}:
        lines.append("Install Pandoc and then run: npm install -g mermaid-filter")
    if "Mermaid browser" in names:
        lines.append(
            "Install the managed Chromium runtime for Mermaid rendering: python -m playwright install chromium"
        )
    if "MarkItDown" in names:
        lines.append("Install MarkItDown: pip install 'markitdown[docx,pptx,xlsx]'")
    if "Quarto" in names:
        lines.append("Install Quarto: https://quarto.org/docs/get-started/")
    if "Playwright/Chromium" in names:
        lines.append("Install Playwright and a browser: pip install playwright && python -m playwright install chromium")
    return "\n".join(lines)


def scan_source_files(
    project_root: Path | str,
    *,
    kind: str = KIND_MD2DOC,
    recursive: bool = True,
    output_dir: Path | None = None,
    excluded_dirs: set[str] | None = None,
) -> list[Path]:
    if kind == KIND_DOC2MD:
        suffixes = OFFICE_SUFFIXES
    elif kind == KIND_QMD2PPT:
        suffixes = {".qmd"}
    elif kind == KIND_HTML2PDF:
        suffixes = HTML_SUFFIXES
    else:
        suffixes = MARKDOWN_SUFFIXES
    return _scan_files(
        project_root,
        suffixes,
        recursive=recursive,
        output_dir=output_dir,
        excluded_dirs=excluded_dirs,
    )


def scan_markdown_files(
    project_root: Path | str,
    *,
    recursive: bool = True,
    output_dir: Path | None = None,
    excluded_dirs: set[str] | None = None,
) -> list[Path]:
    return _scan_files(
        project_root,
        MARKDOWN_SUFFIXES,
        recursive=recursive,
        output_dir=output_dir,
        excluded_dirs=excluded_dirs,
    )


def _scan_files(
    project_root: Path | str,
    suffixes: set[str],
    *,
    recursive: bool = True,
    output_dir: Path | None = None,
    excluded_dirs: set[str] | None = None,
) -> list[Path]:
    root = Path(project_root).expanduser().resolve()
    excluded = set(DEFAULT_EXCLUDED_DIRS if excluded_dirs is None else excluded_dirs)
    if output_dir:
        output_dir = output_dir.expanduser().resolve()
        if output_dir == root:
            output_dir = None

    files: list[Path] = []
    if recursive:
        for current, dirnames, filenames in os.walk(root):
            current_path = Path(current).resolve()
            dirnames[:] = [
                dirname
                for dirname in dirnames
                if dirname not in excluded
                and not _is_same_or_child(current_path / dirname, output_dir)
                and not (current_path / dirname / PROJECT_DIR_NAME).is_dir()
            ]
            for filename in filenames:
                path = current_path / filename
                if path.suffix.lower() in suffixes:
                    files.append(path)
    else:
        for path in root.iterdir():
            if path.is_file() and path.suffix.lower() in suffixes:
                files.append(path.resolve())
    return sorted(files, key=lambda path: path.relative_to(root).as_posix().lower())


def plan_conversions(
    project_root: Path | str,
    sources: Iterable[Path],
    settings: ConvertSettings,
    manifest: BuildManifest | None = None,
    *,
    use_cached_fingerprints: bool = False,
) -> list[PlanItem]:
    root = Path(project_root).expanduser().resolve()
    manifest = manifest or BuildManifest.load(root)
    output_dir = (settings.output_dir or root).expanduser().resolve()
    output_suffix = settings.output_suffix()
    signature = settings_signature(settings, root)
    planned: list[PlanItem] = []

    for source in sources:
        source = Path(source).expanduser().resolve()
        relative = source.relative_to(root).as_posix()
        record = manifest.records.get(relative)
        fingerprint = _plan_fingerprint(
            source,
            record,
            use_cached=use_cached_fingerprints,
        )
        output = output_dir / source.relative_to(root).with_suffix(output_suffix)
        action, reason = _decide_action(
            settings=settings,
            record=record,
            fingerprint=fingerprint,
            output=output,
            signature=signature,
        )
        planned.append(
            PlanItem(
                source=source,
                relative_source=relative,
                output=output,
                action=action,
                reason=reason,
                fingerprint=fingerprint,
                settings_signature=signature,
            )
        )
    return planned


def run_conversions(
    project_root: Path | str,
    sources: Iterable[Path],
    settings: ConvertSettings,
    *,
    on_event: Callable[[ConversionResult], None] | None = None,
    on_start: Callable[[PlanItem], None] | None = None,
    cancel_event: threading.Event | None = None,
) -> list[ConversionResult]:
    root = Path(project_root).expanduser().resolve()
    manifest = BuildManifest.load(root)
    items = plan_conversions(root, sources, settings, manifest)
    needs_convert = [item for item in items if item.action == "convert"]
    if needs_convert:
        _validate_settings(root, settings)
        checks = check_dependencies(settings)
        if settings.kind == KIND_MD2DOC and not _plans_need_mermaid(needs_convert):
            checks = [check for check in checks if check.name != "Mermaid browser"]
        message = missing_dependency_message(checks)
        if message:
            raise RuntimeError(message)

    results: list[ConversionResult] = []
    for item in items:
        if cancel_event is not None and cancel_event.is_set():
            raise ConversionCancelledError("Conversion cancelled by user")

        if item.action == "skip":
            result = ConversionResult(item=item, status="skipped", message=item.reason)
            results.append(result)
            if item.reason == "unchanged":
                record = manifest.records.get(item.relative_source)
                if record and (
                    record.get("source_mtime_ns") != item.fingerprint.mtime_ns
                    or record.get("source_size") != item.fingerprint.size
                ):
                    manifest.record_success(item)
                    manifest.save()
            if on_event:
                on_event(result)
            continue

        if on_start:
            on_start(item)
        result = _run_one(root, item, settings, cancel_event=cancel_event)
        results.append(result)
        if result.status == "converted":
            manifest.record_success(item)
            manifest.save()
        if on_event:
            on_event(result)
    return results


def file_fingerprint(path: Path) -> FileFingerprint:
    return _file_fingerprint_from_stat(path, path.stat())


def _plan_fingerprint(
    source: Path,
    record: dict | None,
    *,
    use_cached: bool,
) -> FileFingerprint:
    stat = source.stat()
    if use_cached:
        cached_sha = str(record.get("source_sha256") or "") if record else ""
        if record and (
            cached_sha
            and record.get("source_size") == stat.st_size
            and record.get("source_mtime_ns") == stat.st_mtime_ns
        ):
            return FileFingerprint(size=stat.st_size, mtime_ns=stat.st_mtime_ns, sha256=cached_sha)
        return FileFingerprint(size=stat.st_size, mtime_ns=stat.st_mtime_ns, sha256="")
    return _file_fingerprint_from_stat(source, stat)


def _file_fingerprint_from_stat(path: Path, stat: os.stat_result) -> FileFingerprint:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return FileFingerprint(size=stat.st_size, mtime_ns=stat.st_mtime_ns, sha256=digest.hexdigest())


def _plans_need_mermaid(items: Iterable[PlanItem]) -> bool:
    return any(_source_contains_mermaid_block(item.source) for item in items)


def _source_contains_mermaid_block(path: Path) -> bool:
    try:
        text = path.read_text(encoding="utf-8-sig")
    except UnicodeDecodeError:
        text = path.read_text(encoding="utf-8", errors="replace")
    except OSError:
        return False
    for line in text.splitlines():
        stripped = line.strip().lower()
        if (stripped.startswith("```") or stripped.startswith("~~~")) and "mermaid" in stripped:
            return True
    return False


def settings_signature(settings: ConvertSettings, project_root: Path | None = None) -> str:
    payload = {
        "kind": settings.kind,
        "output_format": settings.output_format,
        "pandoc_cmd": settings.pandoc_cmd,
        "mermaid_filter_cmd": settings.mermaid_filter_cmd,
        "markitdown_cmd": settings.markitdown_cmd,
        "quarto_cmd": settings.quarto_cmd,
        "extra_pandoc_args": list(settings.extra_pandoc_args),
        "hr_to_pagebreak": settings.hr_to_pagebreak,
        "toc": settings.toc,
        "toc_depth": settings.toc_depth,
        "title_page": settings.title_page,
        "title": settings.title,
        "subtitle": settings.subtitle,
        "author": settings.author,
        "date": settings.date,
        "number_sections": settings.number_sections,
        "reference_docx": settings.reference_docx,
        "reference_docx_stat": _file_stat_signature(settings.reference_docx, project_root),
        "default_font": settings.default_font,
        "default_font_size": settings.default_font_size,
        "table_borders": settings.table_borders,
        "mermaid_format": settings.mermaid_format,
        "mermaid_theme": settings.mermaid_theme,
        "mermaid_background": settings.mermaid_background,
        "mermaid_scale": settings.mermaid_scale,
        "mermaid_min_dpi": settings.mermaid_min_dpi,
        "figure_numbering": settings.figure_numbering,
        "figure_prefix": settings.figure_prefix,
        "figure_caption_position": settings.figure_caption_position,
        "html_viewport_width": settings.html_viewport_width,
        "html_viewport_height": settings.html_viewport_height,
        "html_device_scale_factor": settings.html_device_scale_factor,
        "html_print_background": settings.html_print_background,
        "html_render_delay": settings.html_render_delay,
        "mermaid_image_location": MERMAID_IMAGE_LOCATION_SIGNATURE,
    }
    encoded = json.dumps(payload, sort_keys=True, ensure_ascii=True).encode("utf-8")
    return hashlib.sha256(encoded).hexdigest()


def _run_subprocess_with_cancel(
    cmd: list[str],
    *,
    cwd: Path | str | None = None,
    env: dict[str, str] | None = None,
    capture_output: bool = False,
    text: bool = False,
    encoding: str | None = None,
    errors: str | None = None,
    check: bool = False,
    cancel_event: threading.Event | None = None,
    **kwargs,
) -> subprocess.CompletedProcess:
    if cancel_event is None:
        return subprocess.run(
            cmd,
            cwd=cwd,
            env=env,
            capture_output=capture_output,
            text=text,
            encoding=encoding,
            errors=errors,
            check=check,
            **kwargs,
        )

    stdout_opt = subprocess.PIPE if capture_output else None
    stderr_opt = subprocess.PIPE if capture_output else None

    p = subprocess.Popen(
        cmd,
        cwd=cwd,
        env=env,
        stdout=stdout_opt,
        stderr=stderr_opt,
        text=text,
        encoding=encoding,
        errors=errors,
        **kwargs,
    )

    stdout_chunks: list[str | bytes] = []
    stderr_chunks: list[str | bytes] = []

    try:
        while p.poll() is None:
            if cancel_event.is_set():
                p.terminate()
                try:
                    p.wait(timeout=1.0)
                except subprocess.TimeoutExpired:
                    p.kill()
                    p.wait()
                raise ConversionCancelledError("Conversion cancelled by user")

            try:
                stdout, stderr = p.communicate(timeout=0.1)
                if stdout:
                    stdout_chunks.append(stdout)
                if stderr:
                    stderr_chunks.append(stderr)
            except subprocess.TimeoutExpired as exc:
                if exc.stdout:
                    stdout_chunks.append(exc.stdout)
                if exc.stderr:
                    stderr_chunks.append(exc.stderr)
                continue
    except Exception:
        if p.poll() is None:
            p.terminate()
            try:
                p.wait(timeout=1.0)
            except subprocess.TimeoutExpired:
                p.kill()
                p.wait()
        raise

    stdout_remain, stderr_remain = p.communicate()
    if stdout_remain:
        stdout_chunks.append(stdout_remain)
    if stderr_remain:
        stderr_chunks.append(stderr_remain)

    stdout_res = "".join(stdout_chunks) if text else b"".join(stdout_chunks)
    stderr_res = "".join(stderr_chunks) if text else b"".join(stderr_chunks)

    if check and p.returncode != 0:
        raise subprocess.CalledProcessError(
            p.returncode,
            cmd,
            output=stdout_res,
            stderr=stderr_res,
        )

    return subprocess.CompletedProcess(
        args=cmd,
        returncode=p.returncode,
        stdout=stdout_res,
        stderr=stderr_res,
    )


def _run_one(
    project_root: Path,
    item: PlanItem,
    settings: ConvertSettings,
    cancel_event: threading.Event | None = None,
) -> ConversionResult:
    if settings.kind == KIND_DOC2MD:
        return _run_markitdown(item, settings, cancel_event=cancel_event)
    if settings.kind == KIND_QMD2PPT:
        return _run_quarto(item, settings, cancel_event=cancel_event)
    if settings.kind == KIND_HTML2PDF:
        return _run_html_to_pdf(item, settings, cancel_event=cancel_event)
    item.output.parent.mkdir(parents=True, exist_ok=True)
    cmd = _pandoc_command(project_root, item, settings)
    env = os.environ.copy()
    env.update(_mermaid_environment(settings, project_root))
    env["MERMAID_FILTER_LOC"] = _shorten_windows_path(_ensure_mermaid_image_dir(project_root, item))
    env["MD2DOC_RESOURCE_PATHS"] = os.pathsep.join(_pandoc_resource_paths(item, project_root))
    mermaid_error_path = _reset_mermaid_filter_error_log(item.source.parent)
    completed = _run_subprocess_with_cancel(
        cmd,
        cwd=_shorten_windows_path(item.source.parent),
        capture_output=True,
        text=True,
        check=False,
        env=env,
        cancel_event=cancel_event,
        **hidden_subprocess_kwargs(),
    )
    mermaid_error = _mermaid_filter_error_text(mermaid_error_path)
    if completed.returncode == 0:
        _remove_file(mermaid_error_path)
        if settings.output_format == "docx":
            _center_docx_images(item.output)
        return ConversionResult(item=item, status="converted", message="converted", returncode=0)

    message = _pandoc_failure_message(completed, mermaid_error)
    return ConversionResult(
        item=item,
        status="failed",
        message=message,
        returncode=completed.returncode,
    )


def _run_html_to_pdf(
    item: PlanItem,
    settings: ConvertSettings,
    cancel_event: threading.Event | None = None,
) -> ConversionResult:
    if cancel_event is not None and cancel_event.is_set():
        raise ConversionCancelledError("Conversion cancelled by user")

    item.output.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf", dir=item.output.parent) as tmp:
        temp_path = Path(tmp.name)
    try:
        _render_html_to_single_page_pdf(item.source, temp_path, settings, cancel_event=cancel_event)
        if cancel_event is not None and cancel_event.is_set():
            raise ConversionCancelledError("Conversion cancelled by user")
        if item.output.exists():
            item.output.unlink()
        shutil.move(str(temp_path), str(item.output))
        return ConversionResult(item=item, status="converted", message="converted", returncode=0)
    except ConversionCancelledError:
        raise
    except Exception as exc:
        return ConversionResult(
            item=item,
            status="failed",
            message=str(exc),
            returncode=1,
        )
    finally:
        if temp_path.exists():
            temp_path.unlink()


def _render_html_to_single_page_pdf(
    source: Path,
    output: Path,
    settings: ConvertSettings,
    *,
    cancel_event: threading.Event | None = None,
) -> None:
    if importlib.util.find_spec("playwright") is None:
        raise RuntimeError(
            "Playwright is not installed. Install it with: pip install playwright && python -m playwright install chromium"
        )

    from playwright.sync_api import sync_playwright

    with sync_playwright() as playwright:
        browser = _launch_html_pdf_browser(playwright)
        try:
            if cancel_event is not None and cancel_event.is_set():
                raise ConversionCancelledError("Conversion cancelled by user")
            page = browser.new_page(
                viewport={"width": settings.html_viewport_width, "height": settings.html_viewport_height},
                device_scale_factor=settings.html_device_scale_factor,
            )
            page.emulate_media(media="screen")
            page.goto(source.as_uri(), wait_until="load", timeout=60000)
            _wait_for_html_render_ready(page, settings.html_render_delay)
            width_px, height_px = _measure_html_render_size(page)
            if cancel_event is not None and cancel_event.is_set():
                raise ConversionCancelledError("Conversion cancelled by user")
            page.pdf(
                path=str(output),
                width=f"{width_px}px",
                height=f"{height_px}px",
                margin={"top": "0px", "right": "0px", "bottom": "0px", "left": "0px"},
                print_background=settings.html_print_background,
                prefer_css_page_size=False,
                scale=1,
            )
        finally:
            browser.close()


def _launch_html_pdf_browser(playwright):
    errors: list[str] = []
    for label, options in _html_pdf_launch_candidates(playwright):
        try:
            return playwright.chromium.launch(headless=True, **options)
        except Exception as exc:
            errors.append(f"{label}: {exc}")
    details = "; ".join(errors)
    raise RuntimeError(
        "Could not launch a Chromium browser for HTML to PDF. "
        "Install Microsoft Edge, Google Chrome, or run: python -m playwright install chromium"
        + (f"\n{details}" if details else "")
    )


def _html_pdf_launch_candidates(playwright) -> list[tuple[str, dict[str, str]]]:
    candidates: list[tuple[str, dict[str, str]]] = []
    if os.name == "nt":
        candidates.extend(
            [
                ("Microsoft Edge", {"channel": "msedge"}),
                ("Google Chrome", {"channel": "chrome"}),
            ]
        )
    else:
        candidates.extend(
            [
                ("Google Chrome", {"channel": "chrome"}),
                ("Microsoft Edge", {"channel": "msedge"}),
            ]
        )
    candidates.append(("Playwright Chromium", {}))
    return candidates


def _wait_for_html_render_ready(page, delay: float = 0.0) -> None:
    try:
        page.wait_for_load_state("networkidle", timeout=5000)
    except Exception:
        pass
    page.evaluate(
        """
        async () => {
          if (document.fonts && document.fonts.ready) {
            await document.fonts.ready;
          }
          const images = Array.from(document.images || []);
          await Promise.all(images.map((image) => {
            if (image.complete) {
              return Promise.resolve();
            }
            return new Promise((resolve) => {
              image.addEventListener('load', resolve, { once: true });
              image.addEventListener('error', resolve, { once: true });
            });
          }));
        }
        """
    )
    if delay > 0:
        page.wait_for_timeout(delay * 1000)


def _measure_html_render_size(page) -> tuple[int, int]:
    size = page.evaluate(
        """
        () => {
          const doc = document.documentElement;
          const body = document.body;
          const viewportWidth = window.innerWidth || 0;
          const viewportHeight = window.innerHeight || 0;
          let right = 0;
          let bottom = 0;
          const sizeProperties = ['width', 'height', 'min-width', 'min-height'];

          const visit = (element) => {
            if (!element || !element.getBoundingClientRect) {
              return;
            }
            const rect = element.getBoundingClientRect();
            if (!Number.isFinite(rect.right) || !Number.isFinite(rect.bottom)) {
              return;
            }
            if (rect.width > 0 || rect.height > 0) {
              right = Math.max(right, rect.right + window.scrollX);
              bottom = Math.max(bottom, rect.bottom + window.scrollY);
            }
          };

          const ruleDeclaresSizeFor = (element, rules) => {
            for (const rule of Array.from(rules || [])) {
              if (rule.cssRules && ruleDeclaresSizeFor(element, rule.cssRules)) {
                return true;
              }
              if (!rule.selectorText || !rule.style) {
                continue;
              }
              try {
                if (
                  sizeProperties.some((name) => rule.style.getPropertyValue(name)) &&
                  element.matches(rule.selectorText)
                ) {
                  return true;
                }
              } catch (_error) {
                continue;
              }
            }
            return false;
          };

          const hasDeclaredSize = (element) => {
            if (!element) {
              return false;
            }
            const inlineDeclares = sizeProperties.some((name) => element.style.getPropertyValue(name));
            if (inlineDeclares) {
              return true;
            }
            for (const sheet of Array.from(document.styleSheets || [])) {
              try {
                if (ruleDeclaresSizeFor(element, sheet.cssRules)) {
                  return true;
                }
              } catch (_error) {
                continue;
              }
            }
            return false;
          };

          if (hasDeclaredSize(doc)) {
            visit(doc);
          }
          if (hasDeclaredSize(body)) {
            visit(body);
          }
          if (body) {
            for (const element of Array.from(body.querySelectorAll('*'))) {
              visit(element);
            }
          }

          const scrollWidth = Math.max(
            doc ? doc.scrollWidth : 0,
            body ? body.scrollWidth : 0,
            doc ? doc.offsetWidth : 0,
            body ? body.offsetWidth : 0
          );
          const scrollHeight = Math.max(
            doc ? doc.scrollHeight : 0,
            body ? body.scrollHeight : 0,
            doc ? doc.offsetHeight : 0,
            body ? body.offsetHeight : 0
          );
          if (scrollWidth > viewportWidth) {
            right = Math.max(right, scrollWidth);
          }
          if (scrollHeight > viewportHeight) {
            bottom = Math.max(bottom, scrollHeight);
          }
          if ((right <= 0 || bottom <= 0) && body) {
            visit(body);
          }

          return {
            width: Math.max(1, Math.ceil(right + 1)),
            height: Math.max(1, Math.ceil(bottom + 1))
          };
        }
        """
    )
    return int(size["width"]), int(size["height"])


def _run_markitdown(
    item: PlanItem,
    settings: ConvertSettings,
    cancel_event: threading.Event | None = None,
) -> ConversionResult:
    if _should_use_markitdown_api(settings.markitdown_cmd):
        return _run_markitdown_api(item)

    item.output.parent.mkdir(parents=True, exist_ok=True)
    cmd = _markitdown_command(item, settings)
    completed = _run_subprocess_with_cancel(
        cmd,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        check=False,
        cancel_event=cancel_event,
        **hidden_subprocess_kwargs(),
    )
    if completed.returncode == 0 and item.output.exists():
        return ConversionResult(item=item, status="converted", message="converted", returncode=0)

    if completed.returncode == 0:
        message = "MarkItDown produced no output"
    else:
        message = (completed.stderr or completed.stdout or "MarkItDown failed").strip()
    return ConversionResult(
        item=item,
        status="failed",
        message=message,
        returncode=completed.returncode,
    )


def _run_markitdown_api(item: PlanItem) -> ConversionResult:
    item.output.parent.mkdir(parents=True, exist_ok=True)
    try:
        from markitdown import MarkItDown

        result = MarkItDown().convert(item.source)
        item.output.write_text(getattr(result, "text_content", "") or "", encoding="utf-8")
    except Exception as exc:
        return ConversionResult(
            item=item,
            status="failed",
            message=str(exc) or "MarkItDown failed",
            returncode=1,
        )
    return ConversionResult(item=item, status="converted", message="converted", returncode=0)


def _should_use_markitdown_api(command: str) -> bool:
    args = _command_args(command)
    if not args or Path(args[0]).name.lower() not in {"markitdown", "markitdown.exe", "markitdown.cmd"}:
        return False
    return not _command_exists(_resolve_command(command)[0]) and _markitdown_api_available()


def _markitdown_api_available() -> bool:
    try:
        from markitdown import MarkItDown  # noqa: F401
    except Exception:
        return False
    return True


def _check_markitdown(command: str) -> DependencyCheck:
    args = _resolve_command(command)
    if _command_exists(args[0]):
        return DependencyCheck(name="MarkItDown", command=command, available=True, detail=f"found at {args[0]}")
    if _should_use_markitdown_api(command):
        return DependencyCheck(
            name="MarkItDown",
            command=command,
            available=True,
            detail="available through bundled Python package",
        )
    return DependencyCheck(name="MarkItDown", command=command, available=False, detail=f"{command} was not found")


def _reset_mermaid_filter_error_log(source_dir: Path) -> Path:
    err_path = source_dir / MERMAID_FILTER_ERR_NAME
    _remove_file(err_path)
    return err_path


def _ensure_mermaid_image_dir(project_root: Path, item: PlanItem) -> Path:
    image_key = hashlib.sha256(item.relative_source.encode("utf-8")).hexdigest()[:16]
    image_dir = project_root / PROJECT_DIR_NAME / MERMAID_IMAGE_DIR_NAME / image_key
    image_dir.mkdir(parents=True, exist_ok=True)
    return image_dir


def _mermaid_filter_error_text(err_path: Path) -> str:
    if not err_path.exists():
        return ""
    try:
        content = err_path.read_text(encoding="utf-8", errors="replace").strip()
    except OSError:
        return ""
    if not content:
        _remove_file(err_path)
        return ""
    return content


def _pandoc_failure_message(completed: subprocess.CompletedProcess[str], mermaid_error: str) -> str:
    parts = [
        text
        for text in (
            (completed.stderr or "").strip(),
            (completed.stdout or "").strip(),
            f"{MERMAID_FILTER_ERR_NAME}:\n{mermaid_error}" if mermaid_error else "",
        )
        if text
    ]
    return "\n\n".join(parts) or "Pandoc failed"


def _remove_file(path: Path) -> None:
    try:
        path.unlink()
    except FileNotFoundError:
        pass
    except OSError:
        pass


def _markitdown_command(item: PlanItem, settings: ConvertSettings) -> list[str]:
    markitdown_cmd = _resolve_command(settings.markitdown_cmd)
    return [
        markitdown_cmd[0],
        *markitdown_cmd[1:],
        str(item.source),
        "-o",
        str(item.output),
    ]


LUA_FILTER_CONTENT = r"""-- mermaid-fit.lua
local function url_decode(str)
  str = string.gsub(str, "+", " ")
  str = string.gsub(str, "%%(%x%x)", function(h)
    return string.char(tonumber(h, 16))
  end)
  return str
end

local function split_path(str)
  local t = {}
  local sep = ";"
  if not string.match(package.config, "^\\") then
    sep = ":"
  end
  for chunk in string.gmatch(str, "[^" .. sep .. "]+") do
    table.insert(t, chunk)
  end
  return t
end

local function resolve_filepath(src)
  local f = io.open(src, "rb")
  if f then
    f:close()
    return src
  end
  
  local decoded = url_decode(src)
  f = io.open(decoded, "rb")
  if f then
    f:close()
    return decoded
  end
  
  local paths_str = os.getenv("MD2DOC_RESOURCE_PATHS")
  if paths_str then
    local paths = split_path(paths_str)
    for _, path in ipairs(paths) do
      local full_path = path .. "/" .. decoded
      f = io.open(full_path, "rb")
      if f then
        f:close()
        return full_path
      end
      -- also try backslash
      local full_path_bs = path .. "\\" .. decoded
      f = io.open(full_path_bs, "rb")
      if f then
        f:close()
        return full_path_bs
      end
    end
  end
  return nil
end

local function get_png_dimensions(filepath)
  local file = io.open(filepath, "rb")
  if not file then return nil, nil end
  file:seek("set", 16)
  local bytes = file:read(8)
  file:close()
  if not bytes or #bytes < 8 then return nil, nil end
  local w = bytes:byte(1) * 16777216 + bytes:byte(2) * 65536 + bytes:byte(3) * 256 + bytes:byte(4)
  local h = bytes:byte(5) * 16777216 + bytes:byte(6) * 65536 + bytes:byte(7) * 256 + bytes:byte(8)
  return w, h
end

local function get_svg_dimensions(filepath)
  local file = io.open(filepath, "rb")
  if not file then return nil, nil end
  local content = file:read(4096)
  file:close()
  if not content then return nil, nil end
  
  local svg_tag = string.match(content, "<svg[^>]+>")
  if not svg_tag then return nil, nil end
  
  local w = string.match(svg_tag, 'width%s*=%s*[\'"]%s*(%d+%.?%d*)%s*[a-zA-Z%%]*[\'"]')
  local h = string.match(svg_tag, 'height%s*=%s*[\'"]%s*(%d+%.?%d*)%s*[a-zA-Z%%]*[\'"]')
  if w and h then
    return tonumber(w), tonumber(h)
  end
  
  local vx1, vy1, vx2, vy2 = string.match(svg_tag, 'viewBox%s*=%s*[\'"]%s*(-?%d+%.?%d*)%s+(-?%d+%.?%d*)%s+(%d+%.?%d*)%s+(%d+%.?%d*)%s*[\'"]')
  if vx2 and vy2 then
    return tonumber(vx2), tonumber(vy2)
  end
  return nil, nil
end

function Image(el)
  local filepath = resolve_filepath(el.src)
  if not filepath then return el end
  
  local lower_filepath = filepath:lower()
  local w, h = nil, nil
  if string.match(lower_filepath, "%.png$") then
    w, h = get_png_dimensions(filepath)
  elseif string.match(lower_filepath, "%.svg$") then
    w, h = get_svg_dimensions(filepath)
  end
  
  if w and h then
    local scale = 1.0
    local min_dpi = 0.0
    if string.match(filepath, "mermaid%-images") then
      scale = tonumber(os.getenv("MERMAID_FILTER_SCALE")) or 1.0
      if scale <= 0 then scale = 1.0 end
      if string.match(lower_filepath, "%.png$") then
        min_dpi = tonumber(os.getenv("MERMAID_FILTER_MIN_DPI")) or 0.0
        if min_dpi < 0 then min_dpi = 0.0 end
      end
    end
    
    local max_width_in = 6.0
    local max_height_in = 8.5
    
    local display_width_in = (w / scale) / 96.0
    local display_height_in = (h / scale) / 96.0
    
    local scale_factor = 1.0
    if display_width_in > max_width_in then
      scale_factor = math.min(scale_factor, max_width_in / display_width_in)
    end
    if display_height_in > max_height_in then
      scale_factor = math.min(scale_factor, max_height_in / display_height_in)
    end
    if min_dpi > 0 then
      local max_width_by_dpi = w / min_dpi
      local max_height_by_dpi = h / min_dpi
      if display_width_in > max_width_by_dpi then
        scale_factor = math.min(scale_factor, max_width_by_dpi / display_width_in)
      end
      if display_height_in > max_height_by_dpi then
        scale_factor = math.min(scale_factor, max_height_by_dpi / display_height_in)
      end
    end
    
    el.attributes['width'] = string.format("%.2fin", display_width_in * scale_factor)
    el.attributes['height'] = string.format("%.2fin", display_height_in * scale_factor)
  end
  return el
end
"""

def _ensure_mermaid_fit_lua(project_root: Path) -> Path:
    meta_dir = project_root / PROJECT_DIR_NAME
    meta_dir.mkdir(parents=True, exist_ok=True)
    lua_path = meta_dir / "mermaid-fit.lua"
    lua_path.write_text(LUA_FILTER_CONTENT, encoding="utf-8")
    return lua_path


FIGURE_CAPTION_LUA_TEMPLATE = r"""-- figure-caption.lua
local configured_prefix = __MD2DOC_FIGURE_PREFIX__
local configured_caption_position = __MD2DOC_FIGURE_CAPTION_POSITION__
local figure_count = 0

local function trim(str)
  if not str then return "" end
  return (str:gsub("^%s+", ""):gsub("%s+$", ""))
end

local function xml_escape(str)
  str = tostring(str or "")
  str = str:gsub("&", "&amp;")
  str = str:gsub("<", "&lt;")
  str = str:gsub(">", "&gt;")
  return str
end

local function attr_escape(str)
  str = xml_escape(str)
  str = str:gsub('"', "&quot;")
  return str
end

local function caption_to_text(caption)
  if not caption then return "" end
  local ok, text = pcall(pandoc.utils.stringify, caption)
  if ok then
    return trim(text)
  end
  if caption.long then
    return trim(pandoc.utils.stringify(caption.long))
  end
  return ""
end

local function sequence_id(prefix)
  local value = trim(prefix)
  if value == "" then value = "Figure" end
  return value:gsub("%s+", "_")
end

local function caption_block(caption_text)
  figure_count = figure_count + 1
  local prefix = trim(configured_prefix)
  if prefix == "" then prefix = "Figure" end
  local instr = "SEQ " .. sequence_id(prefix) .. " \\* ARABIC"
  local suffix = ""
  if caption_text ~= "" then
    suffix = " " .. caption_text
  end

  local xml = table.concat({
    '<w:p>',
    '<w:pPr><w:pStyle w:val="ImageCaption"/></w:pPr>',
    '<w:r><w:t xml:space="preserve">', xml_escape(prefix .. " "), '</w:t></w:r>',
    '<w:fldSimple w:instr="', attr_escape(instr), '">',
    '<w:r><w:t>', tostring(figure_count), '</w:t></w:r>',
    '</w:fldSimple>',
    '<w:r><w:t xml:space="preserve">', xml_escape(suffix), '</w:t></w:r>',
    '</w:p>',
  })
  return pandoc.RawBlock("openxml", xml)
end

local function captioned_blocks(image_blocks, caption_text)
  local position = trim(configured_caption_position):lower()
  local blocks = pandoc.List()

  if position == "above" then
    blocks:insert(caption_block(caption_text))
  end
  for _, block in ipairs(image_blocks) do
    blocks:insert(block)
  end
  if position ~= "above" then
    blocks:insert(caption_block(caption_text))
  end
  return blocks
end

function Figure(el)
  if not FORMAT:match("docx") then return nil end
  local caption_text = caption_to_text(el.caption)
  if caption_text == "" then return nil end
  return captioned_blocks(el.content, caption_text)
end

function Para(el)
  if not FORMAT:match("docx") then return nil end
  if #el.content ~= 1 then return nil end

  local image = el.content[1]
  if image.t ~= "Image" then return nil end

  local caption_text = caption_to_text(image.caption)
  if caption_text == "" then return nil end
  return captioned_blocks(pandoc.List({ pandoc.Para({ image }) }), caption_text)
end
"""


def _ensure_figure_caption_lua(project_root: Path, settings: ConvertSettings | None = None) -> Path:
    meta_dir = project_root / PROJECT_DIR_NAME
    meta_dir.mkdir(parents=True, exist_ok=True)
    lua_path = meta_dir / "figure-caption.lua"
    lua_path.write_text(_figure_caption_lua_content(settings or ConvertSettings()), encoding="utf-8")
    return lua_path


def _figure_caption_lua_content(settings: ConvertSettings) -> str:
    return (
        FIGURE_CAPTION_LUA_TEMPLATE.replace(
            "__MD2DOC_FIGURE_PREFIX__",
            _lua_string_literal(settings.figure_prefix.strip() or "Figure"),
        )
        .replace(
            "__MD2DOC_FIGURE_CAPTION_POSITION__",
            _lua_string_literal(settings.figure_caption_position.strip().lower() or "below"),
        )
    )


def _lua_string_literal(value: str) -> str:
    return json.dumps(value, ensure_ascii=False)


HR_TO_PAGEBREAK_LUA_CONTENT = r"""-- hr-to-pagebreak.lua
function HorizontalRule(el)
  if FORMAT:match('latex') then
    return pandoc.RawBlock('latex', '\\newpage')
  elseif FORMAT:match('html') then
    return pandoc.RawBlock('html', '<div style="page-break-after: always;"></div>')
  elseif FORMAT:match('docx') then
    return pandoc.RawBlock('openxml', '<w:p><w:r><w:br w:type="page"/></w:r></w:p>')
  else
    return el
  end
end
"""


def _ensure_hr_to_pagebreak_lua(project_root: Path) -> Path:
    meta_dir = project_root / PROJECT_DIR_NAME
    meta_dir.mkdir(parents=True, exist_ok=True)
    lua_path = meta_dir / "hr-to-pagebreak.lua"
    lua_path.write_text(HR_TO_PAGEBREAK_LUA_CONTENT, encoding="utf-8")
    return lua_path


def _shorten_windows_path(path: Path | str) -> str:
    """Return a short (8.3) alias for *path* when the normal path is long enough
    to risk tripping tools that are not long-path aware (notably Pandoc, whose
    Lua-filter and reference-doc loading fails with ENOENT past MAX_PATH). The
    path must already exist on disk. On non-Windows systems, or when shortening
    is unavailable or does not help, the original path string is returned."""
    text = str(path)
    if os.name != "nt" or len(text) < LONG_PATH_THRESHOLD:
        return text
    short = _windows_short_path(text)
    return short if short and len(short) < len(text) else text


def _shorten_output_path(path: Path) -> str:
    """Short alias for a not-yet-created output file: shorten the (existing)
    parent directory and re-attach the file name so Pandoc can write to it."""
    if os.name != "nt" or len(str(path)) < LONG_PATH_THRESHOLD:
        return str(path)
    parent_short = _shorten_windows_path(path.parent)
    return str(Path(parent_short) / path.name)


def _windows_short_path(text: str) -> str | None:
    try:
        import ctypes
        from ctypes import wintypes
    except Exception:
        return None
    try:
        get_short = ctypes.windll.kernel32.GetShortPathNameW
    except (AttributeError, OSError):
        return None
    get_short.argtypes = [wintypes.LPCWSTR, wintypes.LPWSTR, wintypes.DWORD]
    get_short.restype = wintypes.DWORD
    needed = get_short(text, None, 0)
    if needed == 0:
        return None
    buffer = ctypes.create_unicode_buffer(needed)
    if get_short(text, buffer, needed) == 0:
        return None
    return buffer.value


def _pandoc_resource_paths(item: PlanItem, project_root: Path) -> list[str]:
    return [_shorten_windows_path(item.source.parent), _shorten_windows_path(project_root)]


def _pandoc_command(project_root: Path, item: PlanItem, settings: ConvertSettings) -> list[str]:
    resource_path = os.pathsep.join(_pandoc_resource_paths(item, project_root))
    pandoc_cmd = _resolve_command(settings.pandoc_cmd)
    mermaid_filter_cmd = _resolve_command(settings.mermaid_filter_cmd)
    lua_filter_path = _ensure_mermaid_fit_lua(project_root)
    cmd = [
        pandoc_cmd[0],
        *pandoc_cmd[1:],
        _shorten_windows_path(item.source),
        "-o",
        _shorten_output_path(item.output),
        "--filter",
        mermaid_filter_cmd[0],
        *mermaid_filter_cmd[1:],
    ]
    if settings.output_format == "docx" and settings.figure_numbering:
        figure_lua_path = _ensure_figure_caption_lua(project_root, settings)
        cmd.append(f"--lua-filter={_shorten_windows_path(figure_lua_path)}")
    cmd.append(f"--lua-filter={_shorten_windows_path(lua_filter_path)}")
    if settings.hr_to_pagebreak:
        hr_lua_path = _ensure_hr_to_pagebreak_lua(project_root)
        cmd.append(f"--lua-filter={_shorten_windows_path(hr_lua_path)}")
    cmd.append(f"--resource-path={resource_path}")
    cmd.extend(_pandoc_format_args(project_root, item, settings))
    cmd.extend(settings.extra_pandoc_args)
    return cmd


def _pandoc_format_args(project_root: Path, item: PlanItem, settings: ConvertSettings) -> list[str]:
    args: list[str] = []
    if settings.toc:
        args.extend(["--toc", f"--toc-depth={max(1, int(settings.toc_depth))}"])
    if settings.number_sections:
        args.append("--number-sections")
    if settings.title_page:
        metadata = {
            "title": settings.title.strip() or item.source.stem,
            "subtitle": settings.subtitle.strip(),
            "author": settings.author.strip(),
            "date": settings.date.strip(),
        }
        for key, value in metadata.items():
            if value:
                args.extend(["--metadata", f"{key}={value}"])

    reference_docx = _effective_reference_docx(project_root, settings)
    if reference_docx:
        args.extend(["--reference-doc", _shorten_windows_path(reference_docx)])
    return args


def _mermaid_environment(settings: ConvertSettings, project_root: Path | None = None) -> dict[str, str]:
    env = {
        "MERMAID_FILTER_FORMAT": settings.mermaid_format or "png",
        "MERMAID_FILTER_THEME": settings.mermaid_theme or "default",
        "MERMAID_FILTER_BACKGROUND": settings.mermaid_background or "white",
    }
    if settings.mermaid_scale > 0:
        env["MERMAID_FILTER_SCALE"] = str(settings.mermaid_scale)
    if settings.mermaid_min_dpi > 0:
        env["MERMAID_FILTER_MIN_DPI"] = str(settings.mermaid_min_dpi)

    if _active_env_path("PUPPETEER_EXECUTABLE_PATH") is None:
        browser_path = _known_mermaid_browser_path()
        if browser_path:
            env["PUPPETEER_EXECUTABLE_PATH"] = str(browser_path)

    return env


def _known_mermaid_browser_path() -> Path | None:
    env_browser = _active_env_path("PUPPETEER_EXECUTABLE_PATH")
    if env_browser is not None:
        return env_browser
    return _known_mermaid_managed_browser_path() or _known_system_chromium_browser_path()


def _known_mermaid_managed_browser_path() -> Path | None:
    return _known_puppeteer_browser_path() or _known_playwright_chromium_path()


def _active_env_path(name: str) -> Path | None:
    value = os.environ.get(name, "").strip()
    if not value:
        return None
    path = Path(_strip_quotes(value)).expanduser()
    try:
        path = path.resolve()
    except OSError:
        return None
    return path if path.exists() else None


def _check_mermaid_browser_runtime() -> DependencyCheck:
    env_browser = _active_env_path("PUPPETEER_EXECUTABLE_PATH")
    if env_browser:
        return DependencyCheck(
            name="Mermaid browser",
            command="chromium",
            available=True,
            detail=f"found configured browser at {env_browser}",
        )

    env_value = os.environ.get("PUPPETEER_EXECUTABLE_PATH", "").strip()
    if env_value:
        return DependencyCheck(
            name="Mermaid browser",
            command="chromium",
            available=False,
            detail=f"PUPPETEER_EXECUTABLE_PATH points to a missing browser: {env_value}",
        )

    managed_browser = _known_mermaid_managed_browser_path()
    if managed_browser:
        return DependencyCheck(
            name="Mermaid browser",
            command="chromium",
            available=True,
            detail=f"found managed Chromium at {managed_browser}",
        )

    system_browser = _known_system_chromium_browser_path()
    if system_browser:
        return DependencyCheck(
            name="Mermaid browser",
            command="chromium",
            available=False,
            detail=(
                f"found system browser at {system_browser}, but managed Chromium is not installed. "
                "Run dependency setup to install Playwright Chromium."
            ),
        )

    return DependencyCheck(
        name="Mermaid browser",
        command="chromium",
        available=False,
        detail="No managed Chromium browser was found",
    )



def _generate_default_reference_docx_if_needed(project_root: Path, settings: ConvertSettings) -> None:
    target_path = project_root / PROJECT_DIR_NAME / "reference.docx"
    if target_path.exists():
        return
    try:
        pandoc_cmd = _resolve_command(settings.pandoc_cmd)
        completed = subprocess.run(
            pandoc_cmd + ["--print-default-data-file", "reference.docx"],
            capture_output=True,
            check=False,
            **hidden_subprocess_kwargs(),
        )
        if completed.returncode == 0:
            target_path.parent.mkdir(parents=True, exist_ok=True)
            target_path.write_bytes(completed.stdout)
    except Exception:
        pass


def _validate_settings(project_root: Path, settings: ConvertSettings) -> None:
    if settings.kind in (KIND_DOC2MD, KIND_QMD2PPT, KIND_HTML2PDF):
        return
    if settings.reference_docx:
        reference_docx = _resolve_project_path(project_root, settings.reference_docx)
        if not reference_docx.exists():
            if settings.reference_docx.replace("\\", "/") == f"{PROJECT_DIR_NAME}/reference.docx":
                _generate_default_reference_docx_if_needed(project_root, settings)
            if not reference_docx.exists():
                raise RuntimeError(f"Reference DOCX not found: {reference_docx}")
    if settings.table_borders not in {"template", "bordered", "plain"}:
        raise RuntimeError("Table borders must be one of: template, bordered, plain")
    if settings.mermaid_format not in {"png", "svg", "pdf"}:
        raise RuntimeError("Mermaid format must be one of: png, svg, pdf")
    if settings.figure_caption_position.strip().lower() not in {"above", "below"}:
        raise RuntimeError("Figure caption position must be one of: above, below")


def _effective_reference_docx(project_root: Path, settings: ConvertSettings) -> Path | None:
    if settings.output_format != "docx":
        return None
    if settings.reference_docx.strip():
        return _resolve_project_path(project_root, settings.reference_docx)
    if not _needs_generated_reference_docx(settings):
        return None
    return _ensure_generated_reference_docx(project_root, settings)


def _needs_generated_reference_docx(settings: ConvertSettings) -> bool:
    return bool(
        settings.default_font.strip()
        or settings.default_font_size > 0
        or settings.table_borders in {"bordered", "plain"}
    )


def _ensure_generated_reference_docx(project_root: Path, settings: ConvertSettings) -> Path:
    meta_dir = project_root / PROJECT_DIR_NAME
    meta_dir.mkdir(parents=True, exist_ok=True)
    reference_path = meta_dir / GENERATED_REFERENCE_DOCX
    meta_path = meta_dir / GENERATED_REFERENCE_META
    signature = _generated_reference_signature(settings)
    if reference_path.exists() and meta_path.exists():
        try:
            payload = json.loads(meta_path.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            payload = {}
        if payload.get("signature") == signature:
            return reference_path

    pandoc_cmd = _resolve_command(settings.pandoc_cmd)
    completed = subprocess.run(
        pandoc_cmd + ["--print-default-data-file", "reference.docx"],
        capture_output=True,
        check=False,
        **hidden_subprocess_kwargs(),
    )
    if completed.returncode != 0:
        message = (completed.stderr or b"Unable to generate reference DOCX").decode(
            "utf-8",
            errors="replace",
        )
        raise RuntimeError(message.strip())

    reference_path.write_bytes(completed.stdout)
    _patch_reference_docx(reference_path, settings)
    meta_path.write_text(
        json.dumps({"signature": signature}, indent=2, ensure_ascii=True),
        encoding="utf-8",
    )
    return reference_path


def _generated_reference_signature(settings: ConvertSettings) -> str:
    payload = {
        "default_font": settings.default_font,
        "default_font_size": settings.default_font_size,
        "table_borders": settings.table_borders,
    }
    return hashlib.sha256(json.dumps(payload, sort_keys=True).encode("utf-8")).hexdigest()


def _center_docx_images(docx_path: Path) -> None:
    with zipfile.ZipFile(docx_path, "r") as source:
        entries = {name: source.read(name) for name in source.namelist()}

    document_name = "word/document.xml"
    if document_name not in entries:
        return

    root = ET.fromstring(entries[document_name])
    changed = False
    for paragraph in root.findall(f".//{W}p"):
        if paragraph.find(f".//{W}drawing") is None and paragraph.find(f".//{W}pict") is None:
            continue
        ppr = _ensure_paragraph_properties(paragraph)
        _ensure_child(ppr, f"{W}jc").set(f"{W}val", "center")
        changed = True

    if not changed:
        return

    entries[document_name] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        temp_path = Path(tmp.name)
    try:
        with zipfile.ZipFile(temp_path, "w", zipfile.ZIP_DEFLATED) as target:
            for name, data in entries.items():
                target.writestr(name, data)
        shutil.move(str(temp_path), docx_path)
    finally:
        if temp_path.exists():
            temp_path.unlink()


def _patch_reference_docx(reference_path: Path, settings: ConvertSettings) -> None:
    with zipfile.ZipFile(reference_path, "r") as source:
        entries = {name: source.read(name) for name in source.namelist()}

    styles_name = "word/styles.xml"
    if styles_name not in entries:
        return

    root = ET.fromstring(entries[styles_name])
    if settings.default_font.strip() or settings.default_font_size > 0:
        _patch_default_run_style(root, settings.default_font.strip(), settings.default_font_size)
    if settings.table_borders in {"bordered", "plain"}:
        _patch_table_styles(root, settings.table_borders)
    entries[styles_name] = ET.tostring(root, encoding="utf-8", xml_declaration=True)

    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        temp_path = Path(tmp.name)
    try:
        with zipfile.ZipFile(temp_path, "w", zipfile.ZIP_DEFLATED) as target:
            for name, data in entries.items():
                target.writestr(name, data)
        shutil.move(str(temp_path), reference_path)
    finally:
        if temp_path.exists():
            temp_path.unlink()


def _patch_default_run_style(root: ET.Element, font_name: str, font_size: int) -> None:
    default_rpr = _default_run_properties(root)
    _set_run_properties(default_rpr, font_name, font_size)
    for style in root.findall(f"{W}style"):
        style_type = style.get(f"{W}type")
        style_id = style.get(f"{W}styleId")
        if style_type == "paragraph" and style_id == "Normal":
            _set_run_properties(_ensure_child(style, f"{W}rPr"), font_name, font_size)


def _default_run_properties(root: ET.Element) -> ET.Element:
    doc_defaults = _ensure_child(root, f"{W}docDefaults")
    rpr_default = _ensure_child(doc_defaults, f"{W}rPrDefault")
    return _ensure_child(rpr_default, f"{W}rPr")


def _set_run_properties(rpr: ET.Element, font_name: str, font_size: int) -> None:
    if font_name:
        r_fonts = _ensure_child(rpr, f"{W}rFonts")
        for key in ("ascii", "hAnsi", "eastAsia", "cs"):
            r_fonts.set(f"{W}{key}", font_name)
    if font_size > 0:
        half_points = str(int(font_size) * 2)
        _ensure_child(rpr, f"{W}sz").set(f"{W}val", half_points)
        _ensure_child(rpr, f"{W}szCs").set(f"{W}val", half_points)


def _patch_table_styles(root: ET.Element, border_mode: str) -> None:
    for style in root.findall(f"{W}style"):
        if style.get(f"{W}type") != "table":
            continue
        tbl_pr = _ensure_child(style, f"{W}tblPr")
        existing = tbl_pr.find(f"{W}tblBorders")
        if existing is not None:
            tbl_pr.remove(existing)
        if border_mode == "bordered":
            tbl_pr.append(_table_borders_element())


def _table_borders_element() -> ET.Element:
    borders = ET.Element(f"{W}tblBorders")
    for name in ("top", "left", "bottom", "right", "insideH", "insideV"):
        border = ET.SubElement(borders, f"{W}{name}")
        border.set(f"{W}val", "single")
        border.set(f"{W}sz", "4")
        border.set(f"{W}space", "0")
        border.set(f"{W}color", "auto")
    return borders


def _ensure_child(parent: ET.Element, tag: str) -> ET.Element:
    child = parent.find(tag)
    if child is None:
        child = ET.SubElement(parent, tag)
    return child


def _ensure_paragraph_properties(paragraph: ET.Element) -> ET.Element:
    ppr = paragraph.find(f"{W}pPr")
    if ppr is None:
        ppr = ET.Element(f"{W}pPr")
        paragraph.insert(0, ppr)
    return ppr


def _resolve_project_path(project_root: Path, path_value: str) -> Path:
    candidate = Path(path_value).expanduser()
    if candidate.is_absolute():
        return candidate.resolve()
    return (project_root / candidate).resolve()


def _file_stat_signature(path_value: str, project_root: Path | None = None) -> dict[str, int | str] | None:
    if not path_value:
        return None
    path = Path(path_value).expanduser()
    if not path.is_absolute() and project_root:
        path = (project_root / path).resolve()
    else:
        path = path.resolve()
    if not path.exists():
        return {"path": str(path), "missing": 1}
    stat = path.stat()
    return {"path": str(path), "size": stat.st_size, "mtime_ns": stat.st_mtime_ns}


def _check_html_pdf_runtime() -> DependencyCheck:
    name = "Playwright/Chromium"
    if importlib.util.find_spec("playwright") is None:
        return DependencyCheck(
            name=name,
            command="playwright",
            available=False,
            detail="Playwright Python package was not found",
        )

    system_browser = _known_html_pdf_browser_path()
    if system_browser:
        return DependencyCheck(
            name=name,
            command="playwright",
            available=True,
            detail=f"Playwright installed; found browser at {system_browser}",
        )

    try:
        from playwright.sync_api import sync_playwright

        with sync_playwright() as playwright:
            bundled = Path(playwright.chromium.executable_path)
    except Exception as exc:
        return DependencyCheck(
            name=name,
            command="playwright",
            available=False,
            detail=f"Playwright runtime failed: {exc}",
        )

    if bundled.exists():
        return DependencyCheck(
            name=name,
            command="playwright",
            available=True,
            detail=f"Playwright installed; found Chromium at {bundled}",
        )

    return DependencyCheck(
        name=name,
        command="playwright",
        available=False,
        detail="Playwright is installed, but no Chromium, Edge, or Chrome browser was found",
    )


def _check_command(name: str, command: str, *, allow_version_failure: bool = False) -> DependencyCheck:
    args = _resolve_command(command)
    if not _command_exists(args[0]):
        return DependencyCheck(name=name, command=command, available=False, detail=f"{command} was not found")
    if allow_version_failure:
        return DependencyCheck(name=name, command=command, available=True, detail=f"found at {args[0]}")

    try:
        completed = subprocess.run(
            args + ["--version"],
            capture_output=True,
            text=True,
            check=False,
            **hidden_subprocess_kwargs(),
        )
    except FileNotFoundError:
        return DependencyCheck(name=name, command=command, available=False, detail=f"{command} was not found")
    except OSError as exc:
        return DependencyCheck(name=name, command=command, available=False, detail=str(exc))

    output = (completed.stdout or completed.stderr).strip()
    first_line = output.splitlines()[0] if output else f"{command} returned {completed.returncode}"
    return DependencyCheck(
        name=name,
        command=command,
        available=completed.returncode == 0,
        detail=first_line,
    )


def _command_args(command: str) -> list[str]:
    path = Path(_strip_quotes(command))
    if path.exists():
        return [str(path)]
    return [_strip_quotes(arg) for arg in shlex.split(command, posix=os.name != "nt")]


def _resolve_command(command: str) -> list[str]:
    args = _command_args(command)
    if not args:
        return [command]
    executable = args[0]
    if _command_exists(executable):
        resolved = shutil.which(executable) or executable
        return [resolved, *args[1:]]

    if os.name == "nt":
        for candidate in _windows_tool_candidates(executable):
            if candidate.exists():
                return [str(candidate), *args[1:]]
    return args


def _command_exists(executable: str) -> bool:
    return Path(executable).exists() or shutil.which(executable) is not None


def _strip_quotes(value: str) -> str:
    return value.strip().strip('"').strip("'")


def _windows_tool_candidates(executable: str) -> list[Path]:
    name = Path(executable).name.lower()
    stem = Path(name).stem
    if stem == "pandoc":
        return _windows_pandoc_candidates()
    if stem == "mermaid-filter":
        return _windows_mermaid_filter_candidates()
    if stem == "quarto":
        return _windows_quarto_candidates()
    return []


def _windows_pandoc_candidates() -> list[Path]:
    candidates: list[Path] = []
    local_app_data = _env_path("LOCALAPPDATA")
    program_files = _env_path("ProgramFiles")
    program_files_x86 = _env_path("ProgramFiles(x86)")

    candidates.extend(
        path
        for path in [
            local_app_data / "Pandoc" / "pandoc.exe" if local_app_data else None,
            local_app_data / "Microsoft" / "WinGet" / "Links" / "pandoc.exe" if local_app_data else None,
            program_files / "Pandoc" / "pandoc.exe" if program_files else None,
            program_files_x86 / "Pandoc" / "pandoc.exe" if program_files_x86 else None,
        ]
        if path is not None
    )

    for install_location in _windows_registry_tool_locations("pandoc"):
        candidates.extend(_find_named_files(install_location, "pandoc.exe", limit=5))

    if local_app_data:
        winget_packages = local_app_data / "Microsoft" / "WinGet" / "Packages"
        for package_dir in _safe_glob(winget_packages, "JohnMacFarlane.Pandoc*"):
            candidates.extend(_find_named_files(package_dir, "pandoc.exe", limit=5))
    return _dedupe_paths(candidates)


def _windows_mermaid_filter_candidates() -> list[Path]:
    candidates: list[Path] = []
    app_data = _env_path("APPDATA")
    if app_data:
        npm_dir = app_data / "npm"
        candidates.extend(
            [
                npm_dir / "mermaid-filter.cmd",
                npm_dir / "mermaid-filter.exe",
                npm_dir / "mermaid-filter.ps1",
                npm_dir / "mermaid-filter",
            ]
        )
    return _dedupe_paths(candidates)


def _windows_registry_tool_locations(display_name_fragment: str) -> list[Path]:
    if os.name != "nt":
        return []
    try:
        import winreg
    except ImportError:
        return []

    roots = [
        (winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Uninstall"),
        (winreg.HKEY_LOCAL_MACHINE, r"Software\Microsoft\Windows\CurrentVersion\Uninstall"),
        (winreg.HKEY_LOCAL_MACHINE, r"Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"),
    ]
    locations: list[Path] = []
    needle = display_name_fragment.lower()
    for root, subkey in roots:
        try:
            with winreg.OpenKey(root, subkey) as parent:
                for index in range(winreg.QueryInfoKey(parent)[0]):
                    try:
                        child_name = winreg.EnumKey(parent, index)
                        with winreg.OpenKey(parent, child_name) as child:
                            display_name = _registry_value(winreg, child, "DisplayName").lower()
                            if needle not in display_name:
                                continue
                            install_location = _registry_value(winreg, child, "InstallLocation")
                            display_icon = _registry_value(winreg, child, "DisplayIcon")
                    except OSError:
                        continue
                    if install_location:
                        locations.append(Path(_strip_quotes(install_location)))
                    if display_icon:
                        icon_path = Path(_strip_quotes(display_icon.split(",")[0]))
                        locations.append(icon_path.parent if icon_path.suffix else icon_path)
        except OSError:
            continue
    return _dedupe_paths(locations)


def _registry_value(winreg_module, key, name: str) -> str:
    try:
        value, _kind = winreg_module.QueryValueEx(key, name)
    except OSError:
        return ""
    return str(value)


def _env_path(name: str) -> Path | None:
    value = os.environ.get(name)
    return Path(value) if value else None


def _known_html_pdf_browser_path() -> Path | None:
    system_browser = _known_system_chromium_browser_path()
    if system_browser:
        return system_browser
    return _known_playwright_chromium_path()


def _known_system_chromium_browser_path() -> Path | None:
    if os.name == "nt":
        for path in _windows_html_pdf_browser_candidates():
            if path.exists():
                return path

    for command in (
        "msedge",
        "microsoft-edge",
        "google-chrome",
        "chrome",
        "chromium",
        "chromium-browser",
    ):
        path = shutil.which(command)
        if path:
            return Path(path)
    return None


def _known_playwright_chromium_path() -> Path | None:
    if importlib.util.find_spec("playwright") is not None:
        try:
            from playwright.sync_api import sync_playwright

            with sync_playwright() as playwright:
                bundled = Path(playwright.chromium.executable_path)
            if bundled.exists():
                return bundled
        except Exception:
            pass
    return _known_ms_playwright_chromium_path()


def _known_ms_playwright_chromium_path() -> Path | None:
    roots: list[Path] = []
    browsers_path = os.environ.get("PLAYWRIGHT_BROWSERS_PATH", "").strip()
    if browsers_path:
        roots.append(Path(_strip_quotes(browsers_path)).expanduser())
    local_app_data = _env_path("LOCALAPPDATA")
    if local_app_data:
        roots.append(local_app_data / "ms-playwright")
    try:
        roots.append(Path.home() / ".cache" / "ms-playwright")
    except RuntimeError:
        pass
    filenames = ("chrome.exe", "chromium.exe") if os.name == "nt" else ("chrome", "chromium")
    for root in _dedupe_paths(roots):
        if not root.exists():
            continue
        for chromium_dir in sorted(root.glob("chromium-*"), reverse=True):
            for filename in filenames:
                for path in _find_named_files(chromium_dir, filename, limit=5):
                    if path.exists():
                        return path
    return None


def _known_puppeteer_browser_path() -> Path | None:
    roots: list[Path] = []
    cache_dir = os.environ.get("PUPPETEER_CACHE_DIR", "").strip()
    if cache_dir:
        roots.append(Path(_strip_quotes(cache_dir)).expanduser())
    try:
        roots.append(Path.home() / ".cache" / "puppeteer")
    except RuntimeError:
        pass
    local_app_data = _env_path("LOCALAPPDATA")
    if local_app_data:
        roots.append(local_app_data / "puppeteer")

    filenames = ("chrome.exe", "chromium.exe") if os.name == "nt" else ("chrome", "chromium")
    for root in _dedupe_paths(roots):
        for filename in filenames:
            for path in _find_named_files(root, filename, limit=10):
                if path.exists():
                    return path
    return None


def _windows_html_pdf_browser_candidates() -> list[Path]:
    candidates: list[Path] = []
    program_files = _env_path("ProgramFiles")
    program_files_x86 = _env_path("ProgramFiles(x86)")
    local_app_data = _env_path("LOCALAPPDATA")

    candidates.extend(
        path
        for path in [
            program_files / "Microsoft" / "Edge" / "Application" / "msedge.exe" if program_files else None,
            program_files_x86 / "Microsoft" / "Edge" / "Application" / "msedge.exe" if program_files_x86 else None,
            local_app_data / "Microsoft" / "Edge" / "Application" / "msedge.exe" if local_app_data else None,
            program_files / "Google" / "Chrome" / "Application" / "chrome.exe" if program_files else None,
            program_files_x86 / "Google" / "Chrome" / "Application" / "chrome.exe" if program_files_x86 else None,
            local_app_data / "Google" / "Chrome" / "Application" / "chrome.exe" if local_app_data else None,
        ]
        if path is not None
    )
    return _dedupe_paths(candidates)


def _safe_glob(root: Path, pattern: str) -> list[Path]:
    if not root.exists():
        return []
    return list(root.glob(pattern))


def _find_named_files(root: Path, filename: str, *, limit: int) -> list[Path]:
    if root.is_file():
        return [root] if root.name.lower() == filename.lower() else []
    if not root.exists():
        return []
    matches: list[Path] = []
    for path in root.rglob(filename):
        matches.append(path)
        if len(matches) >= limit:
            break
    return matches


def _dedupe_paths(paths: Iterable[Path]) -> list[Path]:
    seen: set[str] = set()
    deduped: list[Path] = []
    for path in paths:
        key = str(path).lower()
        if key in seen:
            continue
        seen.add(key)
        deduped.append(path)
    return deduped


def _decide_action(
    *,
    settings: ConvertSettings,
    record: dict | None,
    fingerprint: FileFingerprint,
    output: Path,
    signature: str,
) -> tuple[PlanAction, str]:
    if settings.force:
        return "convert", "forced"
    if not settings.skip_unchanged:
        return "convert", "skip disabled"
    if not output.exists():
        return "convert", "output missing"
    if not record:
        output_mtime = output.stat().st_mtime_ns
        if output_mtime >= fingerprint.mtime_ns:
            if _requires_recorded_settings_to_skip(settings):
                return "convert", "conversion settings untracked"
            return "skip", "output is newer than source"
        return "convert", "no history and source is newer"
    if record.get("source_sha256") != fingerprint.sha256:
        return "convert", "source changed"
    if record.get("settings_signature") != signature:
        return "convert", "conversion settings changed"
    return "skip", "unchanged"


def _requires_recorded_settings_to_skip(settings: ConvertSettings) -> bool:
    return bool(
        settings.extra_pandoc_args
        or settings.toc
        or settings.title_page
        or settings.number_sections
        or settings.reference_docx.strip()
        or settings.default_font.strip()
        or settings.default_font_size > 0
        or settings.table_borders in {"bordered", "plain"}
        or settings.mermaid_format != "png"
        or settings.mermaid_theme != "default"
        or settings.mermaid_background != "white"
        or settings.mermaid_scale != 3.0
        or settings.mermaid_min_dpi != 450.0
    )


def _is_same_or_child(path: Path, maybe_parent: Path | None) -> bool:
    if maybe_parent is None:
        return False
    try:
        path.resolve().relative_to(maybe_parent)
        return True
    except ValueError:
        return path.resolve() == maybe_parent


def _run_quarto(
    item: PlanItem,
    settings: ConvertSettings,
    cancel_event: threading.Event | None = None,
) -> ConversionResult:
    item.output.parent.mkdir(parents=True, exist_ok=True)
    import uuid
    temp_filename = f"qmd2ppt_temp_{uuid.uuid4().hex}.pptx"
    temp_path = item.source.parent / temp_filename

    missing_reference_message = _missing_quarto_reference_doc_message(item.source)
    if missing_reference_message:
        return ConversionResult(
            item=item,
            status="failed",
            message=missing_reference_message,
            returncode=1,
        )
    
    quarto_cmd = _resolve_command(settings.quarto_cmd)
    cmd = [
        quarto_cmd[0],
        *quarto_cmd[1:],
        "render",
        item.source.name,
        "--to",
        "pptx",
        "-o",
        temp_filename
    ]
    
    completed = _run_subprocess_with_cancel(
        cmd,
        cwd=item.source.parent,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        check=False,
        cancel_event=cancel_event,
        **hidden_subprocess_kwargs(),
    )
    
    if completed.returncode == 0 and temp_path.exists():
        try:
            if item.output.exists():
                item.output.unlink()
            shutil.move(str(temp_path), str(item.output))
            return ConversionResult(item=item, status="converted", message="converted", returncode=0)
        except Exception as exc:
            return ConversionResult(
                item=item,
                status="failed",
                message=f"Failed to move output file: {exc}",
                returncode=1,
            )
        finally:
            if temp_path.exists():
                temp_path.unlink()
                
    if temp_path.exists():
        temp_path.unlink()
        
    message = (completed.stderr or completed.stdout or "Quarto failed").strip()
    return ConversionResult(
        item=item,
        status="failed",
        message=message,
        returncode=completed.returncode or 1,
    )


def _missing_quarto_reference_doc_message(qmd_path: Path) -> str:
    missing: list[tuple[str, Path]] = []
    for reference_doc in _quarto_reference_docs(qmd_path):
        if _looks_like_external_reference(reference_doc):
            continue
        reference_path = Path(reference_doc).expanduser()
        if not reference_path.is_absolute():
            reference_path = qmd_path.parent / reference_path
        reference_path = reference_path.resolve()
        if not reference_path.exists():
            missing.append((reference_doc, reference_path))

    if not missing:
        return ""

    lines = ["Reference PPTX not found for QMD conversion:"]
    for raw_value, expected_path in missing:
        lines.append(f"- reference-doc: {raw_value}")
        lines.append(f"  expected path: {expected_path}")
    lines.append("Quarto resolves relative reference-doc paths from the QMD folder.")
    lines.append(
        "Copy or rename the PPTX to that path, change reference-doc to a valid path, "
        "or remove reference-doc to use Quarto's default PPTX template."
    )
    return "\n".join(lines)


def _quarto_reference_docs(qmd_path: Path) -> list[str]:
    front_matter = _quarto_front_matter_lines(qmd_path)
    references: list[str] = []
    for line in front_matter:
        stripped = line.strip()
        if not stripped or stripped.startswith("#"):
            continue
        key, separator, value = stripped.partition(":")
        if separator and key.strip() == "reference-doc":
            reference_doc = _clean_yaml_scalar(value)
            if reference_doc:
                references.append(reference_doc)
    return references


def _quarto_front_matter_lines(qmd_path: Path) -> list[str]:
    try:
        text = qmd_path.read_text(encoding="utf-8-sig")
    except UnicodeDecodeError:
        text = qmd_path.read_text(encoding="utf-8", errors="replace")
    lines = text.splitlines()
    if not lines or lines[0].strip() != "---":
        return []
    for index, line in enumerate(lines[1:], start=1):
        if line.strip() in {"---", "..."}:
            return lines[1:index]
    return []


def _clean_yaml_scalar(value: str) -> str:
    value = value.strip()
    if not value:
        return ""

    quote = value[0]
    if quote in {"'", '"'}:
        end_index = value.find(quote, 1)
        if end_index != -1:
            return value[1:end_index].strip()
        return value[1:].strip()

    comment_index = _yaml_comment_index(value)
    if comment_index != -1:
        value = value[:comment_index]
    return value.strip()


def _yaml_comment_index(value: str) -> int:
    for index, char in enumerate(value):
        if char == "#" and (index == 0 or value[index - 1].isspace()):
            return index
    return -1


def _looks_like_external_reference(value: str) -> bool:
    return "://" in value or value.startswith("data:")


def _windows_quarto_candidates() -> list[Path]:
    candidates: list[Path] = []
    program_files = _env_path("ProgramFiles")
    program_files_x86 = _env_path("ProgramFiles(x86)")
    local_app_data = _env_path("LOCALAPPDATA")
    
    candidates.extend(
        path
        for path in [
            program_files / "Quarto" / "bin" / "quarto.exe" if program_files else None,
            program_files_x86 / "Quarto" / "bin" / "quarto.exe" if program_files_x86 else None,
            local_app_data / "Programs" / "Quarto" / "bin" / "quarto.exe" if local_app_data else None,
        ]
        if path is not None
    )
    return _dedupe_paths(candidates)
