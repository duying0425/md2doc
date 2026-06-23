from __future__ import annotations

from dataclasses import dataclass, field
import json
import os
from pathlib import Path
from typing import Any


PROJECT_DIR_NAME = ".md2doc"
PROJECT_CONFIG_NAME = "project.json"

KIND_MD2DOC = "md2doc"
KIND_DOC2MD = "doc2md"
KIND_QMD2PPT = "qmd2ppt"
KIND_HTML2PDF = "html2pdf"
VALID_KINDS = {KIND_MD2DOC, KIND_DOC2MD, KIND_QMD2PPT, KIND_HTML2PDF}


def default_output_format(kind: str) -> str:
    if kind == KIND_DOC2MD:
        return "md"
    if kind == KIND_QMD2PPT:
        return "pptx"
    if kind == KIND_HTML2PDF:
        return "pdf"
    return "docx"


def normalize_kind(value: object) -> str:
    kind = str(value or KIND_MD2DOC)
    return kind if kind in VALID_KINDS else KIND_MD2DOC


def app_data_dir() -> Path:
    base = os.environ.get("APPDATA")
    if base:
        return Path(base) / "md2doc"
    return Path.home() / ".md2doc"


def registry_path() -> Path:
    return app_data_dir() / "projects.json"


@dataclass
class ProjectConfig:
    name: str
    root: Path
    kind: str = KIND_MD2DOC
    output_dir: str = "."
    output_format: str = "docx"
    recursive: bool = True
    extra_pandoc_args: list[str] = field(default_factory=list)
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

    @property
    def meta_dir(self) -> Path:
        return self.root / PROJECT_DIR_NAME

    @property
    def config_path(self) -> Path:
        return self.meta_dir / PROJECT_CONFIG_NAME

    @property
    def output_path(self) -> Path:
        candidate = Path(self.output_dir)
        if candidate.is_absolute():
            return candidate
        return self.root / candidate

    def to_dict(self) -> dict[str, Any]:
        return {
            "name": self.name,
            "root": str(self.root),
            "kind": self.kind,
            "output_dir": self.output_dir,
            "output_format": self.output_format,
            "recursive": self.recursive,
            "extra_pandoc_args": list(self.extra_pandoc_args),
            "toc": self.toc,
            "toc_depth": self.toc_depth,
            "title_page": self.title_page,
            "title": self.title,
            "subtitle": self.subtitle,
            "author": self.author,
            "date": self.date,
            "number_sections": self.number_sections,
            "reference_docx": self.reference_docx,
            "default_font": self.default_font,
            "default_font_size": self.default_font_size,
            "table_borders": self.table_borders,
            "mermaid_format": self.mermaid_format,
            "mermaid_theme": self.mermaid_theme,
            "mermaid_background": self.mermaid_background,
            "mermaid_scale": self.mermaid_scale,
            "mermaid_min_dpi": self.mermaid_min_dpi,
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "ProjectConfig":
        root = Path(data["root"]).expanduser().resolve()
        kind = normalize_kind(data.get("kind"))
        if kind == KIND_DOC2MD:
            # doc2md projects always emit Markdown; ignore any stale Word/HTML/PDF format.
            output_format = "md"
        elif kind == KIND_QMD2PPT:
            output_format = "pptx"
        elif kind == KIND_HTML2PDF:
            output_format = "pdf"
        else:
            output_format = str(data.get("output_format") or "docx")
        scale_val = data.get("mermaid_scale")
        try:
            mermaid_scale = float(scale_val) if scale_val is not None else 3.0
            if mermaid_scale == 0.0:
                mermaid_scale = 3.0
        except (ValueError, TypeError):
            mermaid_scale = 3.0
        min_dpi_val = data.get("mermaid_min_dpi")
        try:
            mermaid_min_dpi = float(min_dpi_val) if min_dpi_val is not None else 450.0
            if mermaid_min_dpi < 0.0:
                mermaid_min_dpi = 450.0
        except (ValueError, TypeError):
            mermaid_min_dpi = 450.0

        return cls(
            name=str(data.get("name") or root.name),
            root=root,
            kind=kind,
            output_dir=str(data.get("output_dir") or "."),
            output_format=output_format,
            recursive=bool(data.get("recursive", True)),
            extra_pandoc_args=list(data.get("extra_pandoc_args") or []),
            toc=bool(data.get("toc", False)),
            toc_depth=int(data.get("toc_depth") or 3),
            title_page=bool(data.get("title_page", False)),
            title=str(data.get("title") or ""),
            subtitle=str(data.get("subtitle") or ""),
            author=str(data.get("author") or ""),
            date=str(data.get("date") or ""),
            number_sections=bool(data.get("number_sections", False)),
            reference_docx=str(data.get("reference_docx") or ""),
            default_font=str(data.get("default_font") or ""),
            default_font_size=int(data.get("default_font_size") or 0),
            table_borders=str(data.get("table_borders") or "template"),
            mermaid_format=str(data.get("mermaid_format") or "png"),
            mermaid_theme=str(data.get("mermaid_theme") or "default"),
            mermaid_background=str(data.get("mermaid_background") or "white"),
            mermaid_scale=mermaid_scale,
            mermaid_min_dpi=mermaid_min_dpi,
        )

    def save(self) -> None:
        self.meta_dir.mkdir(parents=True, exist_ok=True)
        self.config_path.write_text(
            json.dumps(self.to_dict(), indent=2, ensure_ascii=True),
            encoding="utf-8",
        )


def create_project(
    root: Path | str,
    name: str | None = None,
    kind: str = KIND_MD2DOC,
) -> ProjectConfig:
    resolved = Path(root).expanduser().resolve()
    resolved.mkdir(parents=True, exist_ok=True)
    kind = normalize_kind(kind)
    config = ProjectConfig(
        name=name or resolved.name,
        root=resolved,
        kind=kind,
        output_format=default_output_format(kind),
    )
    config.save()
    ProjectRegistry().add(config)
    return config


def load_project(root: Path | str) -> ProjectConfig:
    resolved = Path(root).expanduser().resolve()
    config_file = resolved / PROJECT_DIR_NAME / PROJECT_CONFIG_NAME
    if not config_file.exists():
        return create_project(resolved)
    data = json.loads(config_file.read_text(encoding="utf-8"))
    config = ProjectConfig.from_dict(data)
    # Clean up legacy configs on disk: persist a normalized copy when the stored
    # data predates the project kind, carries a stale format for its kind,
    # or has Mermaid sizing defaults that need migration.
    stored_scale = data.get("mermaid_scale")
    try:
        scale_needs_migration = stored_scale is None or float(stored_scale) == 0.0
    except (ValueError, TypeError):
        scale_needs_migration = True
    stored_min_dpi = data.get("mermaid_min_dpi")
    try:
        min_dpi_needs_migration = stored_min_dpi is None or float(stored_min_dpi) < 0.0
    except (ValueError, TypeError):
        min_dpi_needs_migration = True

    if (
        data.get("kind") != config.kind
        or data.get("output_format") != config.output_format
        or scale_needs_migration
        or min_dpi_needs_migration
    ):
        config.save()
    ProjectRegistry().add(config)
    return config


class ProjectRegistry:
    def __init__(self, path: Path | None = None) -> None:
        self.path = path or registry_path()

    def list(self) -> list[ProjectConfig]:
        if not self.path.exists():
            return []
        try:
            payload = json.loads(self.path.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            return []
        projects: list[ProjectConfig] = []
        for item in payload.get("projects", []):
            try:
                project = ProjectConfig.from_dict(item)
            except (KeyError, TypeError, ValueError):
                continue
            if project.root.exists():
                projects.append(project)
        projects.sort(key=lambda p: (p.name.lower(), str(p.root).lower()))
        return projects

    def add(self, config: ProjectConfig) -> None:
        projects = [project for project in self.list() if project.root != config.root]
        projects.append(config)
        self._save(projects)

    def remove(self, root: Path | str) -> None:
        resolved = Path(root).expanduser().resolve()
        projects = [project for project in self.list() if project.root != resolved]
        self._save(projects)

    def _save(self, projects: list[ProjectConfig]) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        payload = {"projects": [project.to_dict() for project in projects]}
        self.path.write_text(
            json.dumps(payload, indent=2, ensure_ascii=True),
            encoding="utf-8",
        )
