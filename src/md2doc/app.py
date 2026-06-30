from __future__ import annotations

import ctypes
from dataclasses import replace
import os
from pathlib import Path
import queue
import shlex
import threading
import tkinter as tk
import traceback
from tkinter import filedialog, font as tkfont, messagebox, simpledialog, ttk

# Patch tkinter.Variable.__del__ to prevent crashes when variables are garbage
# collected on background threads (causing RuntimeError: main thread is not in main loop).
_original_variable_del = tk.Variable.__del__

def _safe_variable_del(self) -> None:
    try:
        if threading and threading.current_thread() is threading.main_thread():
            _original_variable_del(self)
    except Exception:
        pass

tk.Variable.__del__ = _safe_variable_del

from .converter import (
    ConvertSettings,
    ConversionResult,
    ConversionCancelledError,
    PlanItem,
    check_dependencies,
    missing_dependency_message,
    plan_conversions,
    run_conversions,
    scan_source_files,
    settings_from_project,
)
from .dependencies import ensure_startup_dependencies
from .project import (
    KIND_DOC2MD,
    KIND_HTML2PDF,
    KIND_MD2DOC,
    KIND_QMD2PPT,
    ProjectConfig,
    ProjectRegistry,
    app_data_dir,
    create_project,
    load_project,
)
from . import __version__


SCAN_TABLE_BATCH_SIZE = 250


class ProjectState:
    def __init__(self) -> None:
        self.plan_by_id: dict[str, PlanItem] = {}
        self.iid_by_source: dict[str, str] = {}
        self.item_states: dict[str, tuple[str, str, str]] = {}  # source_path -> (state, reason, tag)
        self.log_content: str = ""
        self.conversion_total: int = 0
        self.conversion_done: int = 0
        self.converted_count: int = 0
        self.skipped_count: int = 0
        self.failed_count: int = 0
        self.progress_value: float = 0.0
        self.status_text: str = "Ready"
        self.scan_active: bool = False
        self.scan_generation: int = 0
        self.scan_worker: threading.Thread | None = None
        self.kind: str = ""
        self.conversion_active: bool = False



class Md2DocApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.ui_scale = _detect_ui_scale(self)
        self._configure_display()
        self.title("md2doc")
        self.geometry(f"{self._px(1120)}x{self._px(720)}")
        self.minsize(self._px(900), self._px(560))

        self.registry = ProjectRegistry()
        self.current_project: ProjectConfig | None = None
        self.project_states: dict[str, ProjectState] = {}
        self.plan_by_id: dict[str, PlanItem] = {}
        self.iid_by_source: dict[str, str] = {}
        self.event_queue: queue.Queue[object] = queue.Queue()
        self.worker: threading.Thread | None = None
        self.scan_worker: threading.Thread | None = None
        self.scan_generation = 0
        self.scan_active = False
        self.conversion_total = 0
        self.conversion_done = 0
        self.converted_count = 0
        self.skipped_count = 0
        self.failed_count = 0
        self.cancel_event = threading.Event()

        self._build_ui()
        self._load_projects()
        self.after(150, self._poll_events)
        self.protocol("WM_DELETE_WINDOW", self._on_closing)

    def report_callback_exception(self, exc_type, exc_value, exc_tb) -> None:
        message = "".join(traceback.format_exception(exc_type, exc_value, exc_tb))
        _write_error_log(message)
        try:
            self._append_log(message)
            self.status_var.set("Unexpected error")
        except Exception:
            pass
        messagebox.showerror("Unexpected error", str(exc_value), parent=self)

    def _build_ui(self) -> None:
        self.columnconfigure(1, weight=1)
        self.rowconfigure(0, weight=1)

        sidebar = ttk.Frame(self, padding=self._px(10))
        sidebar.grid(row=0, column=0, sticky="nsew")
        sidebar.rowconfigure(1, weight=1)

        ttk.Label(sidebar, text="Projects").grid(row=0, column=0, columnspan=2, sticky="w")
        self.project_list = tk.Listbox(sidebar, width=30, height=20, exportselection=False)
        self.project_list.configure(font=self.default_font, activestyle="none")
        self.project_list.grid(row=1, column=0, columnspan=2, sticky="nsew", pady=self._pad(8, 8))
        self.project_list.bind("<<ListboxSelect>>", self._on_project_selected)
        ttk.Button(sidebar, text="New", command=self._new_project).grid(
            row=2,
            column=0,
            sticky="ew",
            padx=self._pad(0, 4),
        )
        ttk.Button(sidebar, text="Remove", command=self._remove_project).grid(
            row=2,
            column=1,
            sticky="ew",
            padx=self._pad(4, 0),
        )

        try:
            from .build_info import BUILD_TIME
        except ImportError:
            BUILD_TIME = "N/A"

        info_frame = ttk.Frame(sidebar)
        info_frame.grid(row=3, column=0, columnspan=2, sticky="ew", pady=self._pad(12, 0))

        ttk.Label(
            info_frame,
            text=f"Version: {__version__}",
            foreground="#666666",
            font=self.mono_font,
        ).grid(row=0, column=0, sticky="w")

        ttk.Label(
            info_frame,
            text=f"Build: {BUILD_TIME}",
            foreground="#666666",
            font=self.mono_font,
        ).grid(row=1, column=0, sticky="w")

        main = ttk.Frame(self, padding=self._pad(0, 10, 10, 10))
        main.grid(row=0, column=1, sticky="nsew")
        main.columnconfigure(0, weight=1)
        main.rowconfigure(2, weight=1)

        header = ttk.Frame(main)
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(1, weight=1)
        self.project_label = ttk.Label(header, text="No project selected", font=self.heading_font)
        self.project_label.grid(row=0, column=0, sticky="w")
        self.path_label = ttk.Label(header, text="", foreground="#555")
        self.path_label.grid(row=1, column=0, columnspan=4, sticky="ew", pady=self._pad(2, 0))

        controls = ttk.Frame(main)
        controls.grid(row=1, column=0, sticky="ew", pady=self._pad(12, 8))
        controls.columnconfigure(5, weight=1)

        ttk.Label(controls, text="Format").grid(row=0, column=0, sticky="w")
        self.format_var = tk.StringVar(value="docx")
        self.format_box = ttk.Combobox(
            controls,
            textvariable=self.format_var,
            values=("docx",),
            state="readonly",
            width=8,
        )
        self.format_box.grid(row=0, column=1, sticky="w", padx=self._pad(6, 16))
        self.format_box.bind("<<ComboboxSelected>>", lambda event: self._on_settings_changed())

        ttk.Label(controls, text="Output").grid(row=0, column=2, sticky="w")
        self.output_var = tk.StringVar(value=".")
        self.output_entry = ttk.Entry(controls, textvariable=self.output_var, width=28)
        self.output_entry.grid(row=0, column=3, sticky="ew", padx=self._pad(6, 16))
        self.output_entry.bind("<Return>", lambda event: self._on_settings_changed())
        self.output_entry.bind("<FocusOut>", lambda event: self._on_settings_changed())

        self.force_var = tk.BooleanVar(value=False)
        self.force_var.trace_add("write", lambda *args: self._on_force_changed())
        ttk.Checkbutton(controls, text="Force", variable=self.force_var).grid(row=0, column=4, sticky="w")

        action_bar = ttk.Frame(main)
        action_bar.grid(row=3, column=0, sticky="ew", pady=self._pad(8, 0))
        self.scan_button = ttk.Button(action_bar, text="Scan", command=self._scan)
        self.scan_button.pack(side="left")
        self.convert_selected_button = ttk.Button(action_bar, text="Convert selected", command=self._convert_selected)
        self.convert_selected_button.pack(side="left", padx=self._pad(8, 0))
        self.convert_all_button = ttk.Button(action_bar, text="Convert all", command=self._convert_all)
        self.convert_all_button.pack(side="left", padx=self._pad(8, 0))
        self.cancel_button = ttk.Button(action_bar, text="Stop", command=self._cancel_conversion, state="disabled")
        self.cancel_button.pack(side="left", padx=self._pad(8, 0))
        self.check_tools_button = ttk.Button(action_bar, text="Check tools", command=self._check_tools)
        self.check_tools_button.pack(side="left", padx=self._pad(8, 0))
        self.settings_button = ttk.Button(action_bar, text="Settings", command=self._open_settings)
        self.settings_button.pack(side="left", padx=self._pad(8, 0))
        self.open_output_button = ttk.Button(action_bar, text="Open output", command=self._open_output)
        self.open_output_button.pack(side="left", padx=self._pad(8, 0))
        self.busy_buttons = [
            self.scan_button,
            self.convert_selected_button,
            self.convert_all_button,
            self.check_tools_button,
            self.settings_button,
        ]

        progress_frame = ttk.Frame(main)
        progress_frame.grid(row=4, column=0, columnspan=2, sticky="ew", pady=self._pad(8, 0))
        progress_frame.columnconfigure(0, weight=1)
        self.status_var = tk.StringVar(value="Ready")
        ttk.Label(progress_frame, textvariable=self.status_var).grid(row=0, column=0, sticky="w")
        self.progress_var = tk.DoubleVar(value=0)
        self.progress_bar = ttk.Progressbar(
            progress_frame,
            orient="horizontal",
            mode="determinate",
            variable=self.progress_var,
            maximum=1,
        )
        self.progress_bar.grid(row=1, column=0, sticky="ew", pady=self._pad(4, 0))

        columns = ("state", "file", "reason")
        self.tree = ttk.Treeview(main, columns=columns, show="headings", selectmode="extended")
        self.tree.heading("state", text="State")
        self.tree.heading("file", text="Markdown")
        self.tree.heading("reason", text="Reason")
        self.tree.column("state", width=self._px(118), anchor="w", stretch=False)
        self.tree.column("file", width=self._px(430), anchor="w")
        self.tree.column("reason", width=self._px(260), anchor="w")
        self.tree.tag_configure("skip", foreground="#2e7d32", background="#f0f7f4", font=self.state_font)
        self.tree.tag_configure("convert", foreground="#6f4e00", background="#fff3c4", font=self.state_font)
        self.tree.tag_configure("queued", foreground="#444", background="#eef1f5", font=self.state_font)
        self.tree.tag_configure("running", foreground="#0b5cad", background="#e6f2ff", font=self.state_font)
        self.tree.tag_configure("done", foreground="#127a3a", background="#e7f6ed", font=self.state_font)
        self.tree.tag_configure("failed", foreground="#b00020", background="#fde8e8", font=self.state_font)
        self.tree.grid(row=2, column=0, sticky="nsew")

        scrollbar = ttk.Scrollbar(main, orient="vertical", command=self.tree.yview)
        scrollbar.grid(row=2, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=scrollbar.set)

        # Context menu for file list
        self.context_menu = tk.Menu(self, tearoff=0)
        self.context_menu.add_command(label="打开文件", command=self._menu_open_source_file)
        self.context_menu.add_command(label="打开文件所在目录", command=self._menu_open_source_dir)
        self.context_menu.add_command(label="打开转换完成的文件", command=self._menu_open_output_file)

        # Bind mouse events to the Treeview
        self.tree.bind("<Button-3>", self._on_tree_right_click)
        self.tree.bind("<Double-1>", self._on_tree_double_click)

        self.log = tk.Text(main, height=8, wrap="word")
        self.log.configure(font=self.mono_font, padx=self._px(6), pady=self._px(4), spacing1=self._px(1))
        self.log.grid(row=5, column=0, columnspan=2, sticky="ew", pady=self._pad(8, 0))
        self.log.configure(state="disabled")

    def _configure_display(self) -> None:
        dpi = max(72.0, float(self.winfo_fpixels("1i")))
        self.tk.call("tk", "scaling", dpi / 72.0)

        self.default_font = tkfont.nametofont("TkDefaultFont")
        self.default_font.configure(family="Segoe UI", size=10)
        self.text_font = tkfont.nametofont("TkTextFont")
        self.text_font.configure(family="Segoe UI", size=10)
        self.heading_font = tkfont.Font(family="Segoe UI", size=12, weight="bold")
        self.table_heading_font = tkfont.Font(family="Segoe UI", size=10, weight="bold")
        self.state_font = tkfont.Font(family="Segoe UI", size=10, weight="bold")
        self.mono_font = tkfont.Font(family="Cascadia Mono", size=9)
        self.option_add("*Font", self.default_font)

        style = ttk.Style(self)
        if os.name == "nt" and "vista" in style.theme_names():
            style.theme_use("vista")
        style.configure(".", font=self.default_font)
        style.configure("TButton", padding=self._pad(10, 5))
        style.configure("TEntry", padding=self._pad(4, 3))
        style.configure("TCombobox", padding=self._pad(4, 3))
        style.configure("Treeview", font=self.default_font, rowheight=self._px(28))
        style.configure("Treeview.Heading", font=self.table_heading_font, padding=self._pad(8, 5))
        style.configure("Horizontal.TProgressbar", thickness=self._px(14))

    def _px(self, value: int | float) -> int:
        return max(1, int(round(value * self.ui_scale)))

    def _pad(self, *values: int | float) -> tuple[int, ...]:
        return tuple(self._px(value) if value else 0 for value in values)

    def _load_projects(self) -> None:
        self.project_list.delete(0, tk.END)
        self.projects = self.registry.list()
        for project in self.projects:
            self.project_list.insert(tk.END, project.name)
        if self.projects:
            self.project_list.selection_set(0)
            self._set_project(self.projects[0])
        self._update_project_list_display()

    def _update_project_list_display(self) -> None:
        if not hasattr(self, "projects"):
            return
        self._updating_list_display = True
        try:
            selection = self.project_list.curselection()
            selected_index = selection[0] if selection else None
            
            for i, project in enumerate(self.projects):
                state = self._get_or_create_state(project.root)
                if getattr(state, "conversion_active", False):
                    display_name = f"{project.name} [Converting...]"
                elif state.scan_active:
                    display_name = f"{project.name} [Scanning...]"
                else:
                    display_name = project.name
                
                current_text = self.project_list.get(i)
                if current_text != display_name:
                    self.project_list.delete(i)
                    self.project_list.insert(i, display_name)
            
            if selected_index is not None:
                self.project_list.selection_set(selected_index)
        finally:
            self._updating_list_display = False

    def _new_project(self) -> None:
        kind = self._choose_project_kind()
        if kind is None:
            return
        folder = filedialog.askdirectory(title="Select a project folder")
        if not folder:
            return
        default_name = Path(folder).name
        name = simpledialog.askstring("Project name", "Name", initialvalue=default_name, parent=self)
        if name is None:
            return
        project = create_project(folder, name.strip() or default_name, kind=kind)
        self._append_log(f"Created {_kind_label(project.kind)} project: {project.root}")
        self._load_projects()

    def _choose_project_kind(self) -> str | None:
        dialog = tk.Toplevel(self)
        dialog.title("New project type")
        dialog.transient(self)
        dialog.grab_set()
        dialog.resizable(False, False)

        frame = ttk.Frame(dialog, padding=self._px(16))
        frame.grid(row=0, column=0, sticky="nsew")
        ttk.Label(frame, text="What should this project convert?").grid(
            row=0, column=0, sticky="w", pady=self._pad(0, 10)
        )

        kind_var = tk.StringVar(value=KIND_MD2DOC)
        ttk.Radiobutton(
            frame,
            text="Markdown  ->  Word   (Pandoc)",
            variable=kind_var,
            value=KIND_MD2DOC,
        ).grid(row=1, column=0, sticky="w", pady=self._pad(2, 0))
        ttk.Radiobutton(
            frame,
            text="Word / PPT / Excel  ->  Markdown   (MarkItDown)",
            variable=kind_var,
            value=KIND_DOC2MD,
        ).grid(row=2, column=0, sticky="w", pady=self._pad(2, 0))
        ttk.Radiobutton(
            frame,
            text="Quarto Markdown (.qmd)  ->  PowerPoint (.pptx)   (Quarto)",
            variable=kind_var,
            value=KIND_QMD2PPT,
        ).grid(row=3, column=0, sticky="w", pady=self._pad(2, 0))
        ttk.Radiobutton(
            frame,
            text="HTML (.html)  ->  single-page PDF (.pdf)   (Chromium)",
            variable=kind_var,
            value=KIND_HTML2PDF,
        ).grid(row=4, column=0, sticky="w", pady=self._pad(2, 0))

        result: dict[str, str | None] = {"kind": None}

        def confirm() -> None:
            result["kind"] = kind_var.get()
            dialog.destroy()

        buttons = ttk.Frame(frame)
        buttons.grid(row=5, column=0, sticky="e", pady=self._pad(16, 0))
        ttk.Button(buttons, text="Continue", command=confirm).grid(row=0, column=0)

        dialog.bind("<Return>", lambda _event: confirm())
        dialog.bind("<Escape>", lambda _event: dialog.destroy())
        self._center_window(dialog)
        self.wait_window(dialog)
        return result["kind"]

    def _center_window(self, window: tk.Toplevel) -> None:
        window.update_idletasks()
        x = self.winfo_rootx() + max(0, (self.winfo_width() - window.winfo_width()) // 2)
        y = self.winfo_rooty() + max(0, (self.winfo_height() - window.winfo_height()) // 2)
        window.geometry(f"+{x}+{y}")

    def _remove_project(self) -> None:
        project = self.current_project
        if not project:
            return
        self.registry.remove(project.root)
        self.current_project = None
        self._clear_table()
        self._load_projects()

    def _on_project_selected(self, _event: tk.Event) -> None:
        if getattr(self, "_updating_list_display", False):
            return
        selection = self.project_list.curselection()
        if not selection:
            return
        self._set_project(self.projects[selection[0]])

    def _get_or_create_state(self, root: Path) -> ProjectState:
        key = str(root.resolve())
        if key not in self.project_states:
            self.project_states[key] = ProjectState()
        return self.project_states[key]

    def _save_project_state(self, root: Path) -> None:
        state = self._get_or_create_state(root)
        state.plan_by_id = self.plan_by_id
        state.iid_by_source = self.iid_by_source
        state.log_content = self.log.get("1.0", tk.END).rstrip("\n")
        state.conversion_total = self.conversion_total
        state.conversion_done = self.conversion_done
        state.converted_count = self.converted_count
        state.skipped_count = self.skipped_count
        state.failed_count = self.failed_count
        state.progress_value = self.progress_var.get()
        state.status_text = self.status_var.get()
        state.scan_active = self.scan_active
        state.scan_generation = self.scan_generation
        state.scan_worker = self.scan_worker
        if self.current_project:
            state.kind = self.current_project.kind

    def _load_project_state(self, root: Path) -> None:
        key = str(root.resolve())
        is_new = key not in self.project_states
        state = self._get_or_create_state(root)

        self.plan_by_id = state.plan_by_id
        self.iid_by_source = state.iid_by_source
        self.conversion_total = state.conversion_total
        self.conversion_done = state.conversion_done
        self.converted_count = state.converted_count
        self.skipped_count = state.skipped_count
        self.failed_count = state.failed_count
        self.scan_generation = state.scan_generation
        self.scan_active = state.scan_active
        self.scan_worker = state.scan_worker

        self.log.configure(state="normal")
        self.log.delete("1.0", tk.END)
        if state.log_content:
            self.log.insert(tk.END, state.log_content + "\n")
            self.log.see(tk.END)
        self.log.configure(state="disabled")

        self.progress_var.set(state.progress_value)
        self.progress_bar.configure(maximum=max(state.conversion_total, 1))
        self.status_var.set(state.status_text)

        for iid in self.tree.get_children():
            self.tree.delete(iid)

        for iid, item in state.plan_by_id.items():
            item_state = state.item_states.get(str(item.source))
            if item_state:
                val_state, val_reason, tag = item_state
            else:
                tag = "convert" if item.action == "convert" else "skip"
                val_state = _state_label(item.action)
                val_reason = _reason_label(item.reason)
            values = (val_state, item.relative_source, val_reason)
            self.tree.insert("", tk.END, iid=iid, values=values, tags=(tag,))

        self._refresh_busy_state()
        self._update_project_list_display()
        if not state.conversion_active:
            self._scan()

    def _set_project(self, project: ProjectConfig) -> None:
        if self.current_project:
            self._save_project_state(self.current_project.root)

        self.current_project = load_project(project.root)
        if self.current_project.output_dir == "output":
            self.current_project.output_dir = "."
            self.current_project.save()
            self.registry.add(self.current_project)
        self.project_label.configure(
            text=f"{self.current_project.name}  -  {_kind_label(self.current_project.kind)}"
        )
        self.path_label.configure(text=str(self.current_project.root))
        self._apply_kind_to_ui(self.current_project.kind)
        self.format_var.set(self.current_project.output_format)
        self.output_var.set(self.current_project.output_dir)
        self._load_project_state(self.current_project.root)

    def _apply_kind_to_ui(self, kind: str) -> None:
        if kind == KIND_DOC2MD:
            self.format_box.configure(values=("md",), state="disabled")
            self.format_var.set("md")
            self.tree.heading("file", text="Office document")
        elif kind == KIND_QMD2PPT:
            self.format_box.configure(values=("pptx",), state="disabled")
            self.format_var.set("pptx")
            self.tree.heading("file", text="Quarto Markdown")
        elif kind == KIND_HTML2PDF:
            self.format_box.configure(values=("pdf",), state="disabled")
            self.format_var.set("pdf")
            self.tree.heading("file", text="HTML")
        else:
            self.format_box.configure(values=("docx",), state="disabled")
            self.format_var.set("docx")
            self.tree.heading("file", text="Markdown")

    def _settings(self) -> ConvertSettings:
        project = self._require_project()
        output_dir = self.output_var.get().strip() or "."
        project.output_format = self.format_var.get()
        project.output_dir = output_dir
        project.save()
        self.registry.add(project)
        return replace(
            settings_from_project(project, force=self.force_var.get()),
            output_dir=(project.root / output_dir).resolve(),
            force=self.force_var.get(),
        )

    def _on_force_changed(self) -> None:
        if self.current_project:
            self._scan()

    def _on_settings_changed(self) -> None:
        if self.current_project:
            self._scan()

    def _scan(self) -> None:
        project = self.current_project
        if not project:
            return
        state = self._get_or_create_state(project.root)
        if state.conversion_active:
            self.status_var.set(state.status_text)
            self._refresh_busy_state()
            self._update_project_list_display()
            return
        try:
            settings = self._settings()
        except Exception as exc:
            messagebox.showerror("Scan failed", str(exc))
            return

        state.scan_generation += 1
        generation = state.scan_generation
        self.scan_generation = generation
        project_root = project.root
        project_name = project.name
        project_kind = project.kind
        state.scan_active = True
        self.scan_active = True
        self._clear_table()
        state.status_text = "Scanning in background; estimates are not final"
        self.status_var.set(state.status_text)
        log_msg = f"Scanning {project_name} in background. Estimates will update when scanning finishes."
        state.log_content += log_msg + "\n"
        self._append_log(log_msg)
        state.progress_value = 0.0
        self.progress_var.set(0)
        self.progress_bar.configure(mode="determinate", maximum=1)
        self._refresh_busy_state()
        self._update_project_list_display()

        def work() -> None:
            try:
                sources = scan_source_files(
                    project_root,
                    kind=settings.kind,
                    recursive=settings.recursive,
                    output_dir=settings.output_dir,
                )
                planned = plan_conversions(
                    project_root,
                    sources,
                    settings,
                    use_cached_fingerprints=True,
                )
                self.event_queue.put(("scan_done", project_root, generation, project_kind, planned))
            except Exception as exc:
                self.event_queue.put(("scan_error", project_root, generation, str(exc)))

        self.scan_worker = threading.Thread(target=work, daemon=True)
        state.scan_worker = self.scan_worker
        self.scan_worker.start()

    def _handle_scan_done(self, project_root: Path, generation: int, project_kind: str, planned: list[PlanItem]) -> None:
        state = self._get_or_create_state(project_root)
        if generation != state.scan_generation:
            return
        state.scan_worker = None
        if state.conversion_active:
            state.scan_active = False
            if self.current_project and self.current_project.root.resolve() == project_root.resolve():
                self.scan_active = False
                self.status_var.set(state.status_text)
                self._refresh_busy_state()
            self._update_project_list_display()
            return

        state.plan_by_id.clear()
        state.iid_by_source.clear()
        state.item_states.clear()
        for index, item in enumerate(planned):
            iid = str(index)
            state.plan_by_id[iid] = item
            state.iid_by_source[str(item.source)] = iid
            tag = "convert" if item.action == "convert" else "skip"
            state.item_states[str(item.source)] = (_state_label(item.action), _reason_label(item.reason), tag)

        convert_count = sum(1 for item in planned if item.action == "convert")
        skip_count = len(planned) - convert_count

        status_text = f"Scanned {len(planned)} file(s): {convert_count} to convert, {skip_count} up to date"
        log_msg = f"Scanned {len(planned)} {_input_label(project_kind)} file(s)."
        state.status_text = status_text
        state.log_content += log_msg + "\n"
        state.scan_active = False

        if self.current_project and self.current_project.root.resolve() == project_root.resolve():
            self.scan_active = False
            self.status_var.set(status_text)
            self._append_log(log_msg)
            self.progress_bar.configure(mode="determinate", maximum=max(len(planned), 1))
            self.progress_var.set(0)
            self._clear_table()
            self._insert_scan_batch(
                project_root,
                generation,
                planned,
                0,
                convert_count,
                skip_count,
                project_kind,
            )
        else:
            self._update_project_list_display()

    def _insert_scan_batch(
        self,
        project_root: Path,
        generation: int,
        planned: list[PlanItem],
        start: int,
        convert_count: int,
        skip_count: int,
        project_kind: str,
    ) -> None:
        if not self.current_project or self.current_project.root.resolve() != project_root.resolve():
            return
        state = self._get_or_create_state(project_root)
        if generation != state.scan_generation:
            return

        end = min(start + SCAN_TABLE_BATCH_SIZE, len(planned))
        for index in range(start, end):
            self._insert_or_update_plan_item(str(index), planned[index])
        self.progress_var.set(end)
        if end < len(planned):
            self.status_var.set(f"Loading scan results {end}/{len(planned)}...")
            self.after(
                1,
                lambda: self._insert_scan_batch(
                    project_root,
                    generation,
                    planned,
                    end,
                    convert_count,
                    skip_count,
                    project_kind,
                ),
            )
            return

        self.status_var.set(f"Scanned {len(planned)} file(s): {convert_count} to convert, {skip_count} up to date")
        self._refresh_busy_state()
        self._update_project_list_display()

    def _handle_scan_error(self, project_root: Path, generation: int, message: str) -> None:
        state = self._get_or_create_state(project_root)
        if generation != state.scan_generation:
            return
        state.scan_worker = None
        state.scan_active = False
        state.status_text = "Scan failed"
        state.log_content += message + "\n"

        if self.current_project and self.current_project.root.resolve() == project_root.resolve():
            self.scan_active = False
            self.progress_bar.configure(mode="determinate", maximum=1)
            self.progress_var.set(0)
            self.status_var.set("Scan failed")
            self._refresh_busy_state()
            self._append_log(message)
            messagebox.showerror("Scan failed", message)
        self._update_project_list_display()

    def _convert_selected(self) -> None:
        if self.scan_active:
            messagebox.showinfo("Busy", "A scan is still running.")
            return
        selection = self.tree.selection()
        if not selection:
            messagebox.showinfo("Convert selected", "Select one or more files first.")
            return
        planned = [self.plan_by_id[iid] for iid in selection]
        self._start_conversion([item.source for item in planned])

    def _convert_all(self) -> None:
        project = self.current_project
        if not project:
            return
        if self.scan_active:
            messagebox.showinfo("Busy", "A scan is still running.")
            return
        planned = list(self.plan_by_id.values())
        if not planned:
            messagebox.showinfo("Convert all", "Scan this project before converting.")
            return
        self._start_conversion([item.source for item in planned])

    def _start_conversion(self, sources: list[Path]) -> None:
        project = self._require_project()
        settings = self._settings()
        if self.worker and self.worker.is_alive():
            messagebox.showinfo("Busy", "A conversion is already running.")
            return
        # Always re-plan using the current settings, so that changed settings
        # (like format, output directory, or force option) are respected.
        planned = plan_conversions(
            project.root,
            sources,
            settings,
            use_cached_fingerprints=True,
        )
        self._prepare_conversion_progress(planned)
        queued = [item for item in planned if item.action == "convert"]
        skipped = [item for item in planned if item.action == "skip"]
        self._record_already_skipped(skipped)
        if not queued:
            status_text = (
                f"Finished: {self.converted_count} converted, "
                f"{self.skipped_count} up to date, {self.failed_count} failed"
            )
            log_msg = (
                f"Finished: {self.converted_count} converted, "
                f"{self.skipped_count} up to date, {self.failed_count} failed."
            )
            self.status_var.set(status_text)
            self._append_log(log_msg)
            state = self._get_or_create_state(project.root)
            state.status_text = status_text
            state.log_content += log_msg + "\n"
            state.conversion_active = False
            self._save_project_state(project.root)
            self._refresh_busy_state()
            self._update_project_list_display()
            return
        queued_sources = [item.source for item in queued]
        self.cancel_event.clear()

        def work() -> None:
            try:
                results = run_conversions(
                    project.root,
                    queued_sources,
                    settings,
                    on_event=lambda res: self.event_queue.put(("result", project.root, res)),
                    on_start=lambda item: self.event_queue.put(("start", project.root, item)),
                    cancel_event=self.cancel_event,
                )
                converted = sum(1 for result in results if result.status == "converted")
                skipped = sum(1 for result in results if result.status == "skipped")
                failed = sum(1 for result in results if result.status == "failed")
                self.event_queue.put(("done", project.root, converted, skipped, failed))
            except ConversionCancelledError:
                self.event_queue.put(("cancelled", project.root))
            except Exception as exc:
                self.event_queue.put(("error", project.root, str(exc)))

        log_msg = f"Starting conversion for {len(queued_sources)} file(s)."
        self._append_log(log_msg)
        state = self._get_or_create_state(project.root)
        state.log_content += log_msg + "\n"
        state.conversion_active = True
        self._save_project_state(project.root)
        
        self.worker = threading.Thread(target=work, daemon=True)
        self.worker.start()
        self._refresh_busy_state()
        self._update_project_list_display()

    def _cancel_conversion(self) -> None:
        if self.cancel_event:
            self.cancel_event.set()
            self._append_log("Stopping conversion...")
            self.status_var.set("Stopping...")
            self.cancel_button.configure(state="disabled")

    def _record_already_skipped(self, items: list[PlanItem]) -> None:
        if not items:
            return
        self.conversion_done += len(items)
        self.skipped_count += len(items)
        self.progress_var.set(self.conversion_done)
        self.status_var.set(
            f"Progress {self.conversion_done}/{self.conversion_total}: "
            f"{self.converted_count} converted, {self.skipped_count} up to date, {self.failed_count} failed"
        )
        self._append_log(f"{len(items)} file(s) already up to date.")

    def _check_tools(self) -> None:
        checks = check_dependencies(self._settings())
        message = "\n".join(
            f"{check.name}: {'ok' if check.available else 'missing'} - {check.detail}"
            for check in checks
        )
        missing = missing_dependency_message(checks)
        if missing:
            message = f"{message}\n\n{missing}"
        messagebox.showinfo("Conversion tools", message)

    def _open_settings(self) -> None:
        project = self._require_project()
        if project.kind == KIND_DOC2MD:
            messagebox.showinfo(
                "Project Settings",
                "Word/PPT/Excel to Markdown projects convert with MarkItDown and have no "
                "document formatting options. Use the Output box to choose where the .md "
                "files are written.",
            )
            return
        if project.kind == KIND_QMD2PPT:
            messagebox.showinfo(
                "Project Settings",
                "Quarto Markdown to PowerPoint projects are configured via the YAML header "
                "inside the .qmd files. Use the Output box to choose where the .pptx "
                "files are written.",
            )
            return
        if project.kind == KIND_HTML2PDF:
            messagebox.showinfo(
                "Project Settings",
                "HTML to PDF projects use the rendered HTML and CSS size to create one "
                "custom-sized PDF page. Use the Output box to choose where the .pdf "
                "files are written.",
            )
            return
        dialog = SettingsDialog(self, project)
        self.wait_window(dialog)
        if dialog.saved:
            self.current_project = load_project(project.root)
            self.format_var.set(self.current_project.output_format)
            self.output_var.set(self.current_project.output_dir)
            self._append_log("Project settings saved.")
            self._scan()

    def _open_output(self) -> None:
        project = self.current_project
        if not project:
            return
        output_dir = (project.root / self.output_var.get()).resolve()
        output_dir.mkdir(parents=True, exist_ok=True)
        try:
            import os

            os.startfile(output_dir)
        except Exception as exc:
            messagebox.showerror("Open output", str(exc))

    def _open_file_path(self, path: Path) -> None:
        try:
            os.startfile(path)
        except Exception as exc:
            messagebox.showerror("Error", f"Failed to open {path.name}: {exc}")

    def _open_containing_dir(self, path: Path) -> None:
        try:
            import subprocess
            # Highlight the file in Explorer
            subprocess.run(['explorer', '/select,', os.path.normpath(path)], check=True)
        except Exception:
            try:
                os.startfile(path.parent)
            except Exception as exc:
                messagebox.showerror("Error", f"Failed to open directory: {exc}")

    def _menu_open_source_file(self) -> None:
        selection = self.tree.selection()
        for iid in selection:
            item = self.plan_by_id.get(iid)
            if item and item.source.exists():
                self._open_file_path(item.source)

    def _menu_open_source_dir(self) -> None:
        selection = self.tree.selection()
        for iid in selection:
            item = self.plan_by_id.get(iid)
            if item and item.source.exists():
                self._open_containing_dir(item.source)

    def _menu_open_output_file(self) -> None:
        selection = self.tree.selection()
        for iid in selection:
            item = self.plan_by_id.get(iid)
            if item and item.output.exists():
                self._open_file_path(item.output)

    def _on_tree_right_click(self, event: tk.Event) -> None:
        iid = self.tree.identify_row(event.y)
        if iid:
            if iid not in self.tree.selection():
                self.tree.selection_set(iid)
            self._update_context_menu_state()
            self.context_menu.post(event.x_root, event.y_root)

    def _update_context_menu_state(self) -> None:
        selection = self.tree.selection()
        if not selection:
            self.context_menu.entryconfigure("打开文件", state="disabled")
            self.context_menu.entryconfigure("打开文件所在目录", state="disabled")
            self.context_menu.entryconfigure("打开转换完成的文件", state="disabled")
            return

        self.context_menu.entryconfigure("打开文件", state="normal")
        self.context_menu.entryconfigure("打开文件所在目录", state="normal")

        any_output_exists = False
        for iid in selection:
            item = self.plan_by_id.get(iid)
            if item and item.output.exists():
                any_output_exists = True
                break

        if any_output_exists:
            self.context_menu.entryconfigure("打开转换完成的文件", state="normal")
        else:
            self.context_menu.entryconfigure("打开转换完成的文件", state="disabled")

    def _on_tree_double_click(self, event: tk.Event) -> None:
        iid = self.tree.identify_row(event.y)
        if not iid:
            return
        item = self.plan_by_id.get(iid)
        if not item:
            return
        
        if item.output.exists():
            self._open_file_path(item.output)
        elif item.source.exists():
            self._open_file_path(item.source)

    def _poll_events(self) -> None:
        while True:
            try:
                event = self.event_queue.get_nowait()
            except queue.Empty:
                break
            if isinstance(event, tuple):
                kind = event[0]
                if kind == "scan_done":
                    self._handle_scan_done(event[1], event[2], event[3], event[4])
                    continue
                if kind == "scan_error":
                    self._handle_scan_error(event[1], event[2], str(event[3]))
                    continue
                if kind == "start":
                    self._mark_item_running(event[1], event[2])
                    continue
                if kind == "result":
                    self._handle_conversion_result(event[1], event[2])
                    continue
                if kind == "done":
                    self._handle_done_event(event[1], event[2], event[3], event[4])
                    continue
                if kind == "cancelled":
                    self._handle_cancelled_event(event[1])
                    continue
                if kind == "error":
                    self._handle_error_event(event[1], str(event[2]))
                    continue
        self.after(150, self._poll_events)

    def _clear_table(self) -> None:
        self.plan_by_id = {}
        self.iid_by_source = {}
        for iid in self.tree.get_children():
            self.tree.delete(iid)

    def _insert_or_update_plan_item(self, iid: str, item: PlanItem) -> None:
        self.plan_by_id[iid] = item
        self.iid_by_source[str(item.source)] = iid
        tag = "convert" if item.action == "convert" else "skip"
        values = (
            _state_label(item.action),
            item.relative_source,
            _reason_label(item.reason),
        )
        if self.current_project:
            state = self._get_or_create_state(self.current_project.root)
            state.item_states[str(item.source)] = (values[0], values[2], tag)

        if self.tree.exists(iid):
            self.tree.item(iid, values=values, tags=(tag,))
        else:
            self.tree.insert("", tk.END, iid=iid, values=values, tags=(tag,))

    def _prepare_conversion_progress(self, planned: list[PlanItem]) -> None:
        self.conversion_total = len(planned)
        self.conversion_done = 0
        self.converted_count = 0
        self.skipped_count = 0
        self.failed_count = 0
        self.progress_bar.configure(maximum=max(self.conversion_total, 1))
        self.progress_var.set(0)
        self.status_var.set(f"Starting: 0/{self.conversion_total}")
        for index, item in enumerate(planned):
            iid = self.iid_by_source.get(str(item.source), self._next_iid())
            self._insert_or_update_plan_item(iid, item)
            if item.action == "convert":
                if self.current_project and self.current_project.kind == KIND_DOC2MD:
                    tool = "MarkItDown"
                elif self.current_project and self.current_project.kind == KIND_QMD2PPT:
                    tool = "Quarto"
                elif self.current_project and self.current_project.kind == KIND_HTML2PDF:
                    tool = "Chromium"
                else:
                    tool = "Pandoc"
                self._set_item_state(item, _state_label("queued"), f"Waiting for {tool}", "queued")

    def _mark_item_running(self, project_root: Path, item: PlanItem) -> None:
        state = self._get_or_create_state(project_root)
        if state.kind == KIND_DOC2MD:
            tool = "MarkItDown"
        elif state.kind == KIND_QMD2PPT:
            tool = "Quarto"
        elif state.kind == KIND_HTML2PDF:
            tool = "Chromium"
        else:
            tool = "Pandoc"
        state_label = _state_label("running")
        reason = f"Running {tool}"
        tag = "running"

        state.item_states[str(item.source)] = (state_label, reason, tag)
        state.status_text = f"Converting {item.relative_source} ({state.conversion_done}/{state.conversion_total})"

        if self.current_project and self.current_project.root.resolve() == project_root.resolve():
            self._set_item_state(item, state_label, reason, tag)
            self.status_var.set(state.status_text)

    def _handle_conversion_result(self, project_root: Path, result: ConversionResult) -> None:
        state = self._get_or_create_state(project_root)
        state.conversion_done += 1

        if result.status == "converted":
            state.converted_count += 1
            state_val = _state_label("done")
            reason_val = "Output generated"
            tag_val = "done"
        elif result.status == "skipped":
            state.skipped_count += 1
            state_val = _state_label("skipped")
            reason_val = _reason_label(result.message)
            tag_val = "skip"
        else:
            state.failed_count += 1
            state_val = _state_label("failed")
            reason_val = result.message
            tag_val = "failed"

        state.item_states[str(result.item.source)] = (state_val, reason_val, tag_val)

        iid = state.iid_by_source.get(str(result.item.source))
        if iid is None:
            index = len(state.plan_by_id)
            while str(index) in state.plan_by_id:
                index += 1
            iid = str(index)
            state.plan_by_id[iid] = result.item
            state.iid_by_source[str(result.item.source)] = iid

        status_text = (
            f"Progress {state.conversion_done}/{state.conversion_total}: "
            f"{state.converted_count} converted, {state.skipped_count} up to date, {state.failed_count} failed"
        )
        log_msg = f"{result.status}: {result.item.relative_source} - {result.message}"
        state.status_text = status_text
        state.log_content += log_msg + "\n"

        if self.current_project and self.current_project.root.resolve() == project_root.resolve():
            self.conversion_done = state.conversion_done
            self.converted_count = state.converted_count
            self.skipped_count = state.skipped_count
            self.failed_count = state.failed_count

            self.progress_var.set(self.conversion_done)
            self._set_item_state(result.item, state_val, reason_val, tag_val)
            self.status_var.set(status_text)
            self._append_log(log_msg)

    def _handle_done_event(self, project_root: Path, converted: int, skipped: int, failed: int) -> None:
        state = self._get_or_create_state(project_root)
        status_text = (
            f"Finished: {state.converted_count} converted, "
            f"{state.skipped_count} up to date, {state.failed_count} failed"
        )
        log_msg = (
            f"Finished: {state.converted_count} converted, "
            f"{state.skipped_count} up to date, {state.failed_count} failed."
        )
        state.status_text = status_text
        state.log_content += log_msg + "\n"
        state.conversion_active = False

        self.worker = None
        self._refresh_busy_state()

        if self.current_project and self.current_project.root.resolve() == project_root.resolve():
            self.status_var.set(status_text)
            self._append_log(log_msg)
        self._update_project_list_display()

    def _handle_cancelled_event(self, project_root: Path) -> None:
        state = self._get_or_create_state(project_root)
        state.status_text = "Conversion cancelled"
        log_msg = "Conversion cancelled by user."
        state.log_content += log_msg + "\n"
        state.conversion_active = False

        self.worker = None
        self._refresh_busy_state()

        if self.current_project and self.current_project.root.resolve() == project_root.resolve():
            self.status_var.set("Conversion cancelled")
            self._append_log(log_msg)
        self._update_project_list_display()

    def _handle_error_event(self, project_root: Path, message: str) -> None:
        state = self._get_or_create_state(project_root)
        state.status_text = "Conversion failed"
        state.log_content += message + "\n"
        state.conversion_active = False

        self.worker = None
        self._refresh_busy_state()

        if self.current_project and self.current_project.root.resolve() == project_root.resolve():
            self.status_var.set("Conversion failed")
            self._append_log(message)
            messagebox.showerror("Conversion failed", message)
        self._update_project_list_display()

    def _set_item_state(self, item: PlanItem, state: str, reason: str, tag: str) -> None:
        iid = self.iid_by_source.get(str(item.source))
        if iid is None or not self.tree.exists(iid):
            iid = self._next_iid()
            self._insert_or_update_plan_item(iid, item)

        if self.current_project:
            p_state = self._get_or_create_state(self.current_project.root)
            p_state.item_states[str(item.source)] = (state, reason, tag)

        self.tree.item(
            iid,
            values=(state, item.relative_source, reason),
            tags=(tag,),
        )
        self.tree.see(iid)

    def _next_iid(self) -> str:
        index = len(self.plan_by_id)
        while self.tree.exists(str(index)):
            index += 1
        return str(index)

    def _set_busy(self, busy: bool) -> None:
        state = "disabled" if busy else "normal"
        for button in self.busy_buttons:
            button.configure(state=state)

    def _refresh_busy_state(self) -> None:
        converting = self._conversion_is_active()
        self._set_busy(self.scan_active or converting)
        if self.worker is not None and self.worker.is_alive():
            self.cancel_button.configure(state="normal")
        else:
            self.cancel_button.configure(state="disabled")

    def _conversion_is_active(self) -> bool:
        if self.worker is not None and self.worker.is_alive():
            return True
        return any(state.conversion_active for state in self.project_states.values())

    def _on_closing(self) -> None:
        converting = self.worker is not None and self.worker.is_alive()
        if converting or self.scan_active:
            if not messagebox.askyesno("Exit", "Conversion or scan is in progress. Are you sure you want to exit?"):
                return
            if converting:
                self._cancel_conversion()
                self.worker.join(timeout=1.0)
        self.destroy()

    def _append_log(self, message: str) -> None:
        self.log.configure(state="normal")
        self.log.insert(tk.END, message + "\n")
        self.log.see(tk.END)
        self.log.configure(state="disabled")

    def _require_project(self) -> ProjectConfig:
        if not self.current_project:
            raise RuntimeError("No project selected")
        return self.current_project


class SettingsDialog(tk.Toplevel):
    def __init__(self, parent: Md2DocApp, project: ProjectConfig) -> None:
        super().__init__(parent)
        self.parent = parent
        self.project = project
        self.saved = False
        self.title("Project Settings")
        self.geometry(f"{parent._px(760)}x{parent._px(620)}")
        self.minsize(parent._px(680), parent._px(540))
        self.transient(parent)
        self.grab_set()

        self.toc_var = tk.BooleanVar(value=project.toc)
        self.toc_depth_var = tk.StringVar(value=str(project.toc_depth))
        self.title_page_var = tk.BooleanVar(value=project.title_page)
        self.title_var = tk.StringVar(value=project.title)
        self.subtitle_var = tk.StringVar(value=project.subtitle)
        self.author_var = tk.StringVar(value=project.author)
        self.date_var = tk.StringVar(value=project.date)
        self.number_sections_var = tk.BooleanVar(value=project.number_sections)
        self.reference_docx_var = tk.StringVar(value=project.reference_docx)
        self.default_font_var = tk.StringVar(value=project.default_font)
        self.default_font_size_var = tk.StringVar(value=str(project.default_font_size or ""))
        self.table_borders_var = tk.StringVar(value=project.table_borders)
        self.mermaid_format_var = tk.StringVar(value=project.mermaid_format)
        self.mermaid_theme_var = tk.StringVar(value=project.mermaid_theme)
        self.mermaid_background_var = tk.StringVar(value=project.mermaid_background)
        self.mermaid_scale_var = tk.StringVar(value=str(project.mermaid_scale or ""))
        self.mermaid_min_dpi_var = tk.StringVar(value=str(project.mermaid_min_dpi))
        self.hr_to_pagebreak_var = tk.BooleanVar(value=project.hr_to_pagebreak)

        self._build()
        self._center()

    def _build(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        notebook = ttk.Notebook(self)
        notebook.grid(row=0, column=0, sticky="nsew", padx=self.parent._px(12), pady=self.parent._px(12))

        document = ttk.Frame(notebook, padding=self.parent._px(12))
        word = ttk.Frame(notebook, padding=self.parent._px(12))
        mermaid = ttk.Frame(notebook, padding=self.parent._px(12))
        advanced = ttk.Frame(notebook, padding=self.parent._px(12))
        notebook.add(document, text="Document")
        notebook.add(word, text="Word")
        notebook.add(mermaid, text="Mermaid")
        notebook.add(advanced, text="Advanced")

        self._build_document_tab(document)
        self._build_word_tab(word)
        self._build_mermaid_tab(mermaid)
        self._build_advanced_tab(advanced)

        buttons = ttk.Frame(self)
        buttons.grid(row=1, column=0, sticky="ew", padx=self.parent._px(12), pady=self.parent._pad(0, 12))
        buttons.columnconfigure(0, weight=1)
        ttk.Button(buttons, text="Restore defaults", command=self._restore_defaults).grid(row=0, column=0, sticky="w")
        ttk.Button(buttons, text="Cancel", command=self.destroy).grid(row=0, column=1, padx=self.parent._pad(0, 8))
        ttk.Button(buttons, text="Save", command=self._save).grid(row=0, column=2)

    def _build_document_tab(self, frame: ttk.Frame) -> None:
        frame.columnconfigure(1, weight=1)
        ttk.Checkbutton(frame, text="Table of contents", variable=self.toc_var).grid(row=0, column=0, sticky="w")
        ttk.Label(frame, text="Depth").grid(row=0, column=1, sticky="e")
        ttk.Spinbox(frame, from_=1, to=6, textvariable=self.toc_depth_var, width=6).grid(
            row=0,
            column=2,
            sticky="w",
            padx=self.parent._pad(6, 0),
        )
        ttk.Checkbutton(frame, text="Number sections", variable=self.number_sections_var).grid(
            row=1,
            column=0,
            sticky="w",
            pady=self.parent._pad(8, 0),
        )
        ttk.Checkbutton(frame, text="Title page", variable=self.title_page_var).grid(
            row=2,
            column=0,
            sticky="w",
            pady=self.parent._pad(16, 8),
        )

        labels = [("Title", self.title_var), ("Subtitle", self.subtitle_var), ("Author", self.author_var), ("Date", self.date_var)]
        for index, (label, variable) in enumerate(labels, start=3):
            ttk.Label(frame, text=label).grid(row=index, column=0, sticky="w", pady=self.parent._pad(4, 0))
            ttk.Entry(frame, textvariable=variable).grid(
                row=index,
                column=1,
                columnspan=2,
                sticky="ew",
                pady=self.parent._pad(4, 0),
            )

    def _build_word_tab(self, frame: ttk.Frame) -> None:
        frame.columnconfigure(1, weight=1)
        ttk.Label(frame, text="Reference DOCX").grid(row=0, column=0, sticky="w")
        ttk.Entry(frame, textvariable=self.reference_docx_var).grid(
            row=0,
            column=1,
            sticky="ew",
            padx=self.parent._pad(8, 8),
        )
        ttk.Button(frame, text="Browse", command=self._browse_reference_docx).grid(row=0, column=2)
        ttk.Button(frame, text="Clear", command=lambda: self.reference_docx_var.set("")).grid(
            row=0,
            column=3,
            padx=self.parent._pad(8, 0),
        )

        ttk.Label(frame, text="Default font").grid(row=1, column=0, sticky="w", pady=self.parent._pad(16, 0))
        ttk.Entry(frame, textvariable=self.default_font_var).grid(
            row=1,
            column=1,
            columnspan=3,
            sticky="ew",
            pady=self.parent._pad(16, 0),
        )

        ttk.Label(frame, text="Font size").grid(row=2, column=0, sticky="w", pady=self.parent._pad(8, 0))
        ttk.Spinbox(frame, from_=0, to=72, textvariable=self.default_font_size_var, width=8).grid(
            row=2,
            column=1,
            sticky="w",
            pady=self.parent._pad(8, 0),
        )

        ttk.Label(frame, text="Table borders").grid(row=3, column=0, sticky="w", pady=self.parent._pad(8, 0))
        ttk.Combobox(
            frame,
            textvariable=self.table_borders_var,
            values=("template", "bordered", "plain"),
            state="readonly",
            width=12,
        ).grid(row=3, column=1, sticky="w", pady=self.parent._pad(8, 0))
        ttk.Checkbutton(
            frame,
            text="Convert horizontal rules to page breaks",
            variable=self.hr_to_pagebreak_var,
        ).grid(row=4, column=0, columnspan=4, sticky="w", pady=self.parent._pad(12, 0))

    def _build_mermaid_tab(self, frame: ttk.Frame) -> None:
        frame.columnconfigure(1, weight=1)
        fields = [
            ("Theme", self.mermaid_theme_var),
            ("Background", self.mermaid_background_var),
            ("Scale", self.mermaid_scale_var),
            ("Min DPI", self.mermaid_min_dpi_var),
        ]
        for row, (label, variable) in enumerate(fields):
            ttk.Label(frame, text=label).grid(row=row, column=0, sticky="w", pady=self.parent._pad(4, 0))
            ttk.Entry(frame, textvariable=variable).grid(
                row=row,
                column=1,
                sticky="ew",
                pady=self.parent._pad(4, 0),
            )

        ttk.Label(frame, text="Format").grid(row=4, column=0, sticky="w", pady=self.parent._pad(8, 0))
        ttk.Combobox(
            frame,
            textvariable=self.mermaid_format_var,
            values=("png", "svg", "pdf"),
            state="readonly",
            width=10,
        ).grid(row=4, column=1, sticky="w", pady=self.parent._pad(8, 0))

    def _build_advanced_tab(self, frame: ttk.Frame) -> None:
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(1, weight=1)
        ttk.Label(frame, text="Extra Pandoc args").grid(row=0, column=0, sticky="w")
        self.extra_args_text = tk.Text(frame, height=8, wrap="word", font=self.parent.mono_font)
        self.extra_args_text.grid(row=1, column=0, sticky="nsew", pady=self.parent._pad(8, 0))
        self.extra_args_text.insert("1.0", " ".join(self.project.extra_pandoc_args))

    def _browse_reference_docx(self) -> None:
        path = filedialog.askopenfilename(
            parent=self,
            title="Select reference DOCX",
            filetypes=(("Word documents", "*.docx"), ("All files", "*.*")),
        )
        if path:
            self.reference_docx_var.set(path)

    def _restore_defaults(self) -> None:
        defaults = ProjectConfig(name=self.project.name, root=self.project.root)
        self.toc_var.set(defaults.toc)
        self.toc_depth_var.set(str(defaults.toc_depth))
        self.title_page_var.set(defaults.title_page)
        self.title_var.set(defaults.title)
        self.subtitle_var.set(defaults.subtitle)
        self.author_var.set(defaults.author)
        self.date_var.set(defaults.date)
        self.number_sections_var.set(defaults.number_sections)
        self.reference_docx_var.set(defaults.reference_docx)
        self.default_font_var.set(defaults.default_font)
        self.default_font_size_var.set(str(defaults.default_font_size or ""))
        self.table_borders_var.set(defaults.table_borders)
        self.hr_to_pagebreak_var.set(defaults.hr_to_pagebreak)
        self.mermaid_format_var.set(defaults.mermaid_format)
        self.mermaid_theme_var.set(defaults.mermaid_theme)
        self.mermaid_background_var.set(defaults.mermaid_background)
        self.mermaid_scale_var.set(str(defaults.mermaid_scale or ""))
        self.mermaid_min_dpi_var.set(str(defaults.mermaid_min_dpi))
        self.extra_args_text.delete("1.0", tk.END)
        self.extra_args_text.insert("1.0", " ".join(defaults.extra_pandoc_args))

    def _save(self) -> None:
        try:
            extra_args = _split_extra_args(self.extra_args_text.get("1.0", "end").strip())
            self.project.toc_depth = _parse_int(self.toc_depth_var.get(), "TOC depth", minimum=1, maximum=6)
            self.project.default_font_size = _parse_int(
                self.default_font_size_var.get(),
                "Font size",
                minimum=0,
                maximum=72,
                allow_empty=True,
            )
            mermaid_scale = _parse_float(
                self.mermaid_scale_var.get(),
                "Mermaid scale",
                minimum=0.0,
                maximum=10.0,
                allow_empty=True,
            )
            if mermaid_scale == 0.0 or self.mermaid_scale_var.get().strip() == "":
                mermaid_scale = 3.0
            mermaid_min_dpi = _parse_float(
                self.mermaid_min_dpi_var.get(),
                "Mermaid min DPI",
                minimum=0.0,
                maximum=2400.0,
                allow_empty=True,
            )
            if self.mermaid_min_dpi_var.get().strip() == "":
                mermaid_min_dpi = 450.0
        except ValueError as exc:
            messagebox.showerror("Settings", str(exc), parent=self)
            return

        self.project.toc = self.toc_var.get()
        self.project.title_page = self.title_page_var.get()
        self.project.title = self.title_var.get().strip()
        self.project.subtitle = self.subtitle_var.get().strip()
        self.project.author = self.author_var.get().strip()
        self.project.date = self.date_var.get().strip()
        self.project.number_sections = self.number_sections_var.get()
        self.project.reference_docx = self.reference_docx_var.get().strip()
        self.project.default_font = self.default_font_var.get().strip()
        self.project.table_borders = self.table_borders_var.get()
        self.project.mermaid_format = self.mermaid_format_var.get()
        self.project.mermaid_theme = self.mermaid_theme_var.get().strip() or "default"
        self.project.mermaid_background = self.mermaid_background_var.get().strip() or "white"
        self.project.mermaid_scale = mermaid_scale
        self.project.mermaid_min_dpi = mermaid_min_dpi
        self.project.hr_to_pagebreak = self.hr_to_pagebreak_var.get()
        self.project.extra_pandoc_args = extra_args
        self.project.save()
        ProjectRegistry().add(self.project)
        self.saved = True
        self.destroy()

    def _center(self) -> None:
        self.update_idletasks()
        parent_x = self.parent.winfo_rootx()
        parent_y = self.parent.winfo_rooty()
        x = parent_x + max(0, (self.parent.winfo_width() - self.winfo_width()) // 2)
        y = parent_y + max(0, (self.parent.winfo_height() - self.winfo_height()) // 2)
        self.geometry(f"+{x}+{y}")


def run_app() -> None:
    enable_high_dpi_awareness()
    dependency_error = None
    if not _skip_startup_dependency_setup():
        dependency_error = _run_dependency_setup_window()

    try:
        app = Md2DocApp()
        if dependency_error:
            app.after(250, lambda: messagebox.showwarning("Dependency setup", dependency_error, parent=app))
        app.mainloop()
    except Exception:
        _write_error_log(traceback.format_exc())
        raise


def _skip_startup_dependency_setup() -> bool:
    return os.environ.get("MD2DOC_SKIP_DEP_INSTALL", "").lower() in {"1", "true", "yes"}


def _write_error_log(message: str) -> None:
    try:
        path = app_data_dir() / "error.log"
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(message, encoding="utf-8")
    except Exception:
        pass


def _run_dependency_setup_window() -> str | None:
    root = tk.Tk()
    root.title("md2doc setup")
    root.geometry("460x150")
    root.resizable(False, False)
    root.withdraw()

    frame = ttk.Frame(root, padding=16)
    frame.grid(row=0, column=0, sticky="nsew")
    frame.columnconfigure(0, weight=1)

    status_var = tk.StringVar(value="Checking conversion tools...")
    ttk.Label(frame, text="Preparing conversion tools").grid(row=0, column=0, sticky="w")
    ttk.Label(frame, textvariable=status_var, wraplength=420).grid(row=1, column=0, sticky="ew", pady=(8, 10))
    progress = ttk.Progressbar(frame, mode="indeterminate")
    progress.grid(row=2, column=0, sticky="ew")

    root.update_idletasks()
    x = root.winfo_screenwidth() // 2 - root.winfo_width() // 2
    y = root.winfo_screenheight() // 2 - root.winfo_height() // 2
    root.geometry(f"+{x}+{y}")
    root.deiconify()

    events: queue.Queue[tuple[str, str | None]] = queue.Queue()

    def progress_callback(message: str) -> None:
        events.put(("progress", message))

    def worker() -> None:
        try:
            ensure_startup_dependencies(on_progress=progress_callback)
        except Exception as exc:
            events.put(("error", str(exc)))
        finally:
            events.put(("done", None))

    error_message: str | None = None

    def poll() -> None:
        nonlocal error_message
        while True:
            try:
                event, message = events.get_nowait()
            except queue.Empty:
                break
            if event == "progress" and message:
                status_var.set(message)
            elif event == "error" and message:
                error_message = message
            elif event == "done":
                root.destroy()
                return
        root.after(100, poll)

    threading.Thread(target=worker, daemon=True).start()
    root.after(100, poll)
    root.mainloop()

    # Explicitly clean up tkinter widgets and variables on the main thread
    # to avoid deferred garbage collection on background threads.
    try:
        del status_var
        del progress
        del root
    except Exception:
        pass

    return error_message


def enable_high_dpi_awareness() -> None:
    if os.name != "nt":
        return
    try:
        # Per-monitor v2 keeps Tk sharp when the app is moved between displays.
        if ctypes.windll.user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4)):
            return
    except (AttributeError, OSError):
        pass
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
        return
    except (AttributeError, OSError):
        pass
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except (AttributeError, OSError):
        pass


def _detect_ui_scale(root: tk.Tk) -> float:
    dpi = max(96.0, float(root.winfo_fpixels("1i")))
    return max(1.0, min(2.5, dpi / 96.0))


def _kind_label(kind: str) -> str:
    if kind == KIND_DOC2MD:
        return "Office to Markdown"
    if kind == KIND_QMD2PPT:
        return "Quarto to PowerPoint"
    if kind == KIND_HTML2PDF:
        return "HTML to PDF"
    return "Markdown to document"


def _input_label(kind: str) -> str:
    if kind == KIND_DOC2MD:
        return "Office"
    if kind == KIND_QMD2PPT:
        return "Quarto"
    if kind == KIND_HTML2PDF:
        return "HTML"
    return "Markdown"


def _state_label(action: str) -> str:
    return {
        "convert": "[TO CONVERT]",
        "skip": "[UP TO DATE]",
        "queued": "[QUEUED]",
        "running": "[RUNNING]",
        "done": "[DONE]",
        "skipped": "[UP TO DATE]",
        "failed": "[FAILED]",
    }.get(action, action)


def _reason_label(reason: str) -> str:
    return {
        "output missing": "Output file does not exist",
        "output is newer than source": "Output is newer than source",
        "no history and source is newer": "Source is newer than output",
        "source changed": "Source changed",
        "conversion settings changed": "Conversion settings changed",
        "conversion settings untracked": "Conversion settings need to be applied",
        "unchanged": "Output is up to date",
        "forced": "Force conversion enabled",
        "skip disabled": "Smart skip disabled",
    }.get(reason, reason)


def _split_extra_args(value: str) -> list[str]:
    if not value:
        return []
    return [_strip_quotes(arg) for arg in shlex.split(value, posix=os.name != "nt")]


def _parse_int(
    value: str,
    label: str,
    *,
    minimum: int | None = None,
    maximum: int | None = None,
    allow_empty: bool = False,
) -> int:
    value = value.strip()
    if not value and allow_empty:
        return 0
    try:
        parsed = int(value)
    except ValueError as exc:
        raise ValueError(f"{label} must be a whole number.") from exc
    if minimum is not None and parsed < minimum:
        raise ValueError(f"{label} must be at least {minimum}.")
    if maximum is not None and parsed > maximum:
        raise ValueError(f"{label} must be at most {maximum}.")
    return parsed


def _parse_float(
    value: str,
    label: str,
    *,
    minimum: float | None = None,
    maximum: float | None = None,
    allow_empty: bool = False,
) -> float:
    value = value.strip()
    if not value and allow_empty:
        return 0.0
    try:
        parsed = float(value)
    except ValueError as exc:
        raise ValueError(f"{label} must be a number.") from exc
    if minimum is not None and parsed < minimum:
        raise ValueError(f"{label} must be at least {minimum}.")
    if maximum is not None and parsed > maximum:
        raise ValueError(f"{label} must be at most {maximum}.")
    return parsed


def _strip_quotes(value: str) -> str:
    return value.strip().strip('"').strip("'")
