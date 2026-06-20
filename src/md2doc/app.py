from __future__ import annotations

import ctypes
from dataclasses import replace
import os
from pathlib import Path
import queue
import shlex
import threading
import tkinter as tk
from tkinter import filedialog, font as tkfont, messagebox, simpledialog, ttk

from .converter import (
    ConvertSettings,
    ConversionResult,
    PlanItem,
    check_dependencies,
    missing_dependency_message,
    plan_conversions,
    run_conversions,
    scan_source_files,
    settings_from_project,
)
from .project import (
    KIND_DOC2MD,
    KIND_MD2DOC,
    ProjectConfig,
    ProjectRegistry,
    create_project,
    load_project,
)


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
        self.plan_by_id: dict[str, PlanItem] = {}
        self.iid_by_source: dict[str, str] = {}
        self.event_queue: queue.Queue[object] = queue.Queue()
        self.worker: threading.Thread | None = None
        self.conversion_total = 0
        self.conversion_done = 0
        self.converted_count = 0
        self.skipped_count = 0
        self.failed_count = 0

        self._build_ui()
        self._load_projects()
        self.after(150, self._poll_events)

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
            values=("docx", "html", "pdf"),
            state="readonly",
            width=8,
        )
        self.format_box.grid(row=0, column=1, sticky="w", padx=self._pad(6, 16))

        ttk.Label(controls, text="Output").grid(row=0, column=2, sticky="w")
        self.output_var = tk.StringVar(value=".")
        self.output_entry = ttk.Entry(controls, textvariable=self.output_var, width=28)
        self.output_entry.grid(row=0, column=3, sticky="ew", padx=self._pad(6, 16))

        self.force_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(controls, text="Force", variable=self.force_var).grid(row=0, column=4, sticky="w")

        action_bar = ttk.Frame(main)
        action_bar.grid(row=3, column=0, sticky="ew", pady=self._pad(8, 0))
        self.scan_button = ttk.Button(action_bar, text="Scan", command=self._scan)
        self.scan_button.pack(side="left")
        self.convert_selected_button = ttk.Button(action_bar, text="Convert selected", command=self._convert_selected)
        self.convert_selected_button.pack(side="left", padx=self._pad(8, 0))
        self.convert_all_button = ttk.Button(action_bar, text="Convert all", command=self._convert_all)
        self.convert_all_button.pack(side="left", padx=self._pad(8, 0))
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
        progress_frame.columnconfigure(1, weight=1)
        self.status_var = tk.StringVar(value="Ready")
        ttk.Label(progress_frame, textvariable=self.status_var, width=34).grid(row=0, column=0, sticky="w")
        self.progress_var = tk.DoubleVar(value=0)
        self.progress_bar = ttk.Progressbar(
            progress_frame,
            orient="horizontal",
            mode="determinate",
            variable=self.progress_var,
            maximum=1,
        )
        self.progress_bar.grid(row=0, column=1, sticky="ew")

        columns = ("state", "file", "reason")
        self.tree = ttk.Treeview(main, columns=columns, show="headings", selectmode="extended")
        self.tree.heading("state", text="State")
        self.tree.heading("file", text="Markdown")
        self.tree.heading("reason", text="Reason")
        self.tree.column("state", width=self._px(118), anchor="w", stretch=False)
        self.tree.column("file", width=self._px(430), anchor="w")
        self.tree.column("reason", width=self._px(260), anchor="w")
        self.tree.tag_configure("skip", foreground="#555", background="#f4f5f7", font=self.state_font)
        self.tree.tag_configure("convert", foreground="#6f4e00", background="#fff3c4", font=self.state_font)
        self.tree.tag_configure("queued", foreground="#444", background="#eef1f5", font=self.state_font)
        self.tree.tag_configure("running", foreground="#0b5cad", background="#e6f2ff", font=self.state_font)
        self.tree.tag_configure("done", foreground="#127a3a", background="#e7f6ed", font=self.state_font)
        self.tree.tag_configure("failed", foreground="#b00020", background="#fde8e8", font=self.state_font)
        self.tree.grid(row=2, column=0, sticky="nsew")

        scrollbar = ttk.Scrollbar(main, orient="vertical", command=self.tree.yview)
        scrollbar.grid(row=2, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=scrollbar.set)

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
            text="Markdown  ->  Word / HTML / PDF   (Pandoc)",
            variable=kind_var,
            value=KIND_MD2DOC,
        ).grid(row=1, column=0, sticky="w", pady=self._pad(2, 0))
        ttk.Radiobutton(
            frame,
            text="Word / PPT / Excel  ->  Markdown   (MarkItDown)",
            variable=kind_var,
            value=KIND_DOC2MD,
        ).grid(row=2, column=0, sticky="w", pady=self._pad(2, 0))

        result: dict[str, str | None] = {"kind": None}

        def confirm() -> None:
            result["kind"] = kind_var.get()
            dialog.destroy()

        buttons = ttk.Frame(frame)
        buttons.grid(row=3, column=0, sticky="e", pady=self._pad(16, 0))
        ttk.Button(buttons, text="Cancel", command=dialog.destroy).grid(row=0, column=0, padx=self._pad(0, 8))
        ttk.Button(buttons, text="Continue", command=confirm).grid(row=0, column=1)

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
        selection = self.project_list.curselection()
        if not selection:
            return
        self._set_project(self.projects[selection[0]])

    def _set_project(self, project: ProjectConfig) -> None:
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
        self._clear_table()
        self._scan()

    def _apply_kind_to_ui(self, kind: str) -> None:
        if kind == KIND_DOC2MD:
            self.format_box.configure(values=("md",), state="disabled")
            self.format_var.set("md")
            self.tree.heading("file", text="Office document")
        else:
            self.format_box.configure(values=("docx", "html", "pdf"), state="readonly")
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

    def _scan(self) -> None:
        project = self.current_project
        if not project:
            return
        try:
            settings = self._settings()
            sources = scan_source_files(
                project.root,
                kind=settings.kind,
                recursive=settings.recursive,
                output_dir=settings.output_dir,
            )
            planned = plan_conversions(project.root, sources, settings)
        except Exception as exc:
            messagebox.showerror("Scan failed", str(exc))
            return
        self._clear_table()
        for index, item in enumerate(planned):
            iid = str(index)
            self._insert_or_update_plan_item(iid, item)
        convert_count = sum(1 for item in planned if item.action == "convert")
        skip_count = len(planned) - convert_count
        self.status_var.set(f"Scanned {len(planned)} file(s): {convert_count} to convert, {skip_count} skipped")
        self._append_log(f"Scanned {len(planned)} {_input_label(project.kind)} file(s).")

    def _convert_selected(self) -> None:
        selection = self.tree.selection()
        if not selection:
            messagebox.showinfo("Convert selected", "Select one or more files first.")
            return
        sources = [self.plan_by_id[iid].source for iid in selection]
        self._start_conversion(sources)

    def _convert_all(self) -> None:
        project = self.current_project
        if not project:
            return
        settings = self._settings()
        sources = scan_source_files(
            project.root,
            kind=settings.kind,
            recursive=settings.recursive,
            output_dir=settings.output_dir,
        )
        self._start_conversion(sources)

    def _start_conversion(self, sources: list[Path]) -> None:
        project = self._require_project()
        settings = self._settings()
        if self.worker and self.worker.is_alive():
            messagebox.showinfo("Busy", "A conversion is already running.")
            return
        planned = plan_conversions(project.root, sources, settings)
        self._prepare_conversion_progress(planned)

        def work() -> None:
            try:
                results = run_conversions(
                    project.root,
                    sources,
                    settings,
                    on_event=self.event_queue.put,
                    on_start=lambda item: self.event_queue.put(("start", item)),
                )
                converted = sum(1 for result in results if result.status == "converted")
                skipped = sum(1 for result in results if result.status == "skipped")
                failed = sum(1 for result in results if result.status == "failed")
                self.event_queue.put(("done", converted, skipped, failed))
            except Exception as exc:
                self.event_queue.put(("error", str(exc)))

        self._append_log(f"Starting conversion for {len(sources)} file(s).")
        self._set_busy(True)
        self.worker = threading.Thread(target=work, daemon=True)
        self.worker.start()

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

    def _poll_events(self) -> None:
        while True:
            try:
                event = self.event_queue.get_nowait()
            except queue.Empty:
                break
            if isinstance(event, tuple):
                kind = event[0]
                if kind == "start":
                    self._mark_item_running(event[1])
                    continue
                if kind == "error":
                    message = str(event[1])
                    self._append_log(message)
                    self.status_var.set("Conversion failed")
                    self._set_busy(False)
                    messagebox.showerror("Conversion failed", message)
                    continue
                if kind == "done":
                    converted, skipped, failed = event[1], event[2], event[3]
                    self.status_var.set(
                        f"Finished: {converted} converted, {skipped} skipped, {failed} failed"
                    )
                    self._append_log(
                        f"Finished: {converted} converted, {skipped} skipped, {failed} failed."
                    )
                    self._set_busy(False)
            else:
                self._handle_conversion_result(event)
        self.after(150, self._poll_events)

    def _clear_table(self) -> None:
        self.plan_by_id.clear()
        self.iid_by_source.clear()
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
                tool = "MarkItDown" if self.current_project and self.current_project.kind == KIND_DOC2MD else "Pandoc"
                self._set_item_state(item, _state_label("queued"), f"Waiting for {tool}", "queued")

    def _mark_item_running(self, item: PlanItem) -> None:
        tool = "MarkItDown" if self.current_project and self.current_project.kind == KIND_DOC2MD else "Pandoc"
        self._set_item_state(item, _state_label("running"), f"Running {tool}", "running")
        self.status_var.set(f"Converting {item.relative_source} ({self.conversion_done}/{self.conversion_total})")

    def _handle_conversion_result(self, result: ConversionResult) -> None:
        self.conversion_done += 1
        self.progress_var.set(self.conversion_done)
        if result.status == "converted":
            self.converted_count += 1
            self._set_item_state(result.item, _state_label("done"), "Output generated", "done")
        elif result.status == "skipped":
            self.skipped_count += 1
            self._set_item_state(result.item, _state_label("skipped"), _reason_label(result.message), "skip")
        else:
            self.failed_count += 1
            self._set_item_state(result.item, _state_label("failed"), result.message, "failed")
        self.status_var.set(
            f"Progress {self.conversion_done}/{self.conversion_total}: "
            f"{self.converted_count} converted, {self.skipped_count} skipped, {self.failed_count} failed"
        )
        self._append_log(f"{result.status}: {result.item.relative_source} - {result.message}")

    def _set_item_state(self, item: PlanItem, state: str, reason: str, tag: str) -> None:
        iid = self.iid_by_source.get(str(item.source))
        if iid is None or not self.tree.exists(iid):
            iid = self._next_iid()
            self._insert_or_update_plan_item(iid, item)
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

    def _build_mermaid_tab(self, frame: ttk.Frame) -> None:
        frame.columnconfigure(1, weight=1)
        fields = [
            ("Theme", self.mermaid_theme_var),
            ("Background", self.mermaid_background_var),
        ]
        for row, (label, variable) in enumerate(fields):
            ttk.Label(frame, text=label).grid(row=row, column=0, sticky="w", pady=self.parent._pad(4, 0))
            ttk.Entry(frame, textvariable=variable).grid(
                row=row,
                column=1,
                sticky="ew",
                pady=self.parent._pad(4, 0),
            )

        ttk.Label(frame, text="Format").grid(row=2, column=0, sticky="w", pady=self.parent._pad(8, 0))
        ttk.Combobox(
            frame,
            textvariable=self.mermaid_format_var,
            values=("png", "svg", "pdf"),
            state="readonly",
            width=10,
        ).grid(row=2, column=1, sticky="w", pady=self.parent._pad(8, 0))

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
        self.mermaid_format_var.set(defaults.mermaid_format)
        self.mermaid_theme_var.set(defaults.mermaid_theme)
        self.mermaid_background_var.set(defaults.mermaid_background)
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
    app = Md2DocApp()
    app.mainloop()


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
    return "Office to Markdown" if kind == KIND_DOC2MD else "Markdown to document"


def _input_label(kind: str) -> str:
    return "Office" if kind == KIND_DOC2MD else "Markdown"


def _state_label(action: str) -> str:
    return {
        "convert": "[CONVERT]",
        "skip": "[SKIP]",
        "queued": "[QUEUED]",
        "running": "[RUNNING]",
        "done": "[DONE]",
        "skipped": "[SKIPPED]",
        "failed": "[FAILED]",
    }.get(action, action)


def _reason_label(reason: str) -> str:
    return {
        "output missing": "Output file does not exist",
        "output is newer than source": "Output is newer than source",
        "no history and source is newer": "Source is newer than output",
        "source changed": "Markdown changed",
        "conversion settings changed": "Conversion settings changed",
        "unchanged": "Unchanged since last conversion",
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


def _strip_quotes(value: str) -> str:
    return value.strip().strip('"').strip("'")
