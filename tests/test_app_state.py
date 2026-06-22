from __future__ import annotations

from pathlib import Path
import tempfile
import unittest

from md2doc.app import Md2DocApp, ProjectState
from md2doc.converter import FileFingerprint, PlanItem


class FakeButton:
    def __init__(self) -> None:
        self.state = "normal"

    def configure(self, **kwargs: object) -> None:
        if "state" in kwargs:
            self.state = str(kwargs["state"])


class FakeThread:
    def __init__(self, alive: bool) -> None:
        self.alive = alive

    def is_alive(self) -> bool:
        return self.alive


class FakeVar:
    def __init__(self, value: object = "") -> None:
        self.value = value

    def set(self, value: object) -> None:
        self.value = value

    def get(self) -> object:
        return self.value


class FakeProgressBar:
    def __init__(self) -> None:
        self.options: dict[str, object] = {}

    def configure(self, **kwargs: object) -> None:
        self.options.update(kwargs)


class FakeLog:
    def configure(self, **_kwargs: object) -> None:
        pass

    def delete(self, *_args: object) -> None:
        pass

    def insert(self, *_args: object) -> None:
        pass

    def see(self, *_args: object) -> None:
        pass


class FakeTree:
    def __init__(self) -> None:
        self.items: dict[str, tuple[tuple[object, ...], tuple[str, ...]]] = {}

    def get_children(self) -> list[str]:
        return list(self.items)

    def delete(self, iid: str) -> None:
        self.items.pop(iid, None)

    def insert(self, _parent: str, _index: object, *, iid: str, values: tuple[object, ...], tags: tuple[str, ...]) -> None:
        self.items[iid] = (values, tags)


def make_plan_item(root: Path, name: str, action: str = "convert") -> PlanItem:
    source = root / name
    return PlanItem(
        source=source,
        relative_source=name,
        output=root / f"{source.stem}.docx",
        action=action,
        reason="output missing",
        fingerprint=FileFingerprint(size=1, mtime_ns=2, sha256="abc"),
        settings_signature="settings",
    )


class AppStateTests(unittest.TestCase):
    def make_app(self) -> Md2DocApp:
        app = object.__new__(Md2DocApp)
        app.project_states = {}
        app.worker = None
        app.scan_active = False
        app.scan_generation = 0
        app.scan_worker = None
        app.busy_buttons = [FakeButton(), FakeButton()]
        app.cancel_button = FakeButton()
        app.status_var = FakeVar()
        app.progress_var = FakeVar(0)
        app.progress_bar = FakeProgressBar()
        app.log = FakeLog()
        app.tree = FakeTree()
        return app

    def test_refresh_busy_state_disables_buttons_while_conversion_active_before_thread_alive(self) -> None:
        app = self.make_app()
        state = ProjectState()
        state.conversion_active = True
        app.project_states["project"] = state
        app.worker = FakeThread(False)

        Md2DocApp._refresh_busy_state(app)

        self.assertEqual([button.state for button in app.busy_buttons], ["disabled", "disabled"])
        self.assertEqual(app.cancel_button.state, "disabled")

    def test_load_project_state_does_not_scan_while_project_is_converting(self) -> None:
        app = self.make_app()
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            state = ProjectState()
            state.conversion_active = True
            state.status_text = "Converting a.md (0/1)"
            app.project_states[str(root.resolve())] = state

            scan_calls: list[str] = []
            app._scan = lambda: scan_calls.append("scan")
            app._update_project_list_display = lambda: None

            Md2DocApp._load_project_state(app, root)

            self.assertEqual(scan_calls, [])
            self.assertEqual(app.status_var.get(), "Converting a.md (0/1)")

    def test_scan_done_does_not_replace_items_while_project_is_converting(self) -> None:
        app = self.make_app()
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            state = ProjectState()
            state.conversion_active = True
            state.scan_active = True
            state.scan_generation = 2
            existing = make_plan_item(root, "old.md")
            state.plan_by_id["0"] = existing
            state.item_states[str(existing.source)] = ("[RUNNING]", "Running Pandoc", "running")
            app.project_states[str(root.resolve())] = state
            app.current_project = None
            app._update_project_list_display = lambda: None

            Md2DocApp._handle_scan_done(app, root, 2, "md2doc", [make_plan_item(root, "new.md")])

            self.assertEqual(state.plan_by_id, {"0": existing})
            self.assertEqual(state.item_states[str(existing.source)], ("[RUNNING]", "Running Pandoc", "running"))
            self.assertFalse(state.scan_active)


if __name__ == "__main__":
    unittest.main()
