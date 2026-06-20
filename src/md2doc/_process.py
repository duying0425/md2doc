from __future__ import annotations

import subprocess
import sys


def hidden_subprocess_kwargs() -> dict[str, object]:
    """Keep child console windows hidden when the GUI runs on Windows."""

    if sys.platform != "win32" or not hasattr(subprocess, "STARTUPINFO"):
        return {}

    startupinfo = subprocess.STARTUPINFO()
    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    startupinfo.wShowWindow = subprocess.SW_HIDE
    return {
        "startupinfo": startupinfo,
        "creationflags": getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000),
    }
