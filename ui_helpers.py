# ui_helpers.py

import os
import ctypes
import subprocess

def ask_and_open(path: str):
    """
    Pops up a simple Yes/No Windows dialog asking
    “Open the tracker file now?” and if “Yes” opens it.
    """
    res = ctypes.windll.user32.MessageBoxW(
        0,
        f"Tracker created at:\n{path}\n\nOpen it now?",
        "Open Tracker?",
        0x00000004 | 0x00000020  # MB_YESNO | MB_ICONQUESTION
    )
    # IDYES = 6
    if res == 6:
        subprocess.Popen(["start", path], shell=True)
