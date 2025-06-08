import os
import tkinter as tk
from tkinter import messagebox

def ask_and_open(path: str) -> None:
    """
    Pops up a Yes/No dialog asking whether to open `path`.
    If the user clicks “Yes”, opens the file.
    """
    root = tk.Tk()
    root.withdraw()
    if messagebox.askyesno(
        title="Open Excel Tracker?",
        message=f"Your tracker was updated.\n\nOpen it now?"
    ):
        os.startfile(path)
    root.destroy()
