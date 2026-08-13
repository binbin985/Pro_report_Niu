from __future__ import annotations

import tkinter as tk
from tkinter import ttk


class Combopicker(ttk.Combobox):
    """Compatibility fallback for projects that expect a ComBoPicker widget.

    This minimal implementation supports the subset used by code_report.py:
    - constructor with values and entrywidth
    - assigning values via widget['values']
    - reading selected text via current_value/get()
    """

    def __init__(self, master=None, values=None, entrywidth="20", **kwargs):
        width = int(entrywidth) if str(entrywidth).isdigit() else 20
        super().__init__(master=master, values=values or [], width=width, **kwargs)
        self._text_var = tk.StringVar()
        self.configure(textvariable=self._text_var)

    @property
    def current_value(self) -> str:
        return self.get().strip()

    def clear(self) -> None:
        self.set("")
