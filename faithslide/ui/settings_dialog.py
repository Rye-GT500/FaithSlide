"""設定視窗（主視窗右上角齒輪開啟）：各設定以分頁呈現，共用儲存／取消。"""
import tkinter as tk
from tkinter import messagebox, ttk

from faithslide.ui.about_tab import AboutTab
from faithslide.ui.alias_tab import AliasTab
from faithslide.ui.speaker_tab import SpeakerTab


class SettingsDialog(tk.Toplevel):
    def __init__(self, parent, on_update_found):
        """on_update_found(info, parent)：在「關於」分頁手動檢查到新版時呼叫。"""
        super().__init__(parent)
        self.title("設定")
        self.geometry("460x520")
        self.transient(parent)
        self.grab_set()   # 開啟期間主視窗不可操作

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=(10, 0))
        self.tabs = [AliasTab(self.notebook), SpeakerTab(self.notebook),
                     AboutTab(self.notebook, lambda info: on_update_found(info, self))]
        for tab in self.tabs:
            self.notebook.add(tab, text=tab.title)

        buttons = ttk.Frame(self, padding=10)
        buttons.pack()
        ttk.Button(buttons, text="儲存", command=self.save).pack(side="left", padx=5)
        ttk.Button(buttons, text="取消", command=self.cancel).pack(side="left", padx=5)

        self.bind("<Escape>", lambda e: self.cancel())
        self.protocol("WM_DELETE_WINDOW", self.cancel)

    def save(self):
        for tab in self.tabs:
            if tab.has_changes() and not tab.save():
                self.notebook.select(tab)
                return
        self.destroy()

    def cancel(self):
        if any(tab.has_changes() for tab in self.tabs) and not messagebox.askyesno(
                "放棄變更", "有尚未儲存的變更，確定要關閉？", parent=self):
            return
        self.destroy()
