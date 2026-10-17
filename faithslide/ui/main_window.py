"""主視窗：頁面切換（F5）、退出（Esc）、右上角齒輪開啟設定、啟動時檢查更新。"""
import logging
import tkinter as tk
import webbrowser
from tkinter import messagebox, ttk

from faithslide import bible_books, updater
from faithslide.ui.background import run_in_background
from faithslide.ui.search_page import SearchPage
from faithslide.ui.settings_dialog import SettingsDialog
from faithslide.ui.tooltip import Tooltip
from faithslide.ui.slide_page import SlidePage

logger = logging.getLogger(__name__)


class MainWindow:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("FaithSlide")
        self.root.geometry("650x650")

        top_bar = ttk.Frame(self.root, padding=(10, 8, 10, 0))
        top_bar.grid(row=0, column=0, sticky="ew")
        ttk.Style(self.root).configure("Gear.TButton", font=("Segoe UI Symbol", 14))
        gear = ttk.Button(top_bar, text="⚙", width=3, style="Gear.TButton", command=self.open_settings)
        gear.pack(side="right")
        Tooltip(gear, "設定（書卷別名、講員）")

        container = ttk.Frame(self.root, padding=(20, 0, 20, 20))
        container.grid(row=1, column=0, sticky="nsew")
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        # 兩頁都建立好，切換時只顯示其中一頁，輸入內容不會因切換而消失
        self.pages = [SlidePage(container), SearchPage(container)]
        for page in self.pages:
            page.grid(row=0, column=0, sticky="nsew")
        self.current = 0

        buttons = ttk.Frame(self.root)
        buttons.grid(row=2, column=0, pady=(0, 15))
        ttk.Button(buttons, text="切換", command=self.change_page).pack(side="left", padx=5)
        ttk.Button(buttons, text="退出", command=self.close).pack(side="left", padx=5)

        self.root.grid_rowconfigure(1, weight=1)
        self.root.grid_columnconfigure(0, weight=1)

        self.root.bind("<F5>", self.change_page)
        self.root.bind("<Escape>", self.close)
        self.root.bind("<Return>", lambda event: self.pages[self.current].on_return(event))
        self._show_page(0)

    def run(self):
        self.root.after(0, self._notify_load_warnings)
        self._check_update()
        self.root.mainloop()

    def change_page(self, event=None):
        self._show_page(1 - self.current)

    def open_settings(self):
        SettingsDialog(self.root)

    def _notify_load_warnings(self):
        warnings = bible_books.load_warnings()
        if warnings:
            details = "\n".join(f"・{w}" for w in warnings)
            messagebox.showwarning("書卷別名設定", f"部分自訂書卷別名未套用：\n{details}\n\n可在右上角「⚙ 設定」的「書卷別名」中修正。")

    def close(self, event=None):
        self.root.destroy()

    def _show_page(self, index):
        self.current = index
        page = self.pages[index]
        page.tkraise()
        page.on_show()

    def _check_update(self):
        run_in_background(
            self.root,
            lambda report_progress: updater.check(),
            on_done=self._ask_update,
            on_error=lambda error: logger.error("檢查更新失敗", exc_info=error),
        )

    def _ask_update(self, info):
        if info and messagebox.askyesno("更新通知", f"發現新版本 v{info.version}！\n是否前往下載？"):
            webbrowser.open(info.url)
