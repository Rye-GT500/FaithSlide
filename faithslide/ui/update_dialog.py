"""有新版本時的對話框：一鍵下載並安裝，或前往下載頁面。"""
import logging
import tkinter as tk
import webbrowser
from tkinter import messagebox, ttk

from faithslide import updater
from faithslide.background import run_in_background
from faithslide.version import CURRENT_VERSION

logger = logging.getLogger(__name__)


class UpdateDialog(tk.Toplevel):
    def __init__(self, parent, info: updater.UpdateInfo, on_install_started):
        """on_install_started：安裝檔啟動後呼叫，用來關閉主程式。"""
        super().__init__(parent)
        self.info = info
        self.on_install_started = on_install_started
        self.title("有新版本")
        self.geometry("460x420")
        self.transient(parent)
        self.grab_set()

        body = ttk.Frame(self, padding=15)
        body.pack(fill="both", expand=True)
        ttk.Label(body, text=f"發現新版本 v{info.version}", font=("微軟正黑體", 13, "bold")).pack(anchor="w")
        ttk.Label(body, text=f"目前版本：v{CURRENT_VERSION}", foreground="gray").pack(anchor="w", pady=(2, 10))

        notes = tk.Text(body, height=10, wrap="char", relief="solid", borderwidth=1, padx=6, pady=4)
        notes.insert("1.0", updater.plain_notes(info.notes) or "（沒有更新說明）")
        notes.config(state="disabled")
        notes.pack(fill="both", expand=True)

        self.progress = ttk.Progressbar(body, mode="determinate")
        self.status_var = tk.StringVar()
        self.status_label = ttk.Label(body, textvariable=self.status_var)

        buttons = ttk.Frame(body)
        buttons.pack(pady=(12, 0))
        self.install_btn = ttk.Button(buttons, text="立即更新", command=self.start_update)
        self.install_btn.pack(side="left", padx=5)
        self.page_btn = ttk.Button(buttons, text="前往下載頁面", command=self.open_page)
        self.page_btn.pack(side="left", padx=5)
        self.later_btn = ttk.Button(buttons, text="稍後", command=self.destroy)
        self.later_btn.pack(side="left", padx=5)

        if not info.can_install:
            self.install_btn.state(["disabled"])
            self.status_var.set("這個版本無法自動更新，請到下載頁面手動下載。")
            self.status_label.pack(anchor="w", pady=(8, 0))

    def open_page(self):
        webbrowser.open(self.info.url)
        self.destroy()

    def start_update(self):
        for button in (self.install_btn, self.page_btn, self.later_btn):
            button.state(["disabled"])
        self.protocol("WM_DELETE_WINDOW", lambda: None)   # 下載中不可關閉
        self.progress.pack(fill="x", pady=(10, 0))
        self.status_label.pack(anchor="w", pady=(4, 0))
        self.status_var.set("下載中…")
        run_in_background(
            self,
            lambda report_progress: updater.download(self.info, report_progress),
            on_done=self._install,
            on_error=self._failed,
            on_progress=self._show_progress,
        )

    def _show_progress(self, done: int, total: int):
        if total:
            self.progress.config(value=done / total * 100)
            self.status_var.set(f"下載中… {done / 1_048_576:.1f} / {total / 1_048_576:.1f} MB")

    def _install(self, installer_path: str):
        self.status_var.set("啟動安裝程式…")
        try:
            updater.install(installer_path)
        except OSError as e:
            # 例如在系統管理員權限確認視窗按了「否」
            logger.warning(f"啟動安裝檔失敗：{e}")
            self._failed(updater.UpdateError("安裝程式沒有啟動。如果剛才取消了權限確認，可以再按一次「立即更新」。"))
            return
        self.on_install_started()

    def _failed(self, error: Exception):
        if not isinstance(error, updater.UpdateError):
            logger.error("更新失敗", exc_info=error)
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        for button in (self.install_btn, self.page_btn, self.later_btn):
            button.state(["!disabled"])
        self.progress.config(value=0)
        self.status_var.set("")
        messagebox.showwarning("更新失敗", f"{error}\n\n也可以按「前往下載頁面」手動下載安裝。", parent=self)
