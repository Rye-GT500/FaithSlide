"""設定視窗：關於分頁（版本、檢查更新、記錄檔）。"""
import os
from tkinter import messagebox, ttk

from faithslide import paths, updater
from faithslide.background import run_in_background
from faithslide.version import CURRENT_VERSION


class AboutTab(ttk.Frame):
    title = "關於"

    def __init__(self, parent, on_update_found):
        """on_update_found(info)：手動檢查發現新版時呼叫（由主視窗開啟更新對話框）。"""
        super().__init__(parent, padding=15)
        self.on_update_found = on_update_found

        ttk.Label(self, text="FaithSlide", font=("微軟正黑體", 14, "bold")).pack(anchor="w")
        ttk.Label(self, text=f"版本 v{CURRENT_VERSION}").pack(anchor="w", pady=(2, 10))
        self.check_btn = ttk.Button(self, text="檢查更新", command=self.check_update)
        self.check_btn.pack(anchor="w")

        ttk.Separator(self).pack(fill="x", pady=15)

        ttk.Label(self, text="記錄檔", font=("微軟正黑體", 11, "bold")).pack(anchor="w")
        ttk.Label(self, text="回報問題時，請附上這個資料夾裡的 FaithSlide.log。",
                  wraplength=380).pack(anchor="w", pady=(2, 6))
        ttk.Button(self, text="開啟記錄檔資料夾", command=self.open_log_folder).pack(anchor="w")
        ttk.Label(self, text=paths.user_dir(), foreground="gray", wraplength=380).pack(anchor="w", pady=(4, 0))

    def check_update(self):
        self.check_btn.state(["disabled"])
        self.check_btn.config(text="檢查中…")
        run_in_background(
            self,
            lambda report_progress: updater.check(raise_errors=True),
            on_done=self._checked,
            on_error=self._check_failed,
        )

    def _checked(self, info):
        self._reset_button()
        if info:
            self.on_update_found(info)
        else:
            messagebox.showinfo("檢查更新", f"目前已是最新版本（v{CURRENT_VERSION}）。", parent=self)

    def _check_failed(self, error):
        self._reset_button()
        messagebox.showwarning("檢查更新", f"檢查更新失敗：{error}", parent=self)

    def _reset_button(self):
        self.check_btn.state(["!disabled"])
        self.check_btn.config(text="檢查更新")

    def open_log_folder(self):
        os.startfile(paths.user_dir())

    # 設定視窗的共用介面：此分頁沒有需要儲存的內容
    def has_changes(self) -> bool:
        return False

    def save(self) -> bool:
        return True
