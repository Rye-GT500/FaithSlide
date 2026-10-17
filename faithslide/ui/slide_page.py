"""PPT 製作頁。"""
import logging
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from faithslide import paths, pipeline
from faithslide.pipeline import GenerationError
from faithslide.ui.background import run_in_background

logger = logging.getLogger(__name__)

WORD_PLACEHOLDER = "請選擇 Word 文件..."
PPT_PLACEHOLDER = "請選擇輸出 PPT 檔案名稱..."
MAX_WARNINGS_SHOWN = 10


class SlidePage(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.word_path_var = tk.StringVar(value=WORD_PLACEHOLDER)
        self.ppt_path_var = tk.StringVar(value=PPT_PLACEHOLDER)
        self.status_var = tk.StringVar(value="狀態：待命中")

        ttk.Label(self, text="Word 輸入:").grid(row=0, column=0, columnspan=2, pady=(20, 0), sticky="s")
        ttk.Entry(self, textvariable=self.word_path_var, width=50, state="readonly").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        ttk.Button(self, text="選擇 Word", command=self.select_word_file).grid(row=1, column=1, padx=5, pady=5, sticky="w")

        ttk.Label(self, text="PPT 輸出:").grid(row=2, column=0, columnspan=2, pady=(20, 0), sticky="s")
        ttk.Entry(self, textvariable=self.ppt_path_var, width=50, state="readonly").grid(row=3, column=0, padx=5, pady=5, sticky="e")
        ttk.Button(self, text="選擇儲存", command=self.select_save_path).grid(row=3, column=1, padx=5, pady=5, sticky="w")

        self.produce_btn = ttk.Button(self, text="製作", command=self.start)
        self.produce_btn.grid(row=4, column=0, columnspan=2, pady=(15, 0))

        self.progress_bar = ttk.Progressbar(self, orient="horizontal", mode="determinate")
        self.progress_bar.grid(row=5, column=0, columnspan=2, padx=5, pady=5, sticky="ew")
        ttk.Label(self, textvariable=self.status_var).grid(row=6, column=0, columnspan=2, padx=5, pady=5)

        for i in range(7):
            self.grid_rowconfigure(i, weight=1)
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)

    def on_show(self):
        self.focus_set()

    def on_return(self, event=None):
        self.start()

    def select_word_file(self):
        path = filedialog.askopenfilename(
            title="選擇 Word 證道文件",
            defaultextension=".docx",
            filetypes=[("Word 檔案", "*.docx"), ("所有檔案", "*.*")],
        )
        if path:
            self.word_path_var.set(path)
            logger.info(f"選取 Word 檔案: {path}")

    def select_save_path(self):
        path = filedialog.asksaveasfilename(
            title="指定輸出 PPT 檔案名稱",
            defaultextension=".pptx",
            filetypes=[("PowerPoint 檔案", "*.pptx"), ("所有檔案", "*.*")],
            initialfile="證道投影片.pptx",
        )
        if path:
            self.ppt_path_var.set(path)
            logger.info(f"選取 PPT 儲存路徑: {path}")

    def start(self):
        if self.produce_btn.instate(["disabled"]):   # 製作中
            return
        word_path, ppt_path = self.word_path_var.get(), self.ppt_path_var.get()
        if word_path == WORD_PLACEHOLDER:
            messagebox.showwarning("輸入錯誤", "請先選擇 Word 檔案")
            return
        if ppt_path == PPT_PLACEHOLDER:
            messagebox.showwarning("輸入錯誤", "請先選擇 PPT 儲存位置")
            return
        self.produce_btn.state(["disabled"])
        self._set_progress(0, "開始製作")
        run_in_background(
            self,
            lambda report_progress: pipeline.generate(word_path, ppt_path, report_progress),
            on_done=lambda result: self._finish(result.warnings),
            on_error=self._finish_with_error,
            on_progress=self._set_progress,
        )

    def _set_progress(self, progress, message):
        percent = progress * 100
        self.progress_bar.config(value=percent)
        self.status_var.set(f"進度 {percent:.2f}% : {message}")

    def _finish(self, warnings):
        self.produce_btn.state(["!disabled"])
        if warnings:
            shown = "\n".join(f"・{w}" for w in warnings[:MAX_WARNINGS_SHOWN])
            more = (f"\n…另有 {len(warnings) - MAX_WARNINGS_SHOWN} 則，詳見記錄檔：\n{paths.log_path()}"
                    if len(warnings) > MAX_WARNINGS_SHOWN else "")
            messagebox.showwarning("製作完成", f"製作完成，但有以下問題：\n{shown}{more}")
        else:
            messagebox.showinfo("", "製作完成")

    def _finish_with_error(self, error):
        self.produce_btn.state(["!disabled"])
        self.status_var.set("狀態：製作失敗")
        if isinstance(error, GenerationError):
            message = str(error)
        else:
            logger.error("製作投影片時發生未預期的錯誤", exc_info=error)
            message = f"製作時發生未預期的錯誤：{error}"
        messagebox.showwarning("錯誤", message)
