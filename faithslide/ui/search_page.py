"""經文查詢頁：單一輸入框，接受全名、簡寫或別名。"""
import logging
import tkinter as tk
from tkinter import messagebox, ttk

from faithslide.reference_parser import ReferenceParseError
from faithslide.verse_search import lookup

logger = logging.getLogger(__name__)

HINT = "例：約3:16-18、約翰福音3:16,18、詩23（整章）；多段以「；」分隔"


class SearchPage(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.query_var = tk.StringVar()

        ttk.Label(self, text="📖 聖經經文查詢", font=("微軟正黑體", 16, "bold")).grid(row=0, column=0, columnspan=2, pady=(0, 20))

        self.entry = ttk.Entry(self, textvariable=self.query_var, font=("微軟正黑體", 12))
        self.entry.grid(row=1, column=0, sticky="ew", padx=(10, 5), pady=5)
        ttk.Button(self, text="查詢", command=self.search).grid(row=1, column=1, sticky="w", padx=(5, 10), pady=5)
        ttk.Label(self, text=HINT, foreground="gray").grid(row=2, column=0, columnspan=2, sticky="w", padx=10)

        self.text_box = tk.Text(self, wrap="word")
        self.text_box.grid(row=3, column=0, columnspan=2, sticky="nsew", padx=10, pady=10)

        self.grid_rowconfigure(3, weight=1)
        self.grid_columnconfigure(0, weight=1)

    def on_show(self):
        self.entry.focus_set()

    def on_return(self, event=None):
        self.search()

    def search(self):
        query = self.query_var.get()
        try:
            lines, problems = lookup(query)
        except ReferenceParseError as e:
            messagebox.showwarning("輸入錯誤", str(e))
            return
        self.text_box.delete(1.0, tk.END)
        self.text_box.insert(tk.END, "\n".join(lines))
        logger.info(f"查詢 {query}")
        if problems:
            messagebox.showwarning("查詢結果", "\n".join(problems))
