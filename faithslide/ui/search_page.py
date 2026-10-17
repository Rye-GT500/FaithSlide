"""經文查詢頁。"""
import logging
import tkinter as tk
from tkinter import messagebox, ttk

from faithslide import bible_api, bible_books
from faithslide.bible_api import BibleApiError

logger = logging.getLogger(__name__)

LABEL_FONT = ("微軟正黑體", 12)


class SearchPage(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.book_var = tk.StringVar()
        self.chapter_var = tk.StringVar()
        self.verse_var = tk.StringVar()

        ttk.Label(self, text="📖 聖經經文查詢", font=("微軟正黑體", 16, "bold")).grid(row=0, column=0, columnspan=2, pady=(0, 20))

        ttk.Label(self, text="書卷：", font=LABEL_FONT).grid(row=1, column=0, sticky="e", padx=5, pady=5)
        books = [book.abbr for book in bible_books.all_books()]
        ttk.Combobox(self, textvariable=self.book_var, values=books, width=15).grid(row=1, column=1, padx=5, pady=5, sticky="w")

        ttk.Label(self, text="章：", font=LABEL_FONT).grid(row=2, column=0, sticky="e", padx=5, pady=5)
        ttk.Entry(self, textvariable=self.chapter_var, width=18).grid(row=2, column=1, padx=5, pady=5, sticky="w")

        ttk.Label(self, text="節：", font=LABEL_FONT).grid(row=3, column=0, sticky="e", padx=5, pady=5)
        ttk.Entry(self, textvariable=self.verse_var, width=18).grid(row=3, column=1, padx=5, pady=5, sticky="w")

        ttk.Button(self, text="查詢", command=self.search).grid(row=4, column=0, columnspan=2, pady=(15, 0))

        self.text_box = tk.Text(self, wrap="word")
        self.text_box.grid(row=6, column=0, columnspan=2, sticky="nsew", padx=10, pady=10)

        for i in range(6):
            self.grid_rowconfigure(i, weight=1)
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)

    def on_return(self, event=None):
        self.search()

    def search(self):
        book_name = self.book_var.get().strip()
        chapter_text = self.chapter_var.get().strip()
        verse_text = self.verse_var.get().strip()

        if not book_name or not chapter_text:
            messagebox.showwarning("輸入錯誤", "請選擇書卷與章節")
            return
        book = bible_books.resolve(book_name)
        if book is None:
            messagebox.showwarning("輸入錯誤", "書卷名稱錯誤")
            return
        if not chapter_text.isdigit():
            messagebox.showwarning("輸入錯誤", "章數必須是數字")
            return
        start, end = None, None
        if verse_text:
            parts = [p.strip() for p in verse_text.split("-", 1)]
            if not all(p.isdigit() for p in parts):
                messagebox.showwarning("輸入錯誤", "節數格式錯誤")
                return
            start, end = int(parts[0]), int(parts[-1])

        try:
            verses = bible_api.get_chapter(book, int(chapter_text))
        except BibleApiError as e:
            self._show([str(e)])
            return

        if start is None:
            self._show(verses)
            return
        by_number = {v.number: v for v in verses}
        selected = []
        for number in range(start, end + 1):
            if number not in by_number:
                messagebox.showwarning("輸入錯誤", "節數錯誤")
                break
            selected.append(by_number[number])
        self._show(selected)
        logger.info(f"查詢 {book.full} 第 {chapter_text} 章 {verse_text}")

    def _show(self, lines):
        self.text_box.delete(1.0, tk.END)
        for line in lines:
            text = line if isinstance(line, str) else f"{line.number}. {line.text}"
            self.text_box.insert(tk.END, text + "\n")
