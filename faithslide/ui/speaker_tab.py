"""設定視窗：講員分頁。"""
import tkinter as tk
from tkinter import messagebox, ttk

from faithslide import settings
from faithslide.sermon_parser import speaker_word_problem


class SpeakerTab(ttk.Frame):
    title = "講員"

    def __init__(self, parent):
        super().__init__(parent, padding=15)
        self.defaults = settings.default(settings.SPEAKER_WORDS)
        self.saved = settings.get(settings.SPEAKER_WORDS)
        self.working = list(self.saved)   # 儲存前不生效

        self.grid_columnconfigure(0, weight=1)

        ttk.Label(self, text="以下文字會從證道內容中刪除，不會出現在投影片上，例如講員的職稱與姓名。",
                  wraplength=380).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))

        self.word_list = tk.Listbox(self, height=10)
        self.word_list.grid(row=1, column=0, sticky="nsew", pady=5)
        side = ttk.Frame(self)
        side.grid(row=1, column=1, sticky="n", padx=(8, 0), pady=5)
        ttk.Button(side, text="移除所選", command=self.remove_selected).pack(fill="x")
        ttk.Button(side, text="還原預設", command=self.restore_default).pack(fill="x", pady=(5, 0))
        self.grid_rowconfigure(1, weight=1)

        self.new_word_var = tk.StringVar()
        entry = ttk.Entry(self, textvariable=self.new_word_var)
        entry.grid(row=2, column=0, sticky="ew", pady=5)
        entry.bind("<Return>", lambda e: self.add())
        ttk.Button(self, text="新增", command=self.add).grid(row=2, column=1, sticky="ew", padx=(8, 0), pady=5)

        ttk.Label(self, text="請輸入兩個字以上；單一個字會把標題中所有相同的字刪掉。",
                  foreground="gray", wraplength=380).grid(row=3, column=0, columnspan=2, sticky="w", pady=(5, 0))

        self.refresh()

    def refresh(self):
        self.word_list.delete(0, tk.END)
        for word in self.working:
            suffix = "（預設）" if word in self.defaults else ""
            self.word_list.insert(tk.END, word + suffix)

    def add(self):
        word = self.new_word_var.get().strip()
        problem = speaker_word_problem(word)
        if problem:
            messagebox.showwarning("無法新增", problem, parent=self)
            return
        if word not in self.working:
            self.working.append(word)
        self.new_word_var.set("")
        self.refresh()

    def remove_selected(self):
        selected = self.word_list.curselection()
        if selected:
            del self.working[selected[0]]
            self.refresh()

    def restore_default(self):
        self.working = list(self.defaults)
        self.refresh()

    def has_changes(self) -> bool:
        return self.working != self.saved

    def save(self) -> bool:
        """儲存失敗時顯示原因並回傳 False。"""
        try:
            settings.save(settings.SPEAKER_WORDS, self.working)
        except OSError as e:
            messagebox.showwarning("無法儲存", f"寫入講員設定失敗：{e}", parent=self)
            return False
        self.saved = list(self.working)
        return True
