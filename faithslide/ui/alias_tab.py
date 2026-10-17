"""設定視窗：書卷別名分頁。"""
import tkinter as tk
from tkinter import messagebox, ttk

from faithslide import bible_books


class AliasTab(ttk.Frame):
    title = "書卷別名"

    def __init__(self, parent):
        super().__init__(parent, padding=15)
        self.books = bible_books.all_books()
        self.defaults = {b.id: b.aliases for b in bible_books.default_books()}
        self.saved = {b.id: list(b.aliases) for b in self.books}
        self.working = {book_id: list(aliases) for book_id, aliases in self.saved.items()}   # 儲存前不生效

        self.grid_columnconfigure(0, weight=1)

        ttk.Label(self, text="書卷：").grid(row=0, column=0, sticky="w")
        self.book_combo = ttk.Combobox(self, state="readonly", values=[f"{b.full}（{b.abbr}）" for b in self.books])
        self.book_combo.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(0, 10))
        self.book_combo.bind("<<ComboboxSelected>>", lambda e: self.refresh())

        self.abbr_label = ttk.Label(self, foreground="gray")
        self.abbr_label.grid(row=2, column=0, columnspan=2, sticky="w")

        self.alias_list = tk.Listbox(self, height=8)
        self.alias_list.grid(row=3, column=0, sticky="nsew", pady=5)
        side = ttk.Frame(self)
        side.grid(row=3, column=1, sticky="n", padx=(8, 0), pady=5)
        ttk.Button(side, text="移除所選", command=self.remove_selected).pack(fill="x")
        ttk.Button(side, text="還原預設", command=self.restore_default).pack(fill="x", pady=(5, 0))
        self.grid_rowconfigure(3, weight=1)

        self.new_alias_var = tk.StringVar()
        entry = ttk.Entry(self, textvariable=self.new_alias_var)
        entry.grid(row=4, column=0, sticky="ew", pady=5)
        entry.bind("<Return>", lambda e: self.add())
        ttk.Button(self, text="新增", command=self.add).grid(row=4, column=1, sticky="ew", padx=(8, 0), pady=5)

        ttk.Label(self, text="別名可用於週報的經文引用與經文查詢，不可與其他書卷的名稱重複。",
                  foreground="gray", wraplength=380).grid(row=5, column=0, columnspan=2, sticky="w", pady=(5, 0))

        self.book_combo.current(0)
        self.refresh()

    @property
    def book(self) -> bible_books.Book:
        return self.books[self.book_combo.current()]

    def refresh(self):
        book = self.book
        self.abbr_label.config(text=f"全名：{book.full}　簡寫：{book.abbr}（不可修改）")
        self.alias_list.delete(0, tk.END)
        for alias in self.working[book.id]:
            suffix = "（預設）" if alias in self.defaults[book.id] else ""
            self.alias_list.insert(tk.END, alias + suffix)

    def add(self):
        name = self.new_alias_var.get().strip()
        book = self.book
        problem = bible_books.alias_problem(name)
        if problem:
            messagebox.showwarning("無法新增", problem, parent=self)
            return
        if name in self.working[book.id]:
            return
        candidate = {**self.working, book.id: self.working[book.id] + [name]}
        rejected = [r for r in bible_books.check_aliases(candidate) if r.name == name]
        if rejected:
            messagebox.showwarning("無法新增", rejected[0].reason, parent=self)
            return
        if len(name) == 1 and not messagebox.askyesno(
                "確認新增", f"單字別名「{name}」可能把證道標題中的文字誤判為書卷，確定要新增？", parent=self):
            return
        self.working[book.id].append(name)
        self.new_alias_var.set("")
        self.refresh()

    def remove_selected(self):
        selected = self.alias_list.curselection()
        if selected:
            del self.working[self.book.id][selected[0]]
            self.refresh()

    def restore_default(self):
        self.working[self.book.id] = list(self.defaults[self.book.id])
        self.refresh()

    def has_changes(self) -> bool:
        return self.working != self.saved

    def save(self) -> bool:
        """儲存失敗時顯示原因並回傳 False。"""
        try:
            bible_books.save_user_aliases(self.working)
        except bible_books.AliasConflictError as e:
            details = "\n".join(f"・{r}" for r in e.rejected)
            messagebox.showwarning("無法儲存", f"以下書卷別名有問題，未儲存：\n{details}", parent=self)
            return False
        except OSError as e:
            messagebox.showwarning("無法儲存", f"寫入書卷別名設定失敗：{e}", parent=self)
            return False
        self.saved = {book_id: list(aliases) for book_id, aliases in self.working.items()}
        return True
