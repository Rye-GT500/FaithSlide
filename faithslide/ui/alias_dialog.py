"""書卷別名設定視窗。"""
import tkinter as tk
from tkinter import messagebox, ttk

from faithslide import bible_books


class AliasDialog(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("書卷別名設定")
        self.geometry("420x460")
        self.transient(parent)
        self.grab_set()   # 開啟期間主視窗不可操作

        self.books = bible_books.all_books()
        self.defaults = {b.id: b.aliases for b in bible_books.default_books()}
        self.working = {b.id: list(b.aliases) for b in self.books}   # 編輯中的別名，儲存前不生效

        body = ttk.Frame(self, padding=15)
        body.pack(fill="both", expand=True)
        body.grid_columnconfigure(0, weight=1)

        ttk.Label(body, text="書卷：").grid(row=0, column=0, sticky="w")
        self.book_var = tk.StringVar()
        self.book_combo = ttk.Combobox(body, textvariable=self.book_var, state="readonly",
                                       values=[f"{b.full}（{b.abbr}）" for b in self.books])
        self.book_combo.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(0, 10))
        self.book_combo.bind("<<ComboboxSelected>>", lambda e: self.refresh())

        self.abbr_label = ttk.Label(body, foreground="gray")
        self.abbr_label.grid(row=2, column=0, columnspan=2, sticky="w")

        self.alias_list = tk.Listbox(body, height=8)
        self.alias_list.grid(row=3, column=0, sticky="nsew", pady=5)
        side = ttk.Frame(body)
        side.grid(row=3, column=1, sticky="n", padx=(8, 0), pady=5)
        ttk.Button(side, text="移除所選", command=self.remove_selected).pack(fill="x")
        ttk.Button(side, text="還原預設", command=self.restore_default).pack(fill="x", pady=(5, 0))
        body.grid_rowconfigure(3, weight=1)

        self.new_alias_var = tk.StringVar()
        entry = ttk.Entry(body, textvariable=self.new_alias_var)
        entry.grid(row=4, column=0, sticky="ew", pady=5)
        entry.bind("<Return>", lambda e: self.add())
        ttk.Button(body, text="新增", command=self.add).grid(row=4, column=1, sticky="ew", padx=(8, 0), pady=5)

        ttk.Label(body, text="別名可用於證道經文引用與經文查詢，不可與其他書卷的名稱重複。",
                  foreground="gray", wraplength=380).grid(row=5, column=0, columnspan=2, sticky="w", pady=(5, 10))

        buttons = ttk.Frame(body)
        buttons.grid(row=6, column=0, columnspan=2)
        ttk.Button(buttons, text="儲存", command=self.save).pack(side="left", padx=5)
        ttk.Button(buttons, text="取消", command=self.destroy).pack(side="left", padx=5)
        self.bind("<Escape>", lambda e: self.destroy())

        self.book_combo.current(0)
        self.refresh()
        entry.focus_set()

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

    def save(self):
        try:
            bible_books.save_user_aliases(self.working)
        except bible_books.AliasConflictError as e:
            details = "\n".join(f"・{r}" for r in e.rejected)
            messagebox.showwarning("無法儲存", f"以下別名有問題，未儲存任何變更：\n{details}", parent=self)
            return
        except OSError as e:
            messagebox.showwarning("無法儲存", f"寫入設定檔失敗：{e}", parent=self)
            return
        messagebox.showinfo("已儲存", "書卷別名已儲存並立即生效", parent=self)
        self.destroy()
