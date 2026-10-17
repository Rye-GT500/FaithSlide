"""滑鼠停留時顯示的提示文字。"""
import tkinter as tk


class Tooltip:
    def __init__(self, widget, text: str):
        self.widget = widget
        self.text = text
        self.window = None
        widget.bind("<Enter>", self.show, add="+")
        widget.bind("<Leave>", self.hide, add="+")

    def show(self, event=None):
        if self.window:
            return
        x = self.widget.winfo_rootx() + self.widget.winfo_width() // 2
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 4
        self.window = tk.Toplevel(self.widget)
        self.window.wm_overrideredirect(True)
        tk.Label(self.window, text=self.text, background="#ffffe0", relief="solid", borderwidth=1, padx=6, pady=2).pack()
        self.window.update_idletasks()
        # 靠右對齊，避免超出主視窗右緣
        self.window.wm_geometry(f"+{x - self.window.winfo_width() + 10}+{y}")

    def hide(self, event=None):
        if self.window:
            self.window.destroy()
            self.window = None
