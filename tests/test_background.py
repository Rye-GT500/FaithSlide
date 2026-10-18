import threading
import time
import tkinter as tk

import pytest

from faithslide.background import run_in_background


# 整個檔案共用一個 Tk：Windows 上關閉 Tk 後再建立新的，偶爾會讀不到 init.tcl / tk.tcl
@pytest.fixture(scope="module")
def root():
    root = tk.Tk()
    root.withdraw()
    yield root
    root.destroy()


def pump(root, until, timeout=5):
    end = time.time() + timeout
    while not until() and time.time() < end:
        root.update()
        time.sleep(0.01)


def test_progress_and_result_delivered_on_main_thread(root):
    main_thread = threading.get_ident()
    events = []

    def work(report_progress):
        report_progress(0.5, "一半")
        return "完成"

    run_in_background(
        root, work,
        on_done=lambda result: events.append(("done", result, threading.get_ident())),
        on_error=lambda error: events.append(("error", error)),
        on_progress=lambda p, m: events.append(("progress", p, m, threading.get_ident())),
    )
    pump(root, lambda: any(e[0] == "done" for e in events))
    assert events == [("progress", 0.5, "一半", main_thread), ("done", "完成", main_thread)]


def test_error_delivered(root):
    errors = []

    def work(report_progress):
        raise ValueError("壞掉")

    run_in_background(root, work, on_done=lambda r: None, on_error=errors.append)
    pump(root, lambda: errors)
    assert isinstance(errors[0], ValueError)
