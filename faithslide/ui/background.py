"""在背景執行緒執行工作，進度與結果透過 queue 交回主執行緒。

Tk 元件只能在主執行緒操作；背景執行緒只放入 queue，由主執行緒以 after 定時取出處理。
"""
import queue
from threading import Thread
from typing import Any, Callable

POLL_MS = 50


def run_in_background(widget,
                      work: Callable[[Callable[..., None]], Any],
                      on_done: Callable[[Any], None],
                      on_error: Callable[[Exception], None],
                      on_progress: Callable[..., None] = lambda *args: None) -> None:
    """work 會收到一個 report_progress(*args) 函式；回呼皆在主執行緒執行。"""
    events = queue.Queue()

    def worker():
        try:
            result = work(lambda *args: events.put(("progress", args)))
        except Exception as e:
            events.put(("error", e))
        else:
            events.put(("done", result))

    def poll():
        try:
            while True:
                kind, payload = events.get_nowait()
                if kind == "progress":
                    on_progress(*payload)
                elif kind == "done":
                    on_done(payload)
                    return
                else:
                    on_error(payload)
                    return
        except queue.Empty:
            widget.after(POLL_MS, poll)

    Thread(target=worker, daemon=True).start()
    widget.after(POLL_MS, poll)
