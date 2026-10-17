"""以現行 FaithSlide.py 製作投影片，供建立基準與比對使用。

- get_verses 改為讀寫 tests/fixtures/api/verses.json，離線模式下不連網。
- messagebox 改為收集訊息，不跳出視窗。
"""
import json
from pathlib import Path
from tkinter import messagebox
from unittest import mock

# FaithSlide 匯入時會在背景檢查更新，若有新版會跳出 askyesno；測試期間一律回答「否」
messagebox.askyesno = lambda *args, **kwargs: False

import FaithSlide as legacy  # noqa: E402

API_CACHE_PATH = Path(__file__).parent / "fixtures" / "api" / "verses.json"


class VerseCache:
    def __init__(self, allow_network: bool):
        self.allow_network = allow_network
        self.data = json.loads(API_CACHE_PATH.read_text(encoding="utf-8")) if API_CACHE_PATH.exists() else {}
        self.missing = []   # 離線時快取中沒有的 (書卷, 章)
        self.empty = []     # 連網查詢卻回傳空結果的 (書卷, 章)
        self.real_get_verses = legacy.get_verses

    def get_verses(self, book_abbr, chapter):
        key = f"{book_abbr}|{chapter}"
        if key in self.data:
            return list(self.data[key])
        if not self.allow_network:
            self.missing.append(key)
            return []
        verses = self.real_get_verses(book_abbr, chapter)
        if not verses:
            self.empty.append(key)
        self.data[key] = verses
        return list(verses)

    def save(self):
        API_CACHE_PATH.parent.mkdir(parents=True, exist_ok=True)
        API_CACHE_PATH.write_text(json.dumps(self.data, ensure_ascii=False, indent=1, sort_keys=True), encoding="utf-8")


def run(docx_path, out_pptx, cache: VerseCache) -> list:
    """製作投影片並回傳過程中所有 messagebox 訊息。"""
    messages = []

    def record(kind):
        return lambda title="", message="", **_: messages.append([kind, title, message])

    # 模擬全新啟動：清除上一次執行殘留的全域狀態
    legacy.main_book = ""
    legacy.word_path_var.set(str(docx_path))
    legacy.ppt_save_var.set(str(out_pptx))

    with mock.patch.object(legacy, "get_verses", cache.get_verses), \
         mock.patch.object(legacy.messagebox, "showwarning", record("warning")), \
         mock.patch.object(legacy.messagebox, "showinfo", record("info")):
        legacy.produce_the_slides()
    return messages
