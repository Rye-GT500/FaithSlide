"""以 tests/fixtures/api/verses.json 取代 bible_api.get_chapter，離線測試用。"""
import json
from pathlib import Path

from faithslide import bible_books
from faithslide.bible_api import BibleApiError, Verse

API_CACHE_PATH = Path(__file__).parent / "fixtures" / "api" / "verses.json"


class FixtureBible:
    def __init__(self):
        # 錄製時以舊版書卷簡稱為鍵：「約壹|2」→ 每節「1. 經文」
        self.chapters = {}
        for key, lines in json.loads(API_CACHE_PATH.read_text(encoding="utf-8")).items():
            abbr, chapter = key.split("|")
            verses = [Verse(int(n), text) for n, text in (line.split(". ", 1) for line in lines)]
            self.chapters[(bible_books.resolve(abbr).id, int(chapter))] = verses
        self.missing = []

    def get_chapter(self, book, chapter):
        key = (book.id, chapter)
        if key not in self.chapters:
            self.missing.append(key)
            raise BibleApiError(f"測試資料缺少 {book.full} 第 {chapter} 章")
        return list(self.chapters[key])
