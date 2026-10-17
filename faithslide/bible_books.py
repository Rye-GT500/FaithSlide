"""書卷資料：載入 books.default.json，提供名稱解析與書名 pattern。

名稱分兩類：
- 全名（full）：只用於顯示與 resolve。
- 簡稱（abbr + aliases）：resolve 之外，也用於從文字中辨識經文引用（name_pattern）。
"""
import json
import os
import re
from dataclasses import dataclass

from faithslide import paths

DEFAULT_BOOKS_PATH = os.path.join(paths.DATA_DIR, "books.default.json")


@dataclass(frozen=True)
class Book:
    id: str
    api: str        # 送給 fhl API 的書卷值
    full: str
    abbr: str
    aliases: tuple[str, ...] = ()

    @property
    def short_names(self) -> tuple[str, ...]:
        return (self.abbr, *self.aliases)

    @property
    def names(self) -> tuple[str, ...]:
        return (self.full, *self.short_names)


@dataclass(frozen=True)
class AliasConflict:
    name: str
    book_ids: tuple[str, ...]


class BookDataError(ValueError):
    pass


_books: list[Book] | None = None
_index: dict[str, Book] = {}


def find_conflicts(books: list[Book]) -> list[AliasConflict]:
    """找出對應到一卷以上書卷的名稱。"""
    owners: dict[str, list[str]] = {}
    for book in books:
        for name in dict.fromkeys(book.names):   # 同一卷書內重複的名稱不算衝突
            owners.setdefault(name, []).append(book.id)
    return [AliasConflict(name, tuple(ids)) for name, ids in owners.items() if len(ids) > 1]


def _parse_books(raw: dict) -> list[Book]:
    try:
        books = [Book(b["id"], b["api"], b["full"], b["abbr"], tuple(b.get("aliases", []))) for b in raw["books"]]
    except (KeyError, TypeError) as e:
        raise BookDataError(f"書卷資料格式錯誤：{e}") from e
    ids = [b.id for b in books]
    if len(ids) != len(set(ids)):
        raise BookDataError("書卷 id 重複")
    return books


def load(default_path: str = DEFAULT_BOOKS_PATH) -> None:
    global _books, _index
    with open(default_path, encoding="utf-8") as f:
        books = _parse_books(json.load(f))
    conflicts = find_conflicts(books)
    if conflicts:
        # 內建資料出現衝突屬於程式錯誤，直接中止
        raise BookDataError(f"內建書卷名稱衝突：{conflicts}")
    _books = books
    _index = {name: book for book in books for name in book.names}


def reload() -> None:
    load()


def _ensure_loaded() -> None:
    if _books is None:
        load()


def all_books() -> list[Book]:
    """依聖經書卷順序回傳。"""
    _ensure_loaded()
    return list(_books)


def resolve(name: str) -> Book | None:
    """接受全名、簡寫或別名。"""
    _ensure_loaded()
    return _index.get(name.strip())


def name_pattern() -> str:
    """所有簡稱組成的正則（不含全名、不含群組）。長的在前，避免「約壹」被「約」搶先比對。"""
    _ensure_loaded()
    names = {name for book in _books for name in book.short_names}
    return "|".join(re.escape(n) for n in sorted(names, key=lambda n: (-len(n), n)))
