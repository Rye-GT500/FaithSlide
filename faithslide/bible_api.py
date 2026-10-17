"""經文查詢：信望愛站 fhl API（和合本）。"""
import logging
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from functools import lru_cache
from typing import Callable

import requests

from faithslide.bible_books import Book

API_URL = "https://bible.fhl.net/json/qb.php"
TIMEOUT_SECONDS = 5
MAX_CONCURRENT_REQUESTS = 4   # 同時查詢的章數上限，避免對信望愛站造成負擔

logger = logging.getLogger(__name__)


@dataclass(frozen=True)
class Verse:
    number: int
    text: str


class BibleApiError(Exception):
    """訊息可直接顯示給使用者。"""


def get_chapter(book: Book, chapter: int) -> list[Verse]:
    """取得整章經文，依節數排序。同一章在程式執行期間只查詢一次。"""
    return list(_fetch_chapter(book.api, book.full, chapter))


def get_chapters(chapters: list[tuple[Book, int]],
                 on_each: Callable[[int, int, Book, int], None] = lambda done, total, book, chapter: None,
                 ) -> dict[tuple[str, int], list[Verse] | BibleApiError]:
    """同時查詢多章，回傳 {(書卷 id, 章): 經文或查詢錯誤}。

    每查完一章呼叫 on_each(已完成, 總數, 書卷, 章)；回呼在呼叫端的執行緒執行。
    """
    unique = list(dict.fromkeys(chapters))
    results = {}
    if not unique:
        return results
    with ThreadPoolExecutor(max_workers=min(MAX_CONCURRENT_REQUESTS, len(unique))) as pool:
        futures = {pool.submit(get_chapter, book, chapter): (book, chapter) for book, chapter in unique}
        for done, future in enumerate(as_completed(futures), start=1):
            book, chapter = futures[future]
            try:
                results[(book.id, chapter)] = future.result()
            except BibleApiError as e:
                results[(book.id, chapter)] = e
            on_each(done, len(unique), book, chapter)
    return results


def clear_cache() -> None:
    _fetch_chapter.cache_clear()


@lru_cache(maxsize=256)
def _fetch_chapter(api_book: str, full_name: str, chapter: int) -> tuple[Verse, ...]:
    # 錯誤會以例外拋出，lru_cache 不會快取失敗結果
    params = {
        "chineses": api_book,   # 書卷
        "chap": chapter,        # 章
        "version": "unv",       # 和合本
        "strong": "0",          # 不要 Strong Number
    }
    try:
        response = requests.get(API_URL, params=params, timeout=TIMEOUT_SECONDS)
    except requests.RequestException as e:
        logger.error(f"經文查詢連線失敗 {full_name} {chapter}: {e}")
        raise BibleApiError("無法連線到經文網站，請檢查網路") from e

    if response.status_code != 200:
        logger.error(f"經文查詢 HTTP {response.status_code} {full_name} {chapter}")
        raise BibleApiError(f"經文網站回應錯誤（HTTP {response.status_code}）")

    try:
        data = response.json()
        if data.get("status") != "success":
            raise BibleApiError(f"經文網站回應錯誤：{data.get('status')}")
        verses = tuple(sorted(
            (Verse(int(r["sec"]), r["bible_text"]) for r in data["record"]),
            key=lambda v: v.number,
        ))
    except (ValueError, KeyError, TypeError) as e:
        logger.error(f"經文查詢回應格式錯誤 {full_name} {chapter}: {e}")
        raise BibleApiError("經文網站回應格式錯誤") from e

    if not verses:
        raise BibleApiError(f"查無經文：{full_name} 第 {chapter} 章")
    return verses
