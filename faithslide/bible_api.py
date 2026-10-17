"""經文查詢：信望愛站 fhl API（和合本）。"""
import logging
from dataclasses import dataclass
from functools import lru_cache

import requests

from faithslide.bible_books import Book

API_URL = "https://bible.fhl.net/json/qb.php"
TIMEOUT_SECONDS = 5

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
