import json
import re

import pytest

from faithslide import bible_books
from faithslide.bible_books import Book, BookDataError, find_conflicts

# 重構前 FaithSlide.py 的 books 字串，name_pattern 必須涵蓋完全相同的名稱
LEGACY_BOOKS = "創|出|利|民|申|書|士|得|撒上|撒下|王上|王下|代上|代下|拉|尼|斯|伯|詩|箴|傳|歌|賽|耶|哀|結|但|何|珥|摩|俄|拿|彌|鴻|哈|番|該|瑪|亞|太|可|路|約|徒|羅|林前|林後|加|弗|腓|西|帖前|帖後|提前|提後|多|門|來|雅|各|彼前|彼後|約壹|約貳|約參|猶|啟|約一|約二|約三"


@pytest.fixture(autouse=True)
def fresh_data():
    bible_books.load()


def test_all_books_in_canonical_order():
    books = bible_books.all_books()
    assert len(books) == 66
    assert [books[0].full, books[38].full, books[39].full, books[-1].full] == ["創世記", "瑪拉基書", "馬太福音", "啟示錄"]


@pytest.mark.parametrize("name, book_id", [
    ("約翰福音", "jhn"),
    ("約", "jhn"),
    ("約壹", "1jn"),
    ("約一", "1jn"),
    ("約翰一書", "1jn"),
    ("各", "jas"),
    (" 羅 ", "rom"),
])
def test_resolve(name, book_id):
    assert bible_books.resolve(name).id == book_id


def test_resolve_unknown():
    assert bible_books.resolve("約X") is None
    assert bible_books.resolve("") is None


@pytest.mark.parametrize("name, api", [("各", "雅"), ("約壹", "約一"), ("約貳", "約二"), ("約參", "約三")])
def test_alias_maps_to_api_value(name, api):
    # 取代舊版 get_verses 裡寫死的轉換
    assert bible_books.resolve(name).api == api


def test_name_pattern_matches_legacy_names():
    names = set(bible_books.name_pattern().split("|"))
    assert names == set(LEGACY_BOOKS.split("|"))


def test_name_pattern_prefers_longer_names():
    pattern = bible_books.name_pattern()
    assert re.match(pattern, "約壹3:1").group() == "約壹"
    assert re.match(pattern, "林前13:4").group() == "林前"
    assert re.match(pattern, "約3:16").group() == "約"


def test_find_conflicts():
    books = [
        Book("jhn", "約", "約翰福音", "約"),
        Book("1jn", "約一", "約翰一書", "約一", ("約",)),
    ]
    conflicts = find_conflicts(books)
    assert [(c.name, c.book_ids) for c in conflicts] == [("約", ("jhn", "1jn"))]


def test_load_rejects_conflicting_default_data(tmp_path):
    path = tmp_path / "books.json"
    path.write_text(json.dumps({"schema_version": 1, "books": [
        {"id": "jhn", "api": "約", "full": "約翰福音", "abbr": "約"},
        {"id": "1jn", "api": "約一", "full": "約翰一書", "abbr": "約一", "aliases": ["約"]},
    ]}, ensure_ascii=False), encoding="utf-8")
    with pytest.raises(BookDataError):
        bible_books.load(str(path))


def test_load_rejects_malformed_data(tmp_path):
    path = tmp_path / "books.json"
    path.write_text(json.dumps({"books": [{"id": "gen"}]}), encoding="utf-8")
    with pytest.raises(BookDataError):
        bible_books.load(str(path))
