import pytest

from faithslide import bible_api
from faithslide.bible_api import BibleApiError, Verse
from faithslide.reference_parser import ReferenceParseError
from faithslide.ui.search_page import lookup


@pytest.fixture(autouse=True)
def chapters(monkeypatch):
    """假經文：每章 3 節；羅馬書查詢失敗。"""
    def get_chapter(book, chapter):
        if book.id == "rom":
            raise BibleApiError("查無經文")
        return [Verse(n, f"{book.abbr}{chapter}:{n}") for n in range(1, 4)]
    monkeypatch.setattr(bible_api, "get_chapter", get_chapter)


def test_range():
    lines, problems = lookup("約3:1-2")
    assert lines == ["【約翰福音三章1-2節】", "1. 約3:1", "2. 約3:2", ""]
    assert problems == []


def test_full_name_and_multiple_ranges():
    lines, _ = lookup("約翰福音 3:1,3")
    assert lines == ["【約翰福音三章1節】", "1. 約3:1", "", "【約翰福音三章3節】", "3. 約3:3", ""]


def test_whole_chapter():
    lines, _ = lookup("詩23")
    assert lines == ["【詩篇二十三篇】", "1. 詩23:1", "2. 詩23:2", "3. 詩23:3", ""]


def test_alias_and_multiple_segments():
    lines, _ = lookup("各1:1；約壹2:2")
    assert lines == ["【雅各書一章1節】", "1. 雅1:1", "", "【約翰一書二章2節】", "2. 約一2:2", ""]


def test_missing_verses_reported():
    lines, problems = lookup("約3:2-5")
    assert lines == ["【約翰福音三章2-5節】", "2. 約3:2", "3. 約3:3", ""]
    assert problems == ["約翰福音三章2-5節：本章沒有第 4 節"]


def test_range_entirely_missing_shows_no_heading():
    lines, problems = lookup("約3:1,99")
    assert lines == ["【約翰福音三章1節】", "1. 約3:1", ""]
    assert problems == ["約翰福音三章99節：本章沒有第 99 節"]


def test_api_error_reported_other_segments_still_shown():
    lines, problems = lookup("羅1:1;約3:1")
    assert lines == ["【約翰福音三章1節】", "1. 約3:1", ""]
    assert problems == ["羅馬書一章：查無經文"]


@pytest.mark.parametrize("query, message", [
    ("", "請輸入"),
    (" ；", "請輸入"),
    ("約X3:1", "找不到書卷"),
    ("3:16", "缺少書卷"),
    ("約3:16-4:2", "跨章"),
])
def test_invalid_input(query, message):
    with pytest.raises(ReferenceParseError, match=message):
        lookup(query)
