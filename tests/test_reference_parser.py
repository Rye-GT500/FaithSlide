import pytest

from faithslide import bible_books
from faithslide.reference_parser import (
    Reference, ReferenceParseError, chinese_numeral, find_references, format_title, parse, text_before_references,
)


def book(name):
    return bible_books.resolve(name)


# --- find_references ---------------------------------------------------------

@pytest.mark.parametrize("text, expected", [
    # README 範例格式
    ("一、這是第一個大標題： 創 1:4", ["創1:4"]),
    ("1. 這是第一個中標題： 創 1:5-6 箴 5:3", ["創1:5-6", "箴5:3"]),
    ("1) 這是第一個小標題： 創 1:7-8, 10", ["創1:7-8,10"]),
    ("2. 這是第二個中標題 創 1:11, 13, 15-17", ["創1:11,13,15-17"]),
    # 沒寫書卷
    ("標題 1:3", ["1:3"]),
    # 書卷只套用到後面第一組章節
    ("標題 約 3:16, 羅 4:1", ["約3:16,", "羅4:1"]),
    # Word run 切開的數字
    ("標題 士 18:3 1, 約 1:14", ["士18:31,", "約1:14"]),
    ("標題 18:1 , 9- 1 0", ["18:1,9-10"]),
    # 長名稱
    ("標題 約壹 3:1 林前 13:4", ["約壹3:1", "林前13:4"]),
    # 跨章節的逗號
    ("標題 約 3:16,18,4:2", ["約3:16,18,4:2"]),
    # 範圍符號前後可有空白；尾端多餘的符號不納入
    ("標題 創 1:5 - 6", ["創1:5-6"]),
    ("標題 創 1:5-6 ", ["創1:5-6"]),
    # 標題中緊貼中文字的書卷字不算
    ("一、約翰的見證 約 1:4", ["約1:4"]),
    ("沒有引用的標題", []),
])
def test_find_references(text, expected):
    assert find_references(text) == expected


def test_text_before_references():
    assert text_before_references("1. 這是中標題： 創 1:5-6") == "1. 這是中標題： "
    assert text_before_references("1. 這是中標題： 1:5-6") == "1. 這是中標題： "
    assert text_before_references("沒有引用的標題") is None


# --- parse -------------------------------------------------------------------

@pytest.mark.parametrize("raw, expected", [
    ("約3:16", [("jhn", 3, ((16, 16),))]),
    ("約3:16-18", [("jhn", 3, ((16, 18),))]),
    ("約3:16-18,20", [("jhn", 3, ((16, 18), (20, 20)))]),
    ("創1:11,13,15-17", [("gen", 1, ((11, 11), (13, 13), (15, 17)))]),
    # find_references 可能留下結尾逗號
    ("士18:24,", [("jdg", 18, ((24, 24),))]),
    ("約3:16,18,4:2", [("jhn", 3, ((16, 16), (18, 18))), ("jhn", 4, ((2, 2),))]),
    ("林前13:4", [("1co", 13, ((4, 4),))]),
    ("約壹3:1", [("1jn", 3, ((1, 1),))]),
    ("各1:2", [("jas", 1, ((2, 2),))]),
    # 搜尋頁輸入：全名、空白、全形符號、整章
    ("約翰福音 3:16", [("jhn", 3, ((16, 16),))]),
    ("約３：１６－１８", [("jhn", 3, ((16, 18),))]),
    ("約 3:16 ，18", [("jhn", 3, ((16, 16), (18, 18)))]),
    ("詩23", [("psa", 23, ())]),
])
def test_parse(raw, expected):
    assert [(r.book.id, r.chapter, r.ranges) for r in parse(raw)] == expected


def test_parse_uses_default_book():
    refs = parse("3:16", default_book=book("羅"))
    assert refs == [Reference(book("羅"), 3, ((16, 16),))]


def test_explicit_book_overrides_default():
    assert parse("約3:16", default_book=book("羅"))[0].book.id == "jhn"


@pytest.mark.parametrize("raw, message", [
    ("3:16", "缺少書卷"),
    ("約X3:16", "找不到書卷「約X」"),
    ("約", "缺少章數"),
    ("約3:16-4:2", "不支援跨章"),
    ("約3:18-16", "節數範圍錯誤"),
    ("約3:0", "節數範圍錯誤"),
    ("約3:16-", "格式錯誤"),
    ("", "缺少書卷"),
])
def test_parse_errors(raw, message):
    with pytest.raises(ReferenceParseError, match=message):
        parse(raw)


# --- 顯示 --------------------------------------------------------------------

@pytest.mark.parametrize("n, expected", [
    (1, "一"), (9, "九"), (10, "十"), (11, "十一"), (20, "二十"), (21, "二十一"), (99, "九十九"),
    (100, "一百"), (101, "一百零一"), (110, "一百一十"), (119, "一百一十九"), (150, "一百五十"),
])
def test_chinese_numeral(n, expected):
    assert chinese_numeral(n) == expected


@pytest.mark.parametrize("n", [0, 1000])
def test_chinese_numeral_out_of_range(n):
    with pytest.raises(ValueError):
        chinese_numeral(n)


def test_format_title():
    ref = Reference(book("約"), 3, ((16, 18), (20, 20)))
    assert format_title(ref, (16, 18)) == "約翰福音三章16-18節"
    assert format_title(ref, (20, 20)) == "約翰福音三章20節"
    assert format_title(ref) == "約翰福音三章"


def test_format_title_psalms_use_pian():
    assert format_title(Reference(book("詩"), 119, ((1, 1),)), (1, 1)) == "詩篇一百一十九篇1節"
