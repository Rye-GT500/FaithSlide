import pytest

from faithslide import bible_books
from faithslide.bible_api import Verse
from faithslide.sermon_parser import MAJOR, MEDIUM, MINOR, find_book_in, parse_reading, parse_sermon

SPEAKERS = ["證道", "牧師", "王小明"]


def outline(sections):
    """把 Section 樹轉成方便比對的巢狀 tuple。"""
    return [(s.level, s.number, s.text, s.references, outline(s.children)) for s in sections]


# --- 證道 --------------------------------------------------------------------

def test_parse_sermon_full_outline():
    text = (
        "神翻轉生命 證道 王小明 牧師 "
        "一、神翻轉生命： "
        "1. 從死亡到生命： 弗 2:1-3,5-6 "
        "2. 從敗壞到美好： 弗 1:3 箴 5:3 "
        "1) 小標題： 1:7-8, 10 "
        "二、活出新生命： 弗 1:4 "
        "1. 認識耶穌： 弗 1:17-19"
    )
    sermon = parse_sermon(text, SPEAKERS)
    assert sermon.title == "神翻轉生命"
    assert sermon.references == []
    assert sermon.warnings == []
    assert outline(sermon.sections) == [
        (MAJOR, "一", "神翻轉生命：", [], [
            (MEDIUM, "1", "從死亡到生命：", ["弗 2:1-3,5-6"], []),
            (MEDIUM, "2", "從敗壞到美好：", ["弗 1:3", "箴 5:3"], [
                (MINOR, "1", "小標題：", ["1:7-8, 10"], []),
            ]),
        ]),
        (MAJOR, "二", "活出新生命：", ["弗 1:4"], [
            (MEDIUM, "1", "認識耶穌：", ["弗 1:17-19"], []),
        ]),
    ]


def test_heading_spaces_from_word_runs_are_removed():
    sermon = parse_sermon("主標題 一、大 標題 1. 悔 改的果子 ： 太 3:8-10", [])
    major = sermon.sections[0]
    assert major.text == "大標題"
    assert major.children[0].text == "悔改的果子："


def test_two_digit_heading_numbers():
    sermon = parse_sermon("主標題 十一、大標題 10. 中標題 3:1 12) 小標題", [])
    major = sermon.sections[0]
    assert (major.number, major.text) == ("十一", "大標題")
    assert (major.children[0].number, major.children[0].text) == ("10", "中標題")
    assert major.children[0].children[0].number == "12"


def test_main_title_references_kept_separately():
    sermon = parse_sermon("主標題 創 1:1-3 一、大標題", [])
    assert sermon.title == "主標題"
    assert sermon.references == ["創 1:1-3"]


def test_sermon_starting_with_heading_has_empty_title():
    sermon = parse_sermon("一、大標題 創 1:1", [])
    assert sermon.title == ""
    assert sermon.sections[0].references == ["創 1:1"]


def test_orphan_headings_are_skipped_with_warning():
    sermon = parse_sermon("主標題 1. 孤兒中標題 1:1 一、大標題 1) 孤兒小標題 1:2", [])
    assert outline(sermon.sections) == [(MAJOR, "一", "大標題", [], [])]
    assert len(sermon.warnings) == 2


def test_empty_sermon():
    sermon = parse_sermon("", SPEAKERS)
    assert sermon.title == ""
    assert sermon.sections == []


# --- 讀經 --------------------------------------------------------------------

def test_reading_titled_passages():
    reading = parse_reading(["士師記十七章", "[1]第一節", "[2]\t第二節", "士師記十八章", "[1]第一節"])
    assert [(p.title, p.verses) for p in reading.passages] == [
        ("士師記十七章", [Verse(1, "第一節"), Verse(2, "\t第二節")]),
        ("士師記十八章", [Verse(1, "第一節")]),
    ]
    assert reading.main_book.id == "jdg"
    assert reading.warnings == []


def test_reading_listed_passages():
    reading = parse_reading(["以弗所書二章10節，一章17-19節", "[10]a", "[17]b", "[18]c", "[19]d"])
    assert [(p.title, [v.number for v in p.verses]) for p in reading.passages] == [
        ("以弗所書二章10節", [10]),
        ("以弗所書一章17-19節", [17, 18, 19]),
    ]
    assert reading.main_book.id == "eph"


def test_reading_listed_passages_with_other_book():
    reading = parse_reading(["約翰福音一章1節，以弗所書二章3節", "[1]a", "[3]b"])
    assert [p.title for p in reading.passages] == ["約翰福音一章1節", "以弗所書二章3節"]
    assert reading.main_book.id == "jhn"


def test_reading_listed_passages_not_enough_lines():
    reading = parse_reading(["以弗所書二章1-3節，一章17節", "[1]a", "[2]b"])
    assert [len(p.verses) for p in reading.passages] == [2]
    assert reading.warnings


def test_reading_empty():
    reading = parse_reading([""])
    assert reading.passages == []
    assert reading.main_book is None
    assert reading.warnings


@pytest.mark.parametrize("text, book_id", [
    ("約翰一書三章", "1jn"),   # 舊版以中文數字切書名，會把「約翰一書」切成「約翰」
    ("約翰福音十五章1-5節", "jhn"),
    ("詩篇二十三篇", "psa"),
    ("耶利米哀歌三章", "lam"),
])
def test_find_book_in(text, book_id):
    assert find_book_in(text).id == book_id


def test_full_names_are_not_substrings_of_each_other():
    # find_book_in 依賴這個性質；新增書卷資料時若被打破，需要重新檢查
    names = [b.full for b in bible_books.all_books()]
    assert [(a, b) for a in names for b in names if a != b and a in b] == []


@pytest.mark.parametrize("word, ok", [
    ("牧師", True), ("王小明", True), ("Pastor Lee", True),
    ("", False), ("主", False), ("第1堂", False), ("一、", False), ("講員:", False),
])
def test_speaker_word_problem(word, ok):
    from faithslide.sermon_parser import speaker_word_problem
    assert (speaker_word_problem(word) is None) == ok


def test_saved_speaker_words_used_by_pipeline_settings():
    from faithslide import settings
    settings.save("speaker_words", ["王小明"])
    sermon = parse_sermon("主標題 王小明 一、大標題", settings.get("speaker_words"))
    assert sermon.title == "主標題"


def test_unrecognized_fragments_warned_with_heading():
    sermon = parse_sermon("主標題 一、大標題 約 3:16 4:1 1. 中標題 約 3：16", [])
    assert sermon.warnings == [
        "「一、大標題」中的「:1」：無法辨識，可能是兩段經文之間少了逗號",
        # 全形冒號的經文沒被辨識，會整段留在標題裡，警告中顯示實際的標題
        "「1.中標題約3：16」中的「3：16」：章節之間請使用半形冒號「:」",
    ]
