"""經文引用：從文字中尋找、解析成結構、轉成顯示用標題。

引用格式：[書卷簡稱]章:節[-節][,節[-節]]...[,章:節...]
- 書卷可省略，由呼叫端提供 default_book（例如證道的主經文書卷）。
- 不支援跨章範圍（如 約3:16-4:2）。
"""
import re
from dataclasses import dataclass
from functools import lru_cache

from faithslide import bible_books
from faithslide.bible_books import Book

_CJK = r"一-鿿"
_DIGITS = ["零", "一", "二", "三", "四", "五", "六", "七", "八", "九"]
# 搜尋輸入常見的全形符號
_NORMALIZE = str.maketrans({"：": ":", "，": ",", "－": "-", "–": "-", "—": "-", "~": "-", "～": "-"})


@dataclass(frozen=True)
class Reference:
    book: Book
    chapter: int
    ranges: tuple[tuple[int, int], ...]   # 空 tuple = 整章


class ReferenceParseError(ValueError):
    """訊息可直接顯示給使用者。"""


_VERSES = r"\d+(?:\s*-\s*\d+)?"
# 章:節[-節]，之後可接「,節[-節]」或「,章:節[-節]」
_CHAPTER_VERSES = rf"\d+\s*:\s*{_VERSES}(?:\s*,\s*(?:\d+\s*:\s*)?{_VERSES})*"
# 同上，但容許數字中間夾空白：Word 的 run 常把「10」切成「1 0」，串接後中間多出空白。
# 代價是「3:16 4:1」這種只以空白分隔、未寫書卷的兩組章節會被合併。
_CHAPTER_VERSES_LOOSE = (
    r"\d+\s*:(?:(?:\s*\d+\s*)(?:-\s*(?:\s*\d+\s*))?)+"
    r"(?:\s*,\s*(?:\d+\s*)*(?::\s*\d+)*(?:-\s*(?:\d+\s*)*)?)*"
)


@lru_cache(maxsize=None)
def _token_regex(books: str) -> re.Pattern:
    # 書卷簡稱（前後不可緊貼中文字，避免標題文字被誤判）或「章:節」序列
    return re.compile(rf"(?<![{_CJK}])(?:{books})+(?![{_CJK}])|{_CHAPTER_VERSES_LOOSE}")


@lru_cache(maxsize=None)
def _first_reference_regex(books: str) -> re.Pattern:
    # 書卷緊接章節，或單獨的章節
    return re.compile(rf"(?<![{_CJK}])(?:{books})+\s*{_CHAPTER_VERSES}|{_CHAPTER_VERSES}")


def _collapse_spaces(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def find_references(text: str) -> list[str]:
    """從一段文字找出所有引用字串，保留週報原文的寫法（連續空白縮成一個），供錯誤訊息引用。

    書卷與章節之間可以隔著其他文字；章節前沒有書卷時，回傳的字串不含書卷。
    書卷只套用到它後面的第一組章節。parse 會去除空白，所以保留空白不影響解析。
    """
    refs = []
    book = None
    for match in _token_regex(bible_books.name_pattern()).finditer(text):
        token = _collapse_spaces(match.group())
        if token[0].isdigit():
            if book:
                gap = "" if book.end() == match.start() else " "
                refs.append(_collapse_spaces(book.group()) + gap + token)
            else:
                refs.append(token)
            book = None
        else:
            book = match
    return refs


def find_unrecognized(text: str) -> list[tuple[str, str]]:
    """找出看起來像經文、卻沒有被辨識的片段，回傳 [(片段, 原因)]。"""
    # 把已辨識的部分換成分隔符號，避免前後文字接在一起
    leftover = _token_regex(bible_books.name_pattern()).sub(" | ", text)
    problems = []
    for match in re.finditer(r"\d+\s*：\s*\d+", leftover):
        problems.append((_collapse_spaces(match.group()), "章節之間請使用半形冒號「:」"))
    for match in re.finditer(r"(?<![\d：])\s*:\s*\d+", leftover):
        problems.append((_collapse_spaces(match.group()), "無法辨識，可能是兩段經文之間少了逗號"))
    return problems


def has_space_between_digits(raw: str) -> bool:
    """原文的數字之間是否夾著空白（可能是 Word 切開的數字，也可能是兩段經文少了逗號）。"""
    return re.search(r"\d\s+\d", raw) is not None


def text_before_references(text: str) -> str | None:
    """回傳第一個引用之前的文字；沒有引用時回傳 None。"""
    match = _first_reference_regex(bible_books.name_pattern()).search(text)
    return text[:match.start()] if match else None


def parse(raw: str, default_book: Book | None = None) -> list[Reference]:
    """解析一個引用字串，每章回傳一個 Reference。"""
    text = raw.translate(_NORMALIZE).replace(" ", "").replace("　", "")
    match = re.match(r"(\D*)(.*)", text)
    book_name, body = match.group(1), match.group(2)

    if book_name:
        book = bible_books.resolve(book_name)
        if book is None:
            raise ReferenceParseError(f"找不到書卷「{book_name}」")
    elif default_book is not None:
        book = default_book
    else:
        raise ReferenceParseError(f"「{raw}」缺少書卷")

    if not body:
        raise ReferenceParseError(f"「{raw}」缺少章數")
    if ":" not in body:
        return [Reference(book, _to_int(body, raw), ())]

    refs = []
    chapter = None
    ranges = []
    for item in body.split(","):
        if ":" in item:
            if chapter is not None:
                refs.append(Reference(book, chapter, tuple(ranges)))
            chapter_text, item = item.split(":", 1)
            chapter = _to_int(chapter_text, raw)
            ranges = []
            if ":" in item:
                raise ReferenceParseError(f"「{raw}」不支援跨章引用")
        elif chapter is None:
            raise ReferenceParseError(f"「{raw}」缺少章數")
        if item:
            ranges.append(_parse_range(item, raw))
    refs.append(Reference(book, chapter, tuple(ranges)))
    return refs


def _parse_range(item: str, raw: str) -> tuple[int, int]:
    if "-" in item:
        start_text, end_text = item.split("-", 1)
        start, end = _to_int(start_text, raw), _to_int(end_text, raw)
    else:
        start = end = _to_int(item, raw)
    if start < 1 or end < start:
        raise ReferenceParseError(f"「{raw}」節數範圍錯誤")
    return start, end


def _to_int(text: str, raw: str) -> int:
    if not text.isdigit():
        raise ReferenceParseError(f"「{raw}」格式錯誤")
    return int(text)


def chinese_numeral(n: int) -> str:
    """1～999 的中文數字，如 10 → 十、101 → 一百零一、150 → 一百五十。"""
    if not 0 < n < 1000:
        raise ValueError(f"超出範圍：{n}")
    hundreds, tens, ones = n // 100, n // 10 % 10, n % 10
    result = ""
    if hundreds:
        result += _DIGITS[hundreds] + "百"
    if tens:
        result += ("" if tens == 1 and not hundreds else _DIGITS[tens]) + "十"
    elif hundreds and ones:
        result += "零"
    if ones:
        result += _DIGITS[ones]
    return result


def format_title(ref: Reference, verse_range: tuple[int, int] | None = None) -> str:
    """例：約翰福音三章16-18節；詩篇以「篇」取代「章」；不給 verse_range 時只到章。"""
    unit = "篇" if ref.book.id == "psa" else "章"
    title = f"{ref.book.full}{chinese_numeral(ref.chapter)}{unit}"
    if verse_range is not None:
        start, end = verse_range
        title += f"{start}節" if start == end else f"{start}-{end}節"
    return title
