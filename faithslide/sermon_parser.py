"""週報內容解析：讀經段落、證道大綱結構。

證道大綱格式（皆為粗體）：
    主標題 [引用...]
    一、大標題 [引用...]
    1. 中標題 [引用...]
    1) 小標題 [引用...]
"""
import logging
import re
from dataclasses import dataclass, field

from faithslide import bible_books
from faithslide.bible_api import Verse
from faithslide.bible_books import Book
from faithslide.reference_parser import find_references, find_unrecognized, text_before_references

logger = logging.getLogger(__name__)

MAJOR, MEDIUM, MINOR = 1, 2, 3

_CHINESE_DIGITS = "一二三四五六七八九十"
# 在每個標題編號前切開；前面不可緊接編號字元，避免「10.」「十一、」被從中間切開
_HEADING_SPLIT = re.compile(
    rf"(?<![{_CHINESE_DIGITS}\d])(?=[{_CHINESE_DIGITS}]+\s*、|\d+\s*\.|\d+\s*\))"
)
_HEADING_MARKER = re.compile(rf"^(?:(?P<major>[{_CHINESE_DIGITS}]+)、|(?P<medium>\d+)\.|(?P<minor>\d+)\))")
_READING_VERSE = re.compile(r"^\[(\d+)\](.*)$", re.S)
_READING_RANGE = re.compile(r"(\d+)(?:\s*-\s*(\d+))?\s*節")


# --- 證道 --------------------------------------------------------------------

@dataclass
class Section:
    level: int                  # MAJOR / MEDIUM / MINOR
    number: str                 # 「一」「1」
    text: str                   # 編號之後的標題文字
    references: list[str] = field(default_factory=list)   # 未解析的引用字串，書卷可能省略
    children: list["Section"] = field(default_factory=list)


@dataclass
class Sermon:
    title: str
    references: list[str] = field(default_factory=list)   # 主標題後的引用
    sections: list[Section] = field(default_factory=list)  # 大標題
    warnings: list[str] = field(default_factory=list)


def speaker_word_problem(word: str) -> str | None:
    """講員字詞會從證道文字中全部刪除，檢查是否安全；沒有問題回傳 None。"""
    if not word:
        return "不可為空白"
    if len(word) < 2:
        # 單一個字會把標題中所有相同的字刪掉
        return "至少需要兩個字，單一個字會刪掉標題中所有相同的字"
    if re.search(r"[\d:：、.)]", word):
        return "不可包含數字或標題、經文使用的符號"
    return None


def parse_sermon(text: str, speaker_words: list[str]) -> Sermon:
    for word in speaker_words:
        text = text.replace(word, "")
    parts = _HEADING_SPLIT.split(text)

    sermon = Sermon(title=_heading_text(parts[0]), references=find_references(parts[0]))
    _check_unrecognized(sermon, sermon.title or "主標題", parts[0])
    for part in parts[1:]:
        heading = _heading_text(part)
        _check_unrecognized(sermon, heading, part)
        match = _HEADING_MARKER.match(heading)
        if not match:
            _warn(sermon, f"無法辨識的標題「{heading}」，已略過")
            continue
        level = MAJOR if match["major"] else MEDIUM if match["medium"] else MINOR
        section = Section(level, match[match.lastgroup], heading[match.end():], find_references(part))

        if level == MAJOR:
            sermon.sections.append(section)
        elif level == MEDIUM:
            if not sermon.sections:
                _warn(sermon, f"中標題「{heading}」出現在大標題之前，已略過")
                continue
            sermon.sections[-1].children.append(section)
        else:
            if not sermon.sections or not sermon.sections[-1].children:
                _warn(sermon, f"小標題「{heading}」出現在中標題之前，已略過")
                continue
            sermon.sections[-1].children[-1].children.append(section)
    return sermon


def _check_unrecognized(sermon: Sermon, heading: str, part: str) -> None:
    for fragment, reason in find_unrecognized(part):
        _warn(sermon, f"「{heading}」中的「{fragment}」：{reason}")


def _heading_text(part: str) -> str:
    """標題文字：第一個引用之前的部分，去除 Word run 串接產生的空白。"""
    before = text_before_references(part)
    return (part if before is None else before).replace(" ", "").strip()


def _warn(target, message: str) -> None:
    logger.warning(message)
    target.warnings.append(message)


# --- 讀經 --------------------------------------------------------------------

@dataclass
class Passage:
    title: str                  # 例：士師記十七章
    verses: list[Verse] = field(default_factory=list)


@dataclass
class Reading:
    passages: list[Passage] = field(default_factory=list)
    main_book: Book | None = None    # 證道引用省略書卷時使用
    warnings: list[str] = field(default_factory=list)


def parse_reading(lines: list[str]) -> Reading:
    """支援兩種格式：

    A. 標題行（士師記十七章）後接經文行（[1]…）；可有多個標題行。
    B. 第一行以「，」列出多段（以弗所書二章10節，一章17-19節），經文行依各段節數依序分配。
    """
    reading = Reading()
    lines = [line for line in lines if line.strip()]
    if not lines:
        _warn(reading, "找不到讀經內容")
        return reading

    if "，" in lines[0]:
        _parse_listed_passages(reading, lines[0], lines[1:])
    else:
        _parse_titled_passages(reading, lines)

    for passage in reading.passages:
        reading.main_book = find_book_in(passage.title)
        if reading.main_book:
            break
    else:
        _warn(reading, f"無法從讀經判斷主經文書卷：{lines[0]}")
    return reading


def _parse_titled_passages(reading: Reading, lines: list[str]) -> None:
    passage = Passage(lines[0])
    reading.passages.append(passage)
    for line in lines[1:]:
        verse = _parse_reading_verse(line)
        if verse:
            passage.verses.append(verse)
        else:
            passage = Passage(line.strip())
            reading.passages.append(passage)


def _parse_listed_passages(reading: Reading, header: str, lines: list[str]) -> None:
    parts = [p.strip() for p in header.split("，") if p.strip()]
    first_book = find_book_in(parts[0])
    titles = [parts[0]] + [
        p if find_book_in(p) or not first_book else first_book.full + p
        for p in parts[1:]
    ]

    remaining = iter(lines)
    for title in titles:
        passage = Passage(title)
        reading.passages.append(passage)
        match = _READING_RANGE.search(title)
        if match:
            start = int(match[1])
            count = (int(match[2]) if match[2] else start) - start + 1
        else:
            _warn(reading, f"讀經「{title}」沒有節數，只取一行經文")
            count = 1
        for _ in range(count):
            line = next(remaining, None)
            if line is None:
                _warn(reading, f"讀經「{title}」的經文行數不足")
                return
            verse = _parse_reading_verse(line)
            if verse:
                passage.verses.append(verse)
            else:
                _warn(reading, f"讀經經文格式錯誤，已略過：{line[:20]}")


def _parse_reading_verse(line: str) -> Verse | None:
    match = _READING_VERSE.match(line.strip("\n"))
    return Verse(int(match[1]), match[2]) if match else None


def find_book_in(text: str) -> Book | None:
    """找出文字中出現的書卷全名（最早出現者優先，同位置取較長者）。"""
    found = [(text.find(b.full), -len(b.full), i, b) for i, b in enumerate(bible_books.all_books()) if b.full in text]
    return min(found)[3] if found else None
