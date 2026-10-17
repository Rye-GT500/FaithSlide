"""週報 Word → 證道投影片 PPT 的完整流程。

單一引用或經文的問題記錄為警告並略過，繼續製作；無法讀取 Word、範本或存檔失敗則拋出 GenerationError。
"""
import logging
from dataclasses import dataclass, field
from typing import Callable

from faithslide import bible_api, paths, settings, word_reader
from faithslide.bible_api import BibleApiError, Verse
from faithslide.bible_books import Book
from faithslide.ppt_builder import SlideBuilder, TemplateError
from faithslide.reference_parser import Reference, ReferenceParseError, format_title, has_space_between_digits, parse
from faithslide.sermon_parser import Section, Sermon, parse_reading, parse_sermon

logger = logging.getLogger(__name__)

ProgressCallback = Callable[[float, str], None]   # (0～1 的進度, 訊息)


class GenerationError(Exception):
    """訊息可直接顯示給使用者。"""


@dataclass
class GenerationResult:
    warnings: list[str] = field(default_factory=list)


# 進度分配：查經文最花時間（每章約 0.7～1.6 秒），占大部分
_PROGRESS_FETCH_START = 0.1
_PROGRESS_BUILD_START = 0.9
_PROGRESS_SAVE_START = 0.97


def generate(word_path: str, ppt_path: str,
             on_progress: ProgressCallback = lambda progress, message: None) -> GenerationResult:
    result = GenerationResult()

    on_progress(0.0, "讀取 Word…")
    try:
        report = word_reader.read(word_path)
    except Exception as e:
        logger.exception(f"讀取 Word 失敗：{word_path}")
        raise GenerationError(f"無法讀取 Word 檔：{e}") from e
    reading = parse_reading(report.reading_lines)
    sermon = parse_sermon(report.sermon_text, settings.get("speaker_words"))
    result.warnings += reading.warnings + sermon.warnings

    try:
        builder = SlideBuilder(paths.TEMPLATE_PPTX)
    except (OSError, TemplateError) as e:
        logger.exception("讀取範本失敗")
        raise GenerationError(f"無法讀取投影片範本：{e}") from e

    # 先解析所有引用並一次查完經文：查詢可以同時進行，進度也能依實際查詢的章數推進
    slides = _VerseSlides(builder, reading.main_book, result)
    slides.fetch(sermon if report.sermon_text else None, on_progress)

    on_progress(_PROGRESS_BUILD_START, "製作投影片…")
    for passage in reading.passages:
        for verse in passage.verses:
            if not builder.add_verse(passage.title, verse):
                _warn_too_long(result, f"讀經「{passage.title}」", verse.number)
    if not report.sermon_text:
        _warn(result, "找不到證道內容")
    else:
        slides.add_sermon(sermon)

    on_progress(_PROGRESS_SAVE_START, "存檔中…")
    try:
        builder.save(ppt_path)
    except PermissionError as e:
        logger.exception(f"存檔失敗：{ppt_path}")
        raise GenerationError(f"無法儲存 PPT，檔案可能正在 PowerPoint 中開啟，請關閉後再試：{ppt_path}") from e
    except OSError as e:
        logger.exception(f"存檔失敗：{ppt_path}")
        raise GenerationError(f"無法儲存 PPT：{e}") from e
    on_progress(1.0, "製作完畢")
    logger.info(f"製作完成：{ppt_path}，警告 {len(result.warnings)} 則")
    return result


class _VerseSlides:
    """依大綱順序加入標題與其引用經文。經文需先以 fetch 查好。"""

    def __init__(self, builder: SlideBuilder, default_book: Book | None, result: GenerationResult):
        self.builder = builder
        self.default_book = default_book
        self.result = result
        self.parsed: dict[str, list[Reference] | ReferenceParseError] = {}
        self.chapters: dict[tuple[str, int], list[Verse] | BibleApiError] = {}

    def fetch(self, sermon: Sermon | None, on_progress: ProgressCallback) -> None:
        """解析證道中的所有引用，同時查詢用到的每一章。"""
        raws = [] if sermon is None else [*sermon.references, *_section_references(sermon.sections)]
        for raw in raws:
            if raw not in self.parsed:
                try:
                    self.parsed[raw] = parse(raw, self.default_book)
                except ReferenceParseError as e:
                    self.parsed[raw] = e
        needed = [(ref.book, ref.chapter) for refs in self.parsed.values() if isinstance(refs, list) for ref in refs]

        def report(done: int, total: int, book: Book, chapter: int) -> None:
            ratio = done / total   # 以內插計算，查完時剛好等於下一階段的起點
            progress = _PROGRESS_FETCH_START * (1 - ratio) + _PROGRESS_BUILD_START * ratio
            on_progress(progress, f"查詢經文：{format_title(Reference(book, chapter, ()))}（{done}/{total}）")

        on_progress(_PROGRESS_FETCH_START, "查詢經文…")
        self.chapters = bible_api.get_chapters(needed, report)

    def add_sermon(self, sermon: Sermon) -> None:
        self.builder.add_main_title(sermon.title)
        self.add_references(sermon.references, sermon.title or "主標題")
        for major in sermon.sections:
            self.add_major(major)
    def add_major(self, major: Section) -> None:
        self.builder.add_major(major)
        self.add_references(major.references, f"{major.number}、{major.text}")
        for medium in major.children:
            self.builder.add_medium(major, medium)
            self.add_references(medium.references, f"{medium.number}. {medium.text}")
            for minor in medium.children:
                self.builder.add_minor(major, medium, minor)
                self.add_references(minor.references, f"{minor.number}) {minor.text}")

    def add_references(self, raws: list[str], heading: str) -> None:
        for raw in raws:
            where = f"「{heading}」中的「{raw}」"
            refs = self.parsed[raw]
            if isinstance(refs, ReferenceParseError):
                _warn(self.result, f"{where}：{refs}")
                continue
            for ref in refs:
                verses = self.chapters[(ref.book.id, ref.chapter)]
                if isinstance(verses, BibleApiError):
                    _warn(self.result, f"{where}：{format_title(ref)} {verses}")
                    continue
                chapter = {v.number: v for v in verses}
                for verse_range in ref.ranges or [(min(chapter), max(chapter))]:
                    title = format_title(ref, verse_range if ref.ranges else None)
                    missing = []
                    for number in range(verse_range[0], verse_range[1] + 1):
                        verse = chapter.get(number)
                        if verse is None:
                            missing.append(number)
                            continue
                        # 經文去除空白後放入投影片
                        if not self.builder.add_verse(title, Verse(number, verse.text.replace(" ", ""))):
                            _warn_too_long(self.result, where, number)
                    if missing:
                        message = f"{where}：{format_title(ref)}沒有第 {missing[0]} 節，已略過不存在的節數"
                        if has_space_between_digits(raw):
                            # 例如「3:16 4:1」被合併成「3:164」
                            message += "。數字之間有空白，可能是兩段經文之間少了逗號"
                        _warn(self.result, message)


def _section_references(sections: list[Section]) -> list[str]:
    """依大綱順序列出所有標題的引用。"""
    return [raw for section in sections for raw in (*section.references, *_section_references(section.children))]


def _warn_too_long(result: GenerationResult, where: str, number: int) -> None:
    _warn(result, f"{where}：第 {number} 節經文太長，最小字體仍放不下一張投影片，請手動分成兩張")


def _warn(result: GenerationResult, message: str) -> None:
    logger.warning(message)
    result.warnings.append(message)
