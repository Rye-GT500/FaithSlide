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
from faithslide.reference_parser import ReferenceParseError, format_title, has_space_between_digits, parse
from faithslide.sermon_parser import Section, Sermon, parse_reading, parse_sermon

logger = logging.getLogger(__name__)

ProgressCallback = Callable[[float, str], None]   # (0～1 的進度, 訊息)


class GenerationError(Exception):
    """訊息可直接顯示給使用者。"""


@dataclass
class GenerationResult:
    warnings: list[str] = field(default_factory=list)


def generate(word_path: str, ppt_path: str,
             on_progress: ProgressCallback = lambda progress, message: None) -> GenerationResult:
    result = GenerationResult()

    try:
        report = word_reader.read(word_path)
    except Exception as e:
        logger.exception(f"讀取 Word 失敗：{word_path}")
        raise GenerationError(f"無法讀取 Word 檔：{e}") from e
    reading = parse_reading(report.reading_lines)
    sermon = parse_sermon(report.sermon_text, settings.get("speaker_words"))
    result.warnings += reading.warnings + sermon.warnings
    on_progress(0.1, "分析 Word 完畢")

    try:
        builder = SlideBuilder(paths.TEMPLATE_PPTX)
    except (OSError, TemplateError) as e:
        logger.exception("讀取範本失敗")
        raise GenerationError(f"無法讀取投影片範本：{e}") from e

    for passage in reading.passages:
        for verse in passage.verses:
            builder.add_verse(passage.title, verse)
    on_progress(0.2, "讀經製作完畢")

    if not report.sermon_text:
        _warn(result, "找不到證道內容")
    else:
        _VerseSlides(builder, reading.main_book, result).add_sermon(sermon, on_progress)

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
    """依大綱順序加入標題與其引用經文。"""

    def __init__(self, builder: SlideBuilder, default_book: Book | None, result: GenerationResult):
        self.builder = builder
        self.default_book = default_book
        self.result = result

    def add_sermon(self, sermon: Sermon, on_progress: ProgressCallback) -> None:
        self.builder.add_main_title(sermon.title)
        self.add_references(sermon.references, sermon.title or "主標題")
        for i, major in enumerate(sermon.sections):
            on_progress(0.2 + 0.75 * i / len(sermon.sections), f"{major.number}、{major.text}")
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
            try:
                refs = parse(raw, self.default_book)
            except ReferenceParseError as e:
                _warn(self.result, f"{where}：{e}")
                continue
            for ref in refs:
                try:
                    chapter = {v.number: v for v in bible_api.get_chapter(ref.book, ref.chapter)}
                except BibleApiError as e:
                    _warn(self.result, f"{where}：{format_title(ref)} {e}")
                    continue
                for verse_range in ref.ranges or [(min(chapter), max(chapter))]:
                    title = format_title(ref, verse_range if ref.ranges else None)
                    missing = []
                    for number in range(verse_range[0], verse_range[1] + 1):
                        verse = chapter.get(number)
                        if verse is None:
                            missing.append(number)
                            continue
                        # 經文去除空白後放入投影片
                        self.builder.add_verse(title, Verse(number, verse.text.replace(" ", "")))
                    if missing:
                        message = f"{where}：{format_title(ref)}沒有第 {missing[0]} 節，已略過不存在的節數"
                        if has_space_between_digits(raw):
                            # 例如「3:16 4:1」被合併成「3:164」
                            message += "。數字之間有空白，可能是兩段經文之間少了逗號"
                        _warn(self.result, message)


def _warn(result: GenerationResult, message: str) -> None:
    logger.warning(message)
    result.warnings.append(message)
