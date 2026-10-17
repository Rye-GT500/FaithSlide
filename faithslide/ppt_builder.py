"""依 template.pptx 的範本投影片產生簡報。

範本前 6 張投影片是各類投影片的樣板，新投影片以複製樣板的方式加在後面，存檔前刪除樣板。
"""
import copy
import re

from pptx import Presentation
from pptx.util import Pt

from faithslide.bible_api import Verse
from faithslide.sermon_parser import Section

# 範本投影片索引
VERSE_LONG_NUMBER = 0     # 經文（節數兩位數以上）
MAIN_TITLE = 1            # 主標題
AFTER_MAIN_TITLE = 2      # 接在主標題之後
MAJOR = 3                 # 大標題
OUTLINE = 4               # 中標題 / 小標題（大綱列表）
VERSE_SHORT_NUMBER = 5    # 經文（節數一位數）
TEMPLATE_SLIDE_COUNT = 6

# 字體大小計算參數（依範本文字框大小實驗調整）
_MAX_WIDTH_UNITS = 854
_MAX_HEIGHT_UNITS = 372
_MIN_FONT_SIZE = 20
_BASE_FONT_SIZE = 72


class TemplateError(Exception):
    pass


class SlideBuilder:
    def __init__(self, template_path: str):
        self.prs = Presentation(template_path)
        if len(self.prs.slides) < TEMPLATE_SLIDE_COUNT:
            raise TemplateError(f"範本投影片不足 {TEMPLATE_SLIDE_COUNT} 張：{template_path}")

    def add_main_title(self, title: str) -> None:
        slide = self._duplicate(MAIN_TITLE)
        _first_paragraph(slide.shapes[1]).runs[0].text = title
        self._duplicate(AFTER_MAIN_TITLE)

    def add_major(self, major: Section) -> None:
        slide = self._duplicate(MAJOR)
        self._set_major_heading(slide, major)

    def add_medium(self, major: Section, current: Section) -> None:
        """大綱列出 major 底下第一個到 current 為止的中標題。"""
        slide = self._duplicate(OUTLINE)
        self._set_major_heading(slide, major)
        p = _first_paragraph(slide.shapes[1])
        for i, medium in enumerate(major.children):
            p.runs[2 * i].text = f"{medium.number}."
            p.runs[2 * i + 1].text = _outline_text(medium)
            if medium is current:
                break
            p.add_run()
            p.add_run()

    def add_minor(self, major: Section, medium: Section, current: Section) -> None:
        """大綱列出所屬中標題，以及其下第一個到 current 為止的小標題。"""
        slide = self._duplicate(OUTLINE)
        self._set_major_heading(slide, major)
        p = _first_paragraph(slide.shapes[1])
        p.runs[0].text = f"{medium.number}."
        p.runs[1].text = _outline_text(medium)
        for i, minor in enumerate(medium.children, start=1):
            p.add_run()
            p.add_run()
            p.runs[2 * i].text = f"({minor.number})"
            p.runs[2 * i + 1].text = _outline_text(minor)
            if minor is current:
                break

    def add_verse(self, title: str, verse: Verse) -> None:
        number = f"{verse.number}."
        slide = self._duplicate(VERSE_SHORT_NUMBER if verse.number < 10 else VERSE_LONG_NUMBER)

        p = _first_paragraph(slide.shapes[0])
        if "詩篇" in title:
            title = title.replace("章", "篇")
        p.runs[0].text = title
        for run in p.runs[1:3]:
            run.text = ""

        p = _first_paragraph(slide.shapes[1])
        text = verse.text.replace("　", " ")
        p.runs[0].text = number
        p.runs[1].text = text
        size = calculate_font_size(text)
        p.runs[0].font.size = size
        p.runs[1].font.size = size

    def save(self, path: str) -> None:
        """刪除樣板投影片後存檔。"""
        slide_ids = self.prs.slides._sldIdLst
        for slide_id in list(slide_ids)[:TEMPLATE_SLIDE_COUNT]:
            self.prs.part.drop_rel(slide_id.rId)
            slide_ids.remove(slide_id)
        self.prs.save(path)

    def _duplicate(self, index: int):
        template = self.prs.slides[index]
        slide = self.prs.slides.add_slide(template.slide_layout)
        for shape in list(slide.shapes):
            if shape.is_placeholder:
                slide.shapes._spTree.remove(shape._element)
        for shape in template.shapes:
            slide.shapes._spTree.insert_element_before(copy.deepcopy(shape.element), "p:extLst")
        return slide

    @staticmethod
    def _set_major_heading(slide, major: Section) -> None:
        p = _first_paragraph(slide.shapes[0])
        p.runs[0].text = f"{major.number}、"
        p.runs[1].text = major.text


def _first_paragraph(shape):
    p = shape.text_frame.paragraphs[0]
    if not p.runs:
        p.add_run()
    return p


def _outline_text(section: Section) -> str:
    return section.text.replace("：", "") + "\n"


def _weighted_length(text: str) -> float:
    """中文字與全形標點算 1，其他算 0.5。"""
    return sum(1.0 if re.match(r'[一-鿿]|["，。？！：「」；]', char) else 0.5 for char in text)


def calculate_font_size(text: str, base_size: float = _BASE_FONT_SIZE) -> Pt:
    """以二分法找出能放進文字框的最大字體。"""
    total_length = _weighted_length(text)
    low, high = _MIN_FONT_SIZE, base_size
    while high - low > 0.01:
        mid = (low + high) / 2
        capacity = int(_MAX_WIDTH_UNITS / mid) * int(_MAX_HEIGHT_UNITS / mid)
        if total_length > capacity:
            high = mid
        else:
            low = mid
    return Pt(min(max(low, _MIN_FONT_SIZE), base_size))
