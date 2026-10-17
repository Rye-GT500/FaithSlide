import zipfile

import pytest
from pptx.util import Pt

from faithslide import paths
from faithslide.bible_api import Verse
from faithslide.ppt_builder import SlideBuilder, TemplateError, calculate_font_size, fits, wrap_lines
from faithslide.sermon_parser import MAJOR, MEDIUM, MINOR, Section
from tests.slide_extract import extract_slides


def texts(slide):
    """每個文字框第一段的 run 文字。"""
    return [[run[0] for run in shape[0]] for shape in slide["shapes"]]


@pytest.fixture
def outline():
    minors = [Section(MINOR, "1", "小一：", []), Section(MINOR, "2", "小二", [])]
    mediums = [Section(MEDIUM, "1", "中一：", []), Section(MEDIUM, "2", "中二：", [], minors)]
    return Section(MAJOR, "一", "大標題：", [], mediums)


def build(tmp_path, steps):
    builder = SlideBuilder(paths.TEMPLATE_PPTX)
    steps(builder)
    out = tmp_path / "out.pptx"
    builder.save(str(out))
    return out, extract_slides(out)


def test_main_title_adds_two_slides(tmp_path):
    _, slides = build(tmp_path, lambda b: b.add_main_title("主標題"))
    assert len(slides) == 2
    assert "主標題" in texts(slides[0])[1]


def test_major(tmp_path, outline):
    _, slides = build(tmp_path, lambda b: b.add_major(outline))
    assert texts(slides[0])[0][:2] == ["一、", "大標題："]


def test_medium_lists_up_to_current(tmp_path, outline):
    _, slides = build(tmp_path, lambda b: [b.add_medium(outline, m) for m in outline.children])
    assert texts(slides[0])[0][:2] == ["一、", "大標題："]
    assert texts(slides[0])[1][:2] == ["1.", "中一\n"]
    assert texts(slides[1])[1][:4] == ["1.", "中一\n", "2.", "中二\n"]


def test_minor_lists_medium_and_minors_up_to_current(tmp_path, outline):
    medium = outline.children[1]
    _, slides = build(tmp_path, lambda b: [b.add_minor(outline, medium, m) for m in medium.children])
    assert texts(slides[0])[1][:4] == ["2.", "中二\n", "(1)", "小一\n"]
    assert texts(slides[1])[1][:6] == ["2.", "中二\n", "(1)", "小一\n", "(2)", "小二\n"]


def test_verse(tmp_path):
    _, slides = build(tmp_path, lambda b: [
        b.add_verse("約翰福音三章16節", Verse(16, "神愛世人，　甚至")),
        b.add_verse("約翰福音三章1節", Verse(1, "有一個法利賽人")),
    ])
    long_number, short_number = slides
    assert texts(long_number)[0][0] == "約翰福音三章16節"
    assert texts(long_number)[1][:2] == ["16.\t", "神愛世人， 甚至"]
    assert texts(short_number)[1][:2] == ["1.\t", "有一個法利賽人"]
    # 節數與經文使用相同字體大小（XML 以 1/100 點儲存）
    sizes = [run[1] for run in long_number["shapes"][1][0][:2]]
    assert sizes[0] == sizes[1]
    assert abs(sizes[0] - calculate_font_size("16.", "神愛世人， 甚至")) <= Pt(0.01)


def test_verse_psalm_title(tmp_path):
    _, slides = build(tmp_path, lambda b: b.add_verse("詩篇二十三章1節", Verse(1, "x")))
    assert texts(slides[0])[0][0] == "詩篇二十三篇1節"


def test_save_removes_template_slides_from_package(tmp_path):
    out, slides = build(tmp_path, lambda b: b.add_verse("約翰福音三章1節", Verse(1, "x")))
    assert len(slides) == 1
    slide_parts = [n for n in zipfile.ZipFile(out).namelist() if n.startswith("ppt/slides/slide")]
    assert len(slide_parts) == 1


def test_template_with_too_few_slides(tmp_path):
    from pptx import Presentation
    path = tmp_path / "empty.pptx"
    Presentation().save(str(path))
    with pytest.raises(TemplateError):
        SlideBuilder(str(path))


@pytest.mark.parametrize("text, expected", [
    ("", Pt(72)),
    ("短", Pt(72)),
    ("字" * 2000, Pt(20)),
])
def test_calculate_font_size_bounds(text, expected):
    assert abs(calculate_font_size("16.", text) - expected) <= Pt(0.01)


def test_calculate_font_size_shrinks_with_length():
    sizes = [calculate_font_size("16.", "字" * n) for n in (10, 50, 100, 200)]
    assert sizes == sorted(sizes, reverse=True)
    assert sizes[0] > sizes[-1]


@pytest.mark.parametrize("n", [20, 40, 60, 80, 120, 200])
def test_calculate_font_size_is_largest_that_fits(n):
    text = "神愛世人，甚至將他的獨生子賜給他們。" * 20
    text = text[:n]
    size = calculate_font_size("16.", text)
    assert fits("16.", text, size)
    if size < Pt(72):
        assert not fits("16.", text, size + Pt(0.5))


def test_wider_number_column_gets_smaller_or_equal_font():
    text = "字" * 70
    assert calculate_font_size("176.", text) <= calculate_font_size("1.", text)


def test_too_long_verse_reported(tmp_path):
    builder = SlideBuilder(paths.TEMPLATE_PPTX)
    assert builder.add_verse("標題", Verse(1, "短")) is True
    assert builder.add_verse("標題", Verse(1, "字" * 2000)) is False


# --- 換行模擬（規則以 PowerPoint 實際排版校準） ---------------------------------

def test_wrap_lines_width_follows_number_column():
    # 72 點、節數「16.」：經文欄從 1.696 個字寬開始，可用 (945.6 - 122.1) / 72 = 11.4 個字
    lines = wrap_lines("16.", "一二三四五六七八九十甲乙丙丁戊", Pt(72))
    assert lines == ["一二三四五六七八九十甲", "乙丙丁戊"]


def test_wrap_lines_one_closing_punctuation_hangs():
    # 第 12 個字是「，」：懸掛在行尾，不換行
    lines = wrap_lines("16.", "一二三四五六七八九十甲，乙丙", Pt(72))
    assert lines == ["一二三四五六七八九十甲，", "乙丙"]


def test_wrap_lines_second_closing_punctuation_moves_with_previous_char():
    # 「。」懸掛後「」」也放不下：連同前一個字「甲」移到下一行，避免下一行以標點開頭
    lines = wrap_lines("16.", "一二三四五六七八九十甲。」乙", Pt(72))
    assert lines == ["一二三四五六七八九十", "甲。」乙"]


def test_wrap_lines_opening_punctuation_not_left_at_line_end():
    lines = wrap_lines("16.", "一二三四五六七八九十「甲乙", Pt(72))
    assert lines == ["一二三四五六七八九十", "「甲乙"]


def test_wrap_lines_narrow_punctuation():
    # 「‧」只有 0.254 個字寬，一行可以多放
    assert wrap_lines("16.", "一二三‧四五六七八九十甲", Pt(72)) == ["一二三‧四五六七八九十甲"]


def _verse_paragraph_props(tmp_path, number, text):
    from pptx import Presentation
    from pptx.oxml.ns import qn
    out, _ = build(tmp_path, lambda b: b.add_verse("標題", Verse(number, text)))
    p = Presentation(str(out)).slides[0].shapes[1].text_frame.paragraphs[0]
    pPr = p._p.find(qn("a:pPr"))
    tabs = [(int(t.get("pos")), t.get("algn")) for t in pPr.findall(qn("a:tabLst") + "/" + qn("a:tab"))]
    return int(pPr.get("marL")), int(pPr.get("indent")), tabs


def _number_em(number):
    from faithslide.ppt_builder import _GLYPH_WIDTH_EM
    return sum(_GLYPH_WIDTH_EM[c] for c in f"{number}.")


@pytest.mark.parametrize("number, text", [(1, "短"), (16, "字" * 40), (176, "字" * 120)])
def test_verse_hanging_indent_matches_number_width(tmp_path, number, text):
    """換行後的文字要和第一行經文對齊：凸排距離與 Tab 定位點都等於節數欄寬，且隨字體大小縮放。"""
    from faithslide.ppt_builder import _NUMBER_GAP_EM
    expected = round(calculate_font_size(f"{number}.", text) * (_number_em(number) + _NUMBER_GAP_EM))
    assert _verse_paragraph_props(tmp_path, number, text) == (expected, -expected, [(expected, "l")])


@pytest.mark.parametrize("mark", ["「", "『", "（", "《", "〈", "【", "〔", "“", "‘"])
def test_opening_punctuation_hangs_left(tmp_path, mark):
    """標點懸掛：Tab 停在節數後，標點字格在經文欄左側，後面的字與換行後的文字同在經文欄。"""
    from faithslide.ppt_builder import _HANGING_GAP_EM
    text = mark + "神愛世人" * 10
    size = calculate_font_size("16.", text)
    tab = round(size * (_number_em(16) + _HANGING_GAP_EM))
    column = round(size * (_number_em(16) + _HANGING_GAP_EM + 1))
    assert _verse_paragraph_props(tmp_path, 16, text) == (column, -column, [(tab, "l")])


def test_verse_text_is_trimmed(tmp_path):
    # 讀經「[1]」後常有 Tab，第二個 Tab 會讓第一行跳到其他定位點
    _, slides = build(tmp_path, lambda b: b.add_verse("標題", Verse(1, "	 　以法蓮山地 ")))
    assert texts(slides[0])[1][:2] == ["1.	", "以法蓮山地"]


def test_one_and_two_digit_verses_use_same_template(tmp_path):
    _, slides = build(tmp_path, lambda b: [b.add_verse("標題", Verse(n, "經文")) for n in (1, 10)])
    assert slides[0]["layout"] == slides[1]["layout"]
    assert len(slides[0]["shapes"]) == len(slides[1]["shapes"])
