import pytest

from faithslide import bible_api, pipeline
from faithslide.bible_api import BibleApiError, Verse
from faithslide.pipeline import GenerationError
from tests.docx_factory import make_report
from tests.slide_extract import extract_slides

READING = "約翰福音三章1-2節\n[1]第一節\n[2]第二節"


@pytest.fixture
def chapters(monkeypatch):
    """假經文：每章 5 節；放入 chapters.unavailable 的 (書卷 id, 章) 會查詢失敗。"""
    class Fake:
        unavailable = set()

        def get_chapter(self, book, chapter):
            if (book.id, chapter) in self.unavailable:
                raise BibleApiError("查無經文")
            return [Verse(n, f"{book.abbr}{chapter}:{n} 經 文") for n in range(1, 6)]

    fake = Fake()
    monkeypatch.setattr(bible_api, "get_chapter", fake.get_chapter)
    return fake


def run(tmp_path, sermon, reading=READING):
    rows = [("讀經", reading)] if reading is not None else []
    if sermon is not None:
        rows.append(("證道", [(sermon, True)]))
    docx = make_report(tmp_path / "in.docx", rows)
    out = tmp_path / "out.pptx"
    progress = []
    result = pipeline.generate(docx, str(out), lambda p, m: progress.append(p))
    return result, extract_slides(out), progress


def verse_texts(slides):
    return [slide["shapes"][1][0][1][0] for slide in slides if len(slide["shapes"]) > 1 and len(slide["shapes"][1][0]) > 1]


def test_generate(tmp_path, chapters):
    result, slides, progress = run(tmp_path, "主標題 一、大標題 3:1-2 1. 中標題 羅 1:3")
    assert result.warnings == []
    # 讀經 2 + 主標題 2 + 大標題 1 + 經文 2 + 中標題 1 + 經文 1
    assert len(slides) == 9
    # 省略書卷的引用使用讀經的主經文書卷；經文去除空白
    assert "約3:1經文" in verse_texts(slides)
    assert "羅1:3經文" in verse_texts(slides)
    assert progress[-1] == 1.0
    assert progress == sorted(progress)


def test_main_title_references_are_made(tmp_path, chapters):
    result, slides, _ = run(tmp_path, "主標題 羅 1:1 一、大標題")
    assert result.warnings == []
    assert "羅1:1經文" in verse_texts(slides)


def test_bad_reference_is_skipped_with_warning(tmp_path, chapters):
    result, slides, _ = run(tmp_path, "主標題 一、大標題 3:4-2 羅 1:1")
    assert any("節數範圍錯誤" in w for w in result.warnings)
    assert "羅1:1經文" in verse_texts(slides)


def test_unavailable_chapter_is_skipped_with_warning(tmp_path, chapters):
    chapters.unavailable = {("rom", 1)}
    result, slides, _ = run(tmp_path, "主標題 一、大標題 羅 1:1 約 3:1")
    assert any("羅馬書一章" in w for w in result.warnings)
    assert "約3:1經文" in verse_texts(slides)


def test_missing_verses_warn_once_per_range(tmp_path, chapters):
    result, slides, _ = run(tmp_path, "主標題 一、大標題 羅 1:4-9")
    assert verse_texts(slides).count("羅1:4經文") == 1
    assert len(result.warnings) == 1
    assert "沒有第 6 節" in result.warnings[0]


def test_missing_reading_and_sermon(tmp_path, chapters):
    result, slides, _ = run(tmp_path, None, reading=None)
    assert slides == []
    assert any("讀經" in w for w in result.warnings)
    assert any("證道" in w for w in result.warnings)


def test_reference_without_book_and_no_reading(tmp_path, chapters):
    result, _, _ = run(tmp_path, "主標題 一、大標題 3:1", reading=None)
    assert any("缺少書卷" in w for w in result.warnings)


def test_unreadable_word_file(tmp_path):
    bad = tmp_path / "bad.docx"
    bad.write_text("not a docx")
    with pytest.raises(GenerationError, match="無法讀取 Word"):
        pipeline.generate(str(bad), str(tmp_path / "out.pptx"))


def test_unwritable_output(tmp_path, chapters):
    docx = make_report(tmp_path / "in.docx", [("讀經", READING)])
    with pytest.raises(GenerationError, match="無法儲存") as error:
        pipeline.generate(docx, str(tmp_path / "no_such_dir" / "out.pptx"))
    assert "PowerPoint" not in str(error.value)


def test_output_locked_by_powerpoint(tmp_path, chapters, monkeypatch):
    docx = make_report(tmp_path / "in.docx", [("讀經", READING)])

    def locked(self, path):
        raise PermissionError(13, "Permission denied", path)
    monkeypatch.setattr(pipeline.SlideBuilder, "save", locked)
    with pytest.raises(GenerationError, match="PowerPoint 中開啟"):
        pipeline.generate(docx, str(tmp_path / "out.pptx"))


def test_warning_names_heading_and_original_text(tmp_path, chapters):
    result, _, _ = run(tmp_path, "主標題 一、大標題 1. 中標題 羅 1:9")
    assert result.warnings == ["「1. 中標題」中的「羅 1:9」：羅馬書一章沒有第 9 節，已略過不存在的節數"]


def test_merged_references_get_missing_comma_hint(tmp_path, chapters):
    # 「3:1 4:1」被合併成「3:14」，超出本章範圍
    result, _, _ = run(tmp_path, "主標題 一、大標題 羅 3:1 4:1")
    assert any("少了逗號" in w and "羅 3:1 4" in w for w in result.warnings)
