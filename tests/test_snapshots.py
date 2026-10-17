"""以 pipeline 離線重跑範例週報，投影片必須與重構前的基準 snapshot 一致。"""
import json

import pytest

from faithslide import bible_api, pipeline
from tests.fixture_api import FixtureBible
from tests.slide_extract import extract_slides
from tests.snapshot_tools import sample_reports, snapshot_path


@pytest.mark.parametrize("docx_path", sample_reports(), ids=lambda p: p.stem)
def test_matches_baseline(docx_path, tmp_path, monkeypatch):
    expected = json.loads(snapshot_path(docx_path).read_text(encoding="utf-8"))
    bible = FixtureBible()
    monkeypatch.setattr(bible_api, "get_chapter", bible.get_chapter)

    out = tmp_path / "out.pptx"
    result = pipeline.generate(str(docx_path), str(out))

    assert bible.missing == [], "API 測試資料缺少章節"
    assert result.warnings == []
    assert extract_slides(out) == expected["slides"]
