"""行為刻意改變、確認新輸出正確後，以目前的 pipeline 重新產生 snapshot：

    python -m tests.update_snapshots

離線執行，經文來自 tests/fixtures/api/verses.json。執行前請先確認 git diff 中的投影片變動都是預期的。
"""
import tempfile
from pathlib import Path
from unittest import mock

from faithslide import bible_api, pipeline
from tests.fixture_api import FixtureBible
from tests.slide_extract import extract_slides
from tests.snapshot_tools import dump, sample_reports, snapshot_path


def main():
    bible = FixtureBible()
    with tempfile.TemporaryDirectory() as tmp, mock.patch.object(bible_api, "get_chapter", bible.get_chapter):
        for docx_path in sample_reports():
            out = Path(tmp) / f"{docx_path.stem}.pptx"
            result = pipeline.generate(str(docx_path), str(out))
            snapshot = {"source": docx_path.name, "warnings": result.warnings, "slides": extract_slides(out)}
            snapshot_path(docx_path).write_text(dump(snapshot), encoding="utf-8")
            print(f"{docx_path.name}: 投影片 {len(snapshot['slides'])}，警告 {len(result.warnings)}")
    if bible.missing:
        print(f"注意：測試資料缺少章節 {bible.missing}")


if __name__ == "__main__":
    main()
