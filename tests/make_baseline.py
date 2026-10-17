"""建立重構前的基準：python -m tests.make_baseline

以現行 FaithSlide.py 處理每份範例週報，錄製 API 回應並寫出 snapshot。
已錄製的章節不會重新連網。
"""
import tempfile
from pathlib import Path

from tests import legacy_runner
from tests.snapshot_tools import SNAPSHOT_DIR, build_snapshot, dump, sample_reports, snapshot_path


def main():
    cache = legacy_runner.VerseCache(allow_network=True)
    SNAPSHOT_DIR.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory() as tmp:
        for docx_path in sample_reports():
            snapshot = build_snapshot(docx_path, Path(tmp) / f"{docx_path.stem}.pptx", cache)
            snapshot_path(docx_path).write_text(dump(snapshot), encoding="utf-8")
            slide_count = len(snapshot["slides"]) if snapshot["slides"] is not None else "未產生 PPT"
            print(f"{docx_path.name}: 投影片 {slide_count}，訊息 {len(snapshot['messages'])}")
    cache.save()
    if cache.empty:
        print(f"注意：以下章節查詢結果為空 {cache.empty}")


if __name__ == "__main__":
    main()
