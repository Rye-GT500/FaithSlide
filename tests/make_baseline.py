"""建立重構前的基準：python -m tests.make_baseline

以現行 FaithSlide.py 處理每份範例週報，錄製 API 回應並寫出 snapshot。
已錄製的章節不會重新連網。
"""
import tempfile
from pathlib import Path

from tests import legacy_runner
from tests.slide_extract import extract_slides
from tests.snapshot_tools import SNAPSHOT_DIR, dump, sample_reports, snapshot_path


def build_snapshot(docx_path: Path, out_pptx: Path, cache: legacy_runner.VerseCache) -> dict:
    messages = legacy_runner.run(docx_path, out_pptx, cache)
    slides = extract_slides(out_pptx) if out_pptx.exists() else None
    return {"source": docx_path.name, "messages": messages, "slides": slides}


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
