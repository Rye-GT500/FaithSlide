"""離線重跑範例週報，輸出必須與基準 snapshot 一致。

目前產生器為現行 FaithSlide.py（驗證基準可重現）；第 7 階段改接 pipeline.generate。
"""
import json

import pytest

from tests import legacy_runner
from tests.snapshot_tools import build_snapshot, sample_reports, snapshot_path


@pytest.mark.parametrize("docx_path", sample_reports(), ids=lambda p: p.stem)
def test_matches_baseline(docx_path, tmp_path):
    expected_path = snapshot_path(docx_path)
    if not expected_path.exists():
        pytest.fail(f"缺少基準，請先執行 python -m tests.make_baseline：{expected_path.name}")

    cache = legacy_runner.VerseCache(allow_network=False)
    actual = build_snapshot(docx_path, tmp_path / "out.pptx", cache)

    assert cache.missing == [], "API 快取缺少章節，請重新錄製基準"
    assert actual == json.loads(expected_path.read_text(encoding="utf-8"))
