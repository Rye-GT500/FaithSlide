import json

import pytest

from faithslide import settings


@pytest.fixture
def user_file(tmp_path):
    path = tmp_path / "settings.json"
    yield path
    settings._settings = None


def test_defaults(user_file):
    settings.load(user_path=str(user_file))
    assert settings.get("speaker_words") == ["證道", "傳道", "牧師", "吳佩倫", "錢致榮"]


def test_user_overrides_default(user_file):
    user_file.write_text(json.dumps({"speaker_words": ["牧師", "王小明"]}, ensure_ascii=False), encoding="utf-8")
    settings.load(user_path=str(user_file))
    assert settings.get("speaker_words") == ["牧師", "王小明"]


@pytest.mark.parametrize("content", [
    "not json",
    "[1, 2]",
    json.dumps({"speaker_words": "牧師"}, ensure_ascii=False),   # 型別錯誤
    json.dumps({"unknown_key": 1}),
])
def test_bad_user_file_falls_back_to_defaults(user_file, content):
    user_file.write_text(content, encoding="utf-8")
    settings.load(user_path=str(user_file))
    assert settings.get("speaker_words") == ["證道", "傳道", "牧師", "吳佩倫", "錢致榮"]
