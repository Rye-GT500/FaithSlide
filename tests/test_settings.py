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


DEFAULT_SPEAKERS = ["證道", "傳道", "牧師", "吳佩倫", "錢致榮"]


def test_get_returns_copy(user_file):
    settings.load(user_path=str(user_file))
    settings.get("speaker_words").append("x")
    assert settings.get("speaker_words") == DEFAULT_SPEAKERS


def test_save_writes_only_overrides_and_takes_effect(user_file):
    settings.load(user_path=str(user_file))
    settings.save("speaker_words", ["牧師", "王小明"])
    assert settings.get("speaker_words") == ["牧師", "王小明"]
    assert json.loads(user_file.read_text(encoding="utf-8")) == {"schema_version": 1, "speaker_words": ["牧師", "王小明"]}
    settings.load(user_path=str(user_file))
    assert settings.get("speaker_words") == ["牧師", "王小明"]


def test_save_default_value_removes_override(user_file):
    settings.load(user_path=str(user_file))
    settings.save("speaker_words", ["王小明"])
    settings.save("speaker_words", DEFAULT_SPEAKERS)
    assert json.loads(user_file.read_text(encoding="utf-8")) == {"schema_version": 1}
    assert settings.default("speaker_words") == DEFAULT_SPEAKERS


@pytest.mark.parametrize("key, value, error", [("unknown", [], KeyError), ("speaker_words", "牧師", TypeError)])
def test_save_rejects_invalid(user_file, key, value, error):
    settings.load(user_path=str(user_file))
    with pytest.raises(error):
        settings.save(key, value)
    assert not user_file.exists()
