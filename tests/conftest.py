import pytest

from faithslide import bible_books, paths, settings


@pytest.fixture(autouse=True)
def isolated_user_dir(tmp_path, monkeypatch):
    """測試不讀寫真正的 %APPDATA%\\FaithSlide；每個測試重新載入書卷與設定。"""
    user_dir = tmp_path / "appdata"
    user_dir.mkdir()
    monkeypatch.setattr(paths, "user_dir", lambda: str(user_dir))
    bible_books.load()
    settings._settings = None
    yield user_dir
    settings._settings = None
