import pytest
import requests

from faithslide import bible_api, bible_books
from faithslide.bible_api import BibleApiError, Verse


class FakeResponse:
    def __init__(self, status_code=200, payload=None, invalid_json=False):
        self.status_code = status_code
        self._payload = payload
        self._invalid_json = invalid_json

    def json(self):
        if self._invalid_json:
            raise ValueError("not json")
        return self._payload


def success(*records):
    return FakeResponse(payload={"status": "success", "record": [{"sec": s, "bible_text": t} for s, t in records]})


@pytest.fixture
def fake_get(monkeypatch):
    """替換 requests.get；以 fake_get.response 設定回應，fake_get.calls 記錄呼叫參數。"""
    class Fake:
        response = success()
        calls = []

        def __call__(self, url, params=None, timeout=None):
            self.calls.append(params)
            if isinstance(self.response, Exception):
                raise self.response
            return self.response

    fake = Fake()
    monkeypatch.setattr(bible_api.requests, "get", fake)
    bible_api.clear_cache()
    yield fake
    bible_api.clear_cache()


def test_get_chapter(fake_get):
    fake_get.response = success((2, "第二節"), (1, "第一節"))
    verses = bible_api.get_chapter(bible_books.resolve("約"), 3)
    assert verses == [Verse(1, "第一節"), Verse(2, "第二節")]
    assert fake_get.calls == [{"chineses": "約", "chap": 3, "version": "unv", "strong": "0"}]


def test_alias_uses_api_value(fake_get):
    fake_get.response = success((1, "x"))
    bible_api.get_chapter(bible_books.resolve("約壹"), 2)
    bible_api.get_chapter(bible_books.resolve("各"), 1)
    assert [c["chineses"] for c in fake_get.calls] == ["約一", "雅"]


def test_sec_as_string(fake_get):
    fake_get.response = success(("1", "x"))
    assert bible_api.get_chapter(bible_books.resolve("約"), 3) == [Verse(1, "x")]


def test_same_chapter_fetched_once(fake_get):
    fake_get.response = success((1, "x"))
    book = bible_books.resolve("約")
    bible_api.get_chapter(book, 3)
    bible_api.get_chapter(bible_books.resolve("約翰福音"), 3)
    assert len(fake_get.calls) == 1


def test_returned_list_is_a_copy(fake_get):
    fake_get.response = success((1, "x"))
    book = bible_books.resolve("約")
    bible_api.get_chapter(book, 3).clear()
    assert bible_api.get_chapter(book, 3) == [Verse(1, "x")]


@pytest.mark.parametrize("response, message", [
    (requests.ConnectionError("down"), "無法連線"),
    (requests.Timeout("slow"), "無法連線"),
    (FakeResponse(status_code=500), "HTTP 500"),
    (FakeResponse(payload={"status": "error"}), "回應錯誤"),
    (FakeResponse(invalid_json=True), "格式錯誤"),
    (FakeResponse(payload={"status": "success"}), "格式錯誤"),
    (success(), "查無經文：約翰福音 第 99 章"),
])
def test_errors(fake_get, response, message):
    fake_get.response = response
    with pytest.raises(BibleApiError, match=message):
        bible_api.get_chapter(bible_books.resolve("約"), 99)


def test_failure_not_cached(fake_get):
    book = bible_books.resolve("約")
    fake_get.response = requests.ConnectionError("down")
    with pytest.raises(BibleApiError):
        bible_api.get_chapter(book, 3)
    fake_get.response = success((1, "x"))
    assert bible_api.get_chapter(book, 3) == [Verse(1, "x")]
