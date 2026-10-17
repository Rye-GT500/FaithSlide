import pytest
import requests

from faithslide import updater
from faithslide.updater import version_key


class FakeResponse:
    def __init__(self, status_code=200, payload=None):
        self.status_code = status_code
        self._payload = payload

    def json(self):
        return self._payload


@pytest.fixture
def respond(monkeypatch):
    def set_response(response):
        def fake_get(url, timeout=None):
            if isinstance(response, Exception):
                raise response
            return response
        monkeypatch.setattr(updater.requests, "get", fake_get)
    monkeypatch.setattr(updater, "CURRENT_VERSION", "1.2.2-beta")
    return set_response


def release(tag):
    return FakeResponse(payload={"tag_name": tag, "html_url": "https://example.com/release"})


@pytest.mark.parametrize("older, newer", [
    ("1.9.0", "1.10.0"),          # 舊版字串比較會判斷錯誤
    ("1.2.2", "1.2.3"),
    ("1.2.2-beta", "1.2.2"),      # 正式版大於預覽版
    ("1.2.2-beta", "1.2.3-beta"),
    ("1.2.2-alpha", "1.2.2-beta"),
    ("1.2", "1.2.1"),
    ("v1.2.2", "1.2.3"),
])
def test_version_order(older, newer):
    assert version_key(older) < version_key(newer)


def test_equal_versions():
    assert version_key("1.2") == version_key("1.2.0")
    assert version_key("v1.2.2-beta") == version_key("1.2.2-beta")


@pytest.mark.parametrize("version", ["", "beta", "1.x", "1.2.2 beta"])
def test_invalid_version(version):
    with pytest.raises(ValueError):
        version_key(version)


@pytest.mark.parametrize("tag, expected", [
    ("v1.2.3-beta", "1.2.3-beta"),
    ("v1.2.2", "1.2.2"),
    ("v1.10.0", "1.10.0"),
])
def test_newer_version(respond, tag, expected):
    respond(release(tag))
    assert updater.check() == updater.UpdateInfo(expected, "https://example.com/release")


@pytest.mark.parametrize("tag", ["v1.2.2-beta", "v1.2.1", "v1.2.2-alpha"])
def test_not_newer(respond, tag):
    respond(release(tag))
    assert updater.check() is None


@pytest.mark.parametrize("response", [
    requests.ConnectionError("down"),
    FakeResponse(status_code=403),
    FakeResponse(payload={"message": "rate limited"}),
    release("nightly"),
])
def test_failures_return_none(respond, response):
    respond(response)
    assert updater.check() is None
