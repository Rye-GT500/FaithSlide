import pytest
import requests

from faithslide import updater


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


def test_newer_version(respond):
    respond(release("v1.2.3-beta"))
    assert updater.check() == updater.UpdateInfo("1.2.3", "https://example.com/release")


@pytest.mark.parametrize("tag", ["v1.2.2-beta", "v1.2.2", "v1.2.1"])
def test_not_newer(respond, tag):
    respond(release(tag))
    assert updater.check() is None


@pytest.mark.parametrize("response", [
    requests.ConnectionError("down"),
    FakeResponse(status_code=403),
    FakeResponse(payload={"message": "rate limited"}),
])
def test_failures_return_none(respond, response):
    respond(response)
    assert updater.check() is None
