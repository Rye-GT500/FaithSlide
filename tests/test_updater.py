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


def release(tag, assets=None, body="說明"):
    return FakeResponse(payload={"tag_name": tag, "html_url": "https://example.com/release", "body": body,
                                 "assets": assets if assets is not None else []})


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
    info = updater.check()
    assert (info.version, info.url, info.notes) == (expected, "https://example.com/release", "說明")
    assert not info.can_install          # 沒有安裝檔時只能開啟下載頁面


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


INSTALLER_BYTES = b"fake installer " * 1000
INSTALLER_SHA256 = __import__("hashlib").sha256(INSTALLER_BYTES).hexdigest()


def installer_asset(digest=f"sha256:{INSTALLER_SHA256}", name="FaithSlide-Setup-1.3.2.exe"):
    return {"name": name, "size": len(INSTALLER_BYTES), "digest": digest,
            "browser_download_url": f"https://example.com/{name}"}


def test_check_reads_installer_asset(respond):
    respond(release("v1.3.2", assets=[{"name": "source.zip"}, installer_asset()]))
    info = updater.check()
    assert info.installer_url == "https://example.com/FaithSlide-Setup-1.3.2.exe"
    assert info.installer_sha256 == INSTALLER_SHA256
    assert info.can_install


def test_check_without_digest_cannot_install(respond):
    respond(release("v1.3.2", assets=[installer_asset(digest=None)]))
    assert not updater.check().can_install


@pytest.mark.parametrize("response, message", [
    (requests.ConnectionError("down"), "無法連線"),
    (FakeResponse(status_code=403), "HTTP 403"),
    (release("nightly"), "無法辨識"),
])
def test_manual_check_raises(respond, response, message):
    respond(response)
    with pytest.raises(updater.UpdateError, match=message):
        updater.check(raise_errors=True)


class FakeStream:
    def __init__(self, data, status_code=200, fail_after=None):
        self.data, self.status_code, self.fail_after = data, status_code, fail_after
        self.headers = {"Content-Length": str(len(data))}

    def __enter__(self):
        return self

    def __exit__(self, *args):
        return False

    def iter_content(self, size):
        for i in range(0, len(self.data), size):
            if self.fail_after is not None and i >= self.fail_after:
                raise requests.ConnectionError("cut")
            yield self.data[i:i + size]


@pytest.fixture
def info():
    return updater.UpdateInfo("1.3.2", "https://example.com/release", "", "https://example.com/setup.exe",
                              len(INSTALLER_BYTES), INSTALLER_SHA256)


def test_download_verifies_and_saves(monkeypatch, isolated_user_dir, info):
    monkeypatch.setattr(updater.requests, "get", lambda *a, **k: FakeStream(INSTALLER_BYTES))
    monkeypatch.setattr(updater, "DOWNLOAD_CHUNK_BYTES", 4096)
    (isolated_user_dir / "updates").mkdir()
    (isolated_user_dir / "updates" / "FaithSlide-Setup-1.0.0.exe").write_bytes(b"old")
    progress = []
    path = updater.download(info, lambda done, total: progress.append((done, total)))
    assert open(path, "rb").read() == INSTALLER_BYTES
    assert progress[-1] == (len(INSTALLER_BYTES), len(INSTALLER_BYTES))
    # 舊的安裝檔與暫存檔都會清掉
    assert sorted(p.name for p in (isolated_user_dir / "updates").iterdir()) == ["FaithSlide-Setup-1.3.2.exe"]


def test_download_rejects_corrupted_file(monkeypatch, isolated_user_dir, info):
    monkeypatch.setattr(updater.requests, "get", lambda *a, **k: FakeStream(INSTALLER_BYTES[:-1] + b"X"))
    with pytest.raises(updater.UpdateError, match="校驗失敗"):
        updater.download(info)
    assert list((isolated_user_dir / "updates").iterdir()) == []


def test_download_interrupted(monkeypatch, isolated_user_dir, info):
    monkeypatch.setattr(updater.requests, "get", lambda *a, **k: FakeStream(INSTALLER_BYTES, fail_after=4096))
    monkeypatch.setattr(updater, "DOWNLOAD_CHUNK_BYTES", 4096)
    with pytest.raises(updater.UpdateError, match="下載中斷"):
        updater.download(info)
    assert list((isolated_user_dir / "updates").iterdir()) == []


def test_download_without_installer_refused(info):
    with pytest.raises(updater.UpdateError):
        updater.download(updater.UpdateInfo("1.3.2", "https://example.com/release"))


def test_install_starts_silent_installer(monkeypatch):
    calls = []
    monkeypatch.setattr(updater.os, "startfile", lambda *args: calls.append(args), raising=False)
    updater.install("C:/tmp/FaithSlide-Setup-1.3.2.exe")
    assert calls == [("C:/tmp/FaithSlide-Setup-1.3.2.exe", "open", "/SILENT /SUPPRESSMSGBOXES /NORESTART /CLOSEAPPLICATIONS")]


def test_plain_notes():
    from faithslide.updater import plain_notes
    markdown = "說明\r\n\r\n## 修正\r\n\r\n- **超出**：已修正\r\n  - 子項目 `code`\r\n"
    assert plain_notes(markdown) == "說明\n\n【修正】\n\n・超出：已修正\n　・子項目 code"
