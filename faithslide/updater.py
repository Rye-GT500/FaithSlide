"""檢查、下載、安裝新版本，不碰 UI。

安裝檔來自 GitHub Release；下載後比對 GitHub 提供的 SHA-256，一致才執行。
"""
import hashlib
import logging
import os
import re
from dataclasses import dataclass
from typing import Callable

import requests

from faithslide import paths
from faithslide.version import CURRENT_VERSION, GITHUB_API_URL

TIMEOUT_SECONDS = 5
DOWNLOAD_TIMEOUT_SECONDS = 30
DOWNLOAD_CHUNK_BYTES = 64 * 1024
UPDATES_DIRNAME = "updates"
_INSTALLER_NAME = re.compile(r"^FaithSlide-Setup-.+\.exe$")
# Inno Setup 靜默安裝：只顯示進度視窗，自動關閉執行中的程式，完成後重新開啟
INSTALLER_ARGUMENTS = "/SILENT /SUPPRESSMSGBOXES /NORESTART /CLOSEAPPLICATIONS"

logger = logging.getLogger(__name__)

_VERSION = re.compile(r"^v?(\d+(?:\.\d+)*)(?:-(\S+))?$")


@dataclass(frozen=True)
class UpdateInfo:
    version: str
    url: str                            # Release 頁面
    notes: str = ""
    installer_url: str | None = None
    installer_size: int = 0
    installer_sha256: str | None = None

    @property
    def can_install(self) -> bool:
        """有安裝檔與校驗碼才能自動更新；否則只能開啟下載頁面。"""
        return bool(self.installer_url and self.installer_sha256)


class UpdateError(Exception):
    """訊息可直接顯示給使用者。"""


def version_key(version: str) -> tuple:
    """可比較的版本鍵：數字逐段比較；數字相同時正式版大於預覽版（1.2.2 > 1.2.2-beta）。"""
    match = _VERSION.match(version.strip())
    if not match:
        raise ValueError(f"無法辨識的版本號：{version}")
    numbers = tuple(int(n) for n in match[1].split("."))
    numbers += (0,) * (3 - len(numbers))   # 1.2 視同 1.2.0
    prerelease = match[2]
    return numbers, prerelease is None, prerelease or ""


def check(raise_errors: bool = False) -> UpdateInfo | None:
    """有新版時回傳版本資訊；沒有新版回傳 None。

    檢查失敗時預設記錄 log 並回傳 None（啟動時的自動檢查）；raise_errors=True 時拋出 UpdateError（手動檢查）。
    """
    try:
        response = requests.get(GITHUB_API_URL, timeout=TIMEOUT_SECONDS)
        if response.status_code != 200:
            raise UpdateError(f"更新伺服器回應錯誤（HTTP {response.status_code}）")
        data = response.json()
        latest = data["tag_name"].strip().removeprefix("v")
        if version_key(latest) <= version_key(CURRENT_VERSION):
            return None
        return _update_info(latest, data)
    except (requests.RequestException, ValueError, KeyError, TypeError, AttributeError, UpdateError) as e:
        logger.warning(f"檢查更新失敗：{e}")
        if raise_errors:
            if isinstance(e, UpdateError):
                raise
            if isinstance(e, requests.RequestException):
                raise UpdateError("無法連線到更新伺服器，請檢查網路") from e
            raise UpdateError(f"無法辨識更新資訊：{e}") from e
    return None


def _update_info(version: str, data: dict) -> UpdateInfo:
    installer = next((a for a in data.get("assets", []) if _INSTALLER_NAME.match(a.get("name", ""))), None)
    sha256 = None
    if installer and str(installer.get("digest", "")).startswith("sha256:"):
        sha256 = installer["digest"].removeprefix("sha256:").lower()
    return UpdateInfo(
        version=version,
        url=data["html_url"],
        notes=data.get("body") or "",
        installer_url=installer["browser_download_url"] if installer else None,
        installer_size=installer.get("size", 0) if installer else 0,
        installer_sha256=sha256,
    )


def download(info: UpdateInfo, on_progress: Callable[[int, int], None] = lambda done, total: None) -> str:
    """下載安裝檔並比對 SHA-256，回傳檔案路徑。校驗不符時刪除檔案並拋出 UpdateError。"""
    if not info.can_install:
        raise UpdateError("這個版本沒有提供可自動安裝的檔案，請到下載頁面手動下載")
    folder = os.path.join(paths.user_dir(), UPDATES_DIRNAME)
    os.makedirs(folder, exist_ok=True)
    for old in os.listdir(folder):   # 清掉之前下載過的安裝檔
        try:
            os.remove(os.path.join(folder, old))
        except OSError:
            pass
    target = os.path.join(folder, f"FaithSlide-Setup-{info.version}.exe")
    partial = target + ".part"

    digest = hashlib.sha256()
    done = 0
    try:
        with requests.get(info.installer_url, stream=True, timeout=DOWNLOAD_TIMEOUT_SECONDS) as response:
            if response.status_code != 200:
                raise UpdateError(f"下載失敗（HTTP {response.status_code}）")
            total = int(response.headers.get("Content-Length") or info.installer_size)
            with open(partial, "wb") as f:
                for chunk in response.iter_content(DOWNLOAD_CHUNK_BYTES):
                    f.write(chunk)
                    digest.update(chunk)
                    done += len(chunk)
                    on_progress(done, total)
    except requests.RequestException as e:
        _remove(partial)
        raise UpdateError("下載中斷，請檢查網路後再試") from e
    except (OSError, UpdateError):
        _remove(partial)
        raise

    if digest.hexdigest() != info.installer_sha256:
        _remove(partial)
        logger.error(f"安裝檔校驗失敗：預期 {info.installer_sha256}，實際 {digest.hexdigest()}")
        raise UpdateError("下載的檔案校驗失敗，可能下載不完整，請再試一次")
    os.replace(partial, target)
    logger.info(f"已下載並校驗安裝檔：{target}")
    return target


def install(installer_path: str) -> None:
    """啟動安裝檔（Windows 會詢問系統管理員權限）；呼叫端隨後應關閉程式。"""
    logger.info(f"啟動安裝檔：{installer_path} {INSTALLER_ARGUMENTS}")
    os.startfile(installer_path, "open", INSTALLER_ARGUMENTS)


def _remove(path: str) -> None:
    try:
        os.remove(path)
    except OSError:
        pass


def plain_notes(markdown: str) -> str:
    """把 Release 說明的 Markdown 轉成適合直接顯示的文字。"""
    lines = []
    for line in markdown.replace("\r\n", "\n").split("\n"):
        line = line.replace("**", "").replace("`", "")
        stripped = line.lstrip()
        if stripped.startswith("#"):
            line = "【" + stripped.lstrip("#").strip() + "】"
        elif stripped.startswith(("- ", "* ")):
            indent = "　" * ((len(line) - len(stripped)) // 2)
            line = f"{indent}・{stripped[2:]}"
        lines.append(line)
    return "\n".join(lines).strip()
