"""檢查 GitHub 最新版本，不碰 UI。"""
import logging
import re
from dataclasses import dataclass

import requests

from faithslide.version import CURRENT_VERSION, GITHUB_API_URL

TIMEOUT_SECONDS = 5

logger = logging.getLogger(__name__)

_VERSION = re.compile(r"^v?(\d+(?:\.\d+)*)(?:-(\S+))?$")


@dataclass(frozen=True)
class UpdateInfo:
    version: str
    url: str


def version_key(version: str) -> tuple:
    """可比較的版本鍵：數字逐段比較；數字相同時正式版大於預覽版（1.2.2 > 1.2.2-beta）。"""
    match = _VERSION.match(version.strip())
    if not match:
        raise ValueError(f"無法辨識的版本號：{version}")
    numbers = tuple(int(n) for n in match[1].split("."))
    numbers += (0,) * (3 - len(numbers))   # 1.2 視同 1.2.0
    prerelease = match[2]
    return numbers, prerelease is None, prerelease or ""


def check() -> UpdateInfo | None:
    """有新版時回傳版本資訊；沒有新版或檢查失敗回傳 None。"""
    try:
        response = requests.get(GITHUB_API_URL, timeout=TIMEOUT_SECONDS)
        if response.status_code != 200:
            logger.warning(f"檢查更新失敗：HTTP {response.status_code}")
            return None
        data = response.json()
        latest = data["tag_name"].strip().removeprefix("v")
        if version_key(latest) > version_key(CURRENT_VERSION):
            return UpdateInfo(latest, data["html_url"])
    except (requests.RequestException, ValueError, KeyError, TypeError, AttributeError) as e:
        logger.warning(f"檢查更新失敗：{e}")
    return None
