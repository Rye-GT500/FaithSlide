"""檢查 GitHub 最新版本，不碰 UI。"""
import logging
from dataclasses import dataclass

import requests

from faithslide.version import CURRENT_VERSION, GITHUB_API_URL

TIMEOUT_SECONDS = 5

logger = logging.getLogger(__name__)


@dataclass(frozen=True)
class UpdateInfo:
    version: str
    url: str


def check() -> UpdateInfo | None:
    """有新版時回傳版本資訊；沒有新版或檢查失敗回傳 None。"""
    try:
        response = requests.get(GITHUB_API_URL, timeout=TIMEOUT_SECONDS)
        if response.status_code != 200:
            logger.warning(f"檢查更新失敗：HTTP {response.status_code}")
            return None
        data = response.json()
        latest = data["tag_name"].replace("v", "").replace("-beta", "")   # tag 格式如 v1.2.2-beta
        # TODO(9a)：目前沿用舊版的字串比較，1.10.0 會被判定小於 1.9.0
        if latest > CURRENT_VERSION.replace("-beta", ""):
            return UpdateInfo(latest, data["html_url"])
    except (requests.RequestException, ValueError, KeyError, TypeError) as e:
        logger.warning(f"檢查更新失敗：{e}")
    return None
