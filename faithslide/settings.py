"""設定：內建 settings.default.json，使用者的 settings.json 逐項覆寫。

| 鍵 | 說明 |
|---|---|
| speaker_words | 證道文字中要移除的講員相關字詞（職稱、姓名） |
"""
import json
import logging
import os

from faithslide import paths

DEFAULT_SETTINGS_PATH = os.path.join(paths.DATA_DIR, "settings.default.json")
USER_SETTINGS_FILENAME = "settings.json"

logger = logging.getLogger(__name__)

_settings: dict | None = None


def _user_settings_path() -> str:
    return os.path.join(paths.user_dir(), USER_SETTINGS_FILENAME)


def load(default_path: str = DEFAULT_SETTINGS_PATH, user_path: str | None = None) -> None:
    global _settings
    with open(default_path, encoding="utf-8") as f:
        settings = json.load(f)

    user_path = user_path or _user_settings_path()
    if os.path.exists(user_path):
        try:
            with open(user_path, encoding="utf-8") as f:
                user = json.load(f)
            if not isinstance(user, dict):
                raise ValueError("最外層必須是物件")
        except (OSError, ValueError) as e:
            # 使用者檔損毀時只用預設值，程式照常啟動
            logger.warning(f"使用者設定檔讀取失敗，改用預設值：{user_path} {e}")
        else:
            for key, value in user.items():
                if key == "schema_version":
                    continue
                if key not in settings:
                    logger.warning(f"忽略未知的使用者設定：{key}")
                elif type(value) is not type(settings[key]):
                    logger.warning(f"忽略型別錯誤的使用者設定：{key}")
                else:
                    settings[key] = value
    _settings = settings


def get(key: str):
    if _settings is None:
        load()
    return _settings[key]
