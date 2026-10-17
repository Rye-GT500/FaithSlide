"""設定：內建 settings.default.json，使用者的 settings.json 逐項覆寫。

| 鍵 | 說明 |
|---|---|
| speaker_words | 證道文字中要移除的講員相關字詞（職稱、姓名） |

使用者設定檔只記錄與預設不同的項目。
"""
import copy
import json
import logging
import os

from faithslide import paths

DEFAULT_SETTINGS_PATH = os.path.join(paths.DATA_DIR, "settings.default.json")
USER_SETTINGS_FILENAME = "settings.json"
SCHEMA_VERSION = 1

logger = logging.getLogger(__name__)

_defaults: dict | None = None
_settings: dict | None = None
_user_path: str | None = None


def _user_settings_path() -> str:
    return os.path.join(paths.user_dir(), USER_SETTINGS_FILENAME)


def load(default_path: str = DEFAULT_SETTINGS_PATH, user_path: str | None = None) -> None:
    global _defaults, _settings, _user_path
    with open(default_path, encoding="utf-8") as f:
        defaults = json.load(f)
    defaults.pop("schema_version", None)
    settings = copy.deepcopy(defaults)

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
    _defaults, _settings, _user_path = defaults, settings, user_path


def _ensure_loaded() -> None:
    if _settings is None:
        load()


def get(key: str):
    """回傳副本，修改不會影響設定。"""
    _ensure_loaded()
    return copy.deepcopy(_settings[key])


def default(key: str):
    _ensure_loaded()
    return copy.deepcopy(_defaults[key])


def save(key: str, value) -> None:
    """儲存一項設定並立即生效；與預設相同時從使用者設定檔移除。"""
    _ensure_loaded()
    if key not in _defaults:
        raise KeyError(f"未知的設定：{key}")
    if type(value) is not type(_defaults[key]):
        raise TypeError(f"設定 {key} 的型別錯誤")

    settings = {**_settings, key: copy.deepcopy(value)}
    overrides = {k: v for k, v in settings.items() if v != _defaults[k]}
    tmp_path = _user_path + ".tmp"
    with open(tmp_path, "w", encoding="utf-8") as f:
        json.dump({"schema_version": SCHEMA_VERSION, **overrides}, f, ensure_ascii=False, indent=1)
    os.replace(tmp_path, _user_path)
    _settings[key] = copy.deepcopy(value)
    logger.info(f"已儲存設定 {key}")
