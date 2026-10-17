import os
import sys

FROZEN = getattr(sys, 'frozen', False)

# PyInstaller 打包時資源解壓在 _MEIPASS
BUNDLE_DIR = sys._MEIPASS if FROZEN else os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

DATA_DIR = os.path.join(BUNDLE_DIR, "faithslide", "data")
TEMPLATE_PPTX = os.path.join(DATA_DIR, "template.pptx")
LOG_FILENAME = "FaithSlide.log"


def user_dir() -> str:
    """使用者資料夾（%APPDATA%\\FaithSlide），不存在時建立。設定檔與 log 都放在這裡。"""
    path = os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")), "FaithSlide")
    os.makedirs(path, exist_ok=True)
    return path


def log_path() -> str:
    # 不放在 exe 旁：安裝在 Program Files 時一般權限無法寫入
    return os.path.join(user_dir(), LOG_FILENAME)
