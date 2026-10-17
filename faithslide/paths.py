import os
import sys

FROZEN = getattr(sys, 'frozen', False)

if FROZEN:
    # PyInstaller 打包：資源解壓在 _MEIPASS，exe 所在資料夾另計
    BUNDLE_DIR = sys._MEIPASS
    APP_DIR = os.path.dirname(sys.executable)
else:
    BUNDLE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    APP_DIR = BUNDLE_DIR

TEMPLATE_PPTX = os.path.join(BUNDLE_DIR, "template.pptx")
DATA_DIR = os.path.join(BUNDLE_DIR, "faithslide", "data")
LOG_PATH = os.path.join(APP_DIR, "FaithSlide.log")


def user_dir() -> str:
    """使用者設定資料夾（%APPDATA%\\FaithSlide），不存在時建立。"""
    path = os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")), "FaithSlide")
    os.makedirs(path, exist_ok=True)
    return path
