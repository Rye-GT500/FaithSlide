import logging
from logging.handlers import RotatingFileHandler

from faithslide import paths

LOG_MAX_BYTES = 1_000_000   # 超過約 1MB 換新檔
LOG_BACKUP_COUNT = 3        # 保留 FaithSlide.log.1～.3


def setup_logging():
    handler = RotatingFileHandler(paths.log_path(), maxBytes=LOG_MAX_BYTES, backupCount=LOG_BACKUP_COUNT,
                                  encoding="utf-8")
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[handler],
    )


if __name__ == "__main__":
    setup_logging()
    from faithslide.ui.main_window import MainWindow
    MainWindow().run()
