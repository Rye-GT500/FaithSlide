import logging

from faithslide import paths


def setup_logging():
    logging.basicConfig(
        filename=paths.LOG_PATH,
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        encoding="utf-8"
    )


if __name__ == "__main__":
    setup_logging()
    from faithslide.ui.main_window import MainWindow
    MainWindow().run()
