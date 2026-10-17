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
    # 過渡期：UI 仍在 FaithSlide.py，第 8 階段改為啟動 faithslide.ui
    import FaithSlide
    FaithSlide.run()
