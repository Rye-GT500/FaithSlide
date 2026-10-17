import logging

import main


def test_log_rotates(isolated_user_dir, monkeypatch):
    monkeypatch.setattr(main, "LOG_MAX_BYTES", 2000)
    root = logging.getLogger()
    saved = root.handlers[:]
    root.handlers.clear()
    try:
        main.setup_logging()
        for i in range(200):
            logging.info("測試記錄 %d %s", i, "x" * 50)
        for handler in root.handlers:
            handler.close()
    finally:
        root.handlers[:] = saved
    names = sorted(p.name for p in isolated_user_dir.iterdir())
    assert names == ["FaithSlide.log", "FaithSlide.log.1", "FaithSlide.log.2", "FaithSlide.log.3"]
    assert all(p.stat().st_size <= 2000 + 200 for p in isolated_user_dir.iterdir())
