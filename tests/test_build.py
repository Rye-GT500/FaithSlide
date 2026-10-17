import pytest

import build
from build import BuildError, Commit, bump, classify, release_notes, suggest_level


@pytest.mark.parametrize("subject, expected", [
    ("新增：查詢頁改為單一輸入框", Commit("新增", "minor", "查詢頁改為單一輸入框")),
    ("修正: 二位數標題被切開", Commit("修正", "patch", "二位數標題被切開")),
    ("重大：週報格式改變", Commit("重大變更", "major", "週報格式改變")),
    ("改善：錯誤訊息", Commit("改善", "patch", "錯誤訊息")),
    ("內部：重構", Commit("內部", "none", "重構")),
    ("偵測更新", Commit("修正", "patch", "偵測更新")),          # 沒有前綴
    ("新增功能說明", Commit("修正", "patch", "新增功能說明")),  # 前綴後沒有冒號不算
])
def test_classify(subject, expected):
    assert classify(subject) == expected


@pytest.mark.parametrize("subjects, level", [
    (["修正：a", "新增：b", "內部：c"], "minor"),
    (["修正：a", "重大：b"], "major"),
    (["內部：a", "內部：b"], "patch"),
    ([], "patch"),
    (["沒有前綴"], "patch"),
])
def test_suggest_level(subjects, level):
    assert suggest_level([classify(s) for s in subjects]) == level


@pytest.mark.parametrize("version, level, pre, expected", [
    ("1.2.2-beta", "minor", None, "1.3.0"),
    ("1.2.2-beta", "patch", None, "1.2.3"),
    ("1.2.2", "major", None, "2.0.0"),
    ("1.3.0", "patch", "beta", "1.3.1-beta"),
    ("1.3.1-beta", "release", None, "1.3.1"),
    ("1.3.0", "none", None, "1.3.0"),
])
def test_bump(version, level, pre, expected):
    assert bump(version, level, pre) == expected


@pytest.mark.parametrize("version, level", [("1.2", "patch"), ("1.2.2", "huge")])
def test_bump_invalid(version, level):
    with pytest.raises(BuildError):
        bump(version, level)


def test_replace_version():
    text = 'CURRENT_VERSION = "1.2.2-beta"\nGITHUB_API_URL = "https://api.github.com/repos/Rye-GT500/FaithSlide/releases/latest"\n'
    new = build.replace_version(text, "1.3.0")
    assert build.read_version(new) == "1.3.0"
    assert new.splitlines()[1] == text.splitlines()[1]
    assert build.github_repo(new) == "Rye-GT500/FaithSlide"


def test_actual_version_file_is_readable():
    text = build.VERSION_FILE.read_text(encoding="utf-8")
    assert build.read_version(text)
    assert build.github_repo(text)


def test_release_notes():
    commits = [classify(s) for s in ["新增：查詢單一輸入框", "內部：重構", "修正：標題切開", "新增：書卷別名"]]
    assert release_notes(commits) == "## 新增\n- 查詢單一輸入框\n- 書卷別名\n\n## 修正\n- 標題切開"
    assert "內部調整" in release_notes([classify("內部：重構")])
