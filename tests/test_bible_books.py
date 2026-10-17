import json
import re

import pytest

from faithslide import bible_books
from faithslide.bible_books import Book, BookDataError, find_conflicts

# 重構前 FaithSlide.py 的 books 字串，name_pattern 必須涵蓋完全相同的名稱
LEGACY_BOOKS = "創|出|利|民|申|書|士|得|撒上|撒下|王上|王下|代上|代下|拉|尼|斯|伯|詩|箴|傳|歌|賽|耶|哀|結|但|何|珥|摩|俄|拿|彌|鴻|哈|番|該|瑪|亞|太|可|路|約|徒|羅|林前|林後|加|弗|腓|西|帖前|帖後|提前|提後|多|門|來|雅|各|彼前|彼後|約壹|約貳|約參|猶|啟|約一|約二|約三"


@pytest.fixture(autouse=True)
def fresh_data():
    bible_books.load()


def test_all_books_in_canonical_order():
    books = bible_books.all_books()
    assert len(books) == 66
    assert [books[0].full, books[38].full, books[39].full, books[-1].full] == ["創世記", "瑪拉基書", "馬太福音", "啟示錄"]


@pytest.mark.parametrize("name, book_id", [
    ("約翰福音", "jhn"),
    ("約", "jhn"),
    ("約壹", "1jn"),
    ("約一", "1jn"),
    ("約翰一書", "1jn"),
    ("各", "jas"),
    (" 羅 ", "rom"),
])
def test_resolve(name, book_id):
    assert bible_books.resolve(name).id == book_id


def test_resolve_unknown():
    assert bible_books.resolve("約X") is None
    assert bible_books.resolve("") is None


@pytest.mark.parametrize("name, api", [("各", "雅"), ("約壹", "約一"), ("約貳", "約二"), ("約參", "約三")])
def test_alias_maps_to_api_value(name, api):
    # 取代舊版 get_verses 裡寫死的轉換
    assert bible_books.resolve(name).api == api


def test_name_pattern_matches_legacy_names():
    names = set(bible_books.name_pattern().split("|"))
    assert names == set(LEGACY_BOOKS.split("|"))


def test_name_pattern_prefers_longer_names():
    pattern = bible_books.name_pattern()
    assert re.match(pattern, "約壹3:1").group() == "約壹"
    assert re.match(pattern, "林前13:4").group() == "林前"
    assert re.match(pattern, "約3:16").group() == "約"


def test_find_conflicts():
    books = [
        Book("jhn", "約", "約翰福音", "約"),
        Book("1jn", "約一", "約翰一書", "約一", ("約",)),
    ]
    conflicts = find_conflicts(books)
    assert [(c.name, c.book_ids) for c in conflicts] == [("約", ("jhn", "1jn"))]


def test_load_rejects_conflicting_default_data(tmp_path):
    path = tmp_path / "books.json"
    path.write_text(json.dumps({"schema_version": 1, "books": [
        {"id": "jhn", "api": "約", "full": "約翰福音", "abbr": "約"},
        {"id": "1jn", "api": "約一", "full": "約翰一書", "abbr": "約一", "aliases": ["約"]},
    ]}, ensure_ascii=False), encoding="utf-8")
    with pytest.raises(BookDataError):
        bible_books.load(str(path))


def test_load_rejects_malformed_data(tmp_path):
    path = tmp_path / "books.json"
    path.write_text(json.dumps({"books": [{"id": "gen"}]}), encoding="utf-8")
    with pytest.raises(BookDataError):
        bible_books.load(str(path))


# --- 使用者自訂別名 ----------------------------------------------------------

def write_user_file(user_dir, aliases):
    path = user_dir / bible_books.USER_BOOKS_FILENAME
    path.write_text(json.dumps({"schema_version": 1, "aliases": aliases}, ensure_ascii=False), encoding="utf-8")
    return path


def test_user_aliases_applied(isolated_user_dir):
    write_user_file(isolated_user_dir, {"jas": {"add": ["雅各"], "remove": ["各"]}})
    bible_books.load()
    assert bible_books.resolve("雅各").id == "jas"
    assert bible_books.resolve("各") is None
    assert "雅各" in bible_books.name_pattern().split("|")
    assert bible_books.load_warnings() == []
    # 預設資料不受影響
    assert [b for b in bible_books.default_books() if b.id == "jas"][0].aliases == ("各",)


def test_user_alias_moved_between_books(isolated_user_dir):
    write_user_file(isolated_user_dir, {"jas": {"remove": ["各"]}, "jhn": {"add": ["各"]}})
    bible_books.load()
    assert bible_books.resolve("各").id == "jhn"


@pytest.mark.parametrize("aliases, reason", [
    ({"1jn": {"add": ["約"]}}, "約翰福音"),       # 與其他書卷簡寫衝突
    ({"jhn": {"add": ["約壹"]}}, "約翰一書"),     # 與其他書卷別名衝突
    ({"jhn": {"add": ["雅各書"]}}, "雅各書"),     # 與其他書卷全名衝突
    ({"jhn": {"add": ["約3"]}}, "數字"),
    ({"jhn": {"add": ["約 翰"]}}, "空白"),
    ({"jhn": {"add": [""]}}, "空白"),
    ({"jhn": {"add": ["約" * 11]}}, "超過"),
    ({"xyz": {"add": ["某"]}}, "未知的書卷"),
])
def test_invalid_user_alias_rejected_on_load(isolated_user_dir, aliases, reason):
    write_user_file(isolated_user_dir, aliases)
    bible_books.load()
    warnings = bible_books.load_warnings()
    assert len(warnings) == 1 and reason in warnings[0]
    # 其餘資料照常載入
    assert bible_books.resolve("約").id == "jhn"


def test_two_books_claiming_same_new_alias(isolated_user_dir):
    write_user_file(isolated_user_dir, {"jhn": {"add": ["新"]}, "rom": {"add": ["新"]}})
    bible_books.load()
    assert bible_books.resolve("新").id == "jhn"
    assert len(bible_books.load_warnings()) == 1


@pytest.mark.parametrize("content", ["not json", "[]", json.dumps({"aliases": {"jhn": ["x"]}})])
def test_corrupt_user_file_falls_back_to_defaults(isolated_user_dir, content):
    (isolated_user_dir / bible_books.USER_BOOKS_FILENAME).write_text(content, encoding="utf-8")
    bible_books.load()
    assert "讀取失敗" in bible_books.load_warnings()[0]
    assert bible_books.resolve("各").id == "jas"


def test_save_user_aliases_writes_diff_and_takes_effect(isolated_user_dir):
    bible_books.save_user_aliases({"jas": ["雅各"], "jhn": []})
    assert bible_books.resolve("雅各").id == "jas"
    saved = json.loads((isolated_user_dir / bible_books.USER_BOOKS_FILENAME).read_text(encoding="utf-8"))
    # 只記錄與預設不同之處；未變動的書卷不寫入
    assert saved == {"schema_version": 1, "aliases": {"jas": {"add": ["雅各"], "remove": ["各"]}}}
    bible_books.load()
    assert bible_books.resolve("雅各").id == "jas"


def test_save_rejected_does_not_write(isolated_user_dir):
    with pytest.raises(bible_books.AliasConflictError) as error:
        bible_books.save_user_aliases({"1jn": ["約壹", "約"]})
    assert error.value.rejected[0].name == "約"
    assert not (isolated_user_dir / bible_books.USER_BOOKS_FILENAME).exists()
    assert bible_books.resolve("約").id == "jhn"


def test_save_restoring_defaults_removes_entries(isolated_user_dir):
    bible_books.save_user_aliases({"jas": ["雅各"]})
    bible_books.save_user_aliases({"jas": ["各"]})
    saved = json.loads((isolated_user_dir / bible_books.USER_BOOKS_FILENAME).read_text(encoding="utf-8"))
    assert saved["aliases"] == {}


def test_check_aliases():
    assert bible_books.check_aliases({"jas": ["各", "雅各"]}) == []
    rejected = bible_books.check_aliases({"jas": ["各", "羅"]})
    assert [(r.book_id, r.name) for r in rejected] == [("jas", "羅")]
