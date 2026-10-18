"""書卷資料：載入 books.default.json 與使用者別名，提供名稱解析與書名 pattern。

名稱分兩類：
- 全名（full）：只用於顯示與 resolve。
- 簡稱（abbr + aliases）：resolve 之外，也用於從文字中辨識經文引用（name_pattern）。

使用者別名存在 %APPDATA%\\FaithSlide\\books.user.json，只記錄與預設的差異：
    {"schema_version": 1, "aliases": {"jas": {"add": ["雅各"], "remove": ["各"]}}}
與其他書卷名稱衝突、或格式不合法的別名一律拒絕。
"""
import json
import logging
import os
import re
from dataclasses import dataclass, replace

from faithslide import paths

DEFAULT_BOOKS_PATH = os.path.join(paths.DATA_DIR, "books.default.json")
USER_BOOKS_FILENAME = "books.user.json"
SCHEMA_VERSION = 1
MAX_ALIAS_LENGTH = 10

logger = logging.getLogger(__name__)

# 數字、空白與引用格式用到的符號會破壞經文引用解析
_FORBIDDEN_IN_ALIAS = re.compile(r"[\d\s:：,，\-－~～;；.。()（）]")


@dataclass(frozen=True)
class Book:
    id: str
    api: str        # 送給 fhl API 的書卷值
    full: str
    abbr: str
    aliases: tuple[str, ...] = ()

    @property
    def short_names(self) -> tuple[str, ...]:
        return (self.abbr, *self.aliases)

    @property
    def names(self) -> tuple[str, ...]:
        return (self.full, *self.short_names)


@dataclass(frozen=True)
class AliasConflict:
    name: str
    book_ids: tuple[str, ...]


@dataclass(frozen=True)
class RejectedAlias:
    book_id: str
    name: str
    reason: str

    def __str__(self):
        return f"「{self.name}」：{self.reason}"


class BookDataError(ValueError):
    pass


class AliasConflictError(ValueError):
    def __init__(self, rejected: list[RejectedAlias]):
        self.rejected = rejected
        super().__init__("；".join(str(r) for r in rejected))


_default_books: list[Book] | None = None
_books: list[Book] | None = None
_index: dict[str, Book] = {}
_user_path: str | None = None
_load_warnings: list[str] = []


# --- 載入 --------------------------------------------------------------------

def find_conflicts(books: list[Book]) -> list[AliasConflict]:
    """找出對應到一卷以上書卷的名稱。"""
    owners: dict[str, list[str]] = {}
    for book in books:
        for name in dict.fromkeys(book.names):   # 同一卷書內重複的名稱不算衝突
            owners.setdefault(name, []).append(book.id)
    return [AliasConflict(name, tuple(ids)) for name, ids in owners.items() if len(ids) > 1]


def _parse_books(raw: dict) -> list[Book]:
    try:
        books = [Book(b["id"], b["api"], b["full"], b["abbr"], tuple(b.get("aliases", []))) for b in raw["books"]]
    except (KeyError, TypeError) as e:
        raise BookDataError(f"書卷資料格式錯誤：{e}") from e
    ids = [b.id for b in books]
    if len(ids) != len(set(ids)):
        raise BookDataError("書卷 id 重複")
    return books


def _user_books_path() -> str:
    return os.path.join(paths.user_dir(), USER_BOOKS_FILENAME)


def load(default_path: str = DEFAULT_BOOKS_PATH, user_path: str | None = None) -> None:
    global _default_books, _user_path, _load_warnings
    with open(default_path, encoding="utf-8") as f:
        defaults = _parse_books(json.load(f))
    conflicts = find_conflicts(defaults)
    if conflicts:
        # 內建資料出現衝突屬於程式錯誤，直接中止
        raise BookDataError(f"內建書卷名稱衝突：{conflicts}")

    _default_books = defaults
    _user_path = user_path or _user_books_path()
    _load_warnings = []
    changes = _read_user_changes(_user_path)
    books, rejected = _apply_changes(defaults, changes)
    for r in rejected:
        _warn_on_load(f"已忽略自訂別名{r}")
    _set_books(books)


def reload() -> None:
    load(user_path=_user_path)


def _read_user_changes(path: str) -> dict:
    if not os.path.exists(path):
        return {}
    try:
        with open(path, encoding="utf-8") as f:
            raw = json.load(f)
        changes = raw["aliases"]
        if not isinstance(changes, dict) or not all(isinstance(c, dict) for c in changes.values()):
            raise ValueError("aliases 格式錯誤")
        return changes
    except (OSError, ValueError, KeyError, TypeError) as e:
        _warn_on_load(f"自訂書卷別名檔讀取失敗，已改用預設值：{e}")
        return {}


def _warn_on_load(message: str) -> None:
    logger.warning(message)
    _load_warnings.append(message)


def _set_books(books: list[Book]) -> None:
    global _books, _index
    _books = books
    _index = {name: book for book in books for name in book.names}


def _ensure_loaded() -> None:
    if _books is None:
        load()


def load_warnings() -> list[str]:
    """載入時被忽略的自訂別名或檔案問題，供 UI 通知使用者。"""
    _ensure_loaded()
    return list(_load_warnings)


# --- 自訂別名 ----------------------------------------------------------------

def alias_problem(name: str) -> str | None:
    """別名本身的格式問題；沒有問題回傳 None。"""
    if not name:
        return "別名不可為空白"
    if len(name) > MAX_ALIAS_LENGTH:
        return f"別名不可超過 {MAX_ALIAS_LENGTH} 個字"
    if _FORBIDDEN_IN_ALIAS.search(name):
        return "別名不可包含數字、空白或標點符號"
    return None


def _apply_changes(defaults: list[Book], changes: dict) -> tuple[list[Book], list[RejectedAlias]]:
    """把使用者差異套用到預設書卷上，回傳 (書卷, 被拒絕的別名)。"""
    by_id = {b.id: b for b in defaults}
    aliases = {b.id: list(b.aliases) for b in defaults}
    rejected = []

    # 先處理所有移除，讓「把別名從 A 書移到 B 書」可以一次完成
    for book_id, change in changes.items():
        for name in change.get("remove", []):
            if book_id in aliases and name in aliases[book_id]:
                aliases[book_id].remove(name)

    owners = {name: b.id for b in defaults for name in (b.full, b.abbr, *aliases[b.id])}
    for book_id, change in changes.items():
        for name in change.get("add", []):
            name = name.strip() if isinstance(name, str) else ""
            if book_id not in by_id:
                rejected.append(RejectedAlias(book_id, name, "未知的書卷"))
                continue
            problem = alias_problem(name)
            owner = owners.get(name)
            if problem:
                rejected.append(RejectedAlias(book_id, name, problem))
            elif owner == book_id:
                continue
            elif owner:
                rejected.append(RejectedAlias(book_id, name, f"已是「{by_id[owner].full}」的名稱"))
            else:
                aliases[book_id].append(name)
                owners[name] = book_id

    return [replace(b, aliases=tuple(aliases[b.id])) for b in defaults], rejected


def _diff_from_defaults(aliases_by_book: dict[str, list[str]]) -> dict:
    changes = {}
    for book in default_books():
        if book.id not in aliases_by_book:
            continue
        desired = [a.strip() for a in aliases_by_book[book.id]]
        change = {
            "add": [a for a in dict.fromkeys(desired) if a not in book.aliases],
            "remove": [a for a in book.aliases if a not in desired],
        }
        if change["add"] or change["remove"]:
            changes[book.id] = change
    unknown = set(aliases_by_book) - {b.id for b in default_books()}
    for book_id in unknown:
        changes[book_id] = {"add": list(aliases_by_book[book_id]), "remove": []}
    return changes


def check_aliases(aliases_by_book: dict[str, list[str]]) -> list[RejectedAlias]:
    """檢查一組別名設定（書卷 id → 完整別名清單，未列出的書卷維持預設）。"""
    return _apply_changes(default_books(), _diff_from_defaults(aliases_by_book))[1]


def new_alias_problem(aliases_by_book: dict[str, list[str]], book_id: str, name: str) -> str | None:
    """在目前的別名設定（書卷 id → 完整別名清單）中替 book_id 新增 name 的問題；沒有問題回傳 None。"""
    problem = alias_problem(name)
    if problem:
        return problem
    candidate = {**aliases_by_book, book_id: [*aliases_by_book.get(book_id, []), name]}
    rejected = [r for r in check_aliases(candidate) if r.name == name]
    return rejected[0].reason if rejected else None


def alias_risk(name: str) -> str | None:
    """格式正確但可能造成誤判的別名，回傳提醒文字，由使用者確認是否仍要新增；沒有風險回傳 None。"""
    if len(name) == 1:
        return f"單字別名「{name}」可能把證道標題中的文字誤判為書卷"
    return None


def save_user_aliases(aliases_by_book: dict[str, list[str]]) -> None:
    """儲存別名設定並立即生效；有任何別名被拒絕時拋出 AliasConflictError，不寫入。"""
    _ensure_loaded()
    changes = _diff_from_defaults(aliases_by_book)
    books, rejected = _apply_changes(default_books(), changes)
    if rejected:
        raise AliasConflictError(rejected)

    tmp_path = _user_path + ".tmp"
    with open(tmp_path, "w", encoding="utf-8") as f:
        json.dump({"schema_version": SCHEMA_VERSION, "aliases": changes}, f, ensure_ascii=False, indent=1)
    os.replace(tmp_path, _user_path)
    logger.info(f"已儲存自訂書卷別名：{changes}")
    _set_books(books)


# --- 查詢 --------------------------------------------------------------------

def default_books() -> list[Book]:
    """未套用使用者別名的內建書卷。"""
    _ensure_loaded()
    return list(_default_books)


def all_books() -> list[Book]:
    """依聖經書卷順序回傳（已套用使用者別名）。"""
    _ensure_loaded()
    return list(_books)


def resolve(name: str) -> Book | None:
    """接受全名、簡寫或別名。"""
    _ensure_loaded()
    return _index.get(name.strip())


def name_pattern() -> str:
    """所有簡稱組成的正則（不含全名、不含群組）。長的在前，避免「約壹」被「約」搶先比對。"""
    _ensure_loaded()
    names = {name for book in _books for name in book.short_names}
    return "|".join(re.escape(n) for n in sorted(names, key=lambda n: (-len(n), n)))
