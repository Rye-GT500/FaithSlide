"""UI 與核心的邊界：UI 只能透過下列公開介面使用核心，核心不可依賴 UI。

UI 改造的規則見 docs/ui-guide.md。這個檔案屬於核心，UI 改造時不可修改；
需要新的核心介面或其他套件時，先與專案負責人討論。
"""
import ast
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
UI_DIR = ROOT / "faithslide" / "ui"

# UI 可使用的核心介面（模組 → 名稱）
UI_ALLOWED_CORE_API = {
    "background": {"run_in_background"},
    "bible_books": {"AliasConflictError", "Book", "alias_risk", "all_books", "default_books",
                    "load_warnings", "new_alias_problem", "save_user_aliases"},
    "paths": {"log_path", "user_dir"},
    "pipeline": {"GenerationError", "generate"},
    "reference_parser": {"ReferenceParseError"},
    "sermon_parser": {"speaker_word_problem"},
    "settings": {"SPEAKER_WORDS", "default", "get", "save"},
    "updater": {"UpdateError", "UpdateInfo", "check", "download", "install", "plain_notes"},
    "verse_search": {"lookup"},
    "version": {"CURRENT_VERSION"},
}
# UI 可使用的套件：Python 標準函式庫（含 tkinter）與本專案
UI_ALLOWED_PACKAGES = set(sys.stdlib_module_names) | {"faithslide"}


def _ui_files():
    return sorted(UI_DIR.rglob("*.py"))


def _core_files():
    return [p for p in sorted((ROOT / "faithslide").rglob("*.py")) if UI_DIR not in p.parents] + [ROOT / "build.py"]


def _imports(tree):
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            yield from (alias.name for alias in node.names)
        elif isinstance(node, ast.ImportFrom) and node.module:
            yield node.module


def _core_api_used(tree):
    """回傳 UI 程式使用的 (核心模組, 名稱)。"""
    module_aliases = {}
    for node in ast.walk(tree):
        if isinstance(node, ast.ImportFrom) and node.module == "faithslide":
            for alias in node.names:
                if alias.name != "ui":
                    module_aliases[alias.asname or alias.name] = alias.name
        elif isinstance(node, ast.ImportFrom) and node.module and node.module.startswith("faithslide."):
            module = node.module.split(".")[1]
            if module != "ui":
                yield from ((module, alias.name) for alias in node.names)
        elif isinstance(node, ast.Import):
            for alias in node.names:
                if alias.name.startswith("faithslide.") and not alias.name.startswith("faithslide.ui"):
                    yield alias.name.split(".")[1], "*"   # import faithslide.xxx：無法逐一檢查，一律不允許
    for node in ast.walk(tree):
        if isinstance(node, ast.Attribute) and isinstance(node.value, ast.Name) and node.value.id in module_aliases:
            yield module_aliases[node.value.id], node.attr


@pytest.mark.parametrize("path", _core_files(), ids=lambda p: str(p.relative_to(ROOT)))
def test_core_does_not_depend_on_ui(path):
    imports = list(_imports(ast.parse(path.read_text(encoding="utf-8"))))
    assert not [m for m in imports if m.split(".")[0] == "tkinter" or m.startswith("faithslide.ui")], \
        "核心模組不可 import tkinter 或 faithslide.ui"


@pytest.mark.parametrize("path", _ui_files(), ids=lambda p: str(p.relative_to(ROOT)))
def test_ui_uses_only_public_core_api(path):
    used = set(_core_api_used(ast.parse(path.read_text(encoding="utf-8"))))
    not_allowed = sorted(f"{module}.{name}" for module, name in used
                         if name not in UI_ALLOWED_CORE_API.get(module, set()))
    assert not not_allowed, f"UI 使用了未開放的核心介面：{not_allowed}（需要時先與專案負責人討論）"


@pytest.mark.parametrize("path", _ui_files(), ids=lambda p: str(p.relative_to(ROOT)))
def test_ui_uses_only_allowed_packages(path):
    packages = {m.split(".")[0] for m in _imports(ast.parse(path.read_text(encoding="utf-8")))}
    not_allowed = sorted(packages - UI_ALLOWED_PACKAGES)
    assert not not_allowed, f"UI 使用了未允許的套件：{not_allowed}（更換或新增 UI 框架需先與專案負責人討論）"


def test_entry_point_exists():
    """main.py 以 MainWindow().run() 啟動程式，UI 必須保留這個入口。"""
    from faithslide.ui.main_window import MainWindow
    assert callable(getattr(MainWindow, "run", None))
