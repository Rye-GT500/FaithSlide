"""打包與發版。用法見 docs/release.md。"""
import argparse
import os
import re
import shutil
import subprocess
import sys
import tempfile
from dataclasses import dataclass
from pathlib import Path

ROOT = Path(__file__).resolve().parent
VERSION_FILE = ROOT / "faithslide" / "version.py"
INSTALLER_SCRIPT = ROOT / "FaithSlide_Installer.iss"
ISCC_CANDIDATES = [
    Path(r"C:\Program Files (x86)\Inno Setup 6\ISCC.exe"),
    Path(r"C:\Program Files\Inno Setup 6\ISCC.exe"),
]
GH_CANDIDATES = [Path(r"C:\Program Files\GitHub CLI\gh.exe")]
APP_NAME = "FaithSlide"

LEVELS = ["none", "patch", "minor", "major"]
# 前綴 → (分類標題, 建議升級幅度)
PREFIXES = {
    "重大": ("重大變更", "major"),
    "新增": ("新增", "minor"),
    "修正": ("修正", "patch"),
    "改善": ("改善", "patch"),
    "內部": ("內部", "none"),
}
UNPREFIXED = ("修正", "patch")
_PREFIX = re.compile(rf"^({'|'.join(PREFIXES)})\s*[:：]\s*")
_VERSION = re.compile(r"^(\d+)\.(\d+)\.(\d+)(?:-(\S+))?$")
_VERSION_LINE = re.compile(r'^CURRENT_VERSION = "([^"]+)"$', re.M)


class BuildError(Exception):
    pass


# --- 版本與 commit 分類（純函式） ------------------------------------------

@dataclass(frozen=True)
class Commit:
    category: str
    level: str
    message: str      # 去掉前綴後的訊息


def classify(subject: str) -> Commit:
    match = _PREFIX.match(subject)
    if match:
        category, level = PREFIXES[match[1]]
        return Commit(category, level, subject[match.end():].strip())
    return Commit(*UNPREFIXED, subject.strip())


def suggest_level(commits: list[Commit]) -> str:
    """取最大的升級幅度；只有內部變更時仍建議 patch。"""
    level = max((c.level for c in commits), key=LEVELS.index, default="none")
    return "patch" if level == "none" else level


def bump(version: str, level: str, pre: str | None = None) -> str:
    """level：major / minor / patch / release（只去掉預覽標記）/ none。"""
    match = _VERSION.match(version)
    if not match:
        raise BuildError(f"無法辨識的版本號：{version}")
    major, minor, patch = (int(n) for n in match.groups()[:3])
    if level == "major":
        major, minor, patch = major + 1, 0, 0
    elif level == "minor":
        minor, patch = minor + 1, 0
    elif level == "patch":
        patch += 1
    elif level == "none":
        return version
    elif level != "release":
        raise BuildError(f"未知的升級幅度：{level}")
    return f"{major}.{minor}.{patch}" + (f"-{pre}" if pre else "")


def read_version(text: str) -> str:
    match = _VERSION_LINE.search(text)
    if not match:
        raise BuildError("version.py 找不到 CURRENT_VERSION")
    return match[1]


def replace_version(text: str, version: str) -> str:
    read_version(text)
    return _VERSION_LINE.sub(f'CURRENT_VERSION = "{version}"', text, count=1)


def release_notes(commits: list[Commit]) -> str:
    """依分類整理 Release 說明；內部變更不列出。"""
    sections = []
    for category, _ in [*PREFIXES.values()]:
        if category == "內部":
            continue
        items = [c.message for c in commits if c.category == category]
        if items:
            sections.append(f"## {category}\n" + "\n".join(f"- {m}" for m in items))
    return "\n\n".join(sections) or "內部調整，無使用者可見的變更。"


def github_repo(version_text: str) -> str:
    """從 GITHUB_API_URL 取出 owner/repo。"""
    match = re.search(r"api\.github\.com/repos/([^/]+/[^/]+)/releases", version_text)
    if not match:
        raise BuildError("version.py 找不到 GITHUB_API_URL")
    return match[1]


# --- 執行 --------------------------------------------------------------------

class Runner:
    def __init__(self, dry_run: bool):
        self.dry_run = dry_run

    def run(self, *cmd, side_effect=True):
        """side_effect=False 的指令（只讀取）在 dry-run 時照常執行。"""
        print("  $ " + " ".join(f'"{c}"' if " " in str(c) else str(c) for c in cmd))
        if self.dry_run and side_effect:
            return ""
        result = subprocess.run([str(c) for c in cmd], cwd=ROOT, capture_output=True, text=True,
                                encoding="utf-8", errors="replace")
        if result.returncode != 0:
            raise BuildError(f"指令失敗：{cmd[0]}\n{result.stdout}\n{result.stderr}")
        return result.stdout.strip()


def git(runner, *args, side_effect=True):
    return runner.run("git", *args, side_effect=side_effect)


def confirm(question: str, assume_yes: bool) -> bool:
    if assume_yes:
        return True
    return input(f"{question} [y/N] ").strip().lower() in ("y", "yes")


def find_tool(name: str, candidates: list[Path]) -> Path | None:
    """先找 PATH，再找預設安裝位置（剛安裝完時，已開啟的終端機 PATH 尚未更新）。"""
    found = shutil.which(name)
    if found:
        return Path(found)
    return next((path for path in candidates if path.exists()), None)


def find_iscc() -> Path:
    iscc = find_tool("ISCC", ISCC_CANDIDATES)
    if not iscc:
        raise BuildError("找不到 Inno Setup 的 ISCC.exe，請先安裝 Inno Setup 6")
    return iscc


def commits_since_last_tag(runner) -> tuple[str | None, list[Commit]]:
    try:
        last_tag = git(runner, "describe", "--tags", "--abbrev=0", "--match", "v*", side_effect=False)
    except BuildError:
        last_tag = None
    log_range = f"{last_tag}..HEAD" if last_tag else "HEAD"
    subjects = git(runner, "log", log_range, "--format=%s", side_effect=False).splitlines()
    return last_tag, [classify(s) for s in subjects if s]


def choose_version(current: str, commits: list[Commit], args) -> str:
    if args.no_bump or args.no_git:
        return current
    if args.bump:
        return bump(current, args.bump, args.pre)
    suggested = suggest_level(commits)
    proposal = bump(current, suggested, args.pre)
    print(f"\n建議升級：{suggested}（{current} → {proposal}）")
    if args.yes:
        return proposal
    answer = input("Enter 確認，或輸入 major / minor / patch / release：").strip().lower()
    return bump(current, answer, args.pre) if answer else proposal


def build_executable(runner):
    sep = os.pathsep
    runner.run(
        sys.executable, "-m", "PyInstaller", "--noconfirm", "--onefile", "--noconsole",
        "--name", APP_NAME,
        "--icon", ROOT / "icon.ico",
        "--add-data", f"{ROOT / 'template.pptx'}{sep}.",
        "--add-data", f"{ROOT / 'faithslide' / 'data'}{sep}faithslide/data",
        "--specpath", ROOT / "build",
        "--distpath", ROOT / "dist",
        "--workpath", ROOT / "build",
        ROOT / "main.py",
    )
    return ROOT / "dist" / f"{APP_NAME}.exe"


def build_installer(runner, version: str) -> Path:
    runner.run(find_iscc(), f"/DAppVersion={version}", INSTALLER_SCRIPT)
    return ROOT / "Output" / f"{APP_NAME}-Setup-{version}.exe"


def main(argv=None):
    parser = argparse.ArgumentParser(description="FaithSlide 打包與發版")
    parser.add_argument("--bump", choices=["major", "minor", "patch", "release"])
    parser.add_argument("--pre", help="預覽標記，例如 beta")
    parser.add_argument("--no-bump", action="store_true", help="不升版")
    parser.add_argument("--notes", type=Path, help="Release 說明檔（Markdown），取代依 commit 產生的說明")
    parser.add_argument("--no-release", action="store_true", help="不 push、不建立 GitHub Release")
    parser.add_argument("--no-git", action="store_true", help="只打包，不改版本、不動 git")
    parser.add_argument("--allow-dirty", action="store_true", help="允許有未 commit 的變更")
    parser.add_argument("--dry-run", action="store_true", help="只顯示會做什麼")
    parser.add_argument("-y", "--yes", action="store_true", help="所有確認都回答是")
    args = parser.parse_args(argv)
    runner = Runner(args.dry_run)

    version_text = VERSION_FILE.read_text(encoding="utf-8")
    current = read_version(version_text)
    release = not (args.no_release or args.no_git)
    if args.notes and not args.notes.exists():
        raise BuildError(f"找不到 Release 說明檔：{args.notes}")

    print("1. 檢查環境")
    if not args.allow_dirty and git(runner, "status", "--porcelain", side_effect=False):
        raise BuildError("工作目錄有未 commit 的變更（可用 --allow-dirty 略過）")
    find_iscc()
    if release:
        gh = find_tool("gh", GH_CANDIDATES)
        if not gh:
            raise BuildError("找不到 GitHub CLI（gh），請先安裝並執行 gh auth login，或改用 --no-release")
        repo = github_repo(version_text)
        runner.run(gh, "auth", "status", side_effect=False)
    if not args.no_git:
        git(runner, "fetch", "--tags", "--quiet", side_effect=False)

    print("\n2. 決定版本號")
    last_tag, commits = commits_since_last_tag(runner)
    print(f"上次發版：{last_tag or '（無）'}，之後有 {len(commits)} 個 commit")
    for c in commits:
        print(f"  [{c.category}] {c.message}")
    version = choose_version(current, commits, args)
    tag = f"v{version}"
    if not args.no_git and git(runner, "tag", "--list", tag, side_effect=False):
        raise BuildError(f"tag {tag} 已存在；只想重新打包請用 --no-git")
    print(f"版本：{current} → {version}")

    print("\n3. 執行測試")
    runner.run(sys.executable, "-m", "pytest", "-q", side_effect=False)

    print("\n4. 打包")
    if version != current and not args.dry_run:
        VERSION_FILE.write_text(replace_version(version_text, version), encoding="utf-8")
    try:
        exe = build_executable(runner)
        installer = build_installer(runner, version)
    except BaseException:
        VERSION_FILE.write_text(version_text, encoding="utf-8")   # 打包失敗時還原版本號
        raise
    print(f"執行檔：{exe}\n安裝檔：{installer}")

    if args.no_git:
        return
    print("\n5. commit 與 tag")
    if version != current:
        git(runner, "add", VERSION_FILE)
        git(runner, "commit", "-m", f"內部：版本 {version}")
    git(runner, "tag", "-a", tag, "-m", f"{APP_NAME} {version}")

    if not release:
        print(f"\n完成。尚未 push；發版時執行 git push origin HEAD {tag}")
        return
    print("\n6. 發佈 GitHub Release")
    if not confirm(f"即將 push 目前分支與 {tag}，並在 {repo} 建立 Release，確定？", args.yes):
        print(f"已略過。之後可執行 git push origin HEAD {tag}，再手動建立 Release")
        return
    git(runner, "push", "origin", "HEAD")
    git(runner, "push", "origin", tag)
    notes_text = args.notes.read_text(encoding="utf-8") if args.notes else release_notes(commits)
    print(f"Release 說明：\n{notes_text}\n")
    with tempfile.NamedTemporaryFile("w", suffix=".md", delete=False, encoding="utf-8") as notes:
        notes.write(notes_text)
    try:
        gh_args = [gh, "release", "create", tag, installer, "--repo", repo,
                   "--title", f"{APP_NAME} {version}", "--notes-file", notes.name]
        if "-" in version:
            gh_args.append("--prerelease")
        runner.run(*gh_args)
    finally:
        os.unlink(notes.name)
    print("\n完成。")


if __name__ == "__main__":
    try:
        main()
    except BuildError as e:
        print(f"\n錯誤：{e}", file=sys.stderr)
        sys.exit(1)
    except KeyboardInterrupt:
        print("\n已中止", file=sys.stderr)
        sys.exit(1)
