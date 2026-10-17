# 版本號與發版流程

## 版本號規則

格式：`主版本.次版本.修補版本[-預覽標記]`，例如 `1.3.0`、`1.3.1-beta`。

FaithSlide 沒有其他程式依賴它，所以以「**使用者需不需要改變做法**」判斷升級幅度：

| 升級 | 判斷標準 | 例子 |
|---|---|---|
| 主版本 major | 使用者必須改變做法，舊的東西不能直接用 | 週報 Word 格式要求改變、`template.pptx` 不相容、設定檔格式改變且舊檔無法讀取 |
| 次版本 minor | 新增功能，舊用法照常可用 | 查詢頁單一輸入框、書卷別名設定 |
| 修補版本 patch | 修 bug、改善體驗，不改用法 | 標題解析錯誤、錯誤訊息改善 |

- 正式版不加標記；需要先給少數人試用的版本加 `-beta`（如 `1.3.1-beta`）。
- 版本比較：數字逐段比較；數字相同時正式版大於預覽版（`1.3.0` > `1.3.0-beta`）。
- 預覽版在 GitHub 上標為 Pre-release，程式內的「檢查更新」不會通知預覽版。

## commit 訊息前綴

commit 第一行以前綴開頭，`build.py` 依此建議升級幅度：

| 前綴 | 意義 | 建議升級 |
|---|---|---|
| `重大：` | 使用者必須改變做法 | major |
| `新增：` | 新功能 | minor |
| `修正：` | 修 bug | patch |
| `改善：` | 體驗、效能、訊息調整 | patch |
| `內部：` | 使用者看不到的變更（重構、測試、文件、打包） | 不升級 |

- 沒有前綴的 commit 視為「修正」。
- 建議取所有 commit 中最大的幅度；只有 `內部：` 時仍建議 patch（發版本身代表有新的安裝檔）。
- 例：`新增：查詢頁改為單一輸入框`、`修正：二位數標題編號被切開`。

## 發版

### 事前準備（只需一次）

1. 安裝 Inno Setup 6（預設路徑 `C:\Program Files (x86)\Inno Setup 6\`）。
2. 安裝 GitHub CLI 並登入：
   ```
   winget install GitHub.cli
   gh auth login
   ```

### 執行

```
python build.py                 # 依 commit 前綴建議升級幅度，確認後執行
python build.py --bump minor    # 直接指定 major / minor / patch / release（只去掉預覽標記）
python build.py --pre beta      # 新版本加上 -beta
python build.py --no-bump       # 不升版，用目前版本重新打包
python build.py --dry-run       # 只顯示會做什麼，不執行
python build.py --no-release    # 打包、commit、tag，但不 push、不建立 GitHub Release
python build.py --notes 說明.md  # 用自己寫的 Release 說明（例如 commit 沒有前綴時）
python build.py --no-git        # 只打包（本機測試用），不改版本、不動 git
```

流程：

1. 檢查 git 工作目錄乾淨、抓取遠端 tag。
2. 列出上次 tag 之後的 commit 與建議升級幅度，確認新版本號。
3. 執行 pytest，失敗則中止。
4. 改寫 `faithslide/version.py`。
5. PyInstaller 打包 → `dist/FaithSlide.exe`。
6. Inno Setup 產生安裝檔 → `Output/FaithSlide-Setup-<版本>.exe`。
7. commit `faithslide/version.py`，建立 tag `v<版本>`。
8. 再次確認後 push 目前分支與 tag，以 `gh` 建立 GitHub Release 並上傳安裝檔；Release 說明由 commit 依前綴分類產生。

第 5、6 步失敗時，`version.py` 會還原。
