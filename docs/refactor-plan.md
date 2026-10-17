# FaithSlide SRP 重構企畫書

> 狀態：**已確認，實施中**（第 0～2 階段完成）
> 範圍：`FaithSlide.py`（967 行，單一檔案）→ `faithslide/` 套件
> 不在範圍內：`BibleDictionary.py`（由作者自行隔離）

---

## 1. 目標

1. 依「修改原因」拆分模組，每個模組只有一個職責。
2. 核心邏輯不依賴 tkinter，可單獨測試。
3. 去除全域狀態（`prs`、`main_book`、各種 `*_var`）。
4. 書卷資料改為 JSON，並開放使用者自訂常用簡稱。
5. 查詢頁支援直接輸入引用字串（如 `約3:16-18`）。

**原則：重構與行為變更分開提交。** 第 0～8 階段的輸出應與現行版本一致；行為變更集中在第 9 階段，各自獨立 commit。

---

## 2. 目錄結構

```
FaitSlide/
├─ main.py                     程式進入點（PyInstaller 打包目標）
├─ faithslide/
│  ├─ __init__.py
│  ├─ version.py               CURRENT_VERSION、GITHUB_API_URL（版本唯一來源）
│  ├─ paths.py                 資源路徑（_MEIPASS）、使用者資料夾（%APPDATA%\FaithSlide）、log 路徑
│  ├─ data/
│  │  ├─ books.default.json    內建書卷資料（唯讀，打包進程式）
│  │  └─ settings.default.json 內建設定（唯讀，打包進程式），如講員名稱清單
│  ├─ settings.py              設定載入：預設設定 + 使用者設定
│  ├─ bible_books.py           書卷資料載入、合併使用者別名、名稱解析、書名 pattern 生成
│  ├─ reference_parser.py      經文引用：尋找、解析、顯示格式（含中文數字）
│  ├─ bible_api.py             fhl API 查詢
│  ├─ word_reader.py           docx → 讀經行、證道粗體文字
│  ├─ sermon_parser.py         證道文字 → 大綱結構（大/中/小標題 + 引用）
│  ├─ ppt_builder.py           模板投影片複製/刪除、各類投影片、字體大小計算
│  ├─ pipeline.py              串接流程 generate(word_path, ppt_path, on_progress)
│  ├─ updater.py               檢查 GitHub 最新版本（不碰 UI）
│  └─ ui/
│     ├─ __init__.py
│     ├─ main_window.py        主視窗、頁面切換、快捷鍵、啟動更新檢查
│     ├─ slide_page.py         PPT 製作頁
│     └─ search_page.py        經文查詢頁
├─ tests/
│  ├─ fixtures/
│  │  ├─ api/                  錄製的 API 回應（離線測試用）
│  │  └─ snapshots/            重構前的輸出基準（JSON）
│  └─ test_*.py
└─ docs/
   └─ refactor-plan.md         本文件
```

使用者自訂檔（不在 repo 內）：`%APPDATA%\FaithSlide\books.user.json`、`%APPDATA%\FaithSlide\settings.json`

---

## 3. 依賴方向

```
ui ──► pipeline ──► word_reader
 │         │──────► sermon_parser ──► reference_parser ──► bible_books
 │         │──────► bible_api ─────────────────────────────► bible_books
 │         └──────► ppt_builder
 ├──► reference_parser / bible_api      （查詢頁）
 ├──► bible_books                       （設定、下拉選單）
 └──► updater ──► version
```

規則：

- 箭頭只能往右或往下，**核心模組不得 import `ui` 或 `tkinter`**。
- `version.py`、`paths.py`、`settings.py` 任何模組都可以引用。

---

## 4. 各模組職責與介面

### 4.1 `bible_books.py`

**資料格式（`books.default.json`）**

```json
{
  "schema_version": 1,
  "books": [
    { "id": "jas", "api": "雅",   "full": "雅各書",   "abbr": "雅",   "aliases": ["各"] },
    { "id": "1jn", "api": "約一", "full": "約翰一書", "abbr": "約一", "aliases": ["約壹"] }
  ]
}
```

| 欄位 | 用途 | 可自訂 |
|---|---|---|
| `id` | 穩定內部鍵 | 否 |
| `api` | 送給 fhl API 的值 | 否 |
| `full` / `abbr` | 顯示用 | 否 |
| `aliases` | 解析時可接受的其他寫法 | **是** |

陣列順序即聖經書卷順序。

**使用者檔（只存差異）**

```json
{ "schema_version": 1, "aliases": { "jas": { "add": ["雅各"], "remove": [] } } }
```

**介面**

```python
@dataclass(frozen=True)
class Book: id: str; api: str; full: str; abbr: str; aliases: tuple[str, ...]

def load() -> None                       # 讀預設檔 + 使用者檔，驗證後建立索引
def reload() -> None                     # 設定變更後呼叫
def all_books() -> list[Book]
def resolve(name: str) -> Book | None    # 接受 full / abbr / alias
def name_pattern() -> str                # 依長度遞減排序、re.escape 後以 | 串接
def save_user_aliases(changes) -> None   # 有衝突時拋出 AliasConflictError，不寫入
def rejected_aliases() -> list[AliasConflict]   # load 時被拒絕的使用者別名，供 UI 通知
```

**驗證規則**

- 同一名稱對應到多卷書：**直接拒絕並通知**。
  - 儲存時：拋出 `AliasConflictError`（列出衝突的別名與書卷），不寫入檔案，由設定 UI 顯示。
  - 載入時（例如使用者手動編輯了 JSON）：略過衝突的別名，其餘照常載入；UI 啟動後讀取 `rejected_aliases()` 顯示通知。
- 使用者檔損毀：記錄 log，只使用預設表，程式照常啟動。
- 常用字別名（如「主」、「神」）：由設定 UI 顯示警告，不在此層阻擋。

**取代的現有程式碼**：`abbr_to_full`、`full_to_abbr`、`books` 字串、`get_verses` 中寫死的 `各/約壹…` 轉換。

### 4.2 `reference_parser.py`

```python
@dataclass(frozen=True)
class Reference:
    book: Book
    chapter: int
    ranges: tuple[tuple[int, int], ...]   # 空 tuple = 整章

class ReferenceError(ValueError): ...      # 錯誤訊息供 UI 直接顯示

def find_references(text: str) -> list[str]
def parse(raw: str, default_book: Book | None = None) -> list[Reference]
def format_title(ref: Reference, verse_part: str) -> str   # 例：「約翰福音三章16-18節」；詩篇改為「篇」
def chinese_numeral(n: int) -> str
```

- **不支援跨章**（如 `約3:16-4:2`）：`parse` 拋出 `ReferenceError`。
- 一個字串可包含多章（如 `約3:16,4:2`），每章各回傳一個 `Reference`。
- `default_book` 用來取代全域變數 `main_book`。

**取代的現有程式碼**：`analyze_word` 中的引用正則、`parse_bible_reference`、`process_reference_block`、`analyze_paragraph` 的範圍解析、`num_to_chinese`、`chinese_number`、`number`。

### 4.3 `bible_api.py`

```python
class BibleApiError(Exception): ...

def get_chapter(book: Book, chapter: int) -> list[Verse]   # Verse(number: int, text: str)
```

- 失敗時拋出例外，**不再回傳空 list 掩蓋錯誤**。
- 可選：同一次執行內快取（同章只查一次）。

**取代的現有程式碼**：`get_verses`、`url`。

### 4.4 `word_reader.py`

```python
@dataclass
class WeeklyReport:
    reading_lines: list[str]   # 「讀經」列，依行切分
    sermon_text: str           # 「證道」列粗體文字，以空白串接

def read(path: str) -> WeeklyReport
```

只負責 docx I/O，不做任何內容判斷。

### 4.5 `sermon_parser.py`

```python
@dataclass
class Section:              # 大/中/小標題共用，以 children 表示階層
    number: str             # 「一、」「1.」「(1)」
    text: str
    references: list[Reference]
    children: list["Section"]

@dataclass
class Sermon:
    title: str
    sections: list[Section]

def parse_sermon(text: str, default_book: Book | None, speaker_words: list[str]) -> Sermon
def parse_reading(lines: list[str]) -> ...        # 讀經段落解析（主經文、書卷推斷）
```

- **標題切分正則**（`一、` / `1.` / `1)`）放在這裡。
- 要移除的講員字串（`證道`、`牧師`、人名…）**移到設定檔**（`settings.default.json` 的 `speaker_words`，使用者可在 `settings.json` 覆寫），由呼叫端傳入 `parse_sermon`，parser 本身不讀檔。

**取代的現有程式碼**：`analyze_word`、`produce_the_slides` 中的大綱建構與讀經解析段落。

### 4.6 `ppt_builder.py`

```python
class SlideBuilder:
    def __init__(self, template_path: str): ...
    def add_main_title(self, title: str): ...
    def add_major(self, section: Section): ...
    def add_medium(self, major: Section, current: Section): ...
    def add_minor(self, major: Section, medium: Section, current: Section): ...
    def add_verse(self, title: str, verse: Verse): ...
    def save(self, path: str): ...                 # 內含刪除 6 張模板投影片
```

- 模板投影片索引（0～5）改成具名常數。
- `calculate_font_size`、`get_weighted_length` 為模組內私有函式。

**取代的現有程式碼**：`duplicate_slide`、`remove_slide`、所有 `*_PPT` 函式、字體計算。

### 4.7 `pipeline.py`

```python
def generate(word_path: str, ppt_path: str,
             on_progress: Callable[[float, str], None] = lambda p, m: None) -> None
```

依序執行：讀 Word → 解析讀經/證道 → 查經文 → 建投影片 → 存檔。錯誤以例外向上拋出，由 UI 決定如何顯示。

**取代的現有程式碼**：`produce_the_slides`、`paragraph_PPT`。

### 4.8 `updater.py` / `version.py`

```python
@dataclass
class UpdateInfo: version: str; url: str

def check() -> UpdateInfo | None    # 失敗或無新版都回傳 None，並記錄 log
```

版本比較改為數值 tuple 比較。

### 4.9 `ui/`

- `main_window.py`：建立 `Tk`、頁面切換（F5）、退出（Esc）。啟動時以背景執行緒呼叫 `updater.check()`，再用 `root.after` 回到主執行緒顯示對話框。
- `slide_page.py`：檔案選擇、進度條。背景執行緒呼叫 `pipeline.generate`，`on_progress` 與錯誤處理一律經 `root.after` 轉回主執行緒。
- `search_page.py`：單一輸入框，見第 9 階段。
- 啟動時若 `bible_books.rejected_aliases()` 不為空，顯示通知。
- **所有 `messagebox` 呼叫只存在於 `ui/`。**

---

## 5. 錯誤處理與 log 規範

| 層 | 做法 |
|---|---|
| 核心模組 | 拋出具名例外（`ReferenceError`、`BibleApiError`…），可在 log 記錄細節 |
| pipeline | 單一引用失敗：記錄並略過，繼續製作（維持現行容錯行為）；整體失敗：拋出 |
| ui | 捕捉例外 → `root.after` → `messagebox` |

log 設定在 `main.py` 統一初始化一次，各模組使用 `logging.getLogger(__name__)`。

---

## 6. 實施階段

每個階段結束時，snapshot 測試都要通過，再提交 commit。

| 階段 | 內容 | 驗收 |
|---|---|---|
| **0. 建立基準** | 寫一支腳本，以現行 `FaithSlide.py` 處理 repo 內的範例週報：將 `get_verses` 改為讀寫 `tests/fixtures/api/`，並以程式擷取輸出 PPT 的每張投影片文字，存成 `tests/fixtures/snapshots/*.json` | 每份範例週報都有一個 snapshot |
| **1. 基礎** | `version.py`、`paths.py`、`main.py`（暫時只啟動舊程式） | 程式可正常啟動 |
| **2. 書卷資料** | `books.default.json`、`bible_books.py`（先不做使用者檔） | 單元測試：resolve、pattern 排序、衝突偵測 |
| **3. 引用解析** | `reference_parser.py` | 單元測試涵蓋 README 範例與現有格式 |
| **4. 經文 API** | `bible_api.py` | 以錄製回應做離線測試 |
| **5. Word / 證道** | `settings.py` + `settings.default.json`（講員名稱）、`word_reader.py`、`sermon_parser.py` | 大綱結構測試 |
| **6. PPT** | `ppt_builder.py` | — |
| **7. 串接** | `pipeline.py` | **snapshot 與基準完全一致** |
| **8. UI** | `ui/` 三個檔案，`main.py` 改為啟動新 UI；刪除 `FaithSlide.py` | 手動測試兩個頁面與快捷鍵 |
| **9. 行為變更** | 見下表，每項獨立 commit | 各自說明 |

### 第 9 階段：行為變更清單

| 項目 | 說明 |
|---|---|
| 9a. 版本比較修正 | 字串比較 → 數值比較 |
| 9b. log 位置 | exe 旁 → `%APPDATA%\FaithSlide\FaithSlide.log`（安裝在 Program Files 時無寫入權限）；README 的回報說明需同步更新 |
| 9c. 使用者自訂別名 | `books.user.json` 讀寫 + 設定 UI |
| 9d. 查詢頁單一輸入框 | 移除書卷下拉選單與章、節欄位。輸入框接受全名、簡寫或別名（如 `約3:16-18`、`約翰福音3:16`）→ `parse` → 依（書卷, 章）分組查詢 → 依 ranges 篩選顯示 |

### 打包同步（第 8 階段完成後）

- PyInstaller：進入點改為 `main.py`，並加上 `--add-data "faithslide/data/books.default.json;faithslide/data"`。
- 版本號目前散落在 `FaithSlide.py`、`PyInstaller.txt`、`FaithSlide_Installer.iss` 三處 → 以 `version.py` 為準。可選：寫一個小 build 腳本自動帶入版本號。

---

## 7. 已確認的決策

- 書卷資料用 JSON；預設檔與使用者檔分開，使用者檔只存差異。
- 經文引用解析獨立成 `reference_parser.py`；標題切分正則留在 `sermon_parser.py`。
- Reference 的顯示格式（含中文數字轉換）放在 `reference_parser.py`。
- 暫不支援跨章引用。
- 更新檢查拆成 `updater.py`（邏輯）與 UI（對話框）。
- `BibleDictionary.py` 由作者自行隔離，不在本計畫範圍內。
- 測試框架使用 `pytest`。
- 查詢頁只用單一輸入框，接受全名、簡寫或別名，不保留下拉選單。
- 別名衝突：直接拒絕並通知使用者。
- 講員名稱清單移到設定檔。
- 第 0 階段以腳本執行現行程式並擷取 JSON 作為基準（已完成）。
