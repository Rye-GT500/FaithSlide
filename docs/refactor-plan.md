# FaithSlide SRP 重構企畫書

> 狀態：**已確認，實施中**（第 0～8 階段完成）
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
│     ├─ alias_dialog.py       書卷別名設定視窗
│     ├─ background.py         背景執行緒工作（queue 交回主執行緒）
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
 │         │              └─────────► bible_api（只用 Verse 型別）
 │         │──────► bible_api ─────────────────────────────► bible_books
 │         └──────► ppt_builder ──► sermon_parser、bible_api（只用 Section、Verse 型別）
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
def default_books() -> list[Book]                          # 未套用使用者別名
def alias_problem(name: str) -> str | None                 # 別名本身的格式問題
def check_aliases(aliases_by_book) -> list[RejectedAlias]  # 書卷 id → 完整別名清單
def save_user_aliases(aliases_by_book) -> None             # 有被拒絕的別名時拋出 AliasConflictError，不寫入
def load_warnings() -> list[str]                           # 載入時被忽略的別名或檔案問題，供 UI 通知
```

**驗證規則**

- 同一名稱對應到多卷書：**直接拒絕並通知**。
  - 儲存時：拋出 `AliasConflictError`（列出衝突的別名與書卷），不寫入檔案，由設定 UI 顯示。
  - 載入時（例如使用者手動編輯了 JSON）：略過衝突的別名，其餘照常載入；UI 啟動後讀取 `load_warnings()` 顯示通知。
- 別名不可包含數字、空白或引用格式用到的標點（會破壞引用解析），長度上限 10 字。
- 套用時先處理所有移除再處理新增，所以可以一次把別名從 A 書移到 B 書。
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

class ReferenceParseError(ValueError): ...   # 錯誤訊息供 UI 直接顯示（避開內建的 ReferenceError）

def find_references(text: str) -> list[str]            # 從一段文字找出引用字串
def text_before_references(text: str) -> str | None   # 第一個引用之前的文字（供標題擷取）
def parse(raw: str, default_book: Book | None = None) -> list[Reference]
def format_title(ref: Reference, verse_range: tuple[int, int] | None = None) -> str   # 例：「約翰福音三章16-18節」；詩篇改為「篇」
def chinese_numeral(n: int) -> str
```

- **不支援跨章**（如 `約3:16-4:2`）：`parse` 拋出 `ReferenceParseError`。
- 一個字串可包含多章（如 `約3:16,4:2`），每章各回傳一個 `Reference`。
- `default_book` 用來取代全域變數 `main_book`。
- `parse` 會正規化全形符號（`：`、`，`、`－`…）與空白，供查詢頁輸入使用。
- **刻意保留的舊版行為**：引用正則容許數字中間夾空白（作者的刻意設計），因為 Word 的 run 常把「10」切成「1 0」。代價是「3:16 4:1」這種只以空白分隔、又沒寫書卷的兩組章節會被合併；這屬於使用者輸入失誤，列為後續 UX 改善（見 9f），非首要。
  - 可討論的替代做法：讓 `word_reader` 在同一段落內直接串接 run（不 strip、不插空白），還原 Word 原文，正則就不必容許數字中間夾空白。
- 已用 4 份範例週報比對：引用切分、標題擷取、逐節展開的標題與節數，都與舊版完全一致。

**取代的現有程式碼**：`analyze_word` 中的引用正則、`parse_bible_reference`、`process_reference_block`、`analyze_paragraph` 的範圍解析、`num_to_chinese`、`chinese_number`、`number`。

### 4.3 `bible_api.py`

```python
class BibleApiError(Exception): ...

def get_chapter(book: Book, chapter: int) -> list[Verse]   # Verse(number: int, text: str)
```

- 失敗時拋出例外，**不再回傳空 list 掩蓋錯誤**；查無經文（空結果）也視為錯誤。
- 同一次執行內快取，同章只查一次，失敗結果不快取；`clear_cache()` 供測試使用。
- 經文以節數查找，不再以位置索引（已確認錄製的 26 章節數皆連續，結果相同）。
- 已以實際 API 比對錄製資料：26 章輸出與舊版 `get_verses` 完全一致。

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
    level: int              # MAJOR / MEDIUM / MINOR
    number: str             # 「一」「1」
    text: str               # 編號之後的標題文字
    references: list[str]   # 未解析的引用字串（書卷可能省略）
    children: list["Section"]

@dataclass
class Sermon:
    title: str
    references: list[str]   # 主標題後的引用
    sections: list[Section]
    warnings: list[str]

@dataclass
class Passage: title: str; verses: list[Verse]

@dataclass
class Reading:
    passages: list[Passage]
    main_book: Book | None  # 證道引用省略書卷時使用
    warnings: list[str]

def parse_sermon(text: str, speaker_words: list[str]) -> Sermon
def parse_reading(lines: list[str]) -> Reading
def find_book_in(text: str) -> Book | None
```

- **標題切分正則**（`一、` / `1.` / `1)`）放在這裡。
- 要移除的講員字串（`證道`、`牧師`、人名…）**移到設定檔**（`settings.default.json` 的 `speaker_words`，使用者可在 `settings.json` 覆寫），由呼叫端傳入 `parse_sermon`，parser 本身不讀檔。
- **與原規劃的差異**：引用保留為未解析字串，由 pipeline 以 `Reading.main_book` 解析。這樣 `parse_sermon` 不需要知道主經文，引用錯誤也統一在 pipeline 處理。
- `Verse` 沿用 `bible_api.Verse`（讀經經文與查詢經文共用同一型別）。
- 讀經支援兩種格式：A. 標題行 + `[n]` 經文行（可多個標題行）；B. 第一行以「，」列出多段，經文行依節數依序分配。
- 解析問題記錄在 `warnings`，不再中斷整份製作。
- 已用 4 份範例週報比對：舊版投影片函式呼叫序列（主標題、大/中/小標題、163 節經文的標題與內容）與新模組推導結果完全一致。

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
- `calculate_font_size` 公開（供測試），`_weighted_length` 為私有。
- 中、小標題以物件識別（`is`）判斷目前項目，不再以字串比對（舊版遇到同名標題會停在第一個）。
- 存檔時連同關聯移除樣板投影片；舊版只從清單移除，6 張樣板仍留在檔案中。
- 已用 4 份範例週報比對：新模組 + `SlideBuilder` 產生的 PPT 與第 0 階段 snapshot 完全一致（版面、文字、字體大小）。

**取代的現有程式碼**：`duplicate_slide`、`remove_slide`、所有 `*_PPT` 函式、字體計算。

### 4.7 `pipeline.py`

```python
class GenerationError(Exception): ...        # 訊息可直接顯示給使用者

@dataclass
class GenerationResult:
    warnings: list[str]

def generate(word_path: str, ppt_path: str,
             on_progress: Callable[[float, str], None] = ...) -> GenerationResult   # 進度 0～1
```

依序執行：讀 Word → 解析讀經/證道 → 查經文 → 建投影片 → 存檔。

- 無法讀取 Word、範本，或存檔失敗（例如檔案正在 PowerPoint 中開啟）：拋出 `GenerationError`。
- 單一引用格式錯誤、查無章節、節數超出本章：記錄於 `GenerationResult.warnings` 並略過，繼續製作。同一範圍的超出節數只警告一次。
- 已驗證：snapshot 測試改為以 pipeline 產生（離線），4 份與基準完全一致、無警告；另以實際 API 連網執行，結果相同。
- 第 8 階段已改接 pipeline，並移除 `FaithSlide.py`、`tests/legacy_runner.py`、`tests/make_baseline.py`（基準已凍結，API 錄製資料保留）。刻意改變輸出時，以 `python -m tests.update_snapshots` 重新產生 snapshot。

- 投影片順序：讀經經文 → 主標題（2 張）→ **主標題後的引用經文**（作者決定：偶爾會出現，予以製作）→ 各大標題、其引用 → 各中標題、其引用 → 各小標題、其引用。
- 證道經文去除空白後再放入投影片（沿用舊版）。

**取代的現有程式碼**：`produce_the_slides`、`paragraph_PPT`。

### 4.8 `updater.py` / `version.py`

```python
@dataclass
class UpdateInfo: version: str; url: str

def check() -> UpdateInfo | None    # 失敗或無新版都回傳 None，並記錄 log
```

版本比較改為數值 tuple 比較。

### 4.9 `ui/`

- `background.py`：`run_in_background(widget, work, on_done, on_error, on_progress)`。背景執行緒只把進度與結果放入 `queue`，主執行緒以 `after` 定時取出處理，Tk 元件只在主執行緒操作。（舊版從背景執行緒呼叫 `root.after`，tkinter 不保證安全；冒煙測試中確實會失敗。）
- `main_window.py`：建立 `Tk`、頁面切換（F5）、退出（Esc）、Enter 交給目前頁面。兩頁同時建立、切換時只顯示其中一頁，輸入內容不會因切換而消失。啟動時以 `run_in_background` 呼叫 `updater.check()`。
- `slide_page.py`：檔案選擇、進度條。製作前檢查是否已選擇檔案；製作中停用「製作」按鈕，避免重複執行。完成時若有警告，列出前 10 則。
- `search_page.py`：單一輸入框（9d），接受全名、簡寫或別名，多段以「；」分隔，整章可省略節數。每個範圍顯示一個標題；不存在的節數以對話框提示。查詢邏輯 `lookup()` 不依賴 Tk，可單獨測試。
- `alias_dialog.py`（9c）：主視窗「書卷別名」按鈕開啟。可新增、移除、還原預設；新增時即時檢查，單字別名需確認；儲存後立即生效。
- 啟動時若 `bible_books.load_warnings()` 不為空，顯示通知。
- **所有 `messagebox` 呼叫只存在於 `ui/`。**

---

## 5. 錯誤處理與 log 規範

| 層 | 做法 |
|---|---|
| 核心模組 | 拋出具名例外（`ReferenceParseError`、`BibleApiError`…），可在 log 記錄細節 |
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
| 9a. 版本比較修正 ✅ | 字串比較 → 數字逐段比較；數字相同時正式版大於預覽版（從 1.2.2-beta 會提示更新到 1.2.2，舊版不會） |
| 9b. log 位置 ✅ | exe 旁 → `%APPDATA%\FaithSlide\FaithSlide.log`（安裝在 Program Files 時無寫入權限）；README 的回報說明需同步更新 |
| 9c. 使用者自訂別名 ✅ | `books.user.json` 讀寫 + 設定 UI |
| 9d. 查詢頁單一輸入框 ✅ | 移除書卷下拉選單與章、節欄位。輸入框接受全名、簡寫或別名（如 `約3:16-18`、`約翰福音3:16`）→ `parse` → 依（書卷, 章）分組查詢 → 依 ranges 篩選顯示 |
| 9e. 舊版 bug ✅（第 7 階段接上新模組後自動修正，snapshot 不受影響） | 「各」作為縮寫時標題組不出來；章數為整十的三位數（如詩篇 100、110、150）中文數字錯誤（「一百零零」、「一百一十零」）；二位數標題編號（`10.`、`十一、`）被從中間切開；主標題後接引用、中標題沒有引用、小標題出現在中標題前時，整份證道或整個段落製作失敗；讀經缺少時整份製作失敗；讀經格式 B 的書名含中文數字（如約翰一書）時被切成「約翰」；無引用的標題保留了 Word run 產生的空白；標題文字或讀經經文含第二個分隔符號（如 `1.a.b`）時被截斷 |
| 9f. 可疑引用提示（非首要） | 偵測使用者輸入失誤（如「3:16 4:1」被合併成「3:164」導致節數超出該章），提示使用者而非默默略過 |

### 打包同步（第 8 階段完成後）

- PyInstaller：進入點已改為 `main.py`，並加上 `--add-data "faithslide/data;faithslide/data"`（已完成）。
- 版本號目前在 `version.py`、`PyInstaller.txt`、`FaithSlide_Installer.iss` 三處 → 以 `version.py` 為準，build 腳本見「後續討論」。

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

## 8. 後續討論

- ✅ build 腳本：`build.py`，以 `version.py` 為唯一版本來源，自動帶入 PyInstaller 與 Inno Setup，並可 commit、tag、建立 GitHub Release。
- ✅ 版本號升級規則：依 commit 前綴建議升級幅度、由作者確認。規則與用法見 `docs/release.md`。
