from time import sleep
from random import uniform
from bs4 import BeautifulSoup
from threading import Thread
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException
import tkinter as tk
from tkinter import ttk, messagebox
import logging
import os

self_path = os.path.abspath(__file__)
base_path = os.path.dirname(self_path)
log_path = os.path.join(base_path, "bible_query.log")

logging.basicConfig(
    filename=log_path,
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    encoding="utf-8"
)

url = "https://bible.fhl.net/index.html"
driver = None
driver_ready = False  # 是否完成初始化


# -----------------------------------------
# 聖經書卷縮寫對照表（可選擇）
# -----------------------------------------
abbr_to_full = {
    "創": "創世記", "出": "出埃及記", "利": "利未記", "民": "民數記", "申": "申命記",
    "書": "約書亞記", "士": "士師記", "得": "路得記", "撒上": "撒母耳記上", "撒下": "撒母耳記下",
    "王上": "列王紀上", "王下": "列王紀下", "代上": "歷代志上", "代下": "歷代志下",
    "拉": "以斯拉記", "尼": "尼希米記", "斯": "以斯帖記", "伯": "約伯記", "詩": "詩篇",
    "箴": "箴言", "傳": "傳道書", "歌": "雅歌", "賽": "以賽亞書", "耶": "耶利米書",
    "哀": "耶利米哀歌", "結": "以西結書", "但": "但以理書", "何": "何西阿書",
    "珥": "約珥書", "摩": "阿摩司書", "俄": "俄巴底亞書", "拿": "約拿書", "彌": "彌迦書",
    "鴻": "那鴻書", "哈": "哈巴谷書", "番": "西番雅書", "該": "哈該書", "亞": "撒迦利亞書",
    "瑪": "瑪拉基書", "太": "馬太福音", "可": "馬可福音", "路": "路加福音", "約": "約翰福音",
    "徒": "使徒行傳", "羅": "羅馬書", "林前": "哥林多前書", "林後": "哥林多後書",
    "加": "加拉太書", "弗": "以弗所書", "腓": "腓立比書", "西": "歌羅西書",
    "帖前": "帖撒羅尼迦前書", "帖後": "帖撒羅尼迦後書", "提前": "提摩太前書",
    "提後": "提摩太後書", "多": "提多書", "門": "腓利門書", "來": "希伯來書",
    "雅": "雅各書", "彼前": "彼得前書", "彼後": "彼得後書", "約一": "約翰一書",
    "約二": "約翰二書", "約三": "約翰三書", "猶": "猶大書", "啟": "啟示錄"
}
full_to_abbr = {v: k for k, v in abbr_to_full.items()}
old_testament_books = [
    "創",  # 創世記
    "出",  # 出埃及記
    "利",  # 利未記
    "民",  # 民數記
    "申",  # 申命記
    "書",  # 約書亞記
    "士",  # 士師記
    "得",  # 路得記
    "撒上",  # 撒母耳記上
    "撒下",  # 撒母耳記下
    "王上",  # 列王紀上
    "王下",  # 列王紀下
    "代上",  # 歷代志上
    "代下",  # 歷代志下
    "拉",  # 以斯拉記
    "尼",  # 尼希米記
    "斯",  # 以斯帖記
    "伯",  # 約伯記
    "詩",  # 詩篇
    "箴",  # 箴言
    "傳",  # 傳道書
    "歌",  # 雅歌
    "賽",  # 以賽亞書
    "耶",  # 耶利米書
    "哀",  # 耶利米哀歌
    "結",  # 以西結書
    "但",  # 但以理書
    "何",  # 何西阿書
    "珥",  # 約珥書
    "摩",  # 阿摩司書
    "俄",  # 俄巴底亞書
    "拿",  # 約拿書
    "彌",  # 彌迦書
    "鴻",  # 那鴻書
    "哈",  # 哈巴谷書
    "番",  # 西番雅書
    "該",  # 哈該書
    "亞",  # 撒迦利亞書
    "瑪"   # 瑪拉基書
]


# -----------------------------------------
# 初始化 driver（背景執行）
# -----------------------------------------
def init_driver():
    global driver, driver_ready
    try:
        option = webdriver.ChromeOptions()
        option.add_argument('--headless')
        option.add_experimental_option('excludeSwitches', ['enable-automation'])
        driver = webdriver.Chrome(options=option)
        driver.get(url)
        driver.implicitly_wait(8)
        sleep(0.5)
        driver_ready = True
        # print("✅ Selenium 已預載完成！")
    except Exception as e:
        messagebox.showwarning("⚠️ 初始化 Selenium 失敗：", e)
        logging.error(f"⚠️ 初始化 Selenium 失敗：{e}")


# -----------------------------------------
# Selenium 操作區
# -----------------------------------------
def tap_button(driver, button):
    try:
        tap = driver.find_element('css selector', button)
        driver.execute_script('arguments[0].click();', tap)
    except TimeoutException:
        pass
    except Exception:
        messagebox.showwarning("點擊錯誤")
        logging.error(f"點擊錯誤：{button}")


def Dropdown(driver, by, name, value, old):
    try:
        select_element = driver.find_elements(by, name)
        BookDropdown = Select(select_element[0 if old else 1])
        BookDropdown.select_by_value(value)
        sleep(uniform(0.1, 0.2))
    except Exception:
        messagebox.showwarning(f"搜尋不到下拉選單 {name}")
        logging.error(f"搜尋不到下拉選單 {name}")


def get_verses(book_abbr, chapter, old):
    Dropdown(driver, "name", "chineses", book_abbr, old)
    Dropdown(driver, "name", "chap", chapter, old)
    tap_button(driver, "#content > div > form:nth-child(10) > input[type=submit]:nth-child(19)")

    sleep(1)
    soup = BeautifulSoup(driver.page_source, 'html.parser')
    verses = []
    all_Verse = soup.find_all("tr")

    for i in all_Verse:
        try:
            td = i.find_all("td")
            if len(td) >= 2:
                VerseNumber = td[0].text
                if ":" in VerseNumber:
                    VerseNumber = VerseNumber.split(":")[1]
                    verses.append(f"{VerseNumber}. {td[1].text.strip()}")
        except Exception as e:
            messagebox.showwarning(f"抓取經文錯誤: {e}")
            logging.error(f"抓取經文錯誤: {e}")

    driver.get(url)
    driver.implicitly_wait(8)
    sleep(0.3)

    return verses


# -----------------------------------------
# GUI 介面
# -----------------------------------------
def run_search():
    book_abbr = book_var.get()
    chapter = chapter_var.get()
    verse = verse_var.get()

    if not book_abbr or not chapter:
        messagebox.showwarning("輸入錯誤", "請選擇書卷與章節")
        logging.error("輸入錯誤", "請選擇書卷與章節")
        return
    if book_abbr not in abbr_to_full.keys():
        book_abbr = full_to_abbr.get(book_abbr, "")
        if not book_abbr:
            messagebox.showwarning("輸入錯誤", "書卷名稱錯誤")
            logging.error("輸入錯誤", "書卷名稱錯誤")
            return
    if book_abbr in old_testament_books:
        old = True
    else:
        old = False
    # messagebox.showinfo("請稍候", f"正在抓取 {abbr_to_full[book_abbr]} 第 {chapter} 章 ...")
    if not driver_ready:
        messagebox.showinfo("請稍候", "Selenium 正在初始化，請稍後再查詢。")
        logging.warning("Selenium 正在初始化")
        return

    verses = get_verses(book_abbr, chapter, old)

    text_box.delete(1.0, tk.END)
    if verse:
        if "-" in verse:
            start, end = map(int, verse.split("-"))
            for v in range(start, end + 1):
                if 1 <= v <= len(verses):
                    text_box.insert(tk.END, verses[v - 1] + "\n")
                else:
                    messagebox.showwarning("輸入錯誤", "節數錯誤")
                    logging.warning("抓取經文失敗: 節數超出範圍")
            logging.info(f"成功抓取 {abbr_to_full[book_abbr]} 第 {chapter} 章 {start}-{end} 節")
        else:
            v= int(verse)
            if 1 <= v <= len(verses):
                text_box.insert(tk.END, verses[v-1] + "\n")
                logging.info(f"成功抓取 {abbr_to_full[book_abbr]} 第 {chapter} 章")

            else:
                messagebox.showwarning("輸入錯誤", "節數錯誤")
                logging.warning("抓取經文失敗: 節數超出範圍")
    else:
        if verses:
            for v in verses:
                text_box.insert(tk.END, v + "\n")
                logging.info(f"成功抓取 {abbr_to_full[book_abbr]} 第 {chapter} 章")

        else:
            text_box.insert(tk.END, "未抓取到經文，請檢查網頁或選擇。")
            logging.error("網頁未回應")

def close_driver():
    driver.quit()
    root.destroy()


root = tk.Tk()
root.title("聖經經文查詢工具")
root.geometry("500x500")

# 外框
frame = ttk.Frame(root, padding=20)
frame.pack(expand=True)

# 標題
title_label = ttk.Label(frame, text="📖 聖經經文查詢", font=("微軟正黑體", 16, "bold"))
title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20))

# 書卷
ttk.Label(frame, text="書卷：", font=("微軟正黑體", 12)).grid(row=1, column=0, sticky="e", padx=5, pady=5)
book_var = tk.StringVar()
book_combo = ttk.Combobox(frame, textvariable=book_var, values=list(abbr_to_full.keys()), width=15)
book_combo.grid(row=1, column=1, padx=5, pady=5)

# 章
ttk.Label(frame, text="章：", font=("微軟正黑體", 12)).grid(row=2, column=0, sticky="e", padx=5, pady=5)
chapter_var = tk.StringVar()
chapter_entry = ttk.Entry(frame, textvariable=chapter_var, width=18)
chapter_entry.grid(row=2, column=1, padx=5, pady=5)

# 節（新加的）
ttk.Label(frame, text="節：", font=("微軟正黑體", 12)).grid(row=3, column=0, sticky="e", padx=5, pady=5)
verse_var = tk.StringVar()
verse_entry = ttk.Entry(frame, textvariable=verse_var, width=18)
verse_entry.grid(row=3, column=1, padx=5, pady=5)

# 查詢按鈕
search_btn = ttk.Button(frame, text="查詢", command=run_search)
search_btn.grid(row=4, column=0, columnspan=2, pady=(15, 0))

# 退出按鈕
search_btn = ttk.Button(frame, text="退出", command=close_driver)
search_btn.grid(row=5, column=0, columnspan=2, pady=(15, 0))

text_box = tk.Text(root, wrap="word")
text_box.pack(fill="both", expand=True, padx=10, pady=10)

# 置中設定
for i in range(5):
    frame.grid_rowconfigure(i, weight=1)
frame.grid_columnconfigure(0, weight=1)
frame.grid_columnconfigure(1, weight=1)

Thread(target=init_driver, daemon=True).start()
root.mainloop()
