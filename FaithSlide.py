from docx import Document
from pptx import Presentation
from pptx.util import Inches
import copy
import os
from threading import Thread
import logging
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.service import Service
from bs4 import BeautifulSoup
from time import sleep
import sys
from random import uniform
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

if getattr(sys, 'frozen', False):
    base_path = sys._MEIPASS
else:
    base_path = os.path.dirname(os.path.abspath(__file__))

log_path = os.path.join(base_path, "FaithSlide.log")

logging.basicConfig(
    filename=log_path,
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    encoding="utf-8"
)
self_path = os.path.abspath(__file__)
base_path = os.path.dirname(self_path)
template_ppt_file = os.path.join(base_path, "template.pptx")
prs = None
book_var = None
chapter_var = None
verse_var = None
text_box = None

url = "https://bible.fhl.net/index.html"
driver = None
driver_ready = False  # 是否完成初始化
# 簡稱 -> 全名
abbr_to_full = {
    "創": "創世記",
    "出": "出埃及記",
    "利": "利未記",
    "民": "民數記",
    "申": "申命記",
    "書": "約書亞記",
    "士": "士師記",
    "得": "路得記",
    "撒上": "撒母耳記上",
    "撒下": "撒母耳記下",
    "王上": "列王紀上",
    "王下": "列王紀下",
    "代上": "歷代志上",
    "代下": "歷代志下",
    "拉": "以斯拉記",
    "尼": "尼希米記",
    "斯": "以斯帖記",
    "伯": "約伯記",
    "詩": "詩篇",
    "箴": "箴言",
    "傳": "傳道書",
    "歌": "雅歌",
    "賽": "以賽亞書",
    "耶": "耶利米書",
    "哀": "耶利米哀歌",
    "結": "以西結書",
    "但": "但以理書",
    "何": "何西阿書",
    "珥": "約珥書",
    "摩": "阿摩司書",
    "俄": "俄巴底亞書",
    "拿": "約拿書",
    "彌": "彌迦書",
    "鴻": "何西阿書",  # 小先知書，部分版本略有不同
    "哈": "哈巴谷書",
    "番": "西番雅書",
    "該": "哈該書",
    "瑪": "撒迦利亞書",
    "亞": "瑪拉基書",
    "太": "馬太福音",
    "可": "馬可福音",
    "路": "路加福音",
    "約": "約翰福音",
    "徒": "使徒行傳",
    "羅": "羅馬書",
    "林前": "哥林多前書",
    "林後": "哥林多後書",
    "加": "加拉太書",
    "弗": "以弗所書",
    "腓": "腓立比書",
    "西": "歌羅西書",
    "帖前": "帖撒羅尼迦前書",
    "帖後": "帖撒羅尼迦後書",
    "提前": "提摩太前書",
    "提後": "提摩太後書",
    "多": "提多書",
    "門": "腓利門書",
    "來": "希伯來書",
    "雅": "雅各書",
    "彼前": "彼得前書",
    "彼後": "彼得後書",
    "約壹": "約翰一書",
    "約貳": "約翰二書",
    "約參": "約翰三書",
    "猶": "猶大書",
    "啟": "啟示錄"
}
# 全名 -> 簡稱
full_to_abbr = {v: k for k, v in abbr_to_full.items()}
chinese_number = ["零", "一", "二", "三", "四", "五", "六", "七", "八", "九", "十"]
number = ["1", "2", "3", "4", "5", "6", "7", "8", "9", "0"]
# 旧约书卷列表
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
search_page = False
main_book = ""

#爬蟲啟動
def init_driver():
    global driver, driver_ready
    try:
        option = webdriver.ChromeOptions()
        option.add_argument('--headless')
        option.add_experimental_option('excludeSwitches', ['enable-automation'])
        if getattr(sys, 'frozen', False):
            driver_path = os.path.join(sys._MEIPASS, "chromedriver.exe")
            driver = webdriver.Chrome(service=Service(driver_path), options=option)
        else:
            driver = webdriver.Chrome(options=option)
        driver.get(url)
        driver.implicitly_wait(8)
        sleep(0.5)
        driver_ready = True
        # print("✅ Selenium 已預載完成！")
    except Exception as e:
        # messagebox.showwarning("⚠️ 初始化 Selenium 失敗：", e)
        logging.error(f"⚠️ 初始化 Selenium 失敗：{e}")
#爬蟲點擊
def tap_button(driver, button):
    try:
        tap = driver.find_element('css selector', button)
        driver.execute_script('arguments[0].click();', tap)
    except TimeoutException:
        pass
    except Exception:
        # messagebox.showwarning("點擊錯誤")
        logging.error(f"點擊錯誤：{button}")
#爬蟲下拉選單
def Dropdown(driver, by, name, value, old):
    try:
        select_element = driver.find_elements(by, name)
        BookDropdown = Select(select_element[0 if old else 1])
        BookDropdown.select_by_value(value)
        sleep(uniform(0.1, 0.2))
    except Exception:
        # messagebox.showwarning(f"搜尋不到下拉選單 {name}")
        logging.error(f"搜尋不到下拉選單 {name}")
#爬蟲抓經文
def get_verses(book_abbr, chapter, old):
    Dropdown(driver, "name", "chineses", book_abbr, old)
    Dropdown(driver, "name", "chap", chapter, old)
    if old:
        tap_button(driver, "#content > div > form:nth-child(10) > input[type=submit]:nth-child(19)")
    else:
        tap_button(driver, "#content > div > form:nth-child(13) > input[type=submit]:nth-child(15)")
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
            # messagebox.showwarning(f"抓取經文錯誤: {e}")
            logging.error(f"抓取經文錯誤: {e}")

    driver.get(url)
    driver.implicitly_wait(8)
    sleep(0.3)

    return verses
#PPT 複製投影片
def duplicate_slide(prs:Presentation, index):
    template_slide = prs.slides[index]
    new_slide = prs.slides.add_slide(template_slide.slide_layout)
    for shape in list(new_slide.shapes):
        if shape.is_placeholder:
            sp = shape
            new_slide.shapes._spTree.remove(sp._element)
    for shape in template_slide.shapes:
        # if not shape.is_placeholder:
            el = shape.element
            new_el = copy.deepcopy(el)
            new_slide.shapes._spTree.insert_element_before(new_el, 'p:extLst')
    return new_slide
#PPT 刪除投影片
def remove_slide(prs:Presentation, index:int) -> None:
    xml_slides = prs.slides._sldIdLst
    slide = list(xml_slides)
    xml_slides.remove(slide[index])
#PPT 經文投影片
def verses_PPT(title, verses):
    if "." not in verses:
        logging.warning(f"經文格式錯誤，無法製作投影片: {title} {verses}")
        return
    num = verses.split(".")[0] + "."
    out_verses = verses.split(".")[1]

    if len(num) == 2:
        new_slide = duplicate_slide(prs, 5)
    else:
        new_slide = duplicate_slide(prs, 0)

    text_frame = new_slide.shapes[0].text_frame
    p = text_frame.paragraphs[0]
    if not p.runs:
        p.add_run()
    p.runs[0].text = title
    for i in range(1, 3):
        try:
            p.runs[i].text = ""
        except:
            break

    text_frame = new_slide.shapes[1].text_frame
    p = text_frame.paragraphs[0]
    if not p.runs:
        p.add_run()
    
    p.runs[0].text = num
    p.runs[1].text = out_verses
#PPT 主標題投影片
def main_title_PPT(title):
    new_slide = duplicate_slide(prs, 1)

    text_frame = new_slide.shapes[1].text_frame
    p = text_frame.paragraphs[0]
    if not p.runs:
        p.add_run()
    p.runs[0].text = title
    new_slide = duplicate_slide(prs, 2)
#PPT 大標題投影片
def major_heading_PPT(major):
    new_slide = duplicate_slide(prs, 3)

    text_frame = new_slide.shapes[0].text_frame
    p = text_frame.paragraphs[0]
    if not p.runs:
        p.add_run()
    p.runs[0].text = major.split("、")[0] + "、"
    p.runs[1].text = major.split("、")[1]
#PPT 中標題投影片
def medium_hearding_PPT(major, medium, medium_list):
    new_slide = duplicate_slide(prs, 4)

    text_frame = new_slide.shapes[0].text_frame
    p = text_frame.paragraphs[0]
    if not p.runs:
        p.add_run()
    p.runs[0].text = major.split("、")[0] + "、"
    p.runs[1].text = major.split("、")[1]

    text_frame = new_slide.shapes[1].text_frame
    p = text_frame.paragraphs[0]
    if not p.runs:
        p.add_run()
    t = 0
    for m in medium_list:
        p.runs[2*t].text = m.split(".")[0] + "."
        p.runs[2*t+1].text = m.split(".")[1].replace("：", "") + "\n"
        if m == medium:
            break
        t += 1
        p.add_run()
        p.add_run()
#PPT 小標題投影片
def minor_heading_PPT(major, medium, minor, minor_list):
    new_slide = duplicate_slide(prs, 4)

    text_frame = new_slide.shapes[0].text_frame
    p = text_frame.paragraphs[0]
    if not p.runs:
        p.add_run()
    p.runs[0].text = major.split("、")[0] + "、"
    p.runs[1].text = major.split("、")[1]

    text_frame = new_slide.shapes[1].text_frame
    p = text_frame.paragraphs[0]
    if not p.runs:
        p.add_run()

    p.runs[0].text = medium.split(".")[0] + "."
    p.runs[1].text = medium.split(".")[1].replace("：", "") + "\n"

    t = 1
    for m in minor_list:
        p.add_run()
        p.add_run()
        p.runs[2*t].text = "(" + m.split(")")[0] + ")"
        p.runs[2*t+1].text = m.split(")")[1].replace("：", "") + "\n"
        if m == minor:
            break
        t += 1
#PPT 經文章節轉中文
def num_to_chinese(title, chapter_and_verse: str) -> str:
    chapter = chapter_and_verse.split(":")[0]
    chinese_chapter = ""
    # print(chapter)
    if len(chapter) == 3:
        chinese_chapter += f"{chinese_number[int(chapter[0])]}百"
        chinese_chapter += f"{chinese_number[int(chapter[1])]}"
        if chinese_chapter[-1] != "零":
            chinese_chapter += "十"
        chinese_chapter += f"{chinese_number[int(chapter[2])]}"
    elif len(chapter) == 2:
        chinese_chapter += f"{chinese_number[int(chapter[0])]}"
        if chinese_chapter == "一":
            chinese_chapter = ""
        chinese_chapter += f"十{chinese_number[int(chapter[1])]}"
    elif len(chapter) == 1:
        chinese_chapter = f"{chinese_number[int(chapter[0])]}"

    title += f"{chinese_chapter}章"
    return title
#PPT 經文節數分析
def analyze_paragraph(title, verse_analyze, verses):
    start = int(verse_analyze.split("-")[0])-1
    try:
        end = int(verse_analyze.split("-")[1].replace(",",""))
    except:
        end = start + 1
    for v in range(start, end):
        verses_PPT(title, verses[v].replace(" ", ""))
        verse = verses[v].replace(" ", "")
        logging.info(f"{title} {verse}")
#PPT 經文章節處理
def process_reference_block(chapter_and_verse, book, old):
    scrape_verses = get_verses(book, chapter_and_verse.split(":")[0], old)     

    title = f"{abbr_to_full[book]}"
    title = num_to_chinese(title, chapter_and_verse)
    chapter_and_verse = chapter_and_verse.replace("，", "")
    verse = chapter_and_verse.split(':')[1]
    if "," in verse:
        verse = verse.split(",")

    if isinstance(verse, list):
        # print(verse, "is verse list")
        for v in verse:
            if v:
                analyze_paragraph(title+f"{v}節", v, scrape_verses)
    else:
        title += f"{verse}節"
        analyze_paragraph(title, verse, scrape_verses)
#PPT 經文書卷解析
def parse_bible_reference(bible):
    # print(bible)
    book = ""
    chapter_and_verse = ""
    for char in bible:
        # print(char)
        if char[0] in number:
            chapter_and_verse += char
            if book == "":
                book = main_book
            
            if book in old_testament_books:
                old = True
            else:
                old = False
            
            while not driver_ready:
                sleep(0.5)
            if chapter_and_verse.count(":") > 1:
                chapter_and_verse = chapter_and_verse.split(",")
                # print(chapter_and_verse, "is chapter and verse")
            if isinstance(chapter_and_verse, list):
                for cav in chapter_and_verse:
                    # print(cav, book, old)
                    process_reference_block(cav, book, old)
            else:
                # print(chapter_and_verse, book, old)
                process_reference_block(chapter_and_verse, book, old)
                # print(chapter_and_verse)
            

            book = ""
            chapter_and_verse = ""
        else:
            book += char
#PPT 段落處理
def paragraph_PPT(heading, verses):
    if heading["minor"]:
        heading_livel = 3
    elif heading["medium"]:
        heading_livel = 2
    else:
        heading_livel = 1
    logging.info(verses)
    if heading_livel == 1:
        major_heading_PPT(heading["major"])
        # for verse in verses[0]:
        # print(verses, "in paragraph_PPT")
        parse_bible_reference(verses[0])
            # verses_PPT(verse)
    elif heading_livel == 2:
        major_heading_PPT(heading["major"])
        parse_bible_reference(verses[0])
        for medium in heading["medium"]:
            medium_hearding_PPT(heading["major"], medium, heading["medium"])
            parse_bible_reference(verses[1][medium])
    elif heading_livel == 3: #確認模板
        major_heading_PPT(heading["major"])
        parse_bible_reference(verses[0])
        for medium in heading["medium"]:
            medium_hearding_PPT(heading["major"], medium, heading["medium"])
            parse_bible_reference(verses[1][medium])
            if medium in heading["minor"].keys():
                for minor in heading["minor"][medium]:
                    minor_heading_PPT(heading["major"], medium, minor, heading["minor"][medium])
                    parse_bible_reference(verses[2][minor])
#關閉驅動程式
def close_driver():
    driver.quit()
    root.destroy()
#分析word製作ppt
def Analyze_and_produce_the_slides():
    # messagebox.showwarning("開始製作投影片，請稍候...")
    global main_book, prs
    prs = Presentation(template_ppt_file)
    # print(log_path, "為日誌檔案位置")
    # 請改成你的 Word 路徑
    wordfile_path = word_path_var.get()
    print(wordfile_path)
    doc = Document(wordfile_path)
    ReadTheBible = []
    sermon = []
    # 逐個表格抓文字
    for t_idx, table in enumerate(doc.tables):
        # print(f"=== 表格 {t_idx+1} ===")
        for r_idx, row in enumerate(table.rows):
            # 取每個儲存格文字，去掉前後空白
            
            # print(row_texts)
            tatil = row.cells[0].text.strip()
            # 只印出有內容的列
            if tatil == "讀經":
                row_texts = [cell.text.strip() for cell in row.cells]
                ReadTheBible = row_texts[1].split("\n")
            elif tatil == "證道": #添加全形分割符號支援
                bold_texts  = ""
                for cell in row.cells:
                    for para in cell.paragraphs:
                        for run in para.runs:
                            text = run.text.strip()
                            if run.bold and text:
                                if bold_texts == "證道":
                                    bold_texts = text
                                else:
                                    if text in ["錢致榮", "牧師", "傳道", "吳佩倫"]:
                                        if bold_texts != "":
                                            # print("遇到講員名稱，結束證道內容擷取", bold_texts)
                                            sermon.append(bold_texts)
                                        bold_texts = ""
                                        continue
                                    for symbol in ["、", ".", ")"]:
                                        if symbol in bold_texts:
                                            # print("遇到分隔符號，結束證道內容擷取", symbol, bold_texts)
                                            first_part = bold_texts.split(symbol)[0]
                                            second_part = first_part[-1][-1] + symbol + bold_texts.split(symbol)[1]
                                            first_part = first_part[:-1]
                                            if first_part:
                                                sermon.append(first_part)
                                            if second_part: 
                                                sermon.append(second_part)
                                            bold_texts = text
                                            break
                                    else:
                                        for book in ["創", "出", "利", "民", "申", "書", "士", "得", "撒上", "撒下", "王上", "王下", "代上", "代下", "拉", "尼", "斯", "伯", "詩", "箴", "傳", "歌", "賽", "耶", "哀", "結", "但", "何", "珥", "摩", "俄", "拿", "彌", "鴻", "哈", "番", "該", "瑪", "亞", "太", "可", "路", "約", "徒", "羅", "林前", "林後", "加", "弗", "腓", "西", "帖前", "帖後", "提前", "提後", "多", "門", "來", "雅", "彼前", "彼後", "約壹", "約貳", "約參", "猶", "啟"]:
                                            # if book in bold_texts:
                                                # print(bold_texts.index(book)+len(book), len(bold_texts))
                                            if book in bold_texts and (bold_texts.index(book)+len(book) == len(bold_texts) or bold_texts[bold_texts.index(book)+1] in ["1", "2", "3", "4", "5", "6", "7", "8", "9", "0"]):
                                                first_part = bold_texts.split(book)[0]
                                                bold_texts = bold_texts.replace(first_part, "")
                                                # print("遇到書卷，結束證道內容擷取", book, first_part, bold_texts)
                                                if first_part:
                                                    sermon.append(first_part)
                                                if bold_texts:
                                                    sermon.append(bold_texts)
                                                bold_texts = text
                                                break
                                        else:
                                            n = ""
                                            min_num_index = len(bold_texts)
                                            for num in ["1", "2", "3", "4", "5", "6", "7", "8", "9", "0"]:
                                                if num in bold_texts:
                                                    idx = bold_texts.index(num)
                                                    if idx < min_num_index:
                                                        min_num_index = idx
                                                        n = num
                                            # print("最小數字索引", min_num_index, n)
                                            if n and n in bold_texts:
                                                # print(bold_texts.split(n))
                                                first_part = bold_texts.split(n)[0]
                                                bold_texts = bold_texts.replace(first_part, "")
                                                # print("遇到數字，結束證道內容擷取", n, first_part)
                                                if first_part:
                                                    sermon.append(first_part)
                                                bold_texts += text

                                            else:
                                                bold_texts += text
                                                continue
    sermon.append(bold_texts)

    if not ReadTheBible:
        logging.warning("讀經抓取失敗")
    else:
        logging.info("讀經:")
        main_verses = ReadTheBible[0]
        
        if "，" in main_verses:
            main_book = ""
            for text in main_verses:
                if text in chinese_number:
                    break
                main_book += text
            main_verses = main_verses.replace("，", " " + main_book).split()
        del ReadTheBible[0]
        for verses_index in range(0, len(ReadTheBible)):
            ReadTheBible[verses_index] = ReadTheBible[verses_index].replace("[", "").replace("]", ".")
        verses_index = 0

        logging.info(ReadTheBible)
        if not isinstance(main_verses, list):
            for verses in ReadTheBible:
                logging.info(verses)
                verses_PPT(main_verses, verses)
        else:
            for verses in main_verses:
                first_num = 0
                second_num = 0
                first_end = False
                for text in verses:
                    if text in number:
                        if not first_end:
                            first_num *= 10
                            first_num += int(text)
                        else:
                            second_num *= 10
                            second_num += int(text)
                    elif text == "-":
                        first_end = True
                # print(verses, first_num, second_num)
                if second_num == 0:
                    second_num = first_num
                for i in range(first_num, second_num+1):
                    logging.info(f"{verses}, {ReadTheBible[verses_index]}")
                    verses_PPT(verses, ReadTheBible[verses_index])

                    verses_index += 1


    for book in full_to_abbr.keys():
        if isinstance(main_verses, list):
            if book in main_verses[0]:
                main_book = full_to_abbr[book]
                break
        else:
            if book in main_verses:
                main_book = full_to_abbr[book]
                break

    logging.info(f"main book {main_book}")

    if not sermon:
        logging.warning("證道抓取失敗")
    else:
        logging.info(f"證道:{sermon}")
        # print(f"證道:{sermon}")
        make_main_title = False
        heading = {"major": "", "medium": [], "minor": {}}
        verses = [[], {}, {}]  # 大標題，主標題，副標題 經文
        subtitle = False
        minor_title = False
        heading_livel = 0
        for text in sermon:
            if subtitle: # 如果上一行是副標題的編號，表示這行是副標題內容
                subtitle = False
                heading["medium"][-1] += text
            elif minor_title: # 如果上一行是小標題的編號，表示這行是小標題內容
                minor_title = False
                last_medium = heading["medium"][-1]
                heading["minor"][last_medium][-1] += text
            else:
                if not make_main_title: # 大標題
                    main_title_PPT(text)
                    make_main_title = True
                else:
                    if "、" in text: # 主標題
                        if heading_livel != 0:# 已有完整段落，製作PPT
                            logging.info(f"{heading}, {verses}")
                            # print(heading, "\n", verses, "complete paragraph")
                            paragraph_PPT(heading, verses)
                            heading = {"major": "", "medium": [], "minor": {}}
                            verses = [[], {}, {}]  # 大標題，主標題，副標題 經文

                        heading_livel = 1
                        heading["major"] = text
                    elif "." in text: # 副標題
                        if heading["major"] == "":
                            logging.info("副標題出現於主標題之前，格式錯誤")
                        else:  
                            heading["medium"].append(text)
                            
                            subtitle = True
                            heading_livel = 2

                    elif ")" in text:  # 小標題，待測試
                        heading_livel = 3
                        if len(heading["medium"]) == 0:
                            logging.info("小標題出現於副標題之前，格式錯誤")
                        if heading["medium"][-1] not in heading["minor"].keys():
                            heading["minor"][heading["medium"][-1]] = []
                        heading["minor"][heading["medium"][-1]].append(text)
                        minor_title = True

                    else:
                        # print(text, "is verse")
                        is_verse = False
                        for t in text:
                            if t in number:
                                is_verse = True
                                break
                        else:
                            if text in abbr_to_full.keys():
                                is_verse = True
                        if is_verse:
                            if heading_livel == 1:
                                verses[0].append(text)
                            elif heading_livel == 2:
                                if heading["medium"][-1] not in verses[1].keys():
                                    verses[1][heading["medium"][-1]] = []
                                verses[1][heading["medium"][-1]].append(text)
                            else:
                                last_medium = heading["medium"][-1]
                                if heading["minor"][last_medium][-1] not in verses[2].keys():
                                    verses[2][heading["minor"][last_medium][-1]] = []
                                verses[2][heading["minor"][last_medium][-1]].append(text)
                                # print("小標題經文待測試")

        logging.info(f"{heading}, {verses}")
        paragraph_PPT(heading, verses)
        # print(heading, "\n", verses, "final paragraph")
        
                    
    # 刪除範本投影片                     
    for _ in range(6):
        remove_slide(prs,0)

    save_path = ppt_save_var.get()
    prs.save(save_path)
    logging.info("製作完成")
    messagebox.showwarning("製作完成")
#清空UI介面
def clear_frame(frame_to_clear):
    for widget in frame_to_clear.winfo_children():
        widget.destroy()
#經文搜尋工具
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
                    break
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
#創建經文查詢UI
def search_verse_UI():
    global book_var, chapter_var, verse_var, text_box
    # 標題
    title_label_search = ttk.Label(frame, text="📖 聖經經文查詢", font=("微軟正黑體", 16, "bold"))
    title_label_search.grid(row=0, column=0, columnspan=2, pady=(0, 20))
    # 書卷
    book_label = ttk.Label(frame, text="書卷：", font=("微軟正黑體", 12))
    book_var = tk.StringVar()
    book_combo = ttk.Combobox(frame, textvariable=book_var, values=list(abbr_to_full.keys()), width=15)
    book_label.grid(row=1, column=0, sticky="e", padx=5, pady=5)
    book_combo.grid(row=1, column=1, padx=5, pady=5, sticky="w")

    # 章
    chapter_label = ttk.Label(frame, text="章：", font=("微軟正黑體", 12))
    chapter_var = tk.StringVar()
    chapter_entry = ttk.Entry(frame, textvariable=chapter_var, width=18)
    chapter_label.grid(row=2, column=0, sticky="e", padx=5, pady=5)
    chapter_entry.grid(row=2, column=1, padx=5, pady=5, sticky="w")

    # 節（新加的）
    verse_label = ttk.Label(frame, text="節：", font=("微軟正黑體", 12))
    verse_var = tk.StringVar()
    verse_entry = ttk.Entry(frame, textvariable=verse_var, width=18)
    verse_label.grid(row=3, column=0, sticky="e", padx=5, pady=5)
    verse_entry.grid(row=3, column=1, padx=5, pady=5, sticky="w")

    # 查詢按鈕
    search_btn = ttk.Button(frame, text="查詢", command=run_search)
    search_btn.grid(row=4, column=0, columnspan=2, pady=(15, 0))

    text_box = tk.Text(frame, wrap="word")
    text_box.grid(row=6, column=0, columnspan=2, sticky="nsew", padx=10, pady=10)

    # 置中設定
    for i in range(6):
        frame.grid_rowconfigure(i, weight=1)
    frame.grid_columnconfigure(0, weight=1)
    frame.grid_columnconfigure(1, weight=1)
#創建PPT的UI 
def produce_the_slide_UI():
    ttk.Label(frame, text="Word 輸入:").grid(row=0, column=0, columnspan=2, pady=(20, 0))
    ttk.Entry(frame, textvariable=word_path_var, width=50, state='readonly').grid(row=1, column=0, padx=5, pady=5, sticky="e")
    ttk.Button(frame, text="選擇 Word", command=select_word_file).grid(row=1, column=1, padx=5, pady=5, sticky="w")

    ttk.Label(frame, text="PPT 輸出:").grid(row=2, column=0, columnspan=2, pady=(20, 0))
    ttk.Entry(frame, textvariable=ppt_save_var, width=50, state='readonly').grid(row=3, column=0, padx=5, pady=5, sticky="e")
    ttk.Button(frame, text="選擇儲存", command=select_save_path).grid(row=3, column=1, padx=5, pady=5, sticky="w")

    # 查詢按鈕
    produce_btn = ttk.Button(frame, text="製作", command=Analyze_and_produce_the_slides)
    produce_btn.grid(row=4, column=0, columnspan=2, pady=(15, 0))
#切換頁面
def change_page():
    global search_page
    print("切換", search_page)
    clear_frame(frame)
    search_page = not search_page
    if search_page:
        search_verse_UI()
    else:
        produce_the_slide_UI()
#選取word檔案位置
def select_word_file():
    """打開檔案對話框，讓使用者選擇 Word 檔案 (.docx)"""
    # filedialog.askopenfilename() 打開選擇檔案的對話框
    path = filedialog.askopenfilename(
        title="選擇 Word 證道文件",
        defaultextension=".docx", # 預設副檔名
        filetypes=[
            ("Word 檔案", "*.docx"),
            ("所有檔案", "*.*")
        ]
    )
    if path:
        # 如果使用者選擇了檔案，將路徑設定到 StringVar 變數中
        word_path_var.set(path)
        logging.info(f"選取 Word 檔案: {path}")
#選取PPT存檔位置
def select_save_path():
    """讓使用者指定輸出 PPT 檔案名稱 (.pptx)"""
    path = filedialog.asksaveasfilename(
        title="指定輸出 PPT 檔案名稱",
        defaultextension=".pptx",
        filetypes=[("PowerPoint 檔案", "*.pptx"), ("所有檔案", "*.*")],
        initialfile="證道投影片.pptx"
    )
    if path:
        ppt_save_var.set(path)
        logging.info(f"選取 PPT 儲存路徑: {path}")

root = tk.Tk()
root.title("FaithSlide")
root.geometry("650x650")
word_path_var = tk.StringVar(value="請選擇 Word 文件...")
ppt_save_var = tk.StringVar(value="請選擇輸出 PPT 檔案名稱...")

# 外框
frame = ttk.Frame(root, padding=20)
frame.grid(row=0, column=0, columnspan=2, sticky="nsew")

change_btn = ttk.Button(root, text="切換", command=change_page)
change_btn.grid(row=1, column=0, pady=(15, 0), sticky="e")

quit_btn = ttk.Button(root, text="退出", command=close_driver)
quit_btn.grid(row=1, column=1, pady=(15, 0), sticky="w")

produce_the_slide_UI()

# --- 確保 root 的權重配置 ---
root.grid_rowconfigure(1, weight=1)      # 讓 Button 所在的第二行 (row=1) 能夠擴展
root.grid_columnconfigure(0, weight=1)   # 讓第一列能擴展
root.grid_columnconfigure(1, weight=1)   # 讓第二列能擴展 (因為 frame 跨越了兩列)
# ----------------------------

if __name__ == "__main__":
    Thread(target=init_driver, daemon=True).start()
    root.mainloop()
    