from docx import Document
from pptx import Presentation
from pptx.util import Inches
import copy
import os
from BibleDictionary import old_testament_books
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
# FaithSlide.py

url = "https://bible.fhl.net/index.html"
driver = None
driver_ready = False  # 是否完成初始化

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

def tap_button(driver, button):
    try:
        tap = driver.find_element('css selector', button)
        driver.execute_script('arguments[0].click();', tap)
    except TimeoutException:
        pass
    except Exception:
        # messagebox.showwarning("點擊錯誤")
        logging.error(f"點擊錯誤：{button}")

def Dropdown(driver, by, name, value, old):
    try:
        select_element = driver.find_elements(by, name)
        BookDropdown = Select(select_element[0 if old else 1])
        BookDropdown.select_by_value(value)
        sleep(uniform(0.1, 0.2))
    except Exception:
        # messagebox.showwarning(f"搜尋不到下拉選單 {name}")
        logging.error(f"搜尋不到下拉選單 {name}")

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

def remove_slide(prs:Presentation, index:int) -> None:
    xml_slides = prs.slides._sldIdLst
    slide = list(xml_slides)
    xml_slides.remove(slide[index])

def verses_PPT(title, verses):
    new_slide = duplicate_slide(prs, 0)

    text_frame = new_slide.shapes[0].text_frame
    p = text_frame.paragraphs[0]
    if not p.runs:
        p.add_run()
    p.runs[0].text = title
    for i in range(1, 3):
        if p.runs[i].text:
            p.runs[i].text = ""
        else:
            break

    text_frame = new_slide.shapes[1].text_frame
    p = text_frame.paragraphs[0]
    if not p.runs:
        p.add_run()
    num = verses.split(".")[0] + "."
    out_verses = verses.split(".")[1]
    p.runs[0].text = num
    p.runs[1].text = out_verses

def main_title_PPT(title):
    new_slide = duplicate_slide(prs, 1)

    text_frame = new_slide.shapes[1].text_frame
    p = text_frame.paragraphs[0]
    if not p.runs:
        p.add_run()
    p.runs[0].text = title
    new_slide = duplicate_slide(prs, 2)

def major_heading_PPT(major):
    new_slide = duplicate_slide(prs, 3)

    text_frame = new_slide.shapes[0].text_frame
    p = text_frame.paragraphs[0]
    if not p.runs:
        p.add_run()
    p.runs[0].text = major.split("、")[0] + "、"
    p.runs[1].text = major.split("、")[1]

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


def minor_heading_PPT(minor):
    pass

def parse_bible_reference(bible):
    book = ""
    chapter_and_verse = ""
    # print(bible)
    for char in bible:
        print(char, "char")
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

            verses = get_verses(book, chapter_and_verse.split(":")[0], old)            

            title = f"{abbr_to_full[book]}"
            for ch in chapter_and_verse.split(":")[0:1]:
                chinese_chapter = ""
                for c in ch:
                    if c in number:
                        if chinese_chapter == "":
                            chinese_chapter += chinese_number[int(c)-1]
                        else:
                            if chinese_chapter == "一":
                                chinese_chapter = ""
                            if "十" in chinese_chapter:
                                chinese_chapter  = chinese_chapter.replace("十", "百")
                            chinese_chapter += f"十{chinese_number[int(c)-1]}"
                title += f"{chinese_chapter}章"
                
            
            title += f"{chapter_and_verse.split(':')[1].replace(",", "")}節"
            # print(title, "title")
            start = int(chapter_and_verse.split(":")[1].split("-")[0])-1
            try:
                end = int(chapter_and_verse.split(":")[1].split("-")[1].replace(",",""))
            except:
                end = start + 1
            for v in range(start, end):
                verses_PPT(title, verses[v].replace(" ", ""))
                print(verses[v].replace(" ", ""))
            # verses_PPT
            print()
            book = ""
            chapter_and_verse = ""
        else:
            book += char

def paragraph_PPT(heading, verses):
    if heading["minor"]:
        heading_livel = 3
    elif heading["medium"]:
        heading_livel = 2
    else:
        heading_livel = 1

    if heading_livel == 1:
        major_heading_PPT(heading["major"])
        # for verse in verses[0]:
        print(verses, "in paragraph_PPT")
        parse_bible_reference(verses[0])
            # verses_PPT(verse)
    elif heading_livel == 2:
        major_heading_PPT(heading["major"])
        parse_bible_reference(verses[0])
        for medium in heading["medium"]:
            medium_hearding_PPT(heading["major"], medium, heading["medium"])
    # elif heading_livel == 3: #確認模板
    #     for minor in heading["minor"]:
    #         minor_heading_PPT(minor)
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
    "約一": "約翰一書",
    "約二": "約翰二書",
    "約三": "約翰三書",
    "猶": "猶大書",
    "啟": "啟示錄"
}

# 全名 -> 簡稱
full_to_abbr = {v: k for k, v in abbr_to_full.items()}
chinese_number = ["一", "二", "三", "四", "五", "六", "七", "八", "九", "十"]
number = ["1", "2", "3", "4", "5", "6", "7", "8", "9", "0"]

Thread(target=init_driver, daemon=True).start()

if getattr(sys, 'frozen', False):
    base_path = sys._MEIPASS
else:
    base_path = os.path.dirname(os.path.abspath(__file__))

log_path = os.path.join(base_path, "bible_query.log")

logging.basicConfig(
    filename=log_path,
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    encoding="utf-8"
)

# 請改成你的 Word 路徑
self_path = os.path.abspath(__file__)
base_path = os.path.dirname(self_path)
# wordfile_path = os.path.join(base_path, "202501005新竹主日週報.docx")
wordfile_path = os.path.join(base_path, "20250928新竹主日週報.docx")
template_ppt_file = os.path.join(base_path, "template.pptx")
prs = Presentation(template_ppt_file)
doc = Document(wordfile_path)
ReadTheBible = []
Promise = []

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
        elif tatil == "證道":
            bold_texts  = []
            for cell in row.cells:
                for para in cell.paragraphs:
                    for run in para.runs:
                        text = run.text.strip()
                        if run.bold and text:
                            bold_texts.append(text)
            Promise = bold_texts[1:]

if not ReadTheBible:
    print("讀經抓取失敗")
else:
    print("讀經:")
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

    # print(ReadTheBible)
    if not isinstance(main_verses, list):
        for verses in ReadTheBible:
            print(verses)
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
                # print(verses, ReadTheBible[verses_index])
                verses_PPT(verses, ReadTheBible[verses_index])

                verses_index += 1

print()

for book in full_to_abbr.keys():
    if book in main_verses:
        main_book = full_to_abbr[book]
        break
print("main book", main_book)

if not Promise:
    print("證道抓取失敗")

else:
    print("證道:", Promise)
    make_main_title = False
    heading = {"major": "", "medium": [], "minor": {}}
    verses = [[], {}, []]  # 大標題，主標題，副標題 經文
    subtitle = False
    minor_title = False
    heading_livel = 0
    for text in Promise:
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
                        print(heading, heading_livel, verses)
                        paragraph_PPT(heading, verses)
                        heading = {"major": "", "medium": [], "minor": {}}
                        verses = [[], {}, []]  # 大標題，主標題，副標題 經文

                    heading_livel = 1
                    heading["major"] = text
                elif "." in text: # 副標題
                    if heading["major"] == "":
                        print("副標題出現於主標題之前，格式錯誤")
                    else:  
                        heading["medium"].append(text)
                        
                        subtitle = True
                        heading_livel = 2

                elif ")" in text:  # 小標題，待測試
                    heading_livel = 3
                    if len(heading["medium"]) == 0:
                        print("小標題出現於副標題之前，格式錯誤")
                    if heading["medium"][-1] not in heading["minor"].keys():
                        heading["minor"][heading["medium"][-1]] = []
                    heading["minor"][heading["medium"][-1]].append(text)
                    minor_title = True

                else:
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

    print(heading, heading_livel, verses)
    paragraph_PPT(heading, verses)
    
                
                        
for _ in range(6):
    remove_slide(prs,0)

save_path = os.path.join(base_path, "test.pptx")
prs.save(save_path)