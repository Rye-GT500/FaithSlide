from docx import Document
import os
# FaithSlide.py

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

# 請改成你的 Word 路徑
self_path = os.path.abspath(__file__)
base_path = os.path.dirname(self_path)
file_path = os.path.join(base_path, "20250928新竹主日週報.docx")
doc = Document(file_path)
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
    print("讀經:", ReadTheBible)
print()
if not Promise:
    print("證道抓取失敗")
else:
    print("證道:", Promise)
