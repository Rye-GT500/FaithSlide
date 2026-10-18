"""經文查詢：解析使用者輸入的經文引用，取回經文並整理成要顯示的文字行。"""
import re

from faithslide import bible_api
from faithslide.bible_api import BibleApiError
from faithslide.reference_parser import ReferenceParseError, format_title, parse


def lookup(query: str) -> tuple[list[str], list[str]]:
    """查詢經文，回傳 (要顯示的文字行, 問題訊息)。輸入格式錯誤時拋出 ReferenceParseError。"""
    segments = [s for s in re.split(r"[;；]", query) if s.strip()]
    if not segments:
        raise ReferenceParseError("請輸入要查詢的經文")
    refs = [ref for segment in segments for ref in parse(segment)]

    lines, problems = [], []
    for ref in refs:
        try:
            chapter = {v.number: v for v in bible_api.get_chapter(ref.book, ref.chapter)}
        except BibleApiError as e:
            problems.append(f"{format_title(ref)}：{e}")
            continue
        for verse_range in ref.ranges or [None]:
            title = format_title(ref, verse_range)
            start, end = verse_range or (min(chapter), max(chapter))
            found = [chapter[n] for n in range(start, end + 1) if n in chapter]
            if len(found) < end - start + 1:
                first_missing = next(n for n in range(start, end + 1) if n not in chapter)
                problems.append(f"{title}：本章沒有第 {first_missing} 節")
            if found:
                # 和合本以全形空白表示敬稱前的空格（如「　神」），顯示時一律改為半形
                lines += [f"【{title}】", *(f"{v.number}. {v.text.replace(chr(0x3000), ' ')}" for v in found), ""]
    return lines, problems
