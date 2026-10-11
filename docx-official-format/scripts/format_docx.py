#!/usr/bin/env python3
"""Format an ordinary .docx document as a simplified Chinese official document."""

from __future__ import annotations

import argparse
import json
import re
import shutil
import subprocess
import tempfile
from dataclasses import dataclass
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION_START
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm, Pt, RGBColor

LATIN_FONT = "Times New Roman"
FULLWIDTH_PUNCT = "“”‘’（）〔〕【】《》〈〉、。，；：？！…—～·％‰℃"


@dataclass(frozen=True)
class TypoRule:
    wrong: str
    suggestion: str
    note: str


@dataclass(frozen=True)
class TypoMatch:
    start: int
    end: int
    wrong: str
    suggestion: str
    note: str
    category: str = "错词"


@dataclass(frozen=True)
class FormatReport:
    typo_count: int = 0
    punctuation_fix_count: int = 0
    punctuation_issue_count: int = 0
    normalized_paragraph_count: int = 0


TYPO_RULES = (
    TypoRule("安排布署", "安排部署", "公文常用写法为“安排部署”。"),
    TypoRule("工作布署", "工作部署", "公文常用写法为“工作部署”。"),
    TypoRule("统一布署", "统一部署", "公文常用写法为“统一部署”。"),
    TypoRule("布署", "部署", "“部署”表示安排、布置工作。"),
    TypoRule("贯切落实", "贯彻落实", "公文常用写法为“贯彻落实”。"),
    TypoRule("贯切执行", "贯彻执行", "公文常用写法为“贯彻执行”。"),
    TypoRule("截止目前", "截至目前", "表示到某个时间点为止，一般用“截至”。"),
    TypoRule("截止今日", "截至今日", "表示到某个时间点为止，一般用“截至”。"),
    TypoRule("截止现在", "截至现在", "表示到某个时间点为止，一般用“截至”。"),
    TypoRule("登陆系统", "登录系统", "进入系统账号时通常用“登录”。"),
    TypoRule("系统登陆", "系统登录", "进入系统账号时通常用“登录”。"),
    TypoRule("账号登陆", "账号登录", "进入系统账号时通常用“登录”。"),
    TypoRule("再接再励", "再接再厉", "成语规范写法为“再接再厉”。"),
    TypoRule("一如继往", "一如既往", "成语规范写法为“一如既往”。"),
    TypoRule("历行节约", "厉行节约", "规范写法为“厉行节约”。"),
    TypoRule("凝心聚利", "凝心聚力", "常用表述为“凝心聚力”。"),
    TypoRule("攻艰克难", "攻坚克难", "常用表述为“攻坚克难”。"),
    TypoRule("尊照执行", "遵照执行", "常用表述为“遵照执行”。"),
    TypoRule("近一步", "进一步", "表示更进一层时通常写作“进一步”。"),
)


def normalize_typo_rules(typo_rules=None) -> list[TypoRule]:
    source_rules = TYPO_RULES if typo_rules is None else typo_rules
    normalized: list[TypoRule] = []
    for rule in source_rules:
        if isinstance(rule, TypoRule):
            normalized.append(rule)
            continue
        if not isinstance(rule, dict) or not rule.get("enabled", True):
            continue
        wrong = str(rule.get("wrong", "")).strip()
        suggestion = str(rule.get("suggestion", "")).strip()
        note = str(rule.get("note", "")).strip()
        if wrong and suggestion:
            normalized.append(TypoRule(wrong=wrong, suggestion=suggestion, note=note))
    return normalized


def load_typo_rules_file(path: Path) -> list[TypoRule]:
    data = json.loads(path.read_text(encoding="utf-8"))
    raw_rules = data.get("rules", data) if isinstance(data, dict) else data
    if not isinstance(raw_rules, list):
        raise ValueError("Typo rules file must contain a list of rules.")
    return normalize_typo_rules(raw_rules)


CJK = r"\u3400-\u4dbf\u4e00-\u9fff"
CHINESE_DIGITS = {
    "〇": 0,
    "零": 0,
    "O": 0,
    "○": 0,
    "一": 1,
    "二": 2,
    "三": 3,
    "四": 4,
    "五": 5,
    "六": 6,
    "七": 7,
    "八": 8,
    "九": 9,
}


def chinese_number_to_int(text: str) -> int | None:
    text = text.strip()
    if not text:
        return None
    if text in CHINESE_DIGITS:
        return CHINESE_DIGITS[text]
    if text == "十":
        return 10
    if text.startswith("十"):
        tail = chinese_number_to_int(text[1:]) if text[1:] else 0
        return 10 + tail if tail is not None else None
    if "十" in text:
        head, tail_text = text.split("十", 1)
        head_value = chinese_number_to_int(head)
        tail_value = chinese_number_to_int(tail_text) if tail_text else 0
        if head_value is None or tail_value is None:
            return None
        return head_value * 10 + tail_value
    value = 0
    for char in text:
        digit = CHINESE_DIGITS.get(char)
        if digit is None:
            return None
        value = value * 10 + digit
    return value


def normalize_date_text(text: str) -> str:
    compact = normalize_text(text)
    numeric = re.fullmatch(r"(\d{4})[./-](\d{1,2})[./-](\d{1,2})", compact)
    if numeric:
        year, month, day = numeric.groups()
        return f"{year}年{int(month)}月{int(day)}日"

    chinese_year = r"[〇零O○一二三四五六七八九]{4}"
    chinese_month_day = r"[〇零一二三四五六七八九十]{1,3}"
    chinese = re.fullmatch(rf"({chinese_year})年({chinese_month_day})月({chinese_month_day})日", compact)
    if chinese:
        year_text, month_text, day_text = chinese.groups()
        year = "".join(str(CHINESE_DIGITS.get(char, 0)) for char in year_text)
        month = chinese_number_to_int(month_text)
        day = chinese_number_to_int(day_text)
        if month and day:
            return f"{year}年{month}月{day}日"
    return text


def count_text_changes(before: str, after: str) -> int:
    return 1 if before != after else 0


def normalize_official_punctuation(text: str, *, is_title: bool = False, is_attachment: bool = False) -> tuple[str, int]:
    """Apply safe GB/T 15834 punctuation normalizations for Chinese official docs."""
    original = text
    fix_count = 0
    new = text.replace("\u00a0", " ").strip()
    fix_count += count_text_changes(text, new)
    text = new

    substitutions = [
        (r"附件\s*[:：]", "附件："),
        (r"([一二三四五六七八九十]+)[.．]\s*", r"\1、"),
        (r"^\(([一二三四五六七八九十]+)\)", r"（\1）"),
        (r"^（([一二三四五六七八九十]+)）[、.．]\s*", r"（\1）"),
        (r"^(\d+)、\s*", r"\1."),
        (r"^\((\d+)\)", r"（\1）"),
        (r"[\[【（(]([12]\d{3})[\]】）)]", r"〔\1〕"),
        (r"\.{3,}", "……"),
        (r"…{2,}", "……"),
        (r"--+", "——"),
        (r"\s+([，。；：？！、）】》])", r"\1"),
        (r"([（【《])\s+", r"\1"),
    ]
    for pattern, repl in substitutions:
        text, changed = re.subn(pattern, repl, text)
        fix_count += changed

    if re.search(rf"[{CJK}]", text):
        cjk_substitutions = [
            (rf"(?<=[{CJK}0-9）】》])\s*,\s*(?=[{CJK}0-9（【《])", "，"),
            (rf"(?<=[{CJK}0-9）】》])\s*;\s*(?=[{CJK}0-9（【《])", "；"),
            (rf"(?<=[{CJK}0-9）】》])\s*:\s*(?=[{CJK}0-9（【《])", "："),
            (rf"(?<=[{CJK}0-9）】》]):$", "："),
            (rf"(?<=[{CJK}0-9）】》])\?", "？"),
            (rf"(?<=[{CJK}0-9）】》])!", "！"),
            (rf"(?<=[{CJK}])\(", "（"),
            (rf"\)(?=[{CJK}，。；：？！、])", "）"),
            (r'"([^"\n]{1,80})"', r"“\1”"),
        ]
        for pattern, repl in cjk_substitutions:
            text, changed = re.subn(pattern, repl, text)
            fix_count += changed

    if is_date_text(text):
        normalized_date = normalize_date_text(text)
        fix_count += count_text_changes(text, normalized_date)
        text = normalized_date

    if is_title:
        title_text = re.sub(r"[，。、；：,.．;:]+$", "", text)
        fix_count += count_text_changes(text, title_text)
        text = title_text

    if is_attachment:
        attachment_text = re.sub(r"[，。；：,.．;:]+$", "", text)
        fix_count += count_text_changes(text, attachment_text)
        text = attachment_text

    return text, fix_count if text != original else 0


def detect_punctuation_issues(text: str) -> list[TypoMatch]:
    matches: list[TypoMatch] = []

    def add(start: int, end: int, wrong: str, suggestion: str, note: str) -> None:
        if start < 0 or end <= start:
            return
        matches.append(TypoMatch(start, end, wrong, suggestion, note, category="疑似标点问题"))

    for match in re.finditer(rf"(?<=[{CJK}])[,;:?!](?=[{CJK}0-9])|(?<=[{CJK}0-9])[,;:?!](?=[{CJK}])", text):
        mapping = {",": "，", ";": "；", ":": "：", "?": "？", "!": "！"}
        wrong = match.group(0)
        add(match.start(), match.end(), wrong, mapping.get(wrong, "中文全角标点"), "中文公文正文应优先使用中文全角标点。")

    for match in re.finditer(r"[？！?!]{4,}", text):
        add(match.start(), match.end(), match.group(0), "最多连续使用三个问号或叹号", "GB/T 15834 不宜过度叠用问号、叹号。")

    for match in re.finditer(r"(……)(等|等等|什么的)", text):
        add(match.start(), match.end(), match.group(0), "省略号和“等”等词择一使用", "GB/T 15834 规定省略号不宜与“等”“等等”等同时使用。")

    if text.count("：") > 1 and not text.startswith("附件："):
        second = text.find("：", text.find("：") + 1)
        add(second, second + 1, "：", "核对冒号层级", "一个句子内部一般不应套用冒号，列举式表述宜分段。")

    pairs = [("“", "”"), ("‘", "’"), ("《", "》"), ("（", "）"), ("〔", "〕"), ("【", "】")]
    for left, right in pairs:
        if text.count(left) != text.count(right):
            index = text.find(left) if text.count(left) > text.count(right) else text.find(right)
            add(index, index + 1, text[index:index + 1], f"补齐与 {left}{right} 对应的另一半", "成对标号应前后配套。")

    return non_overlapping_matches(matches)


def non_overlapping_matches(matches: list[TypoMatch]) -> list[TypoMatch]:
    selected: list[TypoMatch] = []
    occupied: set[int] = set()
    for match in sorted(matches, key=lambda item: (item.start, -(item.end - item.start))):
        span = set(range(match.start, match.end))
        if span & occupied:
            continue
        selected.append(match)
        occupied.update(span)
    return selected


def detect_typos(text: str, typo_rules=None) -> list[TypoMatch]:
    matches: list[TypoMatch] = []
    occupied = [False] * len(text)
    rules = sorted(normalize_typo_rules(typo_rules), key=lambda rule: len(rule.wrong), reverse=True)

    for rule in rules:
        start = 0
        while True:
            index = text.find(rule.wrong, start)
            if index == -1:
                break
            end = index + len(rule.wrong)
            if not any(occupied[index:end]):
                matches.append(
                    TypoMatch(
                        start=index,
                        end=end,
                        wrong=rule.wrong,
                        suggestion=rule.suggestion,
                        note=rule.note,
                        category="疑似错词",
                    )
                )
                for pos in range(index, end):
                    occupied[pos] = True
            start = index + 1

    return sorted(matches, key=lambda match: match.start)


def paragraph_has_drawing(paragraph) -> bool:
    return bool(
        paragraph._element.xpath(
            ".//*[local-name()='drawing' or local-name()='pict' or local-name()='object']"
        )
    )


def run_has_drawing(run) -> bool:
    return bool(
        run._element.xpath(
            ".//*[local-name()='drawing' or local-name()='pict' or local-name()='object']"
        )
    )


def annotate_paragraph_typos(document, paragraph, matches: list[TypoMatch]) -> None:
    if not matches:
        return
    if paragraph_has_drawing(paragraph):
        return

    original_text = paragraph.text
    if not original_text:
        return

    base_run = paragraph.runs[0] if paragraph.runs else paragraph.add_run("")
    font_name = base_run.font.name
    font_size = base_run.font.size
    bold = bool(base_run.bold)
    italic = bool(base_run.italic)

    for run in paragraph.runs:
        run.text = ""

    cursor = 0
    for match in matches:
        if match.start > cursor:
            run = paragraph.add_run(original_text[cursor:match.start])
            copy_run_style(run, font_name=font_name, font_size=font_size, bold=bold, italic=italic)

        typo_run = paragraph.add_run(original_text[match.start:match.end])
        copy_run_style(typo_run, font_name=font_name, font_size=font_size, bold=bold, italic=italic)
        typo_run.font.color.rgb = RGBColor(216, 58, 46)

        comment_text = f"{match.category}：建议改为“{match.suggestion}”。{match.note}"
        document.add_comment(typo_run, text=comment_text, author="公文格式助手", initials="AI")
        cursor = match.end

    if cursor < len(original_text):
        run = paragraph.add_run(original_text[cursor:])
        copy_run_style(run, font_name=font_name, font_size=font_size, bold=bold, italic=italic)


def apply_font_slots(rpr, font_name: str, is_latin: bool = False) -> None:
    """按 OOXML 四槽位分工设置字体。

    西文与数字（拉丁字母、阿拉伯数字、%、小数点、千分位）走 Times New Roman，
    汉字与全角标点走中文字体。`w:hint` 决定全角标点取哪个字形，缺省会退回西文字形。
    """
    rfonts = rpr.get_or_add_rFonts()
    rfonts.set(qn("w:ascii"), LATIN_FONT)
    rfonts.set(qn("w:hAnsi"), LATIN_FONT)
    rfonts.set(qn("w:cs"), LATIN_FONT)
    rfonts.set(qn("w:eastAsia"), font_name)
    rfonts.set(qn("w:hint"), "default" if is_latin else "eastAsia")
    lang = rpr.find(qn("w:lang"))
    if lang is None:
        lang = OxmlElement("w:lang")
        rpr.append(lang)
    lang.set(qn("w:val"), "en-US")
    lang.set(qn("w:eastAsia"), "zh-CN")


def split_mixed_script(text: str):
    """切分为 (片段, 是否纯西文) 序列，全角标点单独成段。

    必须用显式下标 pos 单次推进，不能一边 finditer 一边重绑定被迭代的字符串——
    finditer 惰性迭代最初的字符串对象，重绑定只影响切片来源，会造成文字复制错位。
    """
    pieces = []
    pos = 0
    for m in re.finditer(r"[!-~]+", text):
        if m.start() > pos:
            pieces.extend(_split_fullwidth(text[pos : m.start()]))
        pieces.append((text[m.start() : m.end()], True))
        pos = m.end()
    if pos < len(text):
        pieces.extend(_split_fullwidth(text[pos:]))
    return [p for p in pieces if p[0]]


def _split_fullwidth(segment: str):
    """把全角标点从汉字串里单独切出，保证引号独占 run。"""
    out = []
    buf = ""
    for ch in segment:
        if ch in FULLWIDTH_PUNCT:
            if buf:
                out.append((buf, False))
                buf = ""
            out.append((ch, False))
        else:
            buf += ch
    if buf:
        out.append((buf, False))
    return out


def copy_run_style(run, *, font_name, font_size, bold: bool, italic: bool) -> None:
    if font_name:
        run.font.name = LATIN_FONT
        apply_font_slots(run._element.get_or_add_rPr(), font_name)
    if font_size:
        run.font.size = font_size
    run.bold = bold
    run.italic = italic


def annotate_document_typos(document, typo_rules=None) -> int:
    count = 0
    active_rules = normalize_typo_rules(typo_rules)
    for paragraph in document.paragraphs:
        matches = detect_typos(paragraph.text, active_rules)
        if matches:
            annotate_paragraph_typos(document, paragraph, matches)
            count += len(matches)

    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    matches = detect_typos(paragraph.text, active_rules)
                    if matches:
                        annotate_paragraph_typos(document, paragraph, matches)
                        count += len(matches)
    return count


def annotate_document_punctuation(document) -> int:
    count = 0
    for paragraph in document.paragraphs:
        matches = detect_punctuation_issues(paragraph.text)
        if matches:
            annotate_paragraph_typos(document, paragraph, matches)
            count += len(matches)

    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    matches = detect_punctuation_issues(paragraph.text)
                    if matches:
                        annotate_paragraph_typos(document, paragraph, matches)
                        count += len(matches)
    return count


def annotate_document_quality(document, *, check_typos: bool = False, typo_rules=None) -> tuple[int, int]:
    typo_count = 0
    punctuation_issue_count = 0
    active_rules = normalize_typo_rules(typo_rules) if check_typos else []

    def annotate_paragraph(paragraph) -> None:
        nonlocal typo_count, punctuation_issue_count
        matches = detect_punctuation_issues(paragraph.text)
        punctuation_issue_count += len(matches)
        if check_typos:
            typo_matches = detect_typos(paragraph.text, active_rules)
            matches = non_overlapping_matches(matches + typo_matches)
            typo_count += sum(1 for match in matches if match.category == "疑似错词")
        if matches:
            annotate_paragraph_typos(document, paragraph, matches)

    for paragraph in document.paragraphs:
        annotate_paragraph(paragraph)

    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    annotate_paragraph(paragraph)

    return typo_count, punctuation_issue_count


def set_font(run, font_name: str, size: Pt, bold: bool = False) -> None:
    run.font.name = LATIN_FONT
    run.font.size = size
    run.bold = bold
    run.italic = False
    apply_font_slots(run._element.get_or_add_rPr(), font_name)


def reset_paragraph_text(paragraph, text: str) -> None:
    if paragraph_has_drawing(paragraph):
        return
    for run in paragraph.runs:
        run.text = ""
    if paragraph.runs:
        paragraph.runs[0].text = text
    else:
        paragraph.add_run(text)


def format_paragraph(
    paragraph,
    *,
    alignment,
    font_name: str,
    size: Pt,
    bold: bool = False,
    first_indent=None,
    right_indent=None,
    line_spacing=Pt(31),
    space_before=Pt(0),
    space_after=Pt(0),
) -> None:
    paragraph.alignment = alignment
    pf = paragraph.paragraph_format
    pf.line_spacing = line_spacing
    pf.space_before = space_before
    pf.space_after = space_after
    pf.first_line_indent = first_indent
    pf.left_indent = Pt(0)
    pf.right_indent = right_indent
    if not paragraph.runs:
        paragraph.add_run("")
    if paragraph_has_drawing(paragraph):
        for run in paragraph.runs:
            if not run_has_drawing(run):
                set_font(run, font_name, size, bold)
        return
    # 含中西混排或全角引号的段落重建 run，让西文与引号各自独立成run
    text = paragraph.text
    if any(ch in text for ch in FULLWIDTH_PUNCT) or re.search(r"[!-~]", text):
        rebuild_runs_by_script(paragraph, font_name, size, bold)
        return
    for run in paragraph.runs:
        if run_has_drawing(run):
            continue
        set_font(run, font_name, size, bold)


def rebuild_runs_by_script(paragraph, font_name: str, size: Pt, bold: bool) -> None:
    """按中西文与全角标点切分重建 run。

    两个必要条件（缺一即失败）：
    1. 全角引号独占一个 run，不与汉字正文混在同一个 run
    2. 每个 run 声明 w:hint，中文 run 为 eastAsia、纯西文 run 为 default
    """
    text = paragraph.text
    if not text.strip():
        for run in paragraph.runs:
            set_font(run, font_name, size, bold)
        return
    for run in list(paragraph.runs):
        run._element.getparent().remove(run._element)
    for piece, is_latin in split_mixed_script(text):
        run = paragraph.add_run(piece)
        run.font.size = size
        run.bold = bold
        run.italic = False
        run.font.name = LATIN_FONT
        apply_font_slots(run._element.get_or_add_rPr(), font_name, is_latin)


def format_picture_paragraph(paragraph, *, alignment=WD_ALIGN_PARAGRAPH.CENTER) -> None:
    paragraph.alignment = alignment
    pf = paragraph.paragraph_format
    pf.line_spacing = 1.0
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.first_line_indent = None
    pf.left_indent = Pt(0)
    pf.right_indent = Pt(0)


def is_primary_heading(text: str) -> bool:
    return bool(re.match(r"^[一二三四五六七八九十]+、", text))


def starts_with_secondary_number(text: str) -> bool:
    return bool(re.match(r"^（[一二三四五六七八九十]+）", text))


def is_secondary_heading(text: str) -> bool:
    if not starts_with_secondary_number(text):
        return False
    compact = normalize_text(text)
    return len(compact) <= 25 and not any(mark in compact for mark in "，,；;：:")


def is_tertiary_heading(text: str) -> bool:
    compact = normalize_text(text)
    return bool(re.match(r"^\d+\.", text)) and len(compact) <= 25 and not any(mark in compact for mark in "，,；;：:")


def is_date_text(text: str) -> bool:
    compact = normalize_text(text)
    chinese_year = r"[〇零一二三四五六七八九十]{4}"
    chinese_month_day = r"[〇零一二三四五六七八九十]{1,3}"
    return bool(
        re.fullmatch(r"\d{4}年\d{1,2}月\d{1,2}日", compact)
        or re.fullmatch(r"\d{4}[./-]\d{1,2}[./-]\d{1,2}", compact)
        or re.fullmatch(rf"{chinese_year}年{chinese_month_day}月{chinese_month_day}日", compact)
    )


def normalize_text(text: str) -> str:
    return re.sub(r"\s+", "", text).replace("\u00a0", "")


def is_contact_line(text: str) -> bool:
    compact = normalize_text(text)
    return "联系人：" in compact and "联系电话：" in compact


def is_person_name_text(text: str) -> bool:
    compact = normalize_text(text)
    if not re.fullmatch(r"[\u4e00-\u9fff]{2,3}", compact):
        return False
    single_surnames = (
        "赵钱孙李周吴郑王冯陈褚卫蒋沈韩杨朱秦尤许何吕施张孔曹严华"
        "金魏陶姜戚谢邹喻柏水窦章云苏潘葛奚范彭郎鲁韦昌马苗凤花方"
        "俞任袁柳鲍史唐费廉岑薛雷贺倪汤滕殷罗毕郝邬安常乐于时傅皮"
        "卞齐康伍余元卜顾孟平黄和穆萧尹姚邵湛汪祁毛禹狄米贝明臧计"
        "伏成戴谈宋庞熊纪舒屈项祝董梁杜阮蓝闵席季麻强贾路娄危江童"
        "颜郭梅盛林刁钟徐邱骆高夏蔡田胡凌霍虞万支柯昝管卢莫经房裘"
        "缪干解应宗丁宣邓郁单杭洪包诸左石崔吉龚程邢裴陆荣翁荀羊於"
        "惠甄曲家封芮羿储靳汲邴糜松井段富巫乌焦巴弓牧隗山谷车侯宓"
        "蓬全郗班仰秋仲伊宫宁仇栾暴甘斜厉戎祖武符刘景詹龙叶幸司韶"
        "郜黎蓟溥印宿白怀蒲邰从鄂索咸籍赖卓蔺屠蒙池乔阴胥能苍双闻"
        "莘党翟谭贡劳逄姬申扶堵冉宰郦雍郤璩桑桂濮牛寿通边扈燕冀郏"
        "浦尚农温别庄晏柴瞿阎充慕连茹习宦艾鱼容向古易慎戈廖庾终暨"
        "居衡步都耿满弘匡国文寇广禄阙东欧利师巩聂关荆游竺权逯盖益"
        "桓公"
    )
    compound_surnames = (
        "欧阳",
        "太史",
        "端木",
        "上官",
        "司马",
        "东方",
        "独孤",
        "南宫",
        "万俟",
        "闻人",
        "夏侯",
        "诸葛",
        "尉迟",
        "公羊",
        "赫连",
        "澹台",
        "皇甫",
        "宗政",
        "濮阳",
        "公冶",
        "太叔",
        "申屠",
        "公孙",
        "慕容",
        "仲孙",
        "钟离",
        "长孙",
        "宇文",
        "司徒",
        "鲜于",
        "司空",
        "闾丘",
        "子车",
        "亓官",
        "司寇",
        "巫马",
        "公西",
        "颛孙",
        "壤驷",
        "公良",
        "漆雕",
        "乐正",
        "宰父",
        "谷梁",
        "拓跋",
        "夹谷",
        "轩辕",
        "令狐",
        "段干",
        "百里",
        "呼延",
        "东郭",
        "南门",
        "羊舌",
        "微生",
    )
    non_name_endings = (
        "通知",
        "报告",
        "方案",
        "意见",
        "建议",
        "说明",
        "材料",
        "正文",
        "附件",
        "时间",
        "地点",
        "电话",
        "邮箱",
    )
    if compact.endswith(non_name_endings):
        return False
    if len(compact) == 2:
        return compact[0] in single_surnames
    if len(compact) == 3:
        return compact[0] in single_surnames or compact[:2] in compound_surnames
    return False


def is_issuer_text(text: str) -> bool:
    compact = normalize_text(text)
    if not compact or len(compact) > 30:
        return False
    if is_date_text(compact):
        return False
    if any(mark in compact for mark in "，,。；;：:、！？!?（）()[]【】"):
        return False
    if re.search(r"(局|厅|部|委|办|处|科|中心|学院|学校|大学|公司|集团|协会|委员会|人民政府|办公室|课题组)$", compact):
        return True
    return is_person_name_text(compact)


def detect_signature_block(paragraphs) -> set[int]:
    nonempty = [i for i, p in enumerate(paragraphs) if p.text.strip()]
    if len(nonempty) < 2:
        return set()

    last_index = nonempty[-1]
    issuer_index = nonempty[-2]
    date_text = paragraphs[last_index].text.strip()
    issuer_text = paragraphs[issuer_index].text.strip()

    if not is_date_text(date_text):
        return set()
    if not is_issuer_text(issuer_text):
        return set()
    return {issuer_index, last_index}


def detect_subtitle_indexes(paragraphs, title_indexes: set[int], recipient_index: int) -> set[int]:
    if not title_indexes:
        return set()
    last_title_index = max(title_indexes)
    if recipient_index == -1:
        candidate_range = range(last_title_index + 1, min(last_title_index + 3, len(paragraphs)))
    else:
        candidate_range = range(last_title_index + 1, recipient_index)

    subtitle_indexes = set()
    for index in candidate_range:
        text = paragraphs[index].text.strip()
        if not text:
            continue
        if text.endswith("：") or text.startswith("附件") or is_primary_heading(text):
            break
        if len(normalize_text(text)) <= 40:
            subtitle_indexes.add(index)
    return subtitle_indexes


def detect_title_indexes(paragraphs, nonempty: list[int]) -> set[int]:
    if not nonempty:
        return set()

    first_index = nonempty[0]
    for index in nonempty[1:6]:
        text = paragraphs[index].text.strip()
        if text.endswith("："):
            return set(nonempty[:nonempty.index(index)])
        if is_primary_heading(text) or starts_with_secondary_number(text) or text.startswith("附件"):
            break
    return {first_index}


def normalize_document_punctuation(document, title_indexes: set[int] | None = None) -> tuple[int, int]:
    title_indexes = title_indexes or set()
    punctuation_fix_count = 0
    normalized_paragraph_count = 0

    for index, paragraph in enumerate(document.paragraphs):
        text = paragraph.text
        if not text or paragraph_has_drawing(paragraph):
            continue
        normalized, fixes = normalize_official_punctuation(
            text,
            is_title=index in title_indexes,
            is_attachment=text.strip().startswith("附件"),
        )
        if normalized != text:
            reset_paragraph_text(paragraph, normalized)
            normalized_paragraph_count += 1
            punctuation_fix_count += fixes

    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    text = paragraph.text
                    if not text or paragraph_has_drawing(paragraph):
                        continue
                    normalized, fixes = normalize_official_punctuation(
                        text,
                        is_attachment=text.strip().startswith("附件"),
                    )
                    if normalized != text:
                        reset_paragraph_text(paragraph, normalized)
                        normalized_paragraph_count += 1
                        punctuation_fix_count += fixes

    return punctuation_fix_count, normalized_paragraph_count


def add_page_number_field(paragraph) -> None:
    """插入 `— 1 —` 形式的页码域。

    域结构必须拆成三个独立 run——`begin` / `instrText` / `separate` / `end`
    塞进同一个 run 是非法结构，Word 能容错渲染，但 LibreOffice 等严格实现
    会解析异常，导致「首页不编码」等 titlePg 行为失效。
    """
    paragraph.add_run("— ")

    begin_run = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    begin_run._r.append(begin)

    instr_run = paragraph.add_run()
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    instr_run._r.append(instr)

    sep_run = paragraph.add_run()
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    sep_run._r.append(separate)

    cached_run = paragraph.add_run("1")

    end_run = paragraph.add_run()
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    end_run._r.append(end)

    paragraph.add_run(" —")
    return cached_run


def format_footer_paragraph(paragraph, alignment) -> None:
    paragraph.alignment = alignment
    pf = paragraph.paragraph_format
    pf.line_spacing = Pt(18)
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    for run in paragraph.runs:
        set_font(run, "宋体", Pt(14), False)


def clear_paragraph(paragraph) -> None:
    """彻底清空段落：删掉全部 run 及其内部的域元素残留。

    仅把 `run.text` 置空是不够的——若run 内还挂着 `w:fldChar` /
    `w:instrText`，域依然生效，首页会照样显示页码。
    """
    for run in list(paragraph.runs):
        run._element.getparent().remove(run._element)


def setup_page_numbers(doc) -> None:
    """设置页码：首页不编码，奇数页居右、偶数页居左，样式 `— 1 —`。

    `odd_and_even_pages_header_footer` 会让Word 区分奇偶页footer，
    必须配合 `different_first_page_header_footer` 才能实现首页不编码。
    """
    doc.settings.odd_and_even_pages_header_footer = True
    for section in doc.sections:
        section.different_first_page_header_footer = True

        # 首页：彻底清空，且不设任何对齐（避免继承正文格式）
        first = section.first_page_footer.paragraphs[0]
        clear_paragraph(first)
        first.alignment = WD_ALIGN_PARAGRAPH.LEFT

        odd = section.footer.paragraphs[0]
        clear_paragraph(odd)
        add_page_number_field(odd)
        format_footer_paragraph(odd, WD_ALIGN_PARAGRAPH.RIGHT)

        even = section.even_page_footer.paragraphs[0]
        clear_paragraph(even)
        add_page_number_field(even)
        format_footer_paragraph(even, WD_ALIGN_PARAGRAPH.LEFT)


def format_docx_with_report(
    input_path: Path,
    output_path: Path,
    *,
    title_font: str,
    body_font: str,
    heading_font: str,
    secondary_heading_font: str,
    check_typos: bool = False,
    typo_rules=None,
) -> FormatReport:
    doc = Document(input_path)

    body_size = Pt(16)   # 三号
    title_size = Pt(22)  # 二号
    line_spacing = Pt(30)
    two_chars = Pt(32)
    four_chars = Pt(64)

    for section in doc.sections:
        section.start_type = WD_SECTION_START.NEW_PAGE
        section.page_width = Mm(210)
        section.page_height = Mm(297)
        section.top_margin = Mm(35)
        section.bottom_margin = Mm(29)
        section.left_margin = Mm(25.5)
        section.right_margin = Mm(25.5)
        section.header_distance = Mm(15)
        section.footer_distance = Mm(22)

    normal = doc.styles["Normal"]
    normal.font.size = body_size
    apply_font_slots(normal._element.get_or_add_rPr(), body_font)

    punctuation_fix_count, normalized_paragraph_count = normalize_document_punctuation(doc)

    paragraphs = doc.paragraphs
    nonempty = [i for i, p in enumerate(paragraphs) if p.text.strip()]
    if not nonempty:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        doc.save(output_path)
        return FormatReport()

    title_indexes = detect_title_indexes(paragraphs, nonempty)
    title_fix_count, title_normalized_count = normalize_document_punctuation(doc, title_indexes)
    punctuation_fix_count += title_fix_count
    normalized_paragraph_count += title_normalized_count

    paragraphs = doc.paragraphs
    nonempty = [i for i, p in enumerate(paragraphs) if p.text.strip()]
    title_indexes = detect_title_indexes(paragraphs, nonempty)
    last_title_index = max(title_indexes)
    after_title = [i for i in nonempty if i > last_title_index]
    recipient_index = after_title[0] if after_title and paragraphs[after_title[0]].text.strip().endswith("：") else -1
    subtitle_indexes = detect_subtitle_indexes(paragraphs, title_indexes, recipient_index)
    signature_indexes = detect_signature_block(paragraphs)
    issuer_index = min(signature_indexes) if signature_indexes else -1

    for index, paragraph in enumerate(paragraphs):
        text = paragraph.text.strip()

        if paragraph_has_drawing(paragraph) and not text:
            format_picture_paragraph(paragraph, alignment=paragraph.alignment or WD_ALIGN_PARAGRAPH.CENTER)
        elif not text:
            format_paragraph(
                paragraph,
                alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
                font_name=body_font,
                size=body_size,
                first_indent=None,
                line_spacing=line_spacing,
            )
        elif index in title_indexes:
            format_paragraph(
                paragraph,
                alignment=WD_ALIGN_PARAGRAPH.CENTER,
                font_name=title_font,
                size=title_size,
                bold=True,
                first_indent=None,
                line_spacing=line_spacing,
                space_after=Pt(0),
            )
        elif index in signature_indexes:
            format_paragraph(
                paragraph,
                alignment=WD_ALIGN_PARAGRAPH.RIGHT,
                font_name=body_font,
                size=body_size,
                first_indent=None,
                right_indent=four_chars,
                line_spacing=line_spacing,
                space_before=Pt(60) if index == issuer_index else Pt(0),
            )
        elif index in subtitle_indexes:
            format_paragraph(
                paragraph,
                alignment=WD_ALIGN_PARAGRAPH.CENTER,
                font_name=secondary_heading_font,
                size=body_size,
                first_indent=None,
                line_spacing=line_spacing,
                space_after=Pt(30) if index == max(subtitle_indexes) else Pt(0),
            )
        elif index == recipient_index:
            format_paragraph(
                paragraph,
                alignment=WD_ALIGN_PARAGRAPH.LEFT,
                font_name=body_font,
                size=body_size,
                first_indent=None,
                line_spacing=line_spacing,
            )
        elif is_primary_heading(text):
            format_paragraph(
                paragraph,
                alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
                font_name=heading_font,
                size=body_size,
                first_indent=two_chars,
                line_spacing=line_spacing,
            )
        elif is_secondary_heading(text):
            format_paragraph(
                paragraph,
                alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
                font_name=secondary_heading_font,
                size=body_size,
                first_indent=two_chars,
                line_spacing=line_spacing,
            )
        elif is_tertiary_heading(text):
            format_paragraph(
                paragraph,
                alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
                font_name=body_font,
                size=body_size,
                bold=True,
                first_indent=two_chars,
                line_spacing=line_spacing,
            )
        elif text.startswith("附件"):
            format_paragraph(
                paragraph,
                alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
                font_name=body_font,
                size=body_size,
                first_indent=two_chars,
                line_spacing=line_spacing,
                space_before=Pt(30),
            )
        elif is_contact_line(text):
            format_paragraph(
                paragraph,
                alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
                font_name=body_font,
                size=body_size,
                first_indent=two_chars,
                line_spacing=line_spacing,
            )
        else:
            format_paragraph(
                paragraph,
                alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
                font_name=body_font,
                size=body_size,
                first_indent=two_chars,
                line_spacing=line_spacing,
            )

    for table in doc.tables:
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        for row in table.rows:
            for cell in row.cells:
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                for paragraph in cell.paragraphs:
                    if paragraph_has_drawing(paragraph):
                        format_picture_paragraph(paragraph)
                        continue
                    format_paragraph(
                        paragraph,
                        alignment=WD_ALIGN_PARAGRAPH.CENTER,
                        font_name=body_font,
                        size=body_size,
                        first_indent=None,
                        line_spacing=line_spacing,
                    )

    typo_count, punctuation_issue_count = annotate_document_quality(
        doc,
        check_typos=check_typos,
        typo_rules=typo_rules,
    )

    setup_page_numbers(doc)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(output_path)
    return FormatReport(
        typo_count=typo_count,
        punctuation_fix_count=punctuation_fix_count,
        punctuation_issue_count=punctuation_issue_count,
        normalized_paragraph_count=normalized_paragraph_count,
    )


def format_docx(
    input_path: Path,
    output_path: Path,
    *,
    title_font: str,
    body_font: str,
    heading_font: str,
    secondary_heading_font: str,
    check_typos: bool = False,
    typo_rules=None,
) -> int:
    report = format_docx_with_report(
        input_path,
        output_path,
        title_font=title_font,
        body_font=body_font,
        heading_font=heading_font,
        secondary_heading_font=secondary_heading_font,
        check_typos=check_typos,
        typo_rules=typo_rules,
    )
    return report.typo_count


def convert_legacy_doc_to_docx(input_path: Path, work_dir: Path) -> Path:
    output_path = work_dir / f"{input_path.stem}.docx"
    textutil = shutil.which("textutil")
    if textutil:
        subprocess.run(
            [textutil, "-convert", "docx", "-output", str(output_path), str(input_path)],
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
        )
        if output_path.exists():
            return output_path

    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if soffice:
        subprocess.run(
            [
                soffice,
                "--headless",
                "--convert-to",
                "docx",
                "--outdir",
                str(work_dir),
                str(input_path),
            ],
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
        )
        if output_path.exists():
            return output_path

    raise RuntimeError("Cannot convert .doc input to .docx: neither textutil nor LibreOffice is available.")


def default_output_path(input_path: Path) -> Path:
    return input_path.with_name(f"{input_path.stem}_公文格式.docx")


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", type=Path, help="Input .docx or legacy .doc file")
    parser.add_argument("output", type=Path, nargs="?", help="Output .docx file")
    parser.add_argument("--title-font", default="方正小标宋_GBK")
    parser.add_argument("--body-font", default="方正仿宋_GBK")
    parser.add_argument("--heading-font", default="方正黑体_GBK")
    parser.add_argument("--secondary-heading-font", default="方正楷体_GBK")
    parser.add_argument("--check-typos", action="store_true", help="Mark common typo words in red and add Word comments.")
    parser.add_argument("--typo-rules-file", type=Path, help="JSON file containing typo rules.")
    args = parser.parse_args()

    input_path = args.input.expanduser().resolve()
    output_path = (args.output or default_output_path(input_path)).expanduser().resolve()

    if input_path.suffix.lower() not in {".doc", ".docx"}:
        raise SystemExit("Input file must be a .docx or .doc file.")
    if output_path.suffix.lower() != ".docx":
        raise SystemExit("Output file must be a .docx file.")
    if input_path == output_path:
        raise SystemExit("Output path must not be the same as input path.")
    if not input_path.exists():
        raise SystemExit(f"Input file does not exist: {input_path}")

    with tempfile.TemporaryDirectory(prefix="official-docx-format-") as tmpdir:
        formatter_input = input_path
        if input_path.suffix.lower() == ".doc":
            formatter_input = convert_legacy_doc_to_docx(input_path, Path(tmpdir))

        format_docx(
            formatter_input,
            output_path,
            title_font=args.title_font,
            body_font=args.body_font,
            heading_font=args.heading_font,
            secondary_heading_font=args.secondary_heading_font,
            check_typos=args.check_typos,
            typo_rules=load_typo_rules_file(args.typo_rules_file) if args.typo_rules_file else None,
        )
    print(output_path)


if __name__ == "__main__":
    main()
