#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
中建四局公文格式化 — 段落层级检测 + 标题判断

包含：
- detect_level(): 文本编号前缀 → 层级映射
- is_main_title(): 主标题识别
- HeadingCounter: 标题编号计数器
- apply_heading_format(): 标题格式应用
- resolve_final_level(): 综合层级判定（合并原 _precompute_heading + 主循环逻辑）
"""

import re
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from constants import *


# Pt/Cm 包装
SIZE_ERHAO  = Pt(SIZE_ERHAO)
SIZE_SANHAO = Pt(SIZE_SANHAO)


def set_run_font(run, cn_font, size_pt, bold=False, color=None, num_font=None):
    """设置 run 的字体、字号、加粗、颜色。
    cn_font = 中文字体, num_font = 数字/字母字体（默认仿宋_GB2312）"""
    if num_font is None:
        num_font = FONT_FANGSONG
    run.font.name = num_font
    run.font.size = size_pt
    run.font.bold = bold
    if color:
        run.font.color.rgb = color
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:eastAsia'), cn_font)
    rFonts.set(qn('w:ascii'), num_font)
    rFonts.set(qn('w:hAnsi'), num_font)


def set_para_spacing(para, twips=LINE_SPACING_TWIPS):
    """设置段落行距"""
    pPr = para._p.get_or_add_pPr()
    spacing = pPr.find(qn('w:spacing'))
    if spacing is None:
        spacing = OxmlElement('w:spacing')
        pPr.append(spacing)
    spacing.set(qn('w:line'), str(twips))
    spacing.set(qn('w:lineRule'), 'exact')
    spacing.set(qn('w:before'), '0')
    spacing.set(qn('w:after'), '0')


def set_para_indent(para, first_line_chars=2, char_size_pt=16):
    """设置首行缩进"""
    dxa = int(first_line_chars * char_size_pt * 20)
    pPr = para._p.get_or_add_pPr()
    ind = pPr.find(qn('w:ind'))
    if ind is None:
        ind = OxmlElement('w:ind')
        pPr.append(ind)
    ind.set(qn('w:firstLine'), str(dxa))


def clean_text(text):
    """清理 Markdown 遗留标记"""
    text = re.sub(r'^#{1,6}\s*', '', text)
    text = re.sub(r'\*\*(.*?)\*\*', r'\1', text)
    text = re.sub(r'\*(.*?)\*', r'\1', text)
    text = re.sub(r'[`~_>|\\^]', '', text)
    text = re.sub(r'  +', ' ', text).strip()
    return text


def detect_level(text):
    """根据文本编号前缀判断层级

    一、 → h1, （一） → h2, 1. → h3, （1） → h4, ① → h5, 其他 → body
    """
    t = text.strip()
    if re.match(r'^[一二三四五六七八九十]+、', t):
        return 'h1'
    if re.match(r'^（[一二三四五六七八九十]+）', t):
        return 'h2'
    if re.match(r'^\d+[.、．](?!\d)\s*', t):
        return 'h3'
    if re.match(r'^（\d+）', t):
        return 'h4'
    if re.match(r'^[①②③④⑤⑥⑦⑧⑨⑩]', t):
        return 'h5'
    return 'body'


def is_main_title(text, para_index=None):
    """判断文本是否是公文主标题

    主标题特征：无编号前缀、无句号分号、8-40字、非日期、非数据行、非标签字段、非问候语
    位置限制：只在前 MAIN_TITLE_MAX_POSITION 个段落内检测主标题
    """
    if para_index is not None and para_index > MAIN_TITLE_MAX_POSITION:
        return False
    t = text.strip()
    if not t:
        return False
    if re.match(r'^\d{4}年\d{1,2}月\d{0,2}日?\s*$', t):
        return False
    if is_date_line(t):
        return False
    if re.search(r'月底', t):
        return False
    if has_text_number_prefix(t):
        return False
    if is_pure_data_line(t):
        return False
    if is_labeled_field(t):
        return False
    if is_greeting(t):
        return False
    if '。' in t or '；' in t:
        return False
    if len(t) < MAIN_TITLE_MIN_CHARS:
        return False
    if len(t) > MAIN_TITLE_MAX_CHARS:
        return False
    return True


class HeadingCounter:
    """标题编号计数器

    跟踪各级标题编号，自动重置下级计数。
    """
    def __init__(self):
        self.h1 = self.h2 = self.h3 = self.h4 = self.h5 = 0

    def next(self, level):
        if level == 'h1':
            self.h1 += 1; self.h2 = self.h3 = self.h4 = self.h5 = 0
            return int_to_cn(self.h1) + '、'
        elif level == 'h2':
            self.h2 += 1; self.h3 = self.h4 = self.h5 = 0
            return f'（{int_to_cn(self.h2)}）'
        elif level == 'h3':
            self.h3 += 1; self.h4 = self.h5 = 0
            return f'{self.h3}.'
        elif level == 'h4':
            self.h4 += 1; self.h5 = 0
            return f'（{self.h4}）'
        elif level == 'h5':
            self.h5 += 1
            idx = self.h5 - 1
            return CIRCLE_NUMBERS[idx] if idx < len(CIRCLE_NUMBERS) else f'({self.h5})'
        return ''


def apply_heading_format(para, level, text, prefix='', no_indent=False, preserve_bold=False, bold_runs=None):
    """应用标题格式到段落

    Args:
        para: python-docx 段落对象
        level: 层级 ('title', 'h1'-'h5', 'body')
        text: 显示文本
        prefix: 编号前缀（可选）
        no_indent: 是否取消首行缩进
        preserve_bold: 是否保留原文加粗（单 run 模式）
        bold_runs: [(len, is_bold), ...] 逐 run 加粗信息（多 run 模式）
    """
    para.clear()
    para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    set_para_spacing(para)
    if not no_indent:
        set_para_indent(para, 2)
    display = prefix + text
    font_map = {
        'h1': FONT_HEITI,
        'h2': FONT_KAITI,
        'title': FONT_XIAOBIAOSONG,
    }
    cn_font = font_map.get(level, FONT_FANGSONG)
    # 数字/字母字体随层级：title→方正小标宋, h1→黑体, h2→楷体, 其余→仿宋
    num_font_map = {'title': FONT_XIAOBIAOSONG, 'h1': FONT_HEITI, 'h2': FONT_KAITI}
    num_font = num_font_map.get(level, FONT_FANGSONG)

    # 逐 run 重建（保留局部加粗）
    has_mixed = (bold_runs and len(bold_runs) > 1
                 and any(b for _, b in bold_runs)
                 and not all(b for _, b in bold_runs))
    if has_mixed:
        pos = 0
        first = True
        for rlen, rbold in bold_runs:
            chunk = text[pos:pos + rlen]
            if first:
                chunk = prefix + chunk
                first = False
            if chunk.strip():
                run = para.add_run(chunk)
                set_run_font(run, cn_font, SIZE_SANHAO, bold=rbold, num_font=num_font)
            pos += rlen
        return

    run = para.add_run(display)

    if level == 'title':
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        para.paragraph_format.first_line_indent = None
        set_run_font(run, cn_font, SIZE_ERHAO, bold=False, num_font=num_font)
    else:
        set_run_font(run, cn_font, SIZE_SANHAO, bold=preserve_bold, num_font=num_font)


def resolve_final_level(idx, item, promote_x_to_h1, promote_body_indices,
                         num_to_abstract, abstract_num_defs):
    """综合层级判定 — 合并原 _precompute_heading 闭包和主循环层级判断

    统一入口，消除两处重复的 promote / 降级 / wnl 回退逻辑。

    Args:
        idx: 段落索引
        item: 段落数据元组 ('p', text, is_bold, word_num_level, num_id, num_ilvl)
        promote_x_to_h1: 是否启用 X.→h1 提升模式
        promote_body_indices: 需要提升为 h1 的 body 段落索引集合
        num_to_abstract: numId → abstractNumId 映射
        abstract_num_defs: abstractNumId → {ilvl: (fmt, lvlText)} 映射

    Returns:
        level 字符串 ('title', 'h1'-'h5', 'body', None)
        None 表示该段落不是标题（空段落/表格等）
    """
    if item[0] != 'p':
        return None
    raw_text = item.text
    text = clean_text(raw_text)
    if not text:
        return None

    # 主标题（仅限文档前部）
    if is_main_title(text, para_index=idx):
        return 'title'

    # 基础层级：文本编号前缀
    level = detect_level(text)
    wnl = item.word_num_level

    # Word 编号回退（仅在 detect_level 无法识别时）
    if wnl is not None and level == 'body':
        looks_like_title = (len(text) <= WNL_HEADING_MAX_CHARS and not re.search(r'[。；]', text))
        if looks_like_title:
            level = wnl

    # promote 模式：X. → h1
    if promote_x_to_h1 and level == 'h3':
        is_long_sentence = '。' in text and len(text) > WNL_HEADING_MAX_CHARS
        if (re.match(r'^\d+[.、．]\s*\S', text)
            and not re.match(r'^\d+\.\d+', text)
            and not is_verb_after_number(text)
            and not is_long_sentence):
            level = 'h1'

    # promote 模式：body → h1（夹在编号组之间的短段落）
    if promote_x_to_h1 and level == 'body' and wnl is None and idx in promote_body_indices:
        level = 'h1'

    # promote 模式：Word 编号隐藏了编号的短段落
    if promote_x_to_h1 and wnl is not None and wnl in ('h2', 'h3', 'h4'):
        _has_prefix = has_text_number_prefix(text)
        looks_like_heading = (
            not _has_prefix
            and len(text) <= SHORT_H1_MAX_CHARS
            and not re.search(r'[。；！？]', text)
            and not re.match(r'^\d+[.、．]', text)
            and not re.match(r'^\d{1,2}:\d{2}', text)     # 时间项（09:00…）
            and not re.match(r'^第[一二三四五六七八九十\d]+天', text)  # "第一天"
        )
        if looks_like_heading:
            level = 'h1'

    # Word 编号段落：无文本前缀且长度 >BODY_DEMOTE_MIN_CHARS → 降级为 body
    if wnl is not None and level in ('h1', 'h2', 'h3', 'h4', 'h5'):
        _has_prefix = has_text_number_prefix(text)
        if not _has_prefix and len(text) > BODY_DEMOTE_MIN_CHARS:
            _fmt, _ = get_numbering_info(item, num_to_abstract, abstract_num_defs)
            if _fmt not in ('chineseCounting', 'chineseCountingThousand',
                           'ideographDigital', 'ideographEnclosedCircle'):
                return None

    return level
