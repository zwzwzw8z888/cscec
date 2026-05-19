#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
中建四局公文格式化 — 格式常量 + 全局规则集中声明

所有格式常量和全局规则函数统一在此定义，
其他模块通过 from constants import * 使用。
"""

import re
from typing import NamedTuple

# ────────────────────────── 段落数据结构 ──────────────────────────

class ParagraphItem(NamedTuple):
    """paragraphs_text 中每个段落的标准数据结构"""
    kind: str             # 'p' 段落 或 'tbl' 表格
    text: str             # 段落文本（表格时为 ''）
    is_bold: bool         # 全文是否加粗
    word_num_level: str   # Word 编号层级 (h1~h5) 或 None
    num_id: str           # Word numId 或 '0'
    num_ilvl: str         # Word ilvl 或 '0'
    bold_runs: list       # [(length, is_bold), ...] 逐段加粗信息
    elem: object          # 原始 XML 元素

# ────────────────────────── 格式常量 ──────────────────────────
FONT_FANGSONG     = "仿宋_GB2312"
FONT_HEITI        = "黑体"
FONT_KAITI        = "楷体_GB2312"
FONT_XIAOBIAOSONG = "方正小标宋简体"

SIZE_CHUHAO  = 42  # Pt(42) — 调用处自行包装
SIZE_ERHAO   = 22
SIZE_SANHAO  = 16
SIZE_XIAOSI  = 12

LINE_SPACING_TWIPS = 578  # 28.9磅 = 578 twips (579=28.95磅, 多1twips)

MARGIN_TOP    = 3.7   # cm
MARGIN_BOTTOM = 3.5
MARGIN_LEFT   = 2.8
MARGIN_RIGHT  = 2.6

def int_to_cn(n):
    """数字转中文小写：1→一, 10→十, 21→二十一, 99→九十九"""
    if n <= 0: return ''
    digit_cn = ['', '一','二','三','四','五','六','七','八','九']
    if n <= 9: return digit_cn[n]
    if n == 10: return '十'
    if n <= 19: return '十' + digit_cn[n - 10]
    tens = n // 10
    ones = n % 10
    return digit_cn[tens] + '十' + (digit_cn[ones] if ones else '')

CN_NUMBERS = [int_to_cn(i+1) for i in range(30)]  # 兼容旧代码，实际用int_to_cn
CNUM = {str(i+1): int_to_cn(i+1) for i in range(30)}
CNUM_TO_INT = {int_to_cn(i+1): i+1 for i in range(30)}
CIRCLE_NUMBERS = ['①','②','③','④','⑤','⑥','⑦','⑧','⑨','⑩',
                  '⑪','⑫','⑬','⑭','⑮','⑯','⑰','⑱','⑲','⑳']


# ────────────────────────── 全局规则函数 ──────────────────────────
# 集中管理所有全局规则，单一来源，避免散落导致冲突

def is_date_line(text):
    """全局规则：日期行排除（带括号格式）
    匹配格式：（2026年4月28日）或 (2026年4月28日)
    用途：is_main_title / is_short_title / is_after_main_title 三处统一调用
    """
    return bool(re.match(r'^[（(]\d{4}年\d{1,2}月\d{1,2}日[）)]$', text.strip()))


def is_signature_date(text):
    """全局规则：落款日期行检测（不带括号）
    匹配格式：2026年4月23日
    用途：落款格式处理（右对齐、右空4字）
    """
    return bool(re.match(r'^\d{4}年\d{1,2}月\d{1,2}日\s*$', text.strip()))


def is_table_title(text):
    """全局规则：表格标题识别
    匹配格式：表1、表2、表3-1 等
    用途：正文段落居中 + 不缩进
    """
    return bool(re.match(r'^表\s*\d+', text.strip()))


def is_pure_data_line(text):
    """纯数字/数据行排除
    匹配格式：纯数字、标点、运算符组成的行
    用途：is_main_title / is_short_title 排除
    """
    return bool(re.match(r'^[\d,.\-+/：:；，。、]+$', text.strip()))


def should_preserve_bold(level, is_bold, text):
    """全局规则1：原文加粗的，格式化后也保持加粗。

    标题(h1-h5)及正文均适用。不加过滤条件——对"一是/二是"、
    "1.xxx"、"商务部"等主观加粗一视同仁，原文怎样就怎样。
    """
    return bool(is_bold)


def is_greeting(text):
    """问候语/主送机关识别
    匹配格式：含称呼关键词 + 以冒号结尾的短文本
    包含：问候语（尊敬、各位领导）、主送机关（局属各单位、总部各部门）
    用途：正文段落不缩进
    """
    greeting_kw = '领导|同事|各位|尊敬|您好|下午好|上午好|你好'
    recipient_kw = '单位|部门|公司|分公司|事业部'
    t = text.strip()
    return bool(
        re.match(r'^.{2,30}[：:]$', t) and (
            re.search(greeting_kw, t) or re.search(recipient_kw, t)
        )
    )


def has_text_number_prefix(text):
    """文本编号前缀检测
    匹配：一、 / （一） / 1. / （1） / ① 等格式
    用途：多处层级判断共用
    """
    return bool(
        re.match(r'^[一二三四五六七八九十]+、', text)
        or re.match(r'^（[一二三四五六七八九十]+）', text)
        or re.match(r'^\d+[.、．]\s*', text)
        or re.match(r'^（\d+）', text)
        or re.match(r'^[①②③④⑤⑥⑦⑧⑨⑩]', text)
    )


def is_verb_after_number(text):
    """编号+动词格式检测
    匹配：1.是 / 2.要 / 3.以 等不规范格式
    """
    return bool(re.match(r'^\d+[.、．]\s*[是是以要为将把让使被]', text))


def is_labeled_field(text):
    """标签字段检测：培训地点：XXX / 授课方式：XXX / 会议时间：XXX 等元数据行
    短标签（≤8字）+ 冒号 + 内容 → 不是标题，是正文字段
    """
    t = text.strip()
    m = re.match(r'^(.{1,8})[：:](.{2,})$', t)
    if not m:
        return False
    label, content = m.group(1), m.group(2)
    # 标签部分不含编号前缀（一、/（一）/1. 等已在上层排除）
    if re.match(r'^[一二三四五六七八九十]+、', label):
        return False
    if re.match(r'^（[一二三四五六七八九十]+）', label):
        return False
    return True


# ────────────────────────── 检查规则阈值 ──────────────────────────
# 所有规则中散落的魔数集中为命名常量，方便统一调整

# -- 句末标点检查 (check_punctuation_issues) --
MIN_PARAGRAPH_CHARS = 10          # 最小段落长度，短于此跳过
SHORT_TITLE_MAX_CHARS = 30        # 有编号前缀+短于此→视为标题跳过
COLON_GUIDE_MAX_AFTER = 15        # 冒号引导后内容短于此→视为标题引导跳过
NO_CJK_PUNCT_MAX_CHARS = 20       # 无中文标点且短于此→跳过
AI_BODY_CHECK_MIN = 15            # AI兜底判断最小字数
AI_BODY_CHECK_MAX = 30            # AI兜底判断最大字数

# -- 编号格式检查 (check_word_numbering_format) --
HEADING_LIKE_MAX_CHARS = 40       # 标题外观最大字数（无句号分号）
# Word 编号格式：标准公文格式（不做警告）
WNUM_STANDARD_FORMATS = frozenset({
    'decimal',                       # 1. 2. 3. — 三级标题
    'chineseCounting',               # 一、二、三、
    'chineseCountingThousand',       # 一、二、三、
    'ideographDigital',              # 一、二、三、
    'ideographEnclosedCircle',       # ①②③
    'decimalEnclosedCircleChinese',  # ①②③
})

# -- 缺失二级标题检查 (check_missing_h2) --
MISSING_H2_BODY_MIN_CHARS = 25    # 正文判定最小字数
MISSING_H2_LONG_BODY_MIN = 30     # 长正文排除阈值

# -- 标题标点检查 (check_title_punctuation / check_title_trailing_punct) --
TRAILING_PUNCT_MIN_CHARS = 5      # 标题标点检查最小字数
TRAILING_PUNCT_BODY_LIKE_MIN = 28 # 正文型编号段落排除阈值
TRAILING_PUNCT_LONG_BODY_MIN = 35 # 长正文排除阈值

# -- 层级判定 (resolve_final_level) --
WNL_HEADING_MAX_CHARS = 30        # Word编号回退标题最大字数
BODY_DEMOTE_MIN_CHARS = 25        # 长段降级为正文阈值

# -- X. 提升为 h1 (promote_x_to_h1) --
SHORT_H1_MAX_CHARS = 20           # 短段提升h1最大字数
BOLD_H1_MAX_CHARS = 25            # 粗体短标题提升h1最大字数
POST_TABLE_H1_MAX_CHARS = 20      # 表后短标题提升h1最大字数

# -- 主标题 (is_main_title) --
MAIN_TITLE_MAX_CHARS = 40         # 主标题最大字数
MAIN_TITLE_MIN_CHARS = 8          # 主标题最小字数（短于此的不是标题）
MAIN_TITLE_MAX_POSITION = 6       # 主标题最多出现在第几个段落（公文标题只在文档最开头）

# -- 问候语 (is_greeting) --
GREETING_MAX_CHARS = 30           # 问候语最大字数（正则中内联，记录于此）

# -- 表格布局 (calc_smart_col_widths / rebuild_table) --
TABLE_WIDTH_PCT = 5000            # 表格宽度百分比（5000=100%）
TABLE_SEQ_COL_WEIGHT = 800        # 序号列权重
TABLE_COL_WEIGHTS = (1000, 1500, 2200, 3000, 4000)  # 列宽权重梯度
TABLE_WIDTHS_DXA = 12818          # 表格页面总宽度(dxa)
TABLE_ROW_HEIGHT_TWIPS = 397      # 表行高(twips)
TABLE_CELL_SPACING_TWIPS = 312    # 表段落行距(twips)

# -- 落款格式 --
SIGNATURE_INDENT_PT = 64          # 落款右缩进(pt)
SIGNATURE_BLANK_LINES = 2         # 落款前空行数
SIGNATURE_PROXIMITY = 5           # 落款距文档末尾最大段数

# -- AI 判定 --
AI_API_TIMEOUT = 8                # DeepSeek API 超时秒数


# ────────────────────────── 共享辅助函数 ──────────────────────────

def get_numbering_info(item, num_to_abstract, abstract_num_defs):
    """从段落元组提取 Word 编号格式信息。

    item: paragraphs_text 中的段落元组 (type, text, ..., num_id, num_ilvl, ...)
    num_to_abstract: {numId: abstractNumId}
    abstract_num_defs: {abstractNumId: {ilvl: (fmt, lvlText)}}

    返回: (fmt, lvl_txt) 或 (None, None) 当段落无 Word 编号时
    """
    orig_num_id = item.num_id
    if not orig_num_id or orig_num_id == '0':
        return None, None
    an_id = num_to_abstract.get(orig_num_id)
    if not an_id:
        return None, None
    levels = abstract_num_defs.get(an_id, {})
    nilvl = item.num_ilvl or '0'
    return levels.get(nilvl, (None, None))
