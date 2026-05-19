#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
中建四局公文格式化 — 审查逻辑

包含 6 个 _check_* 函数，用于检测公文格式问题并返回问题列表。
"""

import re
import os
import json
from detect import detect_level, is_main_title
from constants import (
    has_text_number_prefix, CNUM_TO_INT, int_to_cn,
    get_numbering_info, is_labeled_field,
    AI_API_TIMEOUT, MIN_PARAGRAPH_CHARS, SHORT_TITLE_MAX_CHARS,
    COLON_GUIDE_MAX_AFTER, NO_CJK_PUNCT_MAX_CHARS,
    AI_BODY_CHECK_MIN, AI_BODY_CHECK_MAX,
    HEADING_LIKE_MAX_CHARS, MISSING_H2_LONG_BODY_MIN,
    WNUM_STANDARD_FORMATS,
    TRAILING_PUNCT_MIN_CHARS, TRAILING_PUNCT_BODY_LIKE_MIN,
    TRAILING_PUNCT_LONG_BODY_MIN, WNL_HEADING_MAX_CHARS,
    BODY_DEMOTE_MIN_CHARS,
)


# ──── AI 语义判断（可选，配置 DMP_AI_KEY 环境变量启用）────
def ai_is_body(text):
    """调 DeepSeek API 判断编号段落是正文还是标题（单段，已弃用，保留兼容）
    返回 True=正文，False=标题，None=未配置/出错（走原有规则）
    """
    result = ai_batch_is_body([text])
    return result.get(text, None)


def ai_batch_is_body(texts):
    """批量调 DeepSeek API 判断多段编号段落是正文还是标题
    返回 {text: True/False, ...}，未配置则返回空 {}
    """
    if not texts:
        return {}
    api_key = os.environ.get('CSCEC_AI_KEY', '')
    if not api_key:
        return {}
    try:
        import urllib.request
        numbered = [f'{i+1}."{t}"' for i, t in enumerate(texts)]
        prompt = '判断以下编号段落是【正文】还是【标题】。正文：描述动作/目标/措施的完整句（如"提升…能力"）、具体说明。标题：名词性短语、纲要要点、时间地点项（如"培训时间"）、联系信息。每行只答"正文"或"标题"。\n' + '\n'.join(numbered)
        req = urllib.request.Request(
            'https://api.deepseek.com/v1/chat/completions',
            data=json.dumps({
                'model': 'deepseek-chat',
                'messages': [{'role': 'user', 'content': prompt}],
                'max_tokens': len(texts) * 4,
                'temperature': 0
            }).encode(),
            headers={'Content-Type': 'application/json', 'Authorization': f'Bearer {api_key}'}
        )
        resp = json.loads(urllib.request.urlopen(req, timeout=AI_API_TIMEOUT).read())
        lines = resp['choices'][0]['message']['content'].strip().split('\n')
        result = {}
        for i, line in enumerate(lines):
            if i < len(texts):
                result[texts[i]] = '正文' in line
        return result
    except Exception:
        return {}


# ──── 审查函数 ────
def _prefetch_ai_decisions(paragraphs_text):
    """预收集所有需要 AI 判断的文本，批量一次调 API，返回 {text: is_body}"""
    pending = set()
    for item in paragraphs_text:
        if item[0] != 'p':
            continue
        text = item.text.strip()
        if not text or len(text) <= MIN_PARAGRAPH_CHARS:
            continue
        if not has_text_number_prefix(text):
            continue
        if AI_BODY_CHECK_MIN < len(text) <= AI_BODY_CHECK_MAX and not re.search(r'[。；]', text):
            pending.add(text)
    if not pending:
        return {}
    return ai_batch_is_body(list(pending))


def check_punctuation_issues(paragraphs_text, ai_cache=None):
    """句末标点检测：找出未以句号/问号/叹号结尾的正文段落"""
    if ai_cache is None:
        ai_cache = _prefetch_ai_decisions(paragraphs_text)
    issues = []
    for i, item in enumerate(paragraphs_text):
        if item[0] != 'p':
            continue
        text = item.text.strip()
        if not text or len(text) <= MIN_PARAGRAPH_CHARS:
            continue
        if is_main_title(text):
            continue
        if is_labeled_field(text):
            continue
        if re.match(r'^[一二三四五六七八九十]+、', text):
            continue
        if re.match(r'^（[一二三四五六七八九十]+）', text):
            continue
        if re.match(r'^（\d+）', text):
            continue
        if re.match(r'^[①②③④⑤⑥⑦⑧⑨⑩]', text):
            continue
        # 有编号前缀 + 短文本 + 无句末标点 → 是标题，跳过句末标点检查
        if has_text_number_prefix(text) and len(text) <= SHORT_TITLE_MAX_CHARS and not re.search(r'[。；]', text):
            continue
        # 有编号前缀 + 含冒号引导 + 冒号后内容短 → 标题式引导，跳过
        if has_text_number_prefix(text) and not re.search(r'[。；]', text):
            colon_m = re.search(r'[：:]', text)
            if colon_m:
                remaining = len(text) - colon_m.end()
                if remaining <= COLON_GUIDE_MAX_AFTER:
                    continue
        if re.match(r'^[\d,.\-+%：:（）()]+$', text):
            continue
        if re.search(r'[：:]$', text):
            continue
        if len(text) <= SHORT_TITLE_MAX_CHARS and not re.search(r'[\u4e00-\u9fff]', text):
            continue
        if len(text) <= NO_CJK_PUNCT_MAX_CHARS and not re.search(r'[，。；！？]', text):
            continue
        # 含冒号 + 以数字或右括号结尾 → 列表项/联系信息，不查句号
        if re.search(r'[：:]', text) and re.search(r'[\d）\)\]】]$', text):
            continue
        # AI 兜底：15-30字无句号的编号段落（>30字/含逗号的直接当正文查句号）
        if has_text_number_prefix(text) and AI_BODY_CHECK_MIN < len(text) <= AI_BODY_CHECK_MAX and not re.search(r'[。；]', text):
            if '，' not in text and ai_cache.get(text) is False:
                continue  # AI 判为标题，跳过句末标点检查
        last_char = text[-1]
        if last_char not in ('。', '？', '！', '…', '"', '"', ')', '）', '；'):
            issues.append((i, text[:60]))
    return issues


def check_h3_numbering_issues(paragraphs_text):
    """三级标题编号不规范检测：如 '1.是' '2.是' 应为 '一是' '二是'"""
    issues = []
    for i, item in enumerate(paragraphs_text):
        if item[0] != 'p':
            continue
        text = item.text.strip()
        if not text:
            continue
        m = re.match(r'^(\d+)[.、．]\s*(是|且|但|将|要|在|已|以|对|为|从|按|于)\s*(.*)', text)
        if m:
            num = int(m.group(1))
            word = m.group(2)
            issues.append((i, text[:70], num, word))
    return issues


def check_word_numbering_format(paragraphs_text, num_to_abstract, abstract_num_defs):
    """检测Word自动编号使用非标准公文格式（如 a.b.c. I.II.III. 等）

    只对非标准格式报警；decimal（1.2.3.）、chineseCounting（一、二、三、）、
    ideographEnclosedCircle（①②③）等标准公文格式不报。
    """
    issues = []
    for i, item in enumerate(paragraphs_text):
        if item[0] != 'p':
            continue
        text = item.text.strip()
        if not text:
            continue
        level = item.word_num_level
        if level is not None:
            continue
        fmt, lvl_txt = get_numbering_info(item, num_to_abstract, abstract_num_defs)
        if fmt is None or fmt in WNUM_STANDARD_FORMATS:
            continue
        is_like_heading = (
            len(text) <= HEADING_LIKE_MAX_CHARS
            and not re.search(r'[。；！？]', text)
        )
        if is_like_heading:
            issues.append((i, text[:50]))
    return issues


def check_numbering_separator(paragraphs_text, num_to_abstract, abstract_num_defs):
    """检测Word自动编号分隔符不规范（、或．应改为.）"""
    issues = []
    for i, item in enumerate(paragraphs_text):
        if item[0] != 'p':
            continue
        text = item.text.strip()
        if not text:
            continue
        level = item.word_num_level
        if level is not None:
            continue  # 已有标题层级，由 check_title_punctuation 处理
        fmt, lvl_txt = get_numbering_info(item, num_to_abstract, abstract_num_defs)
        if not lvl_txt:
            continue
        if '、' in lvl_txt or '．' in lvl_txt:
            correct = lvl_txt.replace('、', '.').replace('．', '.')
            issues.append((i, text[:40], lvl_txt.strip(), correct.strip()))
    return issues


def check_missing_h2(paragraphs_text, ai_cache=None):
    """检测一级标题下直接使用三级标题的情况，建议补充二级标题。

    注意区分：
    - 短文本数字编号（如"1.科技部"）→ 是标题，一级跳三级应提示缺少二级标题
    - 长文本数字编号（如"1.本年度节后新开项目1个..."）→ 是正文，不提示跳级
    """
    if ai_cache is None:
        ai_cache = _prefetch_ai_decisions(paragraphs_text)
    issues = []
    last_h1_index = None
    _prev_was_body = False  # 前一个数字编号段落是否为正文（用于同组继承）

    for i, item in enumerate(paragraphs_text):
        if item[0] != 'p':
            continue
        level = item.word_num_level
        num_id = item.num_id
        text = item.text.strip() if len(item) > 1 else ''

        detected = detect_level(text)
        effective_level = level if level in ('h1', 'h2', 'h3', 'h4', 'h5') else detected

        is_digit_prefix = bool(re.match(r'^\d+[.、．]', text))
        # 是否为正文：含句号/逗号/顿号/冒号（末尾冒号不算）、很长（>30字）、电话/邮箱/日期/时间等
        is_likely_body = is_digit_prefix and (
            '。' in text or '，' in text or '、' in text
            or ('：' in text and not text.rstrip().endswith('：'))
            or '（' in text  # 括号人名（如"（陈漩莹）"）是正文标志
            or len(text) > MISSING_H2_LONG_BODY_MIN
            or bool(re.search(r'\d{11}', text))
            or ('@' in text and '.' in text)
            or bool(re.search(r'\d{4}年\d{1,2}月\d{1,2}日', text))
            or bool(re.search(r'\d{1,2}:\d{2}', text))
        )
        # AI 兜底：规则拿不准时，调 AI 判断
        if is_digit_prefix and not is_likely_body:
            ai_val = ai_cache.get(text)
            if ai_val is True:
                is_likely_body = True
            elif ai_val is False:
                is_likely_body = False

        # 同组继承：若前一项（同编号格式）已被判为正文，当前项也是正文
        # 例："1.AAA、BBB"（正文）、"2.CCC：DDD"（正文）、"3.EEE参会"（无标点，继承正文）
        if is_digit_prefix and not is_likely_body and i > 0:
            prev = paragraphs_text[i - 1]
            if prev[0] == 'p':
                prev_text = prev[1].strip() if len(prev) > 1 else ''
                prev_digit = bool(re.match(r'^\d+[.、．]', prev_text))
                if prev_digit and _prev_was_body:
                    is_likely_body = True

        # 更新同组继承状态（仅数字编号段落参与）
        if is_digit_prefix:
            _prev_was_body = is_likely_body

        if effective_level == 'h1':
            last_h1_index = i
            _prev_was_body = False  # 新章节重置
        elif effective_level == 'h3' and last_h1_index is not None:
            if is_likely_body:
                continue
            has_h2_between = False
            for j in range(last_h1_index + 1, i):
                if paragraphs_text[j][0] == 'p':
                    between_level = paragraphs_text[j][3] if len(paragraphs_text[j]) > 3 else None
                    between_text = paragraphs_text[j][1].strip() if len(paragraphs_text[j]) > 1 else ''
                    between_detected = detect_level(between_text)
                    between_effective = between_level if between_level in ('h1', 'h2', 'h3', 'h4', 'h5') else between_detected
                    if between_effective == 'h2':
                        has_h2_between = True
                        break
            if not has_h2_between:
                issues.append((i, text[:30]))
        elif last_h1_index is not None and num_id and num_id != '0':
            looks_like_title = (
                len(text) <= WNL_HEADING_MAX_CHARS
                and not re.search(r'[。；！？]', text)
            )
            if looks_like_title:
                issues.append((i, text[:30]))
        elif effective_level == 'h2':
            last_h1_index = None
            _prev_was_body = False

    return issues


def check_title_punctuation(paragraphs_text):
    """检测标题编号后的标点是否符合规范：
    - 一级标题：编号后接顿号（、），如"一、"
    - 二级标题：编号后无标点，如"（一）"
    - 三级标题：编号后接点号（.），如"1."
    返回问题列表：[(段落索引, 标题文本, 错误类型, 建议), ...]
    """
    issues = []

    for i, item in enumerate(paragraphs_text):
        if item[0] != 'p':
            continue

        level = item.word_num_level
        text = item.text.strip() if len(item) > 1 else ''

        if level is None and text:
            if re.match(r'^[一二三四五六七八九十]+[、】]', text):
                level = 'h1'
            elif re.match(r'^（[一二三四五六七八九十]+）', text):
                level = 'h2'
            elif re.match(r'^\d+[.、．](?!\d)\s*\S', text):
                level = 'h3'

        if level not in ('h1', 'h2', 'h3'):
            continue

        if not text:
            continue

        if level == 'h1':
            match = re.match(r'^([一二三四五六七八九十]+)([、．.：:；;]?)', text)
            if match:
                num_part = match.group(1)
                punct = match.group(2)
                if punct != '、':
                    if punct:
                        issues.append((i, text[:40], 'h1_wrong_punct',
                            f'一级标题编号"{num_part}"后应为顿号"、"，实为"{punct}"'))
                    else:
                        issues.append((i, text[:40], 'h1_missing_punct',
                            f'一级标题编号"{num_part}"后缺少顿号"、"'))

        elif level == 'h2':
            match = re.match(r'^（([一二三四五六七八九十]+)）([、．.：:；;]?)', text)
            if match:
                num_part = match.group(1)
                punct = match.group(2)
                if punct:
                    issues.append((i, text[:40], 'h2_extra_punct',
                        f'二级标题"（{num_part}）"后不应有标点，检测到"{punct}"'))

        elif level == 'h3':
            match = re.match(r'^(\d+)([、．.：:；;]?)', text)
            if match:
                num_part = match.group(1)
                punct = match.group(2)
                if punct != '.':
                    if punct:
                        issues.append((i, text[:40], 'h3_wrong_punct',
                            f'三级编号"{num_part}"后应为点号"."，实为"{punct}"'))
                    else:
                        issues.append((i, text[:40], 'h3_missing_punct',
                            f'三级编号"{num_part}"后缺少点号"."'))

    return issues


def check_title_trailing_punct(paragraphs_text, num_to_abstract, abstract_num_defs):
    """标题句末标点检测（从 format_document 中提取）

    规则：只有真正的多级标题末尾有标点才需要提示
    正文编号列表和Word原生编号段落不算标题
    返回：[(idx, text, punct_char), ...]
    """
    issues = []
    for idx, item in enumerate(paragraphs_text):
        if item[0] != 'p':
            continue
        text = item.text.strip()
        if not text or len(text) <= TRAILING_PUNCT_MIN_CHARS:
            continue

        level = detect_level(text)
        wnl = item.word_num_level

        # 有文本编号前缀才算标题候选
        if not has_text_number_prefix(text):
            continue

        # 区分"数字编号的标题"和"数字编号的正文"
        is_digit_prefix = bool(re.match(r'^\d+[.、．]', text))
        if is_digit_prefix:
            content_after_num = re.sub(r'^\d+[.、．]\s*', '', text)
            is_likely_body = (
                len(text) > TRAILING_PUNCT_BODY_LIKE_MIN
                or '。' in text
                or '；' in text
                or (len(content_after_num) > SHORT_TITLE_MAX_CHARS and text.rstrip()[-1] == '。')
            )
            if is_likely_body:
                continue

        # Word编号段落排除
        is_word_num_body = (wnl is not None and not has_text_number_prefix(text) and len(text) > BODY_DEMOTE_MIN_CHARS)
        if is_word_num_body:
            continue

        # 有编号前缀但内容超长的是正文
        if len(text) > TRAILING_PUNCT_LONG_BODY_MIN and ('。' in text or '；' in text):
            continue

        # 排除Word原生十进制编号的正文列表
        fmt, _ = get_numbering_info(item, num_to_abstract, abstract_num_defs)
        if fmt == 'decimal':
            continue

        # h4/h5 分项列举：；和。结尾是标准写法，不报
        is_h4 = bool(re.match(r'^（\d+）', text))
        is_h5 = bool(re.match(r'^[①②③④⑤⑥⑦⑧⑨⑩]', text))
        if (is_h4 or is_h5) and text.rstrip()[-1] in ('；', '。'):
            continue

        # 标题应以非句号结尾
        last_char = text.rstrip()[-1]
        if last_char in ('。', '；', '，', '：', ':'):
            issues.append((idx, text, last_char))

    # ── 后处理：排除分项列举 ──
    # 连续 h4/h5 项，前 N-1 个以 ；结尾、末项以 。结尾 → 合法
    if len(issues) >= 2:
        groups = []
        cur = [issues[0]]
        for prev, item in zip(issues, issues[1:]):
            if item[0] == prev[0] + 1:
                cur.append(item)
            else:
                groups.append(cur)
                cur = [item]
        groups.append(cur)

        keep = []
        for g in groups:
            if len(g) < 2:
                keep.extend(g)
                continue
            endings = [item[2] for item in g]
            # 全；或 前N-1个；+末。→ 分项列举
            if all(e == '；' for e in endings) or \
               (all(e == '；' for e in endings[:-1]) and endings[-1] == '。'):
                continue
            keep.extend(g)
        issues = keep

    return issues
