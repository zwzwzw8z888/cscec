#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
调用 format_core.py 对 Word 文档进行公文格式化
用法：python run_format.py <输入文件.docx> [输出文件.docx]
"""

import sys, os, datetime

# 获取当前脚本所在目录（技能文件夹路径）
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)

from formatter_core import format_document

# 从命令行参数读取文件路径
if len(sys.argv) < 2:
    print("用法：python run_format.py <输入文件.docx> [输出文件.docx]")
    sys.exit(1)

SRC = sys.argv[1]

# 如果未指定输出文件，自动生成
if len(sys.argv) >= 3:
    DST = sys.argv[2]
else:
    base_name = os.path.splitext(os.path.basename(SRC))[0]
    now = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    DST = os.path.join(os.path.dirname(SRC), f"{base_name}-格式化_{now}.docx")

if not os.path.exists(SRC):
    print(f"❌ 源文件不存在: {SRC}")
    sys.exit(1)

print(f"📂 源文件: {SRC}")
print(f"💾 输出到: {DST}")
print(f"⏳ 正在格式化...")

dst_path, warnings = format_document(SRC, DST)

print(f"✅ 格式化完成：{dst_path}")
if warnings:
    print(f"⚠️  共 {len(warnings)} 条审查提示：")
    for w in warnings:
        print(f"  [{w['type']}] {w['detail'][:80]}")
else:
    print("无审查提示。")
