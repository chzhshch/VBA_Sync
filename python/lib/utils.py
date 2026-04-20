#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
工具函数
"""

import re
import sys

def log_info(message):
    """记录信息日志"""
    print(f'[INFO] {message}', file=sys.stderr)

def log_error(message):
    """记录错误日志"""
    print(f'[ERROR] {message}', file=sys.stderr)

def extract_module_name(content, file_name):
    """从文件内容或文件名提取模块名"""
    # 尝试从 Attribute VB_Name 提取
    match = re.search(r'Attribute VB_Name\s*=\s*"([^"]+)"', content)
    if match:
        return match.group(1)
    # 从文件名提取
    return file_name.split('.')[0]

def _filter_attributes(content):
    """过滤隐藏成员属性"""
    lines = content.split('\n')
    filtered_lines = []
    
    for line in lines:
        # 检查是否是成员级属性（包含点号）
        if line.strip().startswith('Attribute ') and '.' in line.split('=')[0].strip():
            continue
        filtered_lines.append(line)
    
    return '\n'.join(filtered_lines)

def detect_new_sub_functions(old_content, new_content):
    """检测新增的 Sub 或 Function"""
    old_pattern = re.compile(r'(Sub|Function)\s+([a-zA-Z_][a-zA-Z0-9_]*)')
    new_pattern = re.compile(r'(Sub|Function)\s+([a-zA-Z_][a-zA-Z0-9_]*)')
    
    old_functions = set(match.group(2) for match in old_pattern.finditer(old_content))
    new_functions = set(match.group(2) for match in new_pattern.finditer(new_content))
    
    return new_functions - old_functions

def format_error_message(error):
    """格式化错误消息"""
    return str(error)
