#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""代码精简工具 - 删除冗余的调试信息和注释"""

import re

def simplify_code(input_file, output_file):
    with open(input_file, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    simplified_lines = []
    skip_next = False

    for i, line in enumerate(lines):
        # 跳过标记为删除的行
        if skip_next:
            skip_next = False
            continue

        # 删除独立的 DEBUG print 语句（整行）
        if re.match(r'^\s+print\(f"\[DEBUG\].*\)\s*$', line):
            continue

        # 简化过长的注释
        line = re.sub(r'"""模拟VLOOKUP查找', '"""查找', line)
        line = re.sub(r'"""从文件夹名称中提取.*"""', '"""提取元数据"""', line)
        line = re.sub(r'"""加载辅助Excel文件用于VLOOKUP功能"""', '"""加载辅助文件"""', line)
        line = re.sub(r'"""批量处理特定周文件夹中的所有Excel文件"""', '"""批量处理周文件夹"""', line)
        line = re.sub(r'"""批量处理包含周子文件夹.*的季度文件夹"""', '"""批量处理季度文件夹"""', line)

        # 删除冗余的单行注释
        if re.match(r'^\s+# 打印.*调试', line):
            continue
        if re.match(r'^\s+# 清理.*名称.*去除', line):
            continue
        if re.match(r'^\s+# 假设第.*列', line):
            continue
        if re.match(r'^\s+# 检查是否有第', line):
            continue
        if re.match(r'^\s+# 遍历DataFrame', line):
            continue
        if re.match(r'^\s+# 精确匹配', line):
            continue
        if re.match(r'^\s+# 如果没找到', line):
            continue
        if re.match(r'^\s+# 匹配开头的数字', line):
            continue

        # 删除多余的空行（连续3个以上空行只保留2个）
        if line.strip() == '':
            if i > 0 and i < len(lines) - 1:
                if lines[i-1].strip() == '' and (i < 2 or lines[i-2].strip() == ''):
                    continue

        simplified_lines.append(line)

    # 写入精简后的文件
    with open(output_file, 'w', encoding='utf-8') as f:
        f.writelines(simplified_lines)

    original_lines = len(lines)
    simplified_count = len(simplified_lines)
    removed = original_lines - simplified_count

    print(f"原始行数: {original_lines}")
    print(f"精简后行数: {simplified_count}")
    print(f"删除行数: {removed} ({removed/original_lines*100:.1f}%)")

if __name__ == '__main__':
    input_file = r'g:\Apple日常\case data promot\case summary\brief cost tracker case data builder'
    output_file = r'g:\Apple日常\case data promot\case summary\brief cost tracker case data builder.simplified'

    simplify_code(input_file, output_file)
    print("\n精简完成！")
    print(f"精简后的文件: {output_file}")
    print("请检查精简后的文件，如果没问题，可以替换原文件。")
