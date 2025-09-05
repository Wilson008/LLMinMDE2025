#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
BibTeX文件解析器 - 将.bib文件转换为Excel表格
完全自动化版本，无需用户输入
"""

import re
import pandas as pd
import os
from typing import Dict, List, Any

def clean_field_value(value: str) -> str:
    """
    清理BibTeX字段值，正确转换LaTeX特殊字符到Unicode字符
    """
    if not value:
        return ""
    
    # LaTeX重音符号映射表
    latex_mappings = {
        # 波浪符 (~)
        r'\\~\{n\}': 'ñ', r'\\~\{N\}': 'Ñ',
        r'\\~\{a\}': 'ã', r'\\~\{A\}': 'Ã',
        r'\\~\{o\}': 'õ', r'\\~\{O\}': 'Õ',
        
        # 尖音符 (')
        r'\\\'?\{a\}': 'á', r'\\\'?\{A\}': 'Á',
        r'\\\'?\{e\}': 'é', r'\\\'?\{E\}': 'É',
        r'\\\'?\{i\}': 'í', r'\\\'?\{I\}': 'Í',
        r'\\\'?\{o\}': 'ó', r'\\\'?\{O\}': 'Ó',
        r'\\\'?\{u\}': 'ú', r'\\\'?\{U\}': 'Ú',
        r'\\\'?\{y\}': 'ý', r'\\\'?\{Y\}': 'Ý',
        r'\\\'?\{c\}': 'ć', r'\\\'?\{C\}': 'Ć',
        
        # 重音符 (`)
        r'\\`\{a\}': 'à', r'\\`\{A\}': 'À',
        r'\\`\{e\}': 'è', r'\\`\{E\}': 'È',
        r'\\`\{i\}': 'ì', r'\\`\{I\}': 'Ì',
        r'\\`\{o\}': 'ò', r'\\`\{O\}': 'Ò',
        r'\\`\{u\}': 'ù', r'\\`\{U\}': 'Ù',
        
        # 分音符/元音变音 (")
        r'\\"\{a\}': 'ä', r'\\"\{A\}': 'Ä',
        r'\\"\{e\}': 'ë', r'\\"\{E\}': 'Ë',
        r'\\"\{i\}': 'ï', r'\\"\{I\}': 'Ï',
        r'\\"\{o\}': 'ö', r'\\"\{O\}': 'Ö',
        r'\\"\{u\}': 'ü', r'\\"\{U\}': 'Ü',
        
        # 扬抑符 (^)
        r'\\\^\{a\}': 'â', r'\\\^\{A\}': 'Â',
        r'\\\^\{e\}': 'ê', r'\\\^\{E\}': 'Ê',
        r'\\\^\{i\}': 'î', r'\\\^\{I\}': 'Î',
        r'\\\^\{o\}': 'ô', r'\\\^\{O\}': 'Ô',
        r'\\\^\{u\}': 'û', r'\\\^\{U\}': 'Û',
        
        # 软音符 cedilla (c)
        r'\\c\{c\}': 'ç', r'\\c\{C\}': 'Ç',
        
        # 其他特殊字符
        r'\\ss\b': 'ß',  # 德语 eszett
        r'\\ae\b': 'æ', r'\\AE\b': 'Æ',  # ae连字
        r'\\oe\b': 'œ', r'\\OE\b': 'Œ',  # oe连字
        r'\\o\b': 'ø', r'\\O\b': 'Ø',    # 斜杠o
        r'\\aa\b': 'å', r'\\AA\b': 'Å',  # 环形a
        
        # 西班牙语倒置标点
        r'\\textquestiondown\b': '¿',
        r'\\textexclamdown\b': '¡',
    }
    
    # 应用LaTeX映射
    for latex_pattern, unicode_char in latex_mappings.items():
        value = re.sub(latex_pattern, unicode_char, value)
    
    # 处理没有花括号的简单情况，如 \'a -> á
    simple_mappings = {
        r'\\~n\b': 'ñ', r'\\~N\b': 'Ñ',
        r'\\\'a\b': 'á', r'\\\'A\b': 'Á', r'\\\'e\b': 'é', r'\\\'E\b': 'É',
        r'\\\'i\b': 'í', r'\\\'I\b': 'Í', r'\\\'o\b': 'ó', r'\\\'O\b': 'Ó',
        r'\\\'u\b': 'ú', r'\\\'U\b': 'Ú',
        r'\\`a\b': 'à', r'\\`A\b': 'À', r'\\`e\b': 'è', r'\\`E\b': 'È',
        r'\\`i\b': 'ì', r'\\`I\b': 'Ì', r'\\`o\b': 'ò', r'\\`O\b': 'Ò',
        r'\\`u\b': 'ù', r'\\`U\b': 'Ù',
        r'\\"a\b': 'ä', r'\\"A\b': 'Ä', r'\\"e\b': 'ë', r'\\"E\b': 'Ë',
        r'\\"i\b': 'ï', r'\\"I\b': 'Ï', r'\\"o\b': 'ö', r'\\"O\b': 'Ö',
        r'\\"u\b': 'ü', r'\\"U\b': 'Ü',
    }
    
    for simple_pattern, unicode_char in simple_mappings.items():
        value = re.sub(simple_pattern, unicode_char, value)
    
    # 移除剩余的花括号
    value = re.sub(r'\{([^{}]*)\}', r'\1', value)
    
    # 处理其他LaTeX命令
    value = value.replace('--', '–')  # en-dash
    value = value.replace('---', '—')  # em-dash
    value = value.replace('``', '"')  # 左双引号
    value = value.replace("''", '"')  # 右双引号
    value = value.replace('`', ''')   # 左单引号
    value = value.replace("'", ''')   # 右单引号
    
    # 清理剩余的反斜杠命令
    value = re.sub(r'\\[a-zA-Z]+\s*', ' ', value)
    value = re.sub(r'\\(.)', r'\1', value)  # 移除转义字符前的反斜杠
    
    # 合并多个空格并去除首尾空格
    value = re.sub(r'\s+', ' ', value)
    value = value.strip()
    
    return value

def parse_bibtex_file(file_path: str) -> List[Dict[str, Any]]:
    """
    解析BibTeX文件并返回文献条目列表
    """
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()
    except UnicodeDecodeError:
        # 如果UTF-8解码失败，尝试其他编码
        with open(file_path, 'r', encoding='latin-1') as f:
            content = f.read()
    
    entries = []
    
    # 查找所有条目的起始位置
    entry_pattern = r'@(\w+)\{([^,\s]+),'
    entry_matches = list(re.finditer(entry_pattern, content))
    
    print(f"找到 {len(entry_matches)} 个文献条目")
    
    for i, match in enumerate(entry_matches):
        entry_type = match.group(1).lower()
        entry_key = match.group(2).strip()
        
        # 确定当前条目的结束位置
        start_pos = match.start()
        if i + 1 < len(entry_matches):
            end_pos = entry_matches[i + 1].start()
        else:
            end_pos = len(content)
        
        entry_content = content[start_pos:end_pos]
        
        # 初始化条目
        entry = {
            'type': entry_type,
            'key': entry_key
        }
        
        # 定义要提取的字段
        field_patterns = {
            'author': r'author\s*=\s*\{([^{}]*(?:\{[^{}]*\}[^{}]*)*)\}',
            'title': r'title\s*=\s*\{([^{}]*(?:\{[^{}]*\}[^{}]*)*)\}',
            'year': r'year\s*=\s*\{([^}]+)\}',
            'journal': r'journal\s*=\s*\{([^{}]*(?:\{[^{}]*\}[^{}]*)*)\}',
            'booktitle': r'booktitle\s*=\s*\{([^{}]*(?:\{[^{}]*\}[^{}]*)*)\}',
            'publisher': r'publisher\s*=\s*\{([^{}]*(?:\{[^{}]*\}[^{}]*)*)\}',
            'volume': r'volume\s*=\s*\{([^}]+)\}',
            'number': r'number\s*=\s*\{([^}]+)\}',
            'pages': r'pages\s*=\s*\{([^}]+)\}',
            'doi': r'doi\s*=\s*\{([^}]+)\}',
            'url': r'url\s*=\s*\{([^}]+)\}',
            'abstract': r'abstract\s*=\s*\{([^{}]*(?:\{[^{}]*\}[^{}]*)*)\}',
            'keywords': r'keywords\s*=\s*\{([^{}]*(?:\{[^{}]*\}[^{}]*)*)\}',
            'location': r'location\s*=\s*\{([^{}]*(?:\{[^{}]*\}[^{}]*)*)\}',
            'series': r'series\s*=\s*\{([^{}]*(?:\{[^{}]*\}[^{}]*)*)\}',
            'isbn': r'isbn\s*=\s*\{([^}]+)\}',
            'issn': r'issn\s*=\s*\{([^}]+)\}',
            'address': r'address\s*=\s*\{([^{}]*(?:\{[^{}]*\}[^{}]*)*)\}',
            'editor': r'editor\s*=\s*\{([^{}]*(?:\{[^{}]*\}[^{}]*)*)\}',
            'organization': r'organization\s*=\s*\{([^{}]*(?:\{[^{}]*\}[^{}]*)*)\}',
            'month': r'month\s*=\s*\{([^}]+)\}',
            'note': r'note\s*=\s*\{([^{}]*(?:\{[^{}]*\}[^{}]*)*)\}',
            'articleno': r'articleno\s*=\s*\{([^}]+)\}',
            'numpages': r'numpages\s*=\s*\{([^}]+)\}',
            'issue_date': r'issue_date\s*=\s*\{([^}]+)\}'
        }
        
        # 提取各个字段
        for field_name, pattern in field_patterns.items():
            match = re.search(pattern, entry_content, re.DOTALL)
            if match:
                field_value = clean_field_value(match.group(1))
                entry[field_name] = field_value
        
        entries.append(entry)
        
        # 显示处理进度
        if (i + 1) % 10 == 0 or i + 1 == len(entry_matches):
            print(f"已处理 {i + 1}/{len(entry_matches)} 个条目")
    
    return entries

def export_to_files(entries: List[Dict[str, Any]], output_excel: str):
    """
    将文献条目导出到Excel和CSV文件
    """
    # 准备数据框
    data = []
    
    for i, entry in enumerate(entries, 1):
        row = {
            '序号': i,
            '类型': entry.get('type', ''),
            '引用键': entry.get('key', ''),
            '标题': entry.get('title', ''),
            '作者': entry.get('author', ''),
            '年份': entry.get('year', ''),
            '期刊': entry.get('journal', ''),
            '会议': entry.get('booktitle', ''),
            '出版商': entry.get('publisher', ''),
            '卷': entry.get('volume', ''),
            '期': entry.get('number', ''),
            '页码': entry.get('pages', ''),
            'DOI': entry.get('doi', ''),
            'URL': entry.get('url', ''),
            '关键词': entry.get('keywords', ''),
            '摘要': entry.get('abstract', ''),
            '地点': entry.get('location', ''),
            '系列': entry.get('series', ''),
            'ISBN': entry.get('isbn', ''),
            'ISSN': entry.get('issn', ''),
            '地址': entry.get('address', ''),
            '编辑': entry.get('editor', ''),
            '组织': entry.get('organization', ''),
            '月份': entry.get('month', ''),
            '备注': entry.get('note', ''),
            '文章号': entry.get('articleno', ''),
            '页数': entry.get('numpages', ''),
            '发行日期': entry.get('issue_date', '')
        }
        data.append(row)
    
    # 创建DataFrame
    df = pd.DataFrame(data)
    
    # 导出Excel文件
    try:
        with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='文献信息', index=False)
            
            # 获取工作表以调整列宽
            workbook = writer.book
            worksheet = writer.sheets['文献信息']
            
            # 自动调整列宽
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                
                # 设置最大列宽限制
                adjusted_width = min(max_length + 2, 50)
                worksheet.column_dimensions[column_letter].width = adjusted_width
        
        print(f"Excel文件已成功导出到: {output_excel}")
        
    except Exception as e:
        print(f"导出Excel文件时出错: {e}")

def print_statistics(entries: List[Dict[str, Any]]):
    """
    打印文献统计信息
    """
    print("\n=== 文献统计信息 ===")
    print(f"总文献数量: {len(entries)}")
    
    # 按类型统计
    type_counts = {}
    for entry in entries:
        entry_type = entry.get('type', 'unknown')
        type_counts[entry_type] = type_counts.get(entry_type, 0) + 1
    
    print("\n按类型分布:")
    for entry_type, count in sorted(type_counts.items()):
        print(f"  {entry_type}: {count} 篇")
    
    # 按年份统计
    year_counts = {}
    for entry in entries:
        year = entry.get('year', 'unknown')
        year_counts[year] = year_counts.get(year, 0) + 1
    
    print("\n按年份分布:")
    for year, count in sorted(year_counts.items()):
        print(f"  {year}: {count} 篇")

def main():
    """
    主函数 - 完全自动化，无需用户输入
    """
    print("BibTeX文件解析器 - 自动化版本")
    print("=" * 50)
    
    # 硬编码的文件路径
    work_directory = r"C:\02.Work\09.Git_Repos\LLMinMDE2025"
    input_file = os.path.join(work_directory, "ScienceDirect_citations.bib")
    output_excel = os.path.join(work_directory, "ScienceDirect_results.xlsx")
    
    print(f"工作目录: {work_directory}")
    print(f"输入文件: acm.bib")
    
    # 检查目录是否存在
    if not os.path.exists(work_directory):
        print(f"错误: 工作目录不存在！")
        print(f"请创建目录: {work_directory}")
        return
    
    # 检查输入文件是否存在
    if not os.path.exists(input_file):
        print(f"错误: 输入文件不存在！")
        print(f"请将 acm.bib 文件放到目录: {work_directory}")
        return
    
    try:
        # 解析BibTeX文件
        print(f"\n正在解析BibTeX文件...")
        entries = parse_bibtex_file(input_file)
        
        if not entries:
            print("未找到任何文献条目！")
            return
        
        # 打印统计信息
        print_statistics(entries)
        
        # 导出文件
        print(f"\n正在导出Excel文件...")
        export_to_files(entries, output_excel)
        
        print("\n处理完成！")
        print(f"生成的文件: {output_excel}")
        
    except Exception as e:
        print(f"处理过程中出现错误: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()