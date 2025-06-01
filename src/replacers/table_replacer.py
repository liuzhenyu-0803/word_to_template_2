"""  
表格替换模块  
负责处理Word文档中表格内容的替换
"""  

import os
import json
import re
from docx.shared import Pt

def replace_values_with_placeholders(doc, match_results_dir, match_files):
    """
    将Word文档中的表格内容替换为占位符，用于生成模板
    
    参数:
        doc: docx.Document对象，要处理的原始文档
        match_results_dir: 匹配结果目录路径
        match_files: 表格匹配结果文件列表
    """
    # 创建一个值到键的反向映射
    value_to_key_map = {}
    
    for match_file in match_files:
        file_path = os.path.join(match_results_dir, match_file)
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                match_data = json.load(f)
                for item in match_data:
                    if 'new_key' in item and 'value' in item:
                        key = item['new_key']
                        value = item['value']
                        # 添加到值到键的映射
                        if value and value.strip() and key and key.strip():  # 确保值和键都不为空
                            value_to_key_map[value] = key
        except Exception as e:
            print(f"读取匹配结果文件 {match_file} 失败: {e}")
    
    print(f"加载了 {len(value_to_key_map)} 个表格值-键映射项")
    
    # 按值的长度降序排序，以便先替换长文本，避免部分匹配问题
    sorted_values = sorted(value_to_key_map.keys(), key=len, reverse=True)
    
    # 遍历所有表格及单元格，查找并替换实际内容为占位符
    for table in doc.tables:
        # 遍历表格中的所有行和单元格
        for row in table.rows:
            for cell in row.cells:
                # 处理单元格中的每个段落
                for para in cell.paragraphs:
                    para_text = para.text.strip()
                    if not para_text:
                        continue
                    
                    replaced = False
                    new_text = para_text
                    
                    # 尝试匹配完整的段落文本
                    if para_text in value_to_key_map:
                        key = value_to_key_map[para_text]
                        new_text = key  # 直接使用key，不添加花括号
                        replaced = True
                    else:
                        # 按值的长度降序尝试部分匹配
                        for value in sorted_values:
                            if value in para_text:
                                key = value_to_key_map[value]
                                # 创建一个安全的正则表达式模式，避免特殊字符问题
                                pattern = re.escape(value)
                                new_text = re.sub(pattern, key, new_text)  # 直接使用key，不添加花括号
                                replaced = True
                    
                    # 如果有替换，更新段落文本
                    if replaced and new_text != para_text:
                        para.clear()
                        para.add_run(new_text)
        
        # 处理嵌套表格(如果有)
        process_nested_tables_for_template(table, value_to_key_map)

def process_nested_tables_for_template(table, value_to_key_map):
    """
    递归处理嵌套表格中的内容，将实际值替换为占位符
    
    参数:
        table: 表格对象
        value_to_key_map: 值到键的映射字典
    """
    # 按值的长度降序排序，以便先替换长文本，避免部分匹配问题
    sorted_values = sorted(value_to_key_map.keys(), key=len, reverse=True)
    
    # 遍历表格中的所有单元格，查找嵌套表格
    for row in table.rows:
        for cell in row.cells:
            if hasattr(cell, 'tables') and cell.tables:
                for nested_table in cell.tables:
                    # 处理嵌套表格的单元格
                    for n_row in nested_table.rows:
                        for n_cell in n_row.cells:
                            for para in n_cell.paragraphs:
                                para_text = para.text.strip()
                                if not para_text:
                                    continue
                                
                                replaced = False
                                new_text = para_text
                                
                                # 尝试匹配完整的段落文本
                                if para_text in value_to_key_map:
                                    key = value_to_key_map[para_text]
                                    new_text = key  # 直接使用key，不添加花括号
                                    replaced = True
                                else:
                                    # 按值的长度降序尝试部分匹配
                                    for value in sorted_values:
                                        if value in para_text:
                                            key = value_to_key_map[value]
                                            # 创建一个安全的正则表达式模式，避免特殊字符问题
                                            pattern = re.escape(value)
                                            new_text = re.sub(pattern, key, new_text)  # 直接使用key，不添加花括号
                                            replaced = True
                                
                                # 如果有替换，更新段落文本
                                if replaced and new_text != para_text:
                                    para.clear()
                                    para.add_run(new_text)
                    
                    # 递归处理更深层次的嵌套表格
                    process_nested_tables_for_template(nested_table, value_to_key_map)
