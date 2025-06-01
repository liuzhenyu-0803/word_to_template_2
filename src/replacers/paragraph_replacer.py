"""  
段落替换模块  
负责处理Word文档中段落内容的替换
"""  

import os
import json
import re
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

def replace_values_with_placeholders(doc, match_results_dir, match_files):
    """
    将Word文档中的段落内容替换为占位符，用于生成模板
    
    参数:
        doc: docx.Document对象，要处理的原始文档
        match_results_dir: 匹配结果目录路径
        match_files: 段落匹配结果文件列表
    """
    # 创建一个值到键的反向映射
    value_to_key_map = {}
    
    # 加载所有段落匹配结果
    for match_file in match_files:
        file_path = os.path.join(match_results_dir, match_file)
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                match_data = json.load(f)
                for item in match_data:
                    if 'field_key' in item and 'field_value' in item:
                        key = item['field_key']
                        value = item['field_value']
                        # 添加到值到键的映射
                        if value and value.strip():  # 确保值不为空
                            value_to_key_map[value] = key
        except Exception as e:
            print(f"读取匹配结果文件 {match_file} 失败: {e}")
    
    print(f"加载了 {len(value_to_key_map)} 个段落值-键映射项")
    
    # 按值的长度降序排序，以便先替换长文本，避免部分匹配问题
    sorted_values = sorted(value_to_key_map.keys(), key=len, reverse=True)
    
    # 处理文档中的所有段落
    for para in doc.paragraphs:
        para_text = para.text.strip()
        
        # 跳过空段落
        if not para_text:
            continue
        
        replaced = False
        new_text = para_text
        
        # 尝试匹配完整的段落文本
        if para_text in value_to_key_map:
            key = value_to_key_map[para_text]
            new_text = f"{{{{{key}}}}}"
            replaced = True
        else:
            # 按值的长度降序尝试部分匹配
            for value in sorted_values:
                if value in para_text:
                    key = value_to_key_map[value]
                    # 创建一个安全的正则表达式模式，避免特殊字符问题
                    pattern = re.escape(value)
                    new_text = re.sub(pattern, f"{{{{{key}}}}}", new_text)
                    replaced = True
        
        # 如果有替换，更新段落文本
        if replaced and new_text != para_text:
            para.clear()
            para.add_run(new_text)
