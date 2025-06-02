"""  
表格替换模块  
负责处理Word文档中表格内容的替换
"""  

import os
import json

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
    
    # 遍历所有表格
    for table in doc.tables:
        process_table_replacement(table, value_to_key_map)

def process_table_replacement(table, value_to_key_map):
    """
    处理表格的内容替换，包括嵌套表格
    
    参数:
        table: 表格对象
        value_to_key_map: 值到键的映射字典
    """
    # 按值的长度降序排序，以便先替换长文本，避免部分匹配问题
    sorted_values = sorted(value_to_key_map.keys(), key=len, reverse=True)
    
    for row in table.rows:
        for cell in row.cells:
            # 获取单元格的原始文本
            original_text = cell.text
            if original_text.strip():
                new_text = original_text
                replaced = False
                
                # 按值的长度降序尝试替换
                for value in sorted_values:
                    if value in new_text:
                        key = value_to_key_map[value]
                        new_text = new_text.replace(value, key)  # 直接使用key，不添加花括号
                        replaced = True
                
                # 如果有替换，更新单元格内容
                if replaced and new_text != original_text:
                    cell.text = new_text
            
            # 处理嵌套表格
            if hasattr(cell, 'tables') and cell.tables:
                for nested_table in cell.tables:
                    process_table_replacement(nested_table, value_to_key_map)

# 直接运行时的入口点
if __name__ == "__main__":
    import sys
    from docx import Document
    
    # 添加项目根目录到系统路径
    current_dir = os.path.dirname(os.path.abspath(__file__))
    project_dir = os.path.dirname(os.path.dirname(current_dir))
    if project_dir not in sys.path:
        sys.path.append(project_dir)
    
    # 使用项目标准路径结构
    original_doc_path = os.path.join(project_dir, "document", "document.docx")
    match_results_dir = os.path.join(project_dir, "document", "match_results")
    template_doc_path = os.path.join(project_dir, "document", "template_table6_test.docx")
    
    # 检查输入文件和目录是否存在
    if not os.path.exists(original_doc_path):
        print(f"错误：原始文档 {original_doc_path} 不存在")
        exit(1)
        
    if not os.path.exists(match_results_dir):
        print(f"错误：匹配结果目录 {match_results_dir} 不存在")
        exit(1)
    
    # 确保输出目录存在
    output_dir = os.path.dirname(template_doc_path)
    if not os.path.exists(output_dir):
        os.makedirs(output_dir, exist_ok=True)
    
    # 读取原始文档
    print(f"正在读取原始文档: {original_doc_path}")
    doc = Document(original_doc_path)
    
    # 只使用 table_6 的匹配结果文件
    table_match_files = ['table_6_matches.json']
    
    # 检查 table_6 匹配结果是否存在
    table_6_path = os.path.join(match_results_dir, 'table_6_matches.json')
    if not os.path.exists(table_6_path):
        print(f"错误：table_6 匹配结果文件不存在: {table_6_path}")
        print("请先运行 table_matcher.py 生成匹配结果")
        exit(1)
    
    # 处理表格替换 - 只替换 table_6 的内容
    print("开始处理 table_6 的表格替换...")
    replace_values_with_placeholders(doc, match_results_dir, table_match_files)
    print("table_6 替换完成")
    
    # 保存处理后的模板文档
    doc.save(template_doc_path)
    print(f"已保存生成的 table_6 测试模板文档: {template_doc_path}")


