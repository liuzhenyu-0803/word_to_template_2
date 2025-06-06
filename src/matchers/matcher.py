"""  
Word文档匹配器 - 封装文档元素匹配功能  
调用专门的匹配器模块完成实际工作  
"""  

import os
import shutil
import json
from . import table_matcher  # 使用相对导入

def match_document(document_extract_dir, key_descriptions_dir, match_results_dir):
    """
    对提取的文档元素进行匹配分析
    
    参数:
        document_extract_dir: 文档元素目录路径，包含提取的文档元素
        key_descriptions_dir: 关键字描述文件所在目录
        match_results_dir: 匹配结果目录路径，用于保存匹配结果
        
    返回:
        dict: 匹配结果统计信息，如匹配的元素数量等
    """
    # 检查并清理输出目录
    if os.path.exists(match_results_dir):
        # 删除目录中的所有内容
        for item in os.listdir(match_results_dir):
            item_path = os.path.join(match_results_dir, item)
            if os.path.isfile(item_path):
                os.remove(item_path)
            elif os.path.isdir(item_path):
                shutil.rmtree(item_path)
    else:
        # 创建输出目录
        os.makedirs(match_results_dir)
    
    # 初始化统计信息
    stats = {}
    
    # 检查输入目录是否存在
    if not os.path.exists(document_extract_dir):
        print(f"错误：文档元素目录 {document_extract_dir} 不存在")
        return stats
    
    # 表格关键字描述文件路径
    table_key_description_path = os.path.join(key_descriptions_dir, "table_key_description.txt")
    
    if not os.path.exists(table_key_description_path):
        print(f"错误：表格关键字描述文件 {table_key_description_path} 不存在")
        return stats
    
    # 处理表格 - 调用table_matcher模块的match_tables函数
    table_stats = table_matcher.match_tables(document_extract_dir, table_key_description_path, match_results_dir)
    
    # 合并统计信息
    stats.update(table_stats)
    
    # 未来可以在这里添加其他类型的元素处理，如段落匹配等
    
    return stats

# 直接运行时的入口点
if __name__ == "__main__":
    project_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    document_parts_dir = os.path.join(project_dir, "document/document_extract")
    key_descriptions_dir = os.path.join(project_dir, "document/key_descriptions")
    match_results_dir = os.path.join(project_dir, "document/match_results")
    
    match_document(document_parts_dir, key_descriptions_dir, match_results_dir)
