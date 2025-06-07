"""  
Word文档提取器 - 封装文档元素提取功能  
调用专门的提取器模块完成实际工作  
"""  

import os
import shutil
from . import paragraph_extractor
from . import table_extractor

def extract_document(html_path, output_dir):
    """
    从HTML文件提取所有内容元素
    
    参数:
        html_path: HTML文件路径
        output_dir: 输出目录路径
        
    返回:
        tuple: (段落数量, 表格数量)
    """
    # 检查并清理输出目录
    if os.path.exists(output_dir):
        # 删除目录中的所有内容
        for item in os.listdir(output_dir):
            item_path = os.path.join(output_dir, item)
            if os.path.isfile(item_path):
                os.remove(item_path)
            elif os.path.isdir(item_path):
                shutil.rmtree(item_path)
    else:
        # 创建输出目录
        os.makedirs(output_dir)
    
    # 处理段落
    # paragraph_count = paragraph_extractor.extract_paragraphs(html_path, output_dir)
    paragraph_count = 0
    
    # 处理表格
    table_count = table_extractor.extract_tables(html_path, output_dir)
    
    return paragraph_count, table_count