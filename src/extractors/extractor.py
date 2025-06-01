"""  
Word文档提取器 - 封装文档元素提取功能  
调用专门的提取器模块完成实际工作  
"""  

import os
import shutil
from . import paragraph_extractor
from . import table_extractor
from . import image_extractor

def extract_document(doc, output_dir):
    """
    提取Word文档的所有内容元素
    
    参数:
        doc: docx.Document对象
        output_dir: 输出目录路径
        
    返回:
        tuple: (段落数量, 表格数量, 唯一表格数量, 图片数量)
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
    paragraph_count = paragraph_extractor.extract_paragraphs(doc, output_dir)
    
    # 处理表格
    table_count, unique_tables = table_extractor.extract_tables(doc, output_dir)
    
    # 处理图片
    image_count, _ = image_extractor.extract_images(doc, output_dir)
    
    return paragraph_count, table_count, unique_tables, image_count