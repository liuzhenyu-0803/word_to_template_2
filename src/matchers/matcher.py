"""  
Word文档匹配器 - 封装文档元素匹配功能  
调用专门的匹配器模块完成实际工作  
"""  

import os
import shutil
# 使用绝对导入
from . import table_matcher # 确保table_matcher被正确导入

def match_document(extract_files: list[str], key_descriptions_dir: str, match_results_dir: str):
    """
    对提取的文档元素进行匹配分析
    
    参数:
        extract_files: 要进行语义识别的提取文件列表
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
    
    # 检查文件列表是否为空
    if not extract_files:
        print("错误：提取文件列表为空")
        return stats
    
    # 表格关键字描述文件路径
    table_key_description_path = os.path.join(key_descriptions_dir, "table_key_description.txt")
    
    if not os.path.exists(table_key_description_path):
        print(f"错误：表格关键字描述文件 {table_key_description_path} 不存在")
        return stats
    
    # 筛选出表格文件进行处理 (假设表格文件以 'table_' 开头并以 '.html' 结尾)
    # 你可能需要根据实际的文件命名规则调整这里的筛选逻辑
    table_files = [f for f in extract_files if os.path.basename(f).startswith('table_') and f.endswith('.html')]

    if not table_files:
        print("在提供的文件列表中未找到需要处理的表格文件。")
        # 根据需求，这里可以选择返回空的统计信息或者进行其他处理
        # return stats

    # 处理表格 - 调用table_matcher模块的match_tables函数
    # 注意：match_tables 函数也需要能够接受文件列表
    if table_files:
        table_stats = table_matcher.match_tables(table_files, table_key_description_path, match_results_dir)
        # 合并统计信息
        stats.update(table_stats)
    
    # 未来可以在这里添加其他类型的元素处理，如段落匹配等
    # 例如，筛选出段落文件并调用相应的匹配函数
    # paragraph_files = [f for f in extract_files if os.path.basename(f).startswith('paragraph_') and f.endswith('.txt')]
    # if paragraph_files:
    #     paragraph_stats = paragraph_matcher.match_paragraphs(paragraph_files, paragraph_key_description_path, match_results_dir)
    #     stats.update(paragraph_stats)
        
    return stats

# 直接运行时的入口点
if __name__ == "__main__":
    project_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    # 示例：假设我们只想处理特定的几个提取文件
    # document_parts_dir = os.path.join(project_dir, "document/document_extract")
    example_extract_files = [
        os.path.join(project_dir, "document/document_extract/table_1.html"),
        os.path.join(project_dir, "document/document_extract/table_2.html"),
        # os.path.join(project_dir, "document/document_extract/paragraph_1.txt") # 如果有段落文件
    ]
    key_descriptions_dir = os.path.join(project_dir, "document/key_descriptions")
    match_results_dir = os.path.join(project_dir, "document/match_results")
    
    if example_extract_files:
        # 确保示例文件存在，否则测试会失败
        for f_path in example_extract_files:
            if not os.path.exists(f_path):
                print(f"警告：测试用的示例文件 {f_path} 不存在，请创建或修改路径。")
                # 可以选择在这里创建空文件用于测试
                # open(f_path, 'a').close()
        
        match_document(example_extract_files, key_descriptions_dir, match_results_dir)
    else:
        print("测试未运行，因为 example_extract_files 为空或文件不存在。")
