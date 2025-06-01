"""  
Word文档处理系统 - 主程序  
整合文档元素提取、匹配和替换功能
"""  

from docx import Document  
import os  
import time
from extractors.extractor import extract_document
from matchers.matcher import match_document
from replacers.replacer import replace_document
from models.model_manager import llm_manager

def process_word_document(original_doc_path, template_doc_path, parts_dir, match_results_dir, key_descriptions_dir):
    """
    完整处理Word文档：提取、匹配和生成模板
    
    参数:
        original_doc_path: 输入Word文档路径
        template_doc_path: 生成的模板文档输出路径
        parts_dir: 文档元素保存目录
        match_results_dir: 匹配结果保存目录
        key_descriptions_dir: 关键词描述目录
    """
    # 确保各目录存在
    os.makedirs(parts_dir, exist_ok=True)
    os.makedirs(match_results_dir, exist_ok=True)
    os.makedirs(key_descriptions_dir, exist_ok=True)
    os.makedirs(os.path.dirname(template_doc_path), exist_ok=True)
    
    # 记录开始时间
    start_time = time.time()
    
    # 步骤1: 提取文档元素
    print("\n===== 步骤1: 提取文档元素 =====")
    doc = Document(original_doc_path)
    paragraph_count, table_count, unique_tables, image_count = extract_document(doc, parts_dir)
    
    print(f"文档共有 {paragraph_count} 个段落和 {unique_tables} 个唯一表格")
    print(f"文档包含 {image_count} 个图片")
    print(f"文档已拆分为 {paragraph_count + table_count + image_count} 个元素")
    print(f"所有元素文件已保存到目录: {parts_dir}")
    
    # 步骤2: 匹配文档元素
    print("\n===== 步骤2: 匹配文档元素 =====")
    match_stats = match_document(parts_dir, match_results_dir, key_descriptions_dir)
    
    if match_stats:
        print(f"总共处理了 {match_stats.get('total_tables_processed', 0)} 个表格")
        print(f"匹配到 {match_stats.get('total_fields_matched', 0)} 个字段")
        print(f"有 {match_stats.get('tables_with_matches', 0)} 个表格包含匹配项")
        print(f"所有匹配结果已保存到目录: {match_results_dir}")
    else:
        print("没有找到匹配结果")
    
    # 步骤3: 生成模板文档
    print("\n===== 步骤3: 生成模板文档 =====")
    replace_document(original_doc_path, match_results_dir, template_doc_path)
    print(f"生成的模板文档已保存: {template_doc_path}")
    
    # 计算总耗时
    total_time = time.time() - start_time
    print(f"\n处理完成，总耗时: {total_time:.2f} 秒")

# 直接运行时的入口点  
if __name__ == "__main__":
    # 初始化LLM（优先使用远程模型，失败则尝试本地模型）
    print("\n===== 初始化语言模型 =====")
    llm_manager.init_remote_model()
    # llm_manager.init_local_model()
    # if not llm_manager.init_remote_model():
    #     print("远程模型初始化失败，尝试使用本地模型...")
    #     if llm_manager.init_local_model():
    #         print("本地模型初始化成功")
    #     else:
    #         print("警告：所有模型初始化都失败，匹配功能可能无法正常工作")
    # else:
    #     print("远程模型初始化成功")
    
    # 设置固定路径
    project_dir = os.path.dirname(os.path.abspath(__file__))
    
    # 输入文档路径
    original_doc_path = os.path.join(project_dir, "document/document.docx")
    
    # 输出模板文档路径
    template_doc_path = os.path.join(project_dir, "document/template.docx")
    
    # 中间处理目录
    parts_dir = os.path.join(project_dir, "document/document_parts")
    match_results_dir = os.path.join(project_dir, "document/match_results")
    key_descriptions_dir = os.path.join(project_dir, "document/key_descriptions")
    
    # 处理文档
    process_word_document(
        original_doc_path,
        template_doc_path,
        parts_dir,
        match_results_dir,
        key_descriptions_dir
    )