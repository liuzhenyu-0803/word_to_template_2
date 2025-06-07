"""  
Word文档智能模板生成系统 - 主程序  
整合文档转换、提取、匹配和替换功能，实现Word文档到模板的自动转换
"""  

import os  
import time
from converter.converter import word_to_html
from extractors.extractor import extract_document
from matchers.matcher import match_document
from replacers.replacer import replace_document
from models.model_manager import llm_manager

def main():
    """
    主流程：按照LLM初始化 → 转换 → 提取 → 匹配 → 替换的顺序执行
    """
    print("===== Word文档智能模板生成系统 =====\n")
    
    # 记录开始时间
    start_time = time.time()
      # 设置路径 - 项目根目录是src的父目录
    src_dir = os.path.dirname(os.path.abspath(__file__))
    project_dir = os.path.dirname(src_dir)  # 向上一级到项目根目录
    doc_dir = os.path.join(project_dir, "document")
    doc_path = os.path.join(doc_dir, "document.docx")
    html_path = os.path.join(doc_dir, "document.html")
    extract_dir = os.path.join(doc_dir, "document_extract")
    key_descriptions_dir = os.path.join(doc_dir, "key_descriptions")
    match_results_dir = os.path.join(doc_dir, "match_results")
    template_doc_path = os.path.join(doc_dir, "template.docx")

    # 确保目录存在
    os.makedirs(extract_dir, exist_ok=True)
    os.makedirs(match_results_dir, exist_ok=True)
    os.makedirs(key_descriptions_dir, exist_ok=True)
    os.makedirs(os.path.dirname(template_doc_path), exist_ok=True)
    
    try:
        # 步骤1: 初始化LLM模型
        print("===== 步骤1: 初始化语言模型 =====")
        llm_manager.init_local_model()
        print("远程模型初始化完成\n")
          # 步骤2: Word文档转换为HTML
        print("===== 步骤2: 文档转换 =====")
        word_to_html(doc_path, html_path)
        print(f"Word文档已转换为HTML: {html_path}\n")
        
        # 步骤3: 提取文档元素
        print("===== 步骤3: 文档元素提取 =====")
        paragraph_count, table_count = extract_document(html_path, extract_dir)
        print(f"文档元素提取完成:")
        print(f"  - 段落数量: {paragraph_count}")
        print(f"  - 表格数量: {table_count}")
        print(f"  - 总元素数: {paragraph_count + table_count}")
        print(f"  - 保存位置: {extract_dir}\n")
        
        # 步骤4: 智能语义匹配
        print("===== 步骤4: 智能语义匹配 =====")

        extracted_files = [
            # os.path.join(project_dir, "document/document_extract/table_5.html"),
            os.path.join(project_dir, "document/document_extract/table_6.html"),
            os.path.join(project_dir, "document/document_extract/table_7.html"),
        ]
        
        if not extracted_files:
            print(f"警告: 在目录中没有找到提取的文档元素文件。\n")
            match_stats = {} # 或者根据需要进行其他处理
        else:
            print(f"将对以下提取的文件进行匹配: {extracted_files}\n")
            match_stats = match_document(extracted_files, key_descriptions_dir, match_results_dir)
        
        if match_stats:
            print(f"匹配结果统计:")
            print(f"  - 处理表格数: {match_stats.get('total_tables_processed', 0)}")
            print(f"  - 匹配字段数: {match_stats.get('total_keys_matched', 0)}")
            print(f"  - 有匹配表格数: {match_stats.get('tables_with_matches', 0)}")
            print(f"  - 保存位置: {match_results_dir}\n")
        else:
            print("未找到任何匹配结果\n")
        
        # 步骤5: 生成模板文档
        print("===== 步骤5: 模板文档生成 =====")
        replace_document(doc_path, match_results_dir, template_doc_path)
        print(f"模板文档生成完成: {template_doc_path}\n")
        
        # 计算总耗时
        total_time = time.time() - start_time
        print(f"===== 处理完成，总耗时: {total_time:.2f} 秒 =====")
        
    except Exception as e:
        print(f"处理过程中发生错误: {e}")
        import traceback
        traceback.print_exc()

# 直接运行时的入口点  
if __name__ == "__main__":
    main()