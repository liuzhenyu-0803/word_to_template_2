"""
HTML文档替换模块
负责协调HTML文件中各种元素的替换操作
"""

import os
import sys
from .table_replacer import replace_tables_in_html, convert_html_to_word
from .paragraph_replacer import replace_paragraphs_in_html

def replace_html_document(html_file_path, match_results_dir, output_html_path=None, output_word_path=None):
    """
    处理HTML文档的完整替换流程
    包括表格、段落等所有元素的替换
    
    参数:
        html_file_path: 输入HTML文件路径
        match_results_dir: 匹配结果目录路径
        output_html_path: 输出HTML文件路径，如果为None则覆盖原文件
        output_word_path: 输出Word文件路径，如果提供则自动转换
        
    返回:
        bool: 是否成功
    """
    try:
        # 查找所有匹配结果文件
        table_match_files = []
        paragraph_match_files = []
        
        if os.path.exists(match_results_dir):
            for file in os.listdir(match_results_dir):
                if file.endswith('_matches.json') and 'table' in file:
                    table_match_files.append(file)
                elif file.endswith('_paragraph_matches.json'):
                    paragraph_match_files.append(file)
        
        print(f"找到表格匹配文件: {table_match_files}")
        print(f"找到段落匹配文件: {paragraph_match_files}")
        
        success = True
        
        # 处理表格替换
        if table_match_files:
            if not replace_tables_in_html(html_file_path, match_results_dir, table_match_files, output_html_path):
                print("表格替换失败")
                success = False
        
        # 处理段落替换
        if paragraph_match_files:
            # 如果已经有输出HTML，则基于输出HTML继续处理
            input_for_paragraph = output_html_path if output_html_path and os.path.exists(output_html_path) else html_file_path
            if not replace_paragraphs_in_html(input_for_paragraph, match_results_dir, paragraph_match_files, output_html_path):
                print("段落替换失败")
                success = False
        
        # 如果需要，转换为Word文档
        if success and output_word_path:
            final_html = output_html_path or html_file_path
            if not convert_html_to_word(final_html, output_word_path):
                print("HTML到Word转换失败")
                success = False
        
        if success:
            print("HTML文档替换完成")
        else:
            print("HTML文档替换过程中发生错误")
        
        return success
        
    except Exception as e:
        print(f"HTML文档替换失败: {e}")
        return False

def create_template_from_original_html(original_html_path, match_results_dir, template_html_path, template_word_path=None):
    """
    从原始HTML文档创建模板
    
    参数:
        original_html_path: 原始HTML文档路径
        match_results_dir: 匹配结果目录路径
        template_html_path: 模板HTML文件路径
        template_word_path: 模板Word文件路径（可选）
        
    返回:
        bool: 是否成功
    """
    print(f"开始从HTML文档创建模板...")
    print(f"原始HTML: {original_html_path}")
    print(f"匹配结果目录: {match_results_dir}")
    print(f"模板HTML: {template_html_path}")
    
    return replace_html_document(
        html_file_path=original_html_path,
        match_results_dir=match_results_dir,
        output_html_path=template_html_path,
        output_word_path=template_word_path
    )

def batch_process_html_documents(html_files, match_results_dir, output_dir):
    """
    批量处理多个HTML文档
    
    参数:
        html_files: HTML文件路径列表
        match_results_dir: 匹配结果目录路径
        output_dir: 输出目录路径
        
    返回:
        dict: 处理结果，{文件名: 是否成功}
    """
    results = {}
    
    # 确保输出目录存在
    os.makedirs(output_dir, exist_ok=True)
    
    for html_file in html_files:
        try:
            # 生成输出文件路径
            base_name = os.path.splitext(os.path.basename(html_file))[0]
            output_html = os.path.join(output_dir, f"{base_name}_template.html")
            output_word = os.path.join(output_dir, f"{base_name}_template.docx")
            
            print(f"\n处理文件: {html_file}")
            
            success = replace_html_document(
                html_file_path=html_file,
                match_results_dir=match_results_dir,
                output_html_path=output_html,
                output_word_path=output_word
            )
            
            results[os.path.basename(html_file)] = success
            
        except Exception as e:
            print(f"处理文件 {html_file} 时出错: {e}")
            results[os.path.basename(html_file)] = False
    
    return results

# 测试功能
if __name__ == "__main__":
    # 测试完整的HTML文档替换流程
    import sys
    import os
    
    # 获取项目路径
    current_dir = os.path.dirname(os.path.abspath(__file__))
    project_dir = os.path.dirname(os.path.dirname(os.path.dirname(current_dir)))
    
    # 测试文件路径
    html_file = os.path.join(project_dir, "document", "document.html")
    match_results_dir = os.path.join(project_dir, "document", "match_results")
    template_html = os.path.join(project_dir, "document", "template_complete.html")
    template_word = os.path.join(project_dir, "document", "template_complete.docx")
    
    print("测试HTML文档完整替换流程...")
    
    # 执行完整的HTML文档替换
    if create_template_from_original_html(html_file, match_results_dir, template_html, template_word):
        print("HTML文档模板创建成功")
    else:
        print("HTML文档模板创建失败")
