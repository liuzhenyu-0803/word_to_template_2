"""
HTML段落替换模块
专门处理HTML文件中段落内容的替换
直接修改HTML文件，不涉及Word对象操作
"""

import os
import json
from bs4 import BeautifulSoup

def replace_paragraphs_in_html(html_file_path, match_results_dir, match_files, output_html_path=None):
    """
    在HTML文件中替换段落内容
    
    参数:
        html_file_path: 输入HTML文件路径
        match_results_dir: 匹配结果目录路径
        match_files: 段落匹配结果文件列表
        output_html_path: 输出HTML文件路径，如果为None则覆盖原文件
        
    返回:
        bool: 是否成功
    """
    try:
        # 读取HTML文件
        with open(html_file_path, 'r', encoding='utf-8') as file:
            html_content = file.read()
        
        soup = BeautifulSoup(html_content, 'html.parser')
        
        # 处理每个匹配文件
        for match_file in match_files:
            file_path = os.path.join(match_results_dir, match_file)
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    match_data = json.load(f)
                
                print(f"处理段落匹配文件 {match_file}，共有 {len(match_data)} 个匹配项")
                
                # 替换段落内容
                replace_paragraph_content_html(soup, match_data)
                
            except Exception as e:
                print(f"处理段落匹配文件 {match_file} 时出错: {e}")
                continue
        
        # 保存修改后的HTML
        output_path = output_html_path or html_file_path
        with open(output_path, 'w', encoding='utf-8') as file:
            file.write(str(soup))
        
        print(f"HTML段落替换完成，已保存到: {output_path}")
        return True
        
    except Exception as e:
        print(f"HTML段落替换失败: {e}")
        return False

def replace_paragraph_content_html(soup, match_data):
    """
    在HTML中替换段落内容
    
    参数:
        soup: BeautifulSoup对象
        match_data: 匹配数据，包含原文本和新键值
    """
    for match in match_data:
        old_text = match.get('old_value', '')
        new_key = match.get('new_key', match.get('old_key', 'UNKNOWN'))
        
        if not old_text:
            continue
        
        # 查找包含该文本的所有元素
        elements = soup.find_all(text=lambda text: text and old_text in text)
        
        for element in elements:
            if element.parent:
                # 替换文本内容
                new_text = element.replace(old_text, f"[{new_key}]")
                element.replace_with(new_text)
                print(f"已替换段落内容: '{old_text}' -> '[{new_key}]'")

def debug_paragraph_structure_html(soup):
    """
    调试函数：显示HTML中段落结构信息
    
    参数:
        soup: BeautifulSoup对象
    """
    print("\n=== 调试信息：HTML段落结构分析 ===")
    
    # 查找所有段落元素
    paragraphs = soup.find_all(['p', 'div'])
    
    print(f"总共找到 {len(paragraphs)} 个段落元素:")
    
    for i, para in enumerate(paragraphs[:10], 1):  # 只显示前10个
        text_content = para.get_text(strip=True)
        if text_content:
            preview = text_content[:50] + "..." if len(text_content) > 50 else text_content
            print(f"  段落 {i}: {preview}")
    
    if len(paragraphs) > 10:
        print(f"  ... (还有 {len(paragraphs) - 10} 个段落)")
    
    print("=== HTML段落调试信息结束 ===\n")

# 测试功能
if __name__ == "__main__":
    # 测试HTML段落替换
    import sys
    import os
    
    # 获取项目路径
    current_dir = os.path.dirname(os.path.abspath(__file__))
    project_dir = os.path.dirname(os.path.dirname(os.path.dirname(current_dir)))
    
    # 测试文件路径
    html_file = os.path.join(project_dir, "document", "document.html")
    match_results_dir = os.path.join(project_dir, "document", "match_results")
    output_html = os.path.join(project_dir, "document", "template_html_with_paragraphs.html")
    
    # 查找段落匹配结果文件
    match_files = []
    if os.path.exists(match_results_dir):
        for file in os.listdir(match_results_dir):
            if file.endswith('_paragraph_matches.json'):  # 假设段落匹配文件的命名模式
                match_files.append(file)
    
    if match_files:
        print(f"找到段落匹配文件: {match_files}")
        
        # 执行HTML段落替换
        if replace_paragraphs_in_html(html_file, match_results_dir, match_files, output_html):
            print("HTML段落替换成功")
        else:
            print("HTML段落替换失败")
    else:
        print("未找到段落匹配结果文件")
