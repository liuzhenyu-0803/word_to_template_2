"""
段落提取模块
负责从HTML文件中提取所有段落内容
"""

import os
from bs4 import BeautifulSoup

def extract_paragraphs(html_path, output_dir):
    """从HTML文件提取所有段落并保存"""
    try:
        # 读取HTML文件
        with open(html_path, 'r', encoding='utf-8') as f:
            html_content = f.read()
    except UnicodeDecodeError:
        try:
            with open(html_path, 'r', encoding='gbk') as f:
                html_content = f.read()
        except:
            print(f"无法读取HTML文件: {html_path}")
            return 0

    soup = BeautifulSoup(html_content, 'html.parser')
    paragraphs = soup.find_all(['p', 'div'])  # 提取<p>和<div>标签内容
    paragraph_count = 0

    for para in paragraphs:
        text = para.get_text().strip()
        if not text:
            continue
            
        paragraph_count += 1
        para_filename = f"paragraph_{paragraph_count}.txt"
        
        # 保存段落内容
        with open(os.path.join(output_dir, para_filename), 'w', encoding='utf-8') as f:
            f.write(text)
    
    return paragraph_count