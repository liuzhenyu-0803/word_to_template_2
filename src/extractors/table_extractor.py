"""
表格提取模块 - 基于HTML解析的实现
从Word文档转换的HTML中提取表格并保存为清理后的HTML文件
"""

import os
import hashlib
from bs4 import BeautifulSoup
import sys

# 添加父目录到路径以便导入converter模块
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from converter.converter import word_to_html

def extract_tables(doc, output_dir):
    """
    通过HTML转换方式提取Word文档中的表格
    
    参数:
        doc: docx.Document对象（为保持接口兼容性）
        output_dir: 输出目录路径
        
    返回:
        tuple: (表格数量, 唯一表格数量)
    """
    # 首先将Word文档转换为HTML
    temp_html_path = os.path.join(output_dir, "temp_document.html")
    
    # 获取原始docx文件路径（假设从doc对象可以获取）
    # 这里需要一个临时方案，我们先从output_dir推导
    base_dir = os.path.dirname(output_dir)
    docx_path = os.path.join(base_dir, "document.docx")
    
    try:
        # 转换Word为HTML
        word_to_html(docx_path, temp_html_path)
        
        # 从HTML中提取表格
        table_count, unique_count = extract_tables_from_html(temp_html_path, output_dir)
        
        # 清理临时HTML文件
        if os.path.exists(temp_html_path):
            os.remove(temp_html_path)
            
        return table_count, unique_count
        
    except Exception as e:
        print(f"表格提取失败: {e}")
        return 0, 0

def extract_tables_from_html(html_file_path, output_dir):
    """
    从HTML文件中提取所有表格，每个表格保存为独立HTML文件
    
    参数:
        html_file_path (str): HTML文件路径
        output_dir (str): 输出目录路径
    
    返回:
        tuple: (总表格数量, 唯一表格数量)
    """
    try:
        # 读取HTML文件
        with open(html_file_path, 'r', encoding='utf-8') as file:
            html_content = file.read()
    except UnicodeDecodeError:
        # 如果UTF-8解码失败，尝试其他编码
        try:
            with open(html_file_path, 'r', encoding='gbk') as file:
                html_content = file.read()
        except:
            print(f"无法读取HTML文件: {html_file_path}")
            return 0, 0
    
    # 解析HTML
    soup = BeautifulSoup(html_content, 'html.parser')
    
    # 查找所有表格
    tables = soup.find_all('table')
    
    if not tables:
        print("未找到任何表格")
        return 0, 0
    
    # 用于去重的哈希集合
    table_hashes = set()
    unique_table_count = 0
    
    for i, table in enumerate(tables, 1):
        print(f"处理第 {i} 个表格...")
        
        # 创建清理后的表格
        clean_table = create_clean_table(table)
        
        if clean_table:
            # 计算表格内容哈希用于去重
            table_content = str(clean_table)
            table_hash = hashlib.md5(table_content.encode('utf-8')).hexdigest()
            
            if table_hash not in table_hashes:
                table_hashes.add(table_hash)
                unique_table_count += 1
                
                # 保存清理后的表格
                output_file = os.path.join(output_dir, f"table_{unique_table_count}.html")
                with open(output_file, 'w', encoding='utf-8') as file:
                    file.write(table_content)
                print(f"表格 {unique_table_count} 已保存到: {output_file}")
            else:
                print(f"表格 {i} 是重复表格，已跳过")
    
    print(f"总共处理了 {len(tables)} 个表格，保存了 {unique_table_count} 个唯一表格")
    return len(tables), unique_table_count

def create_clean_table(original_table):
    """
    创建简化的表格，只保留表格相关标签，去除所有样式和多余属性
    
    参数:
        original_table: BeautifulSoup表格对象
    
    返回:
        BeautifulSoup: 简化的表格对象
    """
    # 创建新的表格元素
    soup = BeautifulSoup('<table></table>', 'html.parser')
    new_table = soup.table
    
    # 处理表格标题（caption）
    caption = original_table.find('caption')
    if caption and caption.get_text(strip=True):
        new_caption = soup.new_tag('caption')
        new_caption.string = caption.get_text(strip=True)
        new_table.append(new_caption)
    
    # 处理表头（thead）
    thead = original_table.find('thead')
    if thead:
        new_thead = soup.new_tag('thead')
        process_table_section(thead, new_thead, soup)
        new_table.append(new_thead)
    
    # 处理表体（tbody）
    tbody = original_table.find('tbody')
    if tbody:
        new_tbody = soup.new_tag('tbody')
        process_table_section(tbody, new_tbody, soup)
        new_table.append(new_tbody)
    else:
        # 如果没有tbody，直接处理表格中的行
        rows = original_table.find_all('tr', recursive=False)
        for row in rows:
            if row.parent.name == 'table':  # 确保是直接子元素
                new_row = create_clean_row(row, soup)
                if new_row:
                    new_table.append(new_row)
    
    # 处理表尾（tfoot）
    tfoot = original_table.find('tfoot')
    if tfoot:
        new_tfoot = soup.new_tag('tfoot')
        process_table_section(tfoot, new_tfoot, soup)
        new_table.append(new_tfoot)
    
    return new_table if new_table.find('tr') else None

def process_table_section(section, new_section, soup):
    """
    处理表格的某个部分（thead, tbody, tfoot）
    
    参数:
        section: 原始部分
        new_section: 新的部分
        soup: BeautifulSoup对象
    """
    rows = section.find_all('tr')
    for row in rows:
        new_row = create_clean_row(row, soup)
        if new_row:
            new_section.append(new_row)

def create_clean_row(original_row, soup):
    """
    创建简化的表格行
    
    参数:
        original_row: 原始行
        soup: BeautifulSoup对象
    
    返回:
        BeautifulSoup: 简化的行对象
    """
    new_row = soup.new_tag('tr')
    
    # 查找所有单元格（th和td）
    cells = original_row.find_all(['td', 'th'])
    
    for cell in cells:
        # 获取单元格文本内容
        cell_text = cell.get_text(strip=True)
        
        # 创建新的单元格
        new_cell = soup.new_tag(cell.name)
        
        # 保留重要的表格属性（colspan, rowspan）
        if cell.get('colspan'):
            new_cell['colspan'] = cell['colspan']
        if cell.get('rowspan'):
            new_cell['rowspan'] = cell['rowspan']
        
        # 设置单元格内容
        if cell_text:
            new_cell.string = cell_text
        
        new_row.append(new_cell)
    
    return new_row if new_row.find(['td', 'th']) else None

# 测试功能
if __name__ == "__main__":
    from docx import Document
    
    # 获取项目路径
    current_dir = os.path.dirname(os.path.abspath(__file__))
    project_dir = os.path.dirname(os.path.dirname(current_dir))
    
    # 测试文档路径
    test_doc_path = os.path.join(project_dir, "document", "document.docx")
    output_dir = os.path.join(project_dir, "document", "document_parts")
    
    # 确保输出目录存在
    os.makedirs(output_dir, exist_ok=True)
    
    # 打开测试文档
    try:
        doc = Document(test_doc_path)
        print(f"成功打开文档: {test_doc_path}")
        
        # 提取表格
        table_count, unique_count = extract_tables(doc, output_dir)
        
        print(f"表格提取完成: 总共 {table_count} 个表格，唯一表格 {unique_count} 个")
        
    except Exception as e:
        print(f"测试失败: {e}")
        import traceback
        traceback.print_exc() 