"""
表格提取模块 - 基于HTML解析的实现
从Word文档转换的HTML中提取表格并保存为清理后的HTML文件
"""

import os
from bs4 import BeautifulSoup

def extract_tables(html_file_path, output_dir):
    """从HTML文件中提取所有表格，每个表格保存为独立HTML文件"""
    try:
        # 读取HTML文件
        try:
            with open(html_file_path, 'r', encoding='utf-8') as file:
                html_content = file.read()
        except UnicodeDecodeError:
            with open(html_file_path, 'r', encoding='gbk') as file:
                html_content = file.read()
            
        # 解析HTML并查找所有表格
        soup = BeautifulSoup(html_content, 'html.parser')
        tables = soup.find_all('table')
        
        if not tables:
            print("未找到任何表格")
            return 0
        
        # 保存所有表格
        table_count = 0
        
        for i, table in enumerate(tables, 1):
            clean_table = create_clean_table(table)
            
            if clean_table:
                table_count += 1
                output_file = os.path.join(output_dir, f"table_{table_count}.html")
                with open(output_file, 'w', encoding='utf-8') as file:
                    file.write(str(clean_table))
                print(f"表格 {table_count} 已保存")
        
        print(f"总共处理了 {len(tables)} 个表格，保存了 {table_count} 个表格")
        return table_count
    except Exception as e:
        print(f"表格提取失败: {e}")
        return 0

def create_clean_table(original_table):
    """创建简化的表格，只保留表格相关标签，去除所有样式和多余属性"""
    soup = BeautifulSoup('<table></table>', 'html.parser')
    new_table = soup.table
    
    # 处理表格标题
    caption = original_table.find('caption')
    if caption and caption.get_text(strip=True):
        new_caption = soup.new_tag('caption')
        new_caption.string = caption.get_text(strip=True)
        new_table.append(new_caption)
    
    # 处理所有表格内容（thead, tbody, tfoot或直接的tr）
    for section_name in ['thead', 'tbody', 'tfoot']:
        section = original_table.find(section_name)
        if section:
            new_section = soup.new_tag(section_name)
            _process_rows(section.find_all('tr'), new_section, soup)
            if new_section.find('tr'):
                new_table.append(new_section)
    
    # 处理没有包装在thead/tbody/tfoot中的直接行
    direct_rows = original_table.find_all('tr', recursive=False)
    for row in direct_rows:
        if row.parent.name == 'table':
            new_row = _create_clean_row(row, soup)
            if new_row:
                new_table.append(new_row)
    
    return new_table if new_table.find('tr') else None

def _process_rows(rows, container, soup):
    """处理行集合"""
    for row in rows:
        new_row = _create_clean_row(row, soup)
        if new_row:
            container.append(new_row)

def _create_clean_row(original_row, soup):
    """创建简化的表格行"""
    new_row = soup.new_tag('tr')
    cells = original_row.find_all(['td', 'th'])
    
    for cell in cells:
        new_cell = soup.new_tag(cell.name)
        
        # 保留重要属性
        for attr in ['colspan', 'rowspan']:
            if cell.get(attr):
                new_cell[attr] = cell[attr]
        
        # 设置单元格内容
        cell_text = cell.get_text(strip=True)
        if cell_text:
            new_cell.string = cell_text
        
        new_row.append(new_cell)
    
    return new_row if new_row.find(['td', 'th']) else None

# 测试功能
if __name__ == "__main__":
    # 获取项目路径
    current_dir = os.path.dirname(os.path.abspath(__file__))
    project_dir = os.path.dirname(os.path.dirname(current_dir))
    
    # 测试文档路径
    html_path = os.path.join(project_dir, "document", "document.html")
    output_dir = os.path.join(project_dir, "document", "document_extract")
    
    # 确保输出目录存在
    os.makedirs(output_dir, exist_ok=True)
    
    # 运行测试
    try:
        print(f"开始处理HTML文件: {html_path}")
        
        # 提取表格
        table_count = extract_tables(html_path, output_dir)
        
        print(f"表格提取完成: 表格 {table_count} 个")
        
    except Exception as e:
        print(f"测试失败: {e}")
        import traceback
        traceback.print_exc()