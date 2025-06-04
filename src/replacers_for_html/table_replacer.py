"""  
HTML表格替换模块  
专门处理HTML文件中表格内容的替换
直接修改HTML文件，不涉及Word对象操作
"""  

import os
import json
import sys
from bs4 import BeautifulSoup

# 添加父目录到路径以便导入converter模块
sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
from src.converter.converter import word_to_html

def get_all_tables_recursive_html(soup):
    """
    递归获取HTML中所有表格，包括嵌套表格
    使用与table_extractor一致的编号系统
    
    参数:
        soup: BeautifulSoup对象
        
    返回:
        dict: {table_id: table_element} 表格映射字典
    """
    table_mapping = {}
    
    def process_tables_recursive(element, parent_path=""):
        """递归处理表格，确保编号与extractor一致"""
        tables = element.find_all('table', recursive=False)
        
        for i, table in enumerate(tables, 1):
            table_id = f"{parent_path}table_{i}" if parent_path else f"table_{i}"
            table_mapping[table_id] = table
            
            # 递归处理嵌套表格
            nested_path = f"{table_id}_"
            process_tables_recursive(table, nested_path)
    
    process_tables_recursive(soup)
    return table_mapping

def load_table_mapping(output_dir):
    """
    加载表格映射文件，获取extractor生成的表格编号对应关系
    
    参数:
        output_dir: 输出目录路径
        
    返回:
        dict: 表格映射字典
    """
    mapping_file = os.path.join(output_dir, "table_mapping.json")
    if os.path.exists(mapping_file):
        try:
            with open(mapping_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            print(f"加载表格映射文件失败: {e}")
    return {}

def debug_table_structure_html(soup, show_content=True):
    """
    调试函数：显示HTML中所有表格的结构信息
    使用统一的表格标识系统
    
    参数:
        soup: BeautifulSoup对象
        show_content: 是否显示表格内容预览
    """
    print("\n=== 调试信息：HTML表格结构分析 ===")
    
    table_mapping = get_all_tables_recursive_html(soup)
    
    if not table_mapping:
        print("未找到任何表格")
        return
    
    print(f"总共找到 {len(table_mapping)} 个表格（包括嵌套表格）:")
    
    for table_id, table in table_mapping.items():
        print(f"\n表格 {table_id}:")
        
        # 获取表格基本信息
        rows = table.find_all('tr')
        print(f"  - 行数: {len(rows)}")
        
        if rows:
            cells_count = []
            for row in rows:
                cells = row.find_all(['td', 'th'])
                cells_count.append(len(cells))
            print(f"  - 列数分布: {cells_count}")
        
        # 显示表格内容预览
        if show_content and rows:
            print(f"  - 内容预览:")
            for i, row in enumerate(rows[:3]):  # 只显示前3行
                cells = row.find_all(['td', 'th'])
                cell_texts = [cell.get_text(strip=True)[:20] + "..." 
                             if len(cell.get_text(strip=True)) > 20 
                             else cell.get_text(strip=True) 
                             for cell in cells]
                print(f"    第{i+1}行: {cell_texts}")
            
            if len(rows) > 3:
                print(f"    ... (还有 {len(rows) - 3} 行)")
        
        # 检查是否包含嵌套表格
        nested_tables = table.find_all('table')
        if nested_tables:
            print(f"  - 包含 {len(nested_tables)} 个嵌套表格")
    
    print("=== HTML表格调试信息结束 ===\n")

def replace_cells_by_position_html(table, match_data):
    """
    根据位置信息在HTML表格中替换单元格内容
    
    参数:
        table: BeautifulSoup表格对象
        match_data: 匹配数据，包含位置信息和替换内容
    """
    print(f"开始替换表格内容，共有 {len(match_data)} 个匹配项")
    
    for match in match_data:
        if 'valuePos' not in match:
            print(f"跳过没有位置信息的匹配项: {match}")
            continue
        
        value_pos = match['valuePos']
        if not isinstance(value_pos, list) or len(value_pos) != 2:
            print(f"位置信息格式错误: {value_pos}")
            continue
        
        row_index, col_index = value_pos
        new_key = match.get('new_key', match.get('old_key', 'UNKNOWN'))
        
        try:
            # 获取所有行
            rows = table.find_all('tr')
            if row_index >= len(rows):
                print(f"行索引 {row_index} 超出范围（共 {len(rows)} 行）")
                continue
            
            # 获取指定行的所有单元格
            target_row = rows[row_index]
            cells = target_row.find_all(['td', 'th'])
            if col_index >= len(cells):
                print(f"列索引 {col_index} 超出范围（第 {row_index} 行共 {len(cells)} 列）")
                continue
            
            # 替换单元格内容
            target_cell = cells[col_index]
            old_content = target_cell.get_text(strip=True)
            new_content = f"[{new_key}]"
            
            # 清空单元格并设置新内容
            target_cell.clear()
            target_cell.string = new_content
            
            print(f"已替换位置 ({row_index}, {col_index}): '{old_content}' -> '{new_content}'")
            
        except Exception as e:
            print(f"替换位置 ({row_index}, {col_index}) 时出错: {e}")

def extract_table_number_from_filename(filename):
    """
    从文件名中提取表格编号
    
    参数:
        filename: 文件名，如 'table_6_matches.json'
        
    返回:
        int: 表格编号，如果提取失败返回None
    """
    import re
    match = re.search(r'table_(\d+)', filename)
    if match:
        return int(match.group(1))
    return None

def replace_tables_in_html(html_file_path, match_results_dir, match_files, output_html_path=None):
    """
    在HTML文件中替换表格内容
    
    参数:
        html_file_path: 输入HTML文件路径
        match_results_dir: 匹配结果目录路径
        match_files: 表格匹配结果文件列表
        output_html_path: 输出HTML文件路径，如果为None则覆盖原文件
        
    返回:
        bool: 是否成功
    """
    try:
        # 读取HTML文件
        with open(html_file_path, 'r', encoding='utf-8') as file:
            html_content = file.read()
        
        soup = BeautifulSoup(html_content, 'html.parser')
        
        # 使用统一的表格标识系统
        html_table_mapping = get_all_tables_recursive_html(soup)
        print(f"从HTML中发现 {len(html_table_mapping)} 个表格（包括嵌套表格）")
        
        # 调试：显示表格结构
        debug_table_structure_html(soup, show_content=False)
        
        # 处理每个匹配文件
        for match_file in match_files:
            file_path = os.path.join(match_results_dir, match_file)
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    match_data = json.load(f)
                
                # 提取表格编号并查找对应的HTML表格
                table_number = extract_table_number_from_filename(match_file)
                if table_number is None:
                    print(f"无法从文件名 {match_file} 提取表格编号，跳过")
                    continue
                
                # 根据表格编号找到对应的HTML表格
                table_id = f"table_{table_number}"
                if table_id not in html_table_mapping:
                    print(f"表格ID {table_id} 在HTML中未找到，跳过")
                    continue
                
                target_html_table = html_table_mapping[table_id]
                print(f"处理表格 {table_id}，共有 {len(match_data)} 个匹配项")
                
                # 根据位置信息替换单元格内容
                replace_cells_by_position_html(target_html_table, match_data)
                
            except Exception as e:
                print(f"处理匹配文件 {match_file} 时出错: {e}")
                continue
        
        # 保存修改后的HTML
        output_path = output_html_path or html_file_path
        with open(output_path, 'w', encoding='utf-8') as file:
            file.write(str(soup))
        
        print(f"HTML表格替换完成，已保存到: {output_path}")
        return True
        
    except Exception as e:
        print(f"HTML表格替换失败: {e}")
        return False

def convert_html_to_word(html_file_path, word_file_path):
    """
    将HTML文件转换为Word文档
    
    参数:
        html_file_path: HTML文件路径
        word_file_path: 输出Word文件路径
        
    返回:
        bool: 是否成功
    """
    try:
        # 方法1：尝试使用win32com转换
        try:
            import win32com.client
            word_app = win32com.client.Dispatch("Word.Application")
            word_app.Visible = False
            
            # 打开HTML文件
            doc = word_app.Documents.Open(os.path.abspath(html_file_path))
            
            # 保存为Word格式
            doc.SaveAs(os.path.abspath(word_file_path), FileFormat=16)  # 16 = docx
            doc.Close()
            word_app.Quit()
            
            print(f"使用Word应用程序成功转换: {word_file_path}")
            return True
            
        except ImportError:
            print("win32com不可用，尝试其他方法")
        except Exception as e:
            print(f"使用Word应用程序转换失败: {e}")
        
        # 方法2：简单的文件复制（Word可以直接打开HTML）
        import shutil
        # 确保输出文件有正确的扩展名
        if not word_file_path.endswith('.docx'):
            word_file_path = word_file_path + '.docx'
        
        shutil.copy2(html_file_path, word_file_path)
        print(f"已复制HTML文件为Word格式: {word_file_path}")
        print("注意：可能需要用Word手动打开并重新保存以确保格式正确")
        return True
        
    except Exception as e:
        print(f"HTML到Word转换失败: {e}")
        return False

# 测试功能
if __name__ == "__main__":
    # 测试HTML表格替换
    import sys
    import os
    
    # 获取项目路径
    current_dir = os.path.dirname(os.path.abspath(__file__))
    project_dir = os.path.dirname(os.path.dirname(os.path.dirname(current_dir)))
    
    # 测试文件路径
    html_file = os.path.join(project_dir, "document", "document.html")
    match_results_dir = os.path.join(project_dir, "document", "match_results")
    output_html = os.path.join(project_dir, "document", "template_html.html")
    output_word = os.path.join(project_dir, "document", "template_from_html.docx")
    
    # 查找匹配结果文件
    match_files = []
    if os.path.exists(match_results_dir):
        for file in os.listdir(match_results_dir):
            if file.endswith('_matches.json'):
                match_files.append(file)
    
    if match_files:
        print(f"找到匹配文件: {match_files}")
        
        # 执行HTML表格替换
        if replace_tables_in_html(html_file, match_results_dir, match_files, output_html):
            print("HTML表格替换成功")
            
            # 转换为Word文档
            if convert_html_to_word(output_html, output_word):
                print("HTML到Word转换成功")
            else:
                print("HTML到Word转换失败")
        else:
            print("HTML表格替换失败")
    else:
        print("未找到匹配结果文件")
