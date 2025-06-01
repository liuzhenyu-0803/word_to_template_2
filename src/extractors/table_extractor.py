import os  
import hashlib  
import csv # 导入 csv 模块，虽然我们这里手动实现，但了解标准库有好处  
import io  # 用于在内存中处理文本流，方便 csv 写入  

def extract_tables(doc, output_dir):  
    """提取所有表格(包含嵌套表格)并保存为 CSV 文件"""  
    table_count = 0  
    table_hashes = set()  
    all_tables = []  

    # 收集所有表格  
    for table in doc.tables:  
        # 处理主表格  
        table_data = extract_table_data(table)  
        table_hash = hash_table_content(table_data)  

        if table_hash not in table_hashes:  
            table_count += 1  
            table_hashes.add(table_hash)  
            all_tables.append({"id": table_count, "data": table_data})  

        # 收集嵌套表格  
        collect_nested_tables(table, table_hashes, all_tables)  

    # 保存所有表格为 CSV  
    unique_table_count = 0  
    for i, table_info in enumerate(all_tables, 1):  
        # 仅处理非空且有内容的表格  
        if table_info["data"] and any(any(cell.strip() for cell in row) for row in table_info["data"]):  
            unique_table_count += 1  
            # 修改文件名后缀为 .csv  
            table_filename = f"table_{unique_table_count}.csv"  
            table_path = os.path.join(output_dir, table_filename)  

            # 使用新的函数格式化为 CSV  
            formatted_csv = format_table_as_csv(table_info["data"])  

            # 保存表格内容  
            try:  
                with open(table_path, 'w', encoding='utf-8', newline='') as f: # 使用 newline='' 避免空行  
                    f.write(formatted_csv)  
            except Exception as e:  
                print(f"Error writing CSV file {table_path}: {e}")  
        else:  
             print(f"Skipping empty or effectively empty table {i}.")  


    # 返回实际保存的唯一表格数量  
    return unique_table_count, len(table_hashes) # 返回实际保存的csv文件数和总的唯一hash数  

def extract_table_data(table):  
    """提取表格数据为二维数组 (list of lists)"""  
    # 确保提取的是字符串，并处理 None 或非字符串类型  
    data = []  
    for row in table.rows:  
        row_data = []  
        for cell in row.cells:  
            # 合并单元格内所有段落的文本，去除首尾空格，换行符保留  
            cell_text = '\n'.join(p.text for p in cell.paragraphs).strip()  
            row_data.append(cell_text if cell_text is not None else "") # 确保是字符串  
        data.append(row_data)  
    return data  

def hash_table_content(table_data):  
    """计算表格内容的哈希值，用于去重"""  
    # 使用更稳定的方式将列表转为字符串以计算哈希  
    import json  
    try:  
        table_string = json.dumps(table_data, sort_keys=True)  
    except TypeError:  
        # 如果包含无法json序列化的类型，回退到简单字符串转换  
        table_string = str(table_data)  
    return hashlib.sha256(table_string.encode('utf-8')).hexdigest()  


def collect_nested_tables(table, table_hashes, all_tables):  
    """递归收集所有嵌套表格"""  
    for row in table.rows:  
        for cell in row.cells:  
            if cell.tables: # 检查单元格中是否有表格  
                for nested_table in cell.tables:  
                    # 提取嵌套表格内容  
                    nested_data = extract_table_data(nested_table)  
                    nested_hash = hash_table_content(nested_data)  

                    # 只添加不重复的表格  
                    if nested_hash not in table_hashes:  
                        table_hashes.add(nested_hash)  
                        all_tables.append({"id": len(all_tables) + 1, "data": nested_data})  

                    # 递归处理更深层次嵌套  
                    collect_nested_tables(nested_table, table_hashes, all_tables)  

# --- 新增函数：将表格数据格式化为 CSV 字符串 ---  
def format_table_as_csv(table_data):  
    """将二维数组格式化为标准的 CSV 格式字符串"""  
    if not table_data:  
        return "" # 返回空字符串表示空表格  

    # 使用 io.StringIO 在内存中构建 CSV  
    output = io.StringIO()  
    # 使用 csv.writer 处理复杂的引用和转义规则  
    writer = csv.writer(output, quoting=csv.QUOTE_MINIMAL) # QUOTE_MINIMAL 只引用包含特殊字符的字段  

    try:  
        for row in table_data:  
            # 确保所有单元格都是字符串  
            str_row = [str(cell) if cell is not None else "" for cell in row]  
            writer.writerow(str_row)  
    except csv.Error as e:  
         print(f"CSV writing error: {e}")  
         # 可以选择返回部分结果或空字符串  
         return output.getvalue() # 返回目前为止成功写入的部分  
    except Exception as e:  
         print(f"An unexpected error occurred during CSV formatting: {e}")  
         return "" # 或者返回部分结果  

    return output.getvalue() 